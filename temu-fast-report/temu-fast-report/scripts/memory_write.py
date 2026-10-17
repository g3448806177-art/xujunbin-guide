"""保留历史OOXML辅助代码；报单记忆的程序写入入口已永久禁用。"""
from collections import defaultdict
from contextlib import closing
from copy import deepcopy
from datetime import datetime
import fcntl
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as X
from PIL import Image

from memory_store import DEFAULT_MEMORY, MemoryBook, NS, compact, fingerprint, rel_path, target_path, save_json
from report_contract import skc_text, record_key, validate_inputs, validate_report, number

M, R, P = NS['m'], NS['r'], NS['p']
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
WPS = 'http://www.wps.cn/officeDocument/2017/etCustomData'
V = 'urn:schemas-microsoft-com:vml'
O = 'urn:schemas-microsoft-com:office:office'
EX = 'urn:schemas-microsoft-com:office:excel'


def q(namespace, name): return '{'+namespace+'}'+name


def child(parent, name, attrs=None, text=None, namespace=M):
    c = X.SubElement(parent,q(namespace,name),attrs or {})
    c.text = text
    return c


def xml_bytes(root):
    return X.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)


def parse(data):
    return X.fromstring(data,parser=X.XMLParser(resolve_entities=False,no_network=True))


def formula_string(value): return str(value).replace('"','""')


class Package:
    def __init__(self, path):
        with ZipFile(path) as z:
            self.parts = {n:z.read(n) for n in z.namelist()}
        self.roots = {}
        self.wb = self.root('xl/workbook.xml')
        self.wb_rels = self.root('xl/_rels/workbook.xml.rels')
        self.types = self.root('[Content_Types].xml')
        self.sheets = {}
        for s in self.wb.find(q(M,'sheets')):
            rel = next(r for r in self.wb_rels if r.get('Id') == s.get(q(R,'id')))
            self.sheets[s.get('name')] = target_path('xl/workbook.xml',rel.get('Target'))
        self.image_hashes = None

    def root(self, part, tag=None, ns=None):
        if part not in self.roots:
            self.roots[part] = parse(self.parts[part]) if part in self.parts else X.Element(tag,nsmap=ns)
        return self.roots[part]

    def relationship(self, part, kind, target):
        root = self.root(rel_path(part),q(P,'Relationships'),{None:P})
        for r in root:
            if r.get('Type') == kind and r.get('Target') == target: return r.get('Id')
        ids = {r.get('Id') for r in root}
        n = 1
        while 'rId'+str(n) in ids: n += 1
        rid = 'rId'+str(n)
        child(root,'Relationship',{'Id':rid,'Type':kind,'Target':target},namespace=P)
        return rid

    def content_type(self, part, kind):
        if not any(c.get('PartName') == '/'+part for c in self.types):
            child(self.types,'Override',{'PartName':'/'+part,'ContentType':kind},namespace=CT)

    def add_sheet(self, name, headers):
        if name in self.sheets: return self.sheets[name]
        n = 1
        while f'xl/worksheets/sheet{n}.xml' in self.parts or f'xl/worksheets/sheet{n}.xml' in self.roots: n += 1
        part = f'xl/worksheets/sheet{n}.xml'
        rid = self.relationship('xl/workbook.xml',R+'/worksheet',f'worksheets/sheet{n}.xml')
        sheet_id = str(max(int(s.get('sheetId')) for s in self.wb.find(q(M,'sheets')))+1)
        child(self.wb.find(q(M,'sheets')),'sheet',{'name':name,'sheetId':sheet_id,q(R,'id'):rid})
        root = self.root(part,q(M,'worksheet'),{None:M,'r':R})
        child(root,'dimension',{'ref':'A1:H1'})
        child(child(root,'sheetViews'),'sheetView',{'workbookViewId':'0'})
        child(root,'sheetFormatPr',{'defaultRowHeight':'22'})
        cols = child(root,'cols')
        for i,width in enumerate([14,18,12,10,32,48,24,22],1): child(cols,'col',{'min':str(i),'max':str(i),'width':str(width),'customWidth':'1'})
        child(root,'sheetData')
        self.content_type(part,'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        self.sheets[name] = part
        for i,value in enumerate(headers): self.set_cell(part,chr(65+i)+'1',value)
        return part

    def row(self, part, row):
        data = self.root(part).find(q(M,'sheetData'))
        found = data.find(f"{{{M}}}row[@r='{row}']")
        if found is not None: return found
        attrs = {'r':str(row),'ht':'195','customHeight':'1'}
        found = X.Element(q(M,'row'),attrs)
        insert_at = next((i for i,c in enumerate(data) if int(c.get('r')) > row),len(data))
        data.insert(insert_at,found)
        return found

    def set_cell(self, part, ref, value, formula=None):
        row = self.row(part,int(re.search(r'[0-9]+',ref).group()))
        cell = row.find(f"{{{M}}}c[@r='{ref}']")
        if cell is None:
            col = re.match('[A-Z]+',ref).group()
            template = self.root(part).find(f".//{{{M}}}c[@r='{col}3']")
            attrs = {'r':ref}
            if template is not None and template.get('s'): attrs['s'] = template.get('s')
            cell = X.Element(q(M,'c'),attrs)
            def colnum(r):
                n=0
                for ch in re.match('[A-Z]+',r).group(): n=n*26+ord(ch)-64
                return n
            at = next((i for i,c in enumerate(row) if colnum(c.get('r')) > colnum(ref)),len(row))
            row.insert(at,cell)
        for c in list(cell): cell.remove(c)
        cell.attrib.pop('t',None)
        if formula is not None:
            cell.set('t','str')
            child(cell,'f',text=formula)
            child(cell,'v',text=str(value or ''))
        elif isinstance(value,(int,float)) and not isinstance(value,bool):
            child(cell,'v',text=str(value))
        elif value is not None and value != '':
            cell.set('t','inlineStr')
            child(child(cell,'is'),'t',{'{http://www.w3.org/XML/1998/namespace}space':'preserve'},text=str(value))
        dim = self.root(part).find(q(M,'dimension'))
        if dim is not None:
            # Main memory layout is A:M; audit is A:H.
            maxrow = max(int(r.get('r')) for r in self.root(part).find(q(M,'sheetData')))
            dim.set('ref',f'A1:{"M" if part == self.sheets.get("Sheet1") else "H"}{maxrow}')

    def image(self, part, ref, path):
        raw = Path(path).read_bytes()
        with Image.open(path) as im:
            ext = 'png' if im.format == 'PNG' else 'jpeg' if im.format == 'JPEG' else None
            if ext is None: raise ValueError('记忆图片仅支持PNG/JPEG，请先转为搜款SKU图片')
            width,height = im.size
        cell_part = next((target_path('xl/workbook.xml',r.get('Target')) for r in self.wb_rels if r.get('Type','').endswith('/cellImage')),'xl/cellimages.xml')
        root = self.root(cell_part,q(WPS,'cellImages'),{'etc':WPS,'xdr':NS['xdr'],'a':NS['a'],'r':R})
        rels = self.root(rel_path(cell_part),q(P,'Relationships'),{None:P})
        if self.image_hashes is None:
            self.image_hashes = {}
            relmap = {r.get('Id'):target_path(cell_part,r.get('Target')) for r in rels}
            for pic in root.findall('.//'+q(NS['xdr'],'pic')):
                nv = pic.find('.//'+q(NS['xdr'],'cNvPr'))
                blip = pic.find('.//'+q(NS['a'],'blip'))
                asset = relmap.get(blip.get(q(R,'embed'))) if blip is not None else None
                if asset in self.parts and nv is not None:
                    self.image_hashes[hashlib.sha256(self.parts[asset]).hexdigest()] = nv.get('name')
        digest = hashlib.sha256(raw).hexdigest()
        name = self.image_hashes.get(digest)
        if name is None:
            name = 'ID_'+digest[:32].upper()
            asset = 'xl/media/memory_'+digest+'.'+ext
            self.parts[asset] = raw
            rid = self.relationship(cell_part,R+'/image','media/'+Path(asset).name)
            holder = child(root,'cellImage',namespace=WPS)
            pic = child(holder,'pic',namespace=NS['xdr'])
            nv = child(pic,'nvPicPr',namespace=NS['xdr'])
            maxid = max([int(n.get('id','0')) for n in root.findall('.//'+q(NS['xdr'],'cNvPr'))]+[0])
            child(nv,'cNvPr',{'id':str(maxid+1),'name':name},namespace=NS['xdr'])
            child(nv,'cNvPicPr',namespace=NS['xdr'])
            fill = child(pic,'blipFill',namespace=NS['xdr'])
            child(fill,'blip',{q(R,'embed'):rid},namespace=NS['a'])
            child(child(fill,'stretch',namespace=NS['a']),'fillRect',namespace=NS['a'])
            shape = child(pic,'spPr',namespace=NS['xdr'])
            xfrm = child(shape,'xfrm',namespace=NS['a'])
            child(xfrm,'off',{'x':'0','y':'0'},namespace=NS['a'])
            scale = min(140/width,140/height)
            child(xfrm,'ext',{'cx':str(int(width*scale*9525)),'cy':str(int(height*scale*9525))},namespace=NS['a'])
            child(child(shape,'prstGeom',{'prst':'rect'},namespace=NS['a']),'avLst',namespace=NS['a'])
            self.image_hashes[digest] = name
            self.content_type(asset,'image/'+ext)
        self.relationship('xl/workbook.xml','http://www.wps.cn/officeDocument/2020/cellImage',cell_part.removeprefix('xl/'))
        self.content_type(cell_part,'application/vnd.wps-officedocument.cellimage+xml')
        self.set_cell(part,ref,'',formula=f'_xlfn.DISPIMG("{name}",1)')

    def comment(self, part, ref, text):
        rels = self.root(rel_path(part),q(P,'Relationships'),{None:P})
        existing = next((r for r in rels if r.get('Type','').endswith('/comments')),None)
        suffix = Path(part).stem
        comment_part = target_path(part,existing.get('Target')) if existing is not None else f'xl/comments/memory_{suffix}.xml'
        root = self.root(comment_part,q(M,'comments'),{None:M})
        authors = root.find(q(M,'authors'))
        if authors is None: authors = child(root,'authors')
        author = next((i for i,a in enumerate(authors) if a.text == '报单记忆'),None)
        if author is None:
            author = len(authors)
            child(authors,'author',text='报单记忆')
        listing = root.find(q(M,'commentList'))
        if listing is None: listing = child(root,'commentList')
        old = listing.find(f"{{{M}}}comment[@ref='{ref}']")
        preserved = ''
        if old is not None:
            old_text = ''.join(old.itertext())
            preserved = old_text.split('\n报单记忆记录\n')[0] if '\n报单记忆记录\n' in old_text else old_text
            listing.remove(old)
        c = child(listing,'comment',{'ref':ref,'authorId':str(author)})
        child(child(c,'text'),'t',text=preserved+'\n报单记忆记录\n'+text)
        if existing is None: self.relationship(part,R+'/comments','../comments/'+Path(comment_part).name)
        self.content_type(comment_part,'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml')
        vrel = next((r for r in rels if r.get('Type','').endswith('/vmlDrawing')),None)
        vpart = target_path(part,vrel.get('Target')) if vrel is not None else f'xl/drawings/memory_{suffix}.vml'
        vml = self.root(vpart,'xml',{'v':V,'o':O,'x':EX})
        rownum = int(re.search('[0-9]+',ref).group())-1
        colnum = 0
        for ch in re.match('[A-Z]+',ref).group(): colnum=colnum*26+ord(ch)-64
        colnum -= 1
        present = any(d.findtext(q(EX,'Row')) == str(rownum) and d.findtext(q(EX,'Column')) == str(colnum) for d in vml.findall('.//'+q(EX,'ClientData')))
        if not present:
            if not vml.findall('.//'+q(V,'shapetype')):
                st = child(vml,'shapetype',{'id':'_x0000_t202','coordsize':'21600,21600',q(O,'spt'):'202','path':'m,l,21600r21600,l21600,xe'},namespace=V)
                child(st,'stroke',{'joinstyle':'miter'},namespace=V)
                child(st,'path',{'gradientshapeok':'t',q(O,'connecttype'):'rect'},namespace=V)
            ids = [int(n.get('id').rsplit('s',1)[1]) for n in vml.findall('.//'+q(V,'shape')) if re.fullmatch('_x0000_s[0-9]+',n.get('id',''))]
            shape = child(vml,'shape',{'id':'_x0000_s'+str(max(ids+[1024])+1),'type':'#_x0000_t202','style':'position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:300pt;height:150pt;z-index:1;visibility:hidden','fillcolor':'#ffffe1',q(O,'insetmode'):'auto'},namespace=V)
            child(shape,'fill',{'color2':'#ffffe1'},namespace=V)
            child(shape,'shadow',{'on':'t','color':'black','obscured':'t'},namespace=V)
            child(shape,'path',{q(O,'connecttype'):'none'},namespace=V)
            child(shape,'textbox',{'style':'mso-direction-alt:auto'},namespace=V)
            client = child(shape,'ClientData',{'ObjectType':'Note'},namespace=EX)
            child(client,'MoveWithCells',namespace=EX); child(client,'SizeWithCells',namespace=EX)
            child(client,'AutoFill',text='False',namespace=EX)
            child(client,'Row',text=str(rownum),namespace=EX); child(client,'Column',text=str(colnum),namespace=EX)
        if vrel is None:
            rid = self.relationship(part,R+'/vmlDrawing','../drawings/'+Path(vpart).name)
            root_sheet = self.root(part)
            legacy = X.Element(q(M,'legacyDrawing'),{q(R,'id'):rid})
            before = next((i for i,e in enumerate(root_sheet) if X.QName(e).localname in ('legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst')),len(root_sheet))
            root_sheet.insert(before,legacy)
        if not any(c.get('Extension') == 'vml' for c in self.types): child(self.types,'Default',{'Extension':'vml','ContentType':'application/vnd.openxmlformats-officedocument.vmlDrawing'},namespace=CT)

    def save(self, path):
        # Unmodified ZIP entries keep exactly the original uncompressed bytes.
        with ZipFile(path,'w',ZIP_DEFLATED) as z:
            for name,raw in self.parts.items():
                z.writestr(name,xml_bytes(self.roots[name]) if name in self.roots else raw)
            for name,root in self.roots.items():
                if name not in self.parts: z.writestr(name,xml_bytes(root))


def sync_memory(*_args, **_kwargs):
    """保留旧导入名，但永久禁止报单程序写入记忆库。"""
    raise RuntimeError('记忆库仅允许人工录入，自动同步已禁用。请人工确认后直接修改报单记忆.xlsx。')


def _legacy_sync_memory(records, data, report, memory=DEFAULT_MEMORY):
    data = validate_inputs(records,data)
    validation = validate_report(report,records,data)
    memory = Path(memory).resolve()
    if not memory.exists(): raise ValueError('记忆原表不存在，不能从旧缓存恢复；请恢复或新建原表后重试')
    if any((memory.parent/name).exists() for name in ('~$'+memory.name,'.~'+memory.name)):
        raise ValueError('记忆表正被表格软件占用，请关闭后重试续写')
    run_id = hashlib.sha256(compact({'records':records,'matches':data,'report':str(Path(report).resolve())}).encode()).hexdigest()
    ledger_path = memory.with_name('.'+memory.stem+'.已续写.json')
    with open(memory.with_name('.'+memory.stem+'.lock'),'a+b') as lock:
        fcntl.flock(lock.fileno(),fcntl.LOCK_EX|fcntl.LOCK_NB)
        ledger = json.loads(ledger_path.read_text()) if ledger_path.exists() else []
        if run_id in ledger:
            return {'status':'already_synced','report':validation}
        signature = fingerprint(memory)
        with closing(MemoryBook(memory)) as book:
            if 'Sheet1' not in book.sheets: raise ValueError('记忆库缺少Sheet1')
            book.positions()  # Verify layout before mutation.
            cells = book.cells('Sheet1',lambda ref:bool(re.fullmatch('[A-M][0-9]+',ref)))
            existing_positions = book.positions()
            existing_hits = book.details(existing_positions)
        package = Package(memory)
        part = package.sheets['Sheet1']
        def cell(col,row): return cells.get(f'{col}{row}',{}).get('value','')
        ends = {'purchase':max([int(re.search('[0-9]+',r).group()) for r,c in cells.items() if r[0] in 'ABCDEFGHI' and int(re.search('[0-9]+',r).group()) >= 3 and (c['value'] != '' or c['formula'])]+[2]),
                'online':max([int(re.search('[0-9]+',r).group()) for r,c in cells.items() if r[0] in 'KLM' and int(re.search('[0-9]+',r).group()) >= 3 and (c['value'] != '' or c['formula'])]+[2])}
        positions = {}
        for skc,items in existing_hits.items():
            for item in items:
                if item['kind'] != 'purchase': continue
                row = int(re.search('[0-9]+$',item['ref']).group())
                bits = str(item.get('remarks') or '').split('*')
                if len(bits) == 4:
                    # Do not guess missing specs or merge different suppliers.
                    positions[(skc,str(item['shop']).strip(),bits[0],bits[1],bits[2])] = row
        now = datetime.now().astimezone().isoformat(timespec='seconds')
        grouped = {}
        for r in records:
            key = record_key(r)
            if key not in data['matched']: continue
            m = data['matched'][key]
            identity = (skc_text(r.get('skcId')),str(m['shop']).strip(),str(m['webNo']),str(m['webColor']),str(r.get('size','')))
            # Empty SKC never merges unrelated records.
            if not identity[0]: identity += (key,)
            if identity not in grouped: grouped[identity] = [r,m,0]
            grouped[identity][2] += int(r['quantity'])
        result = {'status':'synced','updated':0,'added':0,'onlineAdded':0,'auditAdded':0,'report':validation}
        for identity,(r,m,qty) in grouped.items():
            row = positions.get(identity)
            if row is None:
                ends['purchase'] += 1; row = ends['purchase']; result['added'] += 1
            else: result['updated'] += 1
            values = {'A':float(number(m['cost'],'单价')*qty),'B':qty,'C':data.get('reporter',''),'E':m['address'],'G':f"{m['webNo']}*{m['webColor']}*{r.get('size','')}*{qty}",'H':r.get('productAttributes',''),'I':skc_text(r.get('skcId'))}
            for col,value in values.items(): package.set_cell(part,f'{col}{row}',value)
            package.set_cell(part,f'D{row}',m['shop'],formula=f'HYPERLINK("{formula_string(m["url"])}","{formula_string(m["shop"])}")')
            package.image(part,f'F{row}',r['imagePath'])
            evidence = m['verification']
            note = {'状态':'自动匹配','商品链接':m['url'],'记录时间':now,'核验时间':evidence['checkedAt'],'来源':Path(report).name,'颜色':m['webColor'],'尺码':r.get('size',''),'依据':evidence['reason']}
            package.comment(part,f'D{row}','\n'.join(f'{k}: {v}' for k,v in note.items()))
        for r in records:
            key = record_key(r)
            if key in data['unmatched']:
                # Upsert online entries by exact SKC + source color/size comments.
                skc = skc_text(r.get('skcId'))
                online_row = None
                for old in existing_hits.get(skc,[]):
                    if old['kind'] == 'online' and old.get('reason') == f"{r.get('color','')}*{r.get('size','')}":
                        online_row = int(re.search('[0-9]+$',old['ref']).group()); break
                if online_row is None:
                    ends['online'] += 1; online_row = ends['online']; result['onlineAdded'] += 1
                    if skc:
                        existing_hits.setdefault(skc,[]).append({'kind':'online','ref':f'Sheet1!M{online_row}','reason':f"{r.get('color','')}*{r.get('size','')}"})
                package.image(part,f'K{online_row}',r['imagePath'])
                package.set_cell(part,f'L{online_row}',r.get('productAttributes',''))
                package.set_cell(part,f'M{online_row}',skc)
                package.comment(part,f'M{online_row}',f"状态: 范围内无同款\n记录时间: {now}\n来源: {Path(report).name}\n依据: {r.get('color','')}*{r.get('size','')}\n搜索: SKU前5，主图前5（有主图时）；本次未找到，下次仍需搜索")
            elif key in data['pending'] or key in data['search_failed']:
                state = 'pending' if key in data['pending'] else 'search_failed'
                audit_part = package.add_sheet('记忆审核',['状态','SKC','颜色','尺码','来源','原因','记录时间','SKU图片'])
                row = max(int(c.get('r')) for c in package.root(audit_part).find(q(M,'sheetData')))+1
                values = ['待审核' if state == 'pending' else '搜索失败',skc_text(r.get('skcId')),r.get('color',''),r.get('size',''),f'{Path(report).name} / {key}',data[state][key]['reason'],now]
                for col,value in zip('ABCDEFG',values): package.set_cell(audit_part,f'{col}{row}',value)
                package.image(audit_part,f'H{row}',r['imagePath'])
                result['auditAdded'] += 1
        fd,tmp = tempfile.mkstemp(prefix=memory.stem+'.',suffix='.xlsx',dir=memory.parent)
        os.close(fd)
        try:
            package.save(tmp)
            with ZipFile(tmp) as z:
                if z.testzip() is not None: raise ValueError('记忆文件压缩校验失败')
            with closing(MemoryBook(tmp)) as check: check.positions()
            if fingerprint(memory) != signature: raise ValueError('记忆表在续写期间被修改，已停止保存，避免覆盖人工删除')
            backup_dir = memory.parent/'.报单记忆备份'
            backup_dir.mkdir(exist_ok=True)
            shutil.copy2(memory,backup_dir/(datetime.now().strftime('%Y%m%d-%H%M%S-%f')+'.xlsx'))
            os.replace(tmp,memory)
            save_json(ledger_path,ledger+[run_id])
        finally:
            if os.path.exists(tmp): os.unlink(tmp)
        return result
