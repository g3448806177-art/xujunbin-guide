#!/usr/bin/env python3
"""SKC-only, batch lookup. SQLite stores positions only; images are opt-in."""
import argparse
from collections import defaultdict
from contextlib import closing
from decimal import Decimal, InvalidOperation
import hashlib
import io
import json
import os
from pathlib import Path
import posixpath
import re
import sqlite3
import tempfile
from time import perf_counter
import xml.etree.ElementTree as ET
from zipfile import ZipFile, BadZipFile

from report_contract import skc_text, record_key
from local_paths import memory_path

DEFAULT_MEMORY = memory_path()
NS = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'p':'http://schemas.openxmlformats.org/package/2006/relationships',
      'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
      'xdr':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
INDEX_VERSION = 2


def fingerprint(path):
    p = Path(path).resolve()
    s = p.stat()
    return [str(p), s.st_size, s.st_mtime_ns, s.st_ctime_ns, s.st_ino]


def rel_path(part):
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part)+'.rels')


def target_path(part, target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def compact(value):
    return json.dumps(value, ensure_ascii=False, separators=(',', ':'))


def save_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name+'.tmp')
    tmp.write_text(compact(value)+'\n', encoding='utf-8')
    os.replace(tmp, path)


def column_index(column):
    value = 0
    for char in column:
        value = value * 26 + ord(char) - ord('A') + 1
    return value


def ref_column(ref):
    match = re.fullmatch(r'([A-Z]+)[0-9]+', ref)
    return match.group(1) if match else ''


class MemoryBook:
    def __init__(self, path):
        self.path = Path(path)
        self.z = ZipFile(path)
        self.names = set(self.z.namelist())
        self.detail_cells = 0
        rels = self.relationships('xl/workbook.xml')
        root = ET.fromstring(self.z.read('xl/workbook.xml'))
        self.sheets = {s.attrib['name']: rels[s.attrib['{'+NS['r']+'}id']]['path'] for s in root.findall('m:sheets/m:sheet', NS)}

    def close(self):
        self.z.close()

    def relationships(self, part):
        name = rel_path(part)
        if name not in self.names:
            return {}
        root = ET.fromstring(self.z.read(name))
        return {r.attrib['Id']:{**r.attrib, 'path':target_path(part,r.attrib['Target'])} for r in root}

    def cells(self, sheet, predicate):
        # Filter cell references before decoding any value. Never read media here.
        picked = {}
        with self.z.open(self.sheets[sheet]) as stream:
            for event, elem in ET.iterparse(stream, events=('end',)):
                if elem.tag == '{'+NS['m']+'}c':
                    if predicate(elem.attrib['r']):
                        picked[elem.attrib['r']] = elem
                elif elem.tag == '{'+NS['m']+'}row':
                    elem.clear()
        ids = {int(c.find('m:v',NS).text) for c in picked.values() if c.attrib.get('t') == 's' and c.find('m:v',NS) is not None}
        strings = {}
        if ids and 'xl/sharedStrings.xml' in self.names:
            with self.z.open('xl/sharedStrings.xml') as stream:
                index = 0
                for event, elem in ET.iterparse(stream, events=('end',)):
                    if elem.tag == '{'+NS['m']+'}si':
                        if index in ids:
                            strings[index] = ''.join(t.text or '' for t in elem.iter('{'+NS['m']+'}t'))
                        index += 1
                        elem.clear()
                        if len(strings) == len(ids):
                            break
        decoded = {}
        for ref, c in picked.items():
            v, f = c.find('m:v',NS), c.find('m:f',NS)
            raw = v.text if v is not None and v.text is not None else ''
            kind = c.attrib.get('t')
            if kind == 's': value = strings.get(int(raw), '') if raw else ''
            elif kind == 'inlineStr': value = ''.join(t.text or '' for t in c.findall('.//m:t',NS))
            elif kind in ('str','e','b'): value = raw
            elif raw:
                try:
                    n = Decimal(raw)
                    value = int(n) if n.is_finite() and n == n.to_integral_value() else float(n)
                except (InvalidOperation, ValueError): value = raw
            else: value = ''
            decoded[ref] = {'value':value, 'formula':f.text or '' if f is not None else '', 'type':kind}
        return decoded

    def layout(self, sheet):
        """按第2行的SKC表头识别手工维护后的采购／网购区。"""
        headers = self.cells(sheet, lambda ref: bool(re.fullmatch(r'[A-Z]+2', ref)))
        by_column = {ref_column(ref): str(cell.get('value', '')).strip() for ref, cell in headers.items()}
        skc_columns = sorted((column for column, value in by_column.items() if value.upper() == 'SKC'), key=column_index)
        if len(skc_columns) < 2:
            raise ValueError(f'记忆库工作表{sheet}未找到采购和网购两个SKC表头')
        purchase_skc, online_skc = skc_columns[0], skc_columns[1]

        def column_named(name, start, end):
            candidates = [column for column, value in by_column.items() if value == name and start <= column_index(column) < end]
            return min(candidates, key=column_index) if candidates else ''

        purchase_end = column_index(purchase_skc)
        online_end = column_index(online_skc)
        layout = {
            'purchase': {
                'skc': purchase_skc,
                'cost': column_named('成本', 1, purchase_end),
                'quantity': column_named('件数', 1, purchase_end),
                'shop': column_named('店铺', 1, purchase_end),
                'address': column_named('地址', 1, purchase_end),
                'remarks': column_named('网站备注', 1, purchase_end),
                'image': column_named('图片', 1, purchase_end),
            },
            'online': {
                'skc': online_skc,
                'image': column_named('图片', purchase_end + 1, online_end),
            },
        }
        required = ('cost', 'quantity', 'shop', 'address', 'remarks', 'image')
        if any(not layout['purchase'][field] for field in required) or not layout['online']['image']:
            raise ValueError(f'记忆库工作表{sheet}表头不完整，无法识别采购或网购字段')
        return layout

    def positions(self):
        positions = []
        valid_layout = False
        for sheet in self.sheets:
            if sheet == '记忆审核':
                continue
            try:
                layout = self.layout(sheet)
            except ValueError:
                continue
            sections = [(layout['purchase']['skc'], 'purchase'), (layout['online']['skc'], 'online')]
            valid_layout = True
            section_columns = {column for column, _ in sections}
            cols = self.cells(sheet, lambda ref: ref_column(ref) in section_columns)
            for col, kind in sections:
                for ref, cell in cols.items():
                    if ref_column(ref) != col or int(re.search(r'[0-9]+$', ref).group()) < 3 or cell['formula'] or cell['type'] in ('b','e'):
                        continue
                    skc = skc_text(cell['value'])
                    if skc:
                        positions.append((skc,sheet,int(ref[1:]),kind))
        if not valid_layout:
            raise ValueError('记忆库未找到采购和网购两个SKC表头，请核对第2行')
        return positions

    def details(self, positions):
        by_sheet = defaultdict(list)
        for skc,sheet,row,kind in positions:
            by_sheet[sheet].append((skc,row,kind))
        hits = defaultdict(list)
        for sheet, rows in by_sheet.items():
            layout = self.layout(sheet)
            refs = set()
            for skc,row,kind in rows:
                if kind == 'purchase':
                    refs.update(f'{layout["purchase"][field]}{row}' for field in ('cost', 'quantity', 'shop', 'address', 'remarks', 'skc'))
                else:
                    refs.add(f'{layout["online"]["skc"]}{row}')
            cells = self.cells(sheet, lambda ref: ref in refs)
            self.detail_cells += len(cells)
            rels = self.relationships(self.sheets[sheet])
            comments = {}
            for rel in rels.values():
                if rel['Type'].endswith('/comments'):
                    root = ET.fromstring(self.z.read(rel['path']))
                    for c in root.findall('m:commentList/m:comment',NS):
                        if c.attrib['ref'] in refs:
                            comments[c.attrib['ref']] = ''.join(t.text or '' for t in c.findall('.//m:t',NS))
            links = {}
            with self.z.open(self.sheets[sheet]) as stream:
                for _, c in ET.iterparse(stream,events=('end',)):
                    if c.tag == '{'+NS['m']+'}hyperlink' and c.attrib.get('ref') in refs:
                        rel = rels.get(c.attrib.get('{'+NS['r']+'}id'),{})
                        links[c.attrib['ref']] = rel.get('Target','')
                    c.clear()
            for skc,row,kind in rows:
                def val(col): return cells.get(f'{col}{row}',{}).get('value','')
                section = layout[kind]
                anchor = f'{section["skc"]}{row}'
                if skc_text(val(section['skc'])) != skc:
                    raise ValueError('记忆库读取期间已变更，不能使用旧位置')
                shop_ref = f'{layout["purchase"]["shop"]}{row}' if kind == 'purchase' else anchor
                note = comments.get(shop_ref,comments.get(anchor,''))
                meta = dict(line.split(': ',1) for line in note.splitlines() if ': ' in line)
                state = meta.get('状态','历史记录' if kind == 'purchase' else '范围内无同款')
                url = links.get(shop_ref,'')
                f = cells.get(shop_ref,{}).get('formula','')
                match = re.match(r'(?:_xlfn\.)?HYPERLINK\("((?:[^"\n]|"")+)"',f,re.I)
                if match: url = match.group(1).replace('""','"')
                url = url or meta.get('商品链接','')
                if not str(url).startswith(('https://','http://')): url = ''
                detail = {'ref':f'{sheet}!{anchor}','kind':kind,'status':state}
                if kind == 'purchase':
                    remarks = val(layout['purchase']['remarks'])
                    parts = str(remarks).split('*')
                    detail.update(
                        shop=val(layout['purchase']['shop']), address=val(layout['purchase']['address']), url=url, remarks=remarks,
                        historicalCost=val(layout['purchase']['cost']), historicalQuantity=val(layout['purchase']['quantity']),
                        historyWebNo=parts[0] if parts else '',
                    )
                for field,label in [('checkedAt','核验时间'),('reason','依据')]:
                    if meta.get(label): detail[field] = meta[label]
                detail['imageRef'] = f'{sheet}!{section["image"]}{row}'
                detail['action'] = 'reuse_memory'
                hits[skc].append(detail)
        return dict(hits)


def ensure_index(memory, cache, signature):
    cache = Path(cache)
    try:
        with closing(sqlite3.connect(f'file:{cache}?mode=ro',uri=True)) as con:
            stored = dict(con.execute('SELECT key,value FROM meta'))
            if stored.get('signature') == compact(signature) and stored.get('version') == str(INDEX_VERSION):
                con.execute('SELECT skc,sheet,row,section FROM positions LIMIT 1').fetchall()
                return 'hit'
    except (sqlite3.DatabaseError, OSError):
        pass
    with closing(MemoryBook(memory)) as book:
        positions = book.positions()
    if fingerprint(memory) != signature:
        raise ValueError('记忆库读取期间已变更')
    cache.parent.mkdir(parents=True,exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=cache.name+'.',dir=cache.parent)
    os.close(fd)
    try:
        with closing(sqlite3.connect(tmp)) as con:
            con.executescript('CREATE TABLE meta(key TEXT PRIMARY KEY,value TEXT); CREATE TABLE positions(skc TEXT,sheet TEXT,row INTEGER,section TEXT,PRIMARY KEY(skc,sheet,row,section));')
            con.executemany('INSERT INTO meta VALUES (?,?)',[('version',str(INDEX_VERSION)),('signature',compact(signature))])
            con.executemany('INSERT INTO positions VALUES (?,?,?,?)',positions)
            con.commit()
        os.replace(tmp,cache)
    finally:
        if os.path.exists(tmp): os.unlink(tmp)
    return 'rebuilt'


def lookup(records, memory=DEFAULT_MEMORY, cache=None):
    started = perf_counter()
    memory = Path(memory)
    cache = Path(cache) if cache else memory.with_name('.'+memory.stem+'.skc.sqlite')
    if cache.resolve() == memory.resolve():
        raise ValueError('索引路径不能与记忆原表相同')
    requested = sorted({skc_text(r.get('skcId')) for r in records} - {''})
    result = {'hits':{},'memoryCopyRecordIds':[],'memoryCopyByRecord':{},'webRecordIds':[]}
    stats = {'inputRows':len(records),'uniqueSkcs':len(requested),'hitSkcs':0,'hitRows':0,'detailCells':0,'cache':'skipped'}
    if requested:
        for attempt in range(2):
            try:
                sig = fingerprint(memory)
                stats['cache'] = ensure_index(memory,cache,sig)
                positions = []
                with closing(sqlite3.connect(f'file:{cache}?mode=ro',uri=True)) as con:
                    for start in range(0,len(requested),500):
                        batch = requested[start:start+500]
                        positions.extend(con.execute(f'SELECT skc,sheet,row,section FROM positions WHERE skc IN ({",".join("?" for _ in batch)}) ORDER BY skc,sheet,row,section',batch))
                if positions:
                    with closing(MemoryBook(memory)) as book:
                        result['hits'] = book.details(positions)
                        stats['detailCells'] = book.detail_cells
                if fingerprint(memory) != sig:
                    raise ValueError('记忆库读取期间已变更')
                break
            except (OSError, ValueError, sqlite3.DatabaseError, BadZipFile, ET.ParseError, KeyError) as e:
                result['hits'] = {}
                if isinstance(e, sqlite3.DatabaseError):
                    try: cache.unlink(missing_ok=True)
                    except OSError: pass
                if attempt == 1:
                    result['warning'] = f'记忆查询失败，改为网页找款: {e}'
                    stats['cache'] = 'unavailable'
    for r in records:
        key = record_key(r)
        skc = skc_text(r.get('skcId'))
        hits = result['hits'].get(skc,[])
        if hits:
            # 完整SKC命中即复用记忆，不再打开历史链接或上传图片搜款。
            # 多条记录时，优先采购记录、有效链接与最新核验；完整候选仍保留在hits中。
            selected = max(hits, key=lambda h: (
                h.get('kind') == 'purchase',
                bool(h.get('url')),
                str(h.get('checkedAt') or ''),
                str(h.get('ref') or ''),
            ))
            result['memoryCopyRecordIds'].append(key)
            result['memoryCopyByRecord'][key] = {'skc':skc, **selected}
        else:
            result['webRecordIds'].append(key)
    stats.update(hitSkcs=len(result['hits']),hitRows=sum(map(len,result['hits'].values())),elapsedMs=round((perf_counter()-started)*1000,2))
    stats['memoryCopyRows'] = len(result['memoryCopyRecordIds'])
    stats['webSearchRows'] = len(result['webRecordIds'])
    stats['resultBytes'] = len(compact(result).encode('utf-8'))
    result['summary'] = stats
    return result


def extract_image(memory, skc, ref, output):
    """Only an explicitly requested, currently matching SKC row may load media."""
    skc = skc_text(skc)
    if not skc:
        raise ValueError('提取记忆图片必须提供完整SKC')
    with closing(MemoryBook(memory)) as book:
        found = []
        for item in book.positions():
            _, sheet, row, kind = item
            anchor = book.layout(sheet)[kind]['skc']
            if item[0] == skc and f'{sheet}!{anchor}{row}' == ref:
                found.append(item)
        if len(found) != 1:
            raise ValueError('图片引用未命中当前SKC，可能已被删除或移动')
        _,sheet,row,kind = found[0]
        image_cell = f'{book.layout(sheet)[kind]["image"]}{row}'
        cell = book.cells(sheet,lambda c:c == image_cell).get(image_cell,{})
        match = re.search(r'DISPIMG\("([^"]+)"',cell.get('formula',''),re.I)
        if not match:
            raise ValueError('命中记录不是WPS单元格图片；请直接复核商品链接或源SKU图')
        part = next((r['path'] for r in book.relationships('xl/workbook.xml').values() if r['Type'].endswith('/cellImage')),None)
        if not part: raise ValueError('记忆图片关系缺失')
        root = ET.fromstring(book.z.read(part))
        rels = book.relationships(part)
        for pic in root.findall('.//xdr:pic',NS):
            nv = pic.find('xdr:nvPicPr/xdr:cNvPr',NS)
            if nv is not None and nv.attrib.get('name') == match.group(1):
                blip = pic.find('.//a:blip',NS)
                asset = rels[blip.attrib['{'+NS['r']+'}embed']]['path']
                output = Path(output).with_suffix(Path(asset).suffix)
                output.parent.mkdir(parents=True,exist_ok=True)
                output.write_bytes(book.z.read(asset))
                return str(output)
    raise ValueError('命中记录的图片资源缺失')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command',required=True)
    find = sub.add_parser('lookup')
    find.add_argument('--records',required=True)
    find.add_argument('--out',required=True)
    find.add_argument('--cache')
    image = sub.add_parser('image')
    image.add_argument('--skc',required=True)
    image.add_argument('--ref',required=True)
    image.add_argument('--out',required=True)
    # 接受旧参数后明确拒绝，避免旧自动续写命令意外写入记忆库。
    sync = sub.add_parser('sync')
    sync.add_argument('--records',required=True)
    sync.add_argument('--matches',required=True)
    sync.add_argument('--report',required=True)
    for p in (find,image,sync): p.add_argument('--memory',default=str(DEFAULT_MEMORY))
    args = parser.parse_args()
    if args.command == 'lookup':
        result = lookup(json.loads(Path(args.records).read_text()),args.memory,args.cache)
        save_json(args.out,result)
        print(compact({'memoryLookup':result['summary'],'output':str(Path(args.out).resolve()),**({'warning':result['warning']} if 'warning' in result else {})}))
    elif args.command == 'image':
        print(compact({'image':extract_image(args.memory,args.skc,args.ref,args.out)}))
    else:
        raise SystemExit('记忆库已设置为仅人工录入：禁止脚本同步。请人工确认后直接在报单记忆.xlsx中添加或修改记录。')


if __name__ == '__main__':
    main()
