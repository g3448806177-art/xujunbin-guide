#!/usr/bin/env python3
"""TEMU 快速报单 - prepare 阶段：从备货表提取记录+商品图+属性索引。

用法:
  python3 prepare.py --source <备货表.xlsx> [--work-dir <工作目录>] [--reporter <报单名>]
                     [--attr-dir <属性库目录>]

默认路径(用户环境):
  属性库: 默认使用安装时在桌面建立的“店铺属性库”文件夹；
          可用 --attr-dir 显式指定。按完整SKC_ID索引。
  工作目录: ./work (相对当前目录)
"""
import argparse
import json
import os
import re
import sys
import zipfile
import hashlib
from pathlib import Path
from urllib.parse import urlparse

import openpyxl
from openpyxl.utils import get_column_letter
from report_contract import skc_text, record_key
from memory_store import DEFAULT_MEMORY, lookup as lookup_memory, save_json
from material_format import normalize_material_attributes
from attribute_source import resolve_attr_dir, attribute_signature
from local_paths import reporter_name, check_reporter

# ============ 参数解析 ============
def parse_args():
    p = argparse.ArgumentParser(description="TEMU报单prepare阶段")
    p.add_argument("--source", required=True, help="备货表xlsx路径")
    p.add_argument("--work-dir", default="./work", help="工作目录(输出records.json/images/)")
    p.add_argument("--reporter", default=None, help="报单名；默认使用安装时设置的报单名(config/settings.json)")
    p.add_argument("--stats", default=None, help="已停用；仅兼容旧命令，传入路径将被忽略")
    p.add_argument("--attr-dir", default=None, help="商品属性库目录；默认使用桌面店铺属性库")
    p.add_argument("--order-file", default=None, help="导出订单包裹表xlsx路径(可选，半托管订单鞋需搜款网找同款，SKC列标记半托管订单)")
    p.add_argument("--memory", default=str(DEFAULT_MEMORY), help="只按SKC查询的报单记忆表")
    p.add_argument("--memory-cache", default=None, help="可选SQLite索引路径；默认在记忆表旁")
    p.add_argument("--verbose", action="store_true", help="输出逐行诊断；默认仅输出摘要")
    return p.parse_args()


ARGS = parse_args()
SOURCE_XLSX = os.path.abspath(ARGS.source)
PRODUCT_ATTR_DIR = resolve_attr_dir(ARGS.attr_dir)
WORK_DIR = os.path.abspath(ARGS.work_dir)
try:
    # 报单名由使用者自定义：命令行 --reporter 优先，其次是安装时保存的设置。
    REPORTER = check_reporter(ARGS.reporter or reporter_name())
except ValueError as exc:
    print(f"ERROR: {exc}", file=sys.stderr)
    sys.exit(2)

os.makedirs(WORK_DIR, exist_ok=True)
os.makedirs(os.path.join(WORK_DIR, "images"), exist_ok=True)
CACHE_DIR = os.path.join(WORK_DIR, "cache")
os.makedirs(CACHE_DIR, exist_ok=True)
# 索引是派生缓存，默认写入本次工作目录，避免对只读记忆原表目录产生写入依赖。
MEMORY_CACHE = os.path.abspath(ARGS.memory_cache) if ARGS.memory_cache else os.path.join(CACHE_DIR, "报单记忆.sqlite")

# ============ 索引缓存 ============
def load_cached(cache_file, signature):
    """签名匹配则返回缓存数据，否则返回None"""
    if not os.path.exists(cache_file):
        return None
    try:
        with open(cache_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("signature") == signature:
            return data.get("payload")
    except Exception:
        pass
    return None

def save_cached(cache_file, signature, payload):
    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump({"signature": signature, "payload": payload}, f, ensure_ascii=False)

# ============ 工具函数 ============
def clean_text(v):
    if v is None:
        return ""
    return str(v).strip()

def numeric(v):
    if isinstance(v, (int, float)) and v == v:  # not NaN
        return v
    text = clean_text(v)
    if not text:
        return None
    text = text.replace(",", "").rstrip("%")
    try:
        val = float(text)
        return val
    except ValueError:
        return None

def normalize_host(value):
    text = clean_text(value)
    if not text:
        return ""
    try:
        if "://" not in text:
            text = "https://" + text
        return urlparse(text).hostname.lower().lstrip("www.")
    except Exception:
        return ""

def normalize_header(v):
    return re.sub(r"\s+", "", clean_text(v)).lower()

def pick_column(headers, names, fallback=-1):
    norm = [normalize_header(h) for h in headers]
    for name in names:
        if normalize_header(name) in norm:
            return norm.index(normalize_header(name))
    return fallback

def find_header_row(rows, required):
    for i, row in enumerate(rows):
        cells = [normalize_header(c) for c in row]
        if all(normalize_header(h) in cells for h in required):
            return i
    return -1

def normalized_identifier(value):
    text = clean_text(value).upper()
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"(\d+)\.0+(?=[-_]|$)", r"\1", text)
    return text

def identifier_variants(value):
    base = normalized_identifier(value)
    if not base or base in ("/", "-"):
        return []
    variants = [base]
    without_size = re.sub(r"[-_](3[0-9]|4[0-9]|50)$", "", base)
    if without_size and without_size != base:
        variants.append(without_size)
    return list(dict.fromkeys(variants))

# ============ 解析备货表记录 ============
def long_id_text(v):
    """把 SKC_ID 长数字(可能被openpyxl读成int/float/带.0字符串)归一为纯数字字符串。"""
    return skc_text(v)


def attach_source(record, path, sheet):
    source = os.path.abspath(path)
    record["sourceFile"] = source
    record["sourceSheet"] = sheet
    prefix = hashlib.sha256((source + "\0" + sheet).encode()).hexdigest()[:16]
    record["recordId"] = f"{prefix}:{record['sourceRow']}"
    image = record.get("imagePath")
    digest = hashlib.sha256(Path(image).read_bytes()).hexdigest() if image else ""
    skc = skc_text(record.get("skcId"))
    record["groupKey"] = f"skc:{skc}:image:{digest}" if skc and digest else "record:" + record["recordId"]

def parse_record(row, source_row, cols, reporter):
    spec = clean_text(row[cols["spec"]])
    raw_sku_code = clean_text(row[cols["skuCode"]])
    skc_number = clean_text(row[cols["skcCode"]]) if cols["skcCode"] >= 0 else ""
    skc_id = long_id_text(row[cols["skcId"]]) if cols.get("skcId", -1) >= 0 else ""

    size = ""
    # Replenishment exports may append a width label, e.g. "45 wide".
    # The numeric part is the shoe size.
    wide_size = re.search(r"(?<!\d)(\d{2})\s*wide\b", spec, re.IGNORECASE)
    trailing_size = re.search(r"[-_](\d{2})$", spec)
    leading_size = re.search(r"^(\d{2})-", spec)
    if wide_size:
        size = wide_size.group(1)
    elif trailing_size:
        size = trailing_size.group(1)
    elif leading_size:
        size = leading_size.group(1)

    item_number = ""
    cost = None
    if raw_sku_code or skc_number:
        remainder = raw_sku_code or skc_number
        code_size = re.search(r"[-_](\d{2})$", remainder)
        if code_size and (not size or code_size.group(1) == size):
            if not size:
                size = code_size.group(1)
            remainder = remainder[: -len(code_size.group(0))]
        encoded_cost = re.search(r"-(\d+(?:\.\d{1,2})?)$", remainder)
        if encoded_cost:
            cost = float(encoded_cost.group(1))
            remainder = remainder[: -len(encoded_cost.group(0))]
        item_number = remainder
    else:
        fallback = re.search(r"^(?:\d{2}-)?([A-Za-z0-9_-]+?)(?=[\u3400-\u9fff]|$)", spec)
        if fallback:
            item_number = fallback.group(1)

    # 从SKU规格中提取颜色文字时不做颜色归一化。找到第一个明确颜色词后，
    # 保留它后面与SKU颜色连写的文字（如“卡黑色”“黑色内增高”“黄色内增高”）；
    # 终止码数/宽度标记后停止。来源SKU颜色必须原样进入网站备注。
    known_colors = ["深蓝色", "浅蓝色", "咖啡色", "卡其色", "墨绿色", "军绿色",
                    "黄棕色", "红棕色", "深棕色", "浅棕色", "米白色", "米黄色",
                    "黑色", "卡其", "深蓝", "浅蓝", "宝蓝", "蓝色", "黄色", "白色",
                    "灰色", "棕色", "红色", "绿色", "米白", "米色", "深棕", "咖色",
                    "驼色", "杏色", "粉色", "橙色", "紫色", "酒红", "军绿", "墨绿"]
    color = ""
    color_terms = sorted(set(known_colors + ["卡黑色", "黑色内增高", "黄色内增高"]), key=len, reverse=True)
    color_hits = [(spec.find(c), -len(c), c) for c in color_terms if spec.find(c) >= 0]
    if color_hits:
        start = min(color_hits)[0]
        color = spec[start:]
        color = re.sub(r"\s*\d{2}\s*wide\s*$", "", color, flags=re.IGNORECASE)
        color = re.sub(r"[-_ ]\d{2}\s*$", "", color).strip()
        # 规格如“0102棕色-42 wide”中的连字符只分隔颜色与尺码，
        # 尺码移除后不能残留到来源颜色字段。
        color = color.rstrip("-_ ").strip()
    if not color:
        color = re.sub(r"^\d{2}-", "", spec)
        color = re.sub(r"[-_]\d{2}$", "", color)
        if item_number:
            color = color.replace(item_number, "").strip()
        # 去掉残留的货号/字母数字-符号片段（如 42-p-3咖色 里的 p-3），只保留中文颜色
        color = re.sub(r"[A-Za-z0-9_.\-]+", "", color).strip()
        # 去掉括号备注
        color = re.sub(r"[（(].*?[）)]", "", color).strip()

    quantity = numeric(row[cols["quantity"]]) or 0
    return {
        "sourceRow": source_row,
        "size": size,
        "color": color,
        "itemNumber": item_number,
        "cost": cost,
        "quantity": quantity,
        "reporter": reporter,
        "rawSkuCode": raw_sku_code,
        "skcNumber": skc_number,
        "skcId": skc_id,
        "productAttributes": "",
        "productAttributeMatch": "not_checked",
        "groupKey": f"item:{item_number}" if item_number else f"row:{source_row}",
        "imagePath": None,
        "mainImagePath": None,
    }

# ============ 解析订单包裹表记录（半托管订单）============
def parse_order_spec(spec_text):
    """从'产品规格(中文)'解析颜色和码数。
    支持格式：
    - '标签尺寸：CN 43（美国9号）/蓝色' → 蓝色, 43
    - '【棕色】 / 标签尺码：中国40（美国7.5）' → 棕色, 40
    """
    text = clean_text(spec_text)
    if not text:
        return "", ""
    known_colors = ["深蓝色", "浅蓝色", "咖啡色", "卡其色", "墨绿色", "军绿色",
                    "黄棕色", "红棕色", "深棕色", "浅棕色", "米白色", "米黄色",
                    "黑色", "卡其", "深蓝", "浅蓝", "宝蓝", "蓝色", "黄色", "白色",
                    "灰色", "棕色", "红色", "绿色", "米白", "米色", "深棕", "咖色",
                    "驼色", "杏色", "粉色", "橙色", "紫色", "酒红", "军绿", "墨绿"]
    color = ""
    color_terms = sorted(set(known_colors + ["卡黑色", "黑色内增高", "黄色内增高"]), key=len, reverse=True)
    hits = [(text.find(c), -len(c), c) for c in color_terms if text.find(c) >= 0]
    if hits:
        start = min(hits)[0]
        color = text[start:]
        color = re.sub(r"\s*\d{2}\s*wide\s*$", "", color, flags=re.IGNORECASE)
        color = re.sub(r"[-_ ]\d{2}\s*$", "", color).strip()
        color = color.rstrip("-_ ").strip()
    bracket_match = re.search(r"【([^】]+)】", text)
    if not color and bracket_match:
        for c in sorted(set(known_colors), key=len, reverse=True):
            if c in bracket_match.group(1):
                color = c
                break
    if not color:
        slash_match = re.search(r"/([^/]+)$", text)
        if slash_match:
            for c in sorted(set(known_colors), key=len, reverse=True):
                if c in slash_match.group(1):
                    color = c
                    break
    if not color:
        for c in sorted(set(known_colors), key=len, reverse=True):
            if c in text:
                color = c
                break
    size = ""
    cn_match = re.search(r"CN\s*(\d+)", text, re.IGNORECASE)
    if cn_match:
        size = cn_match.group(1)
    if not size:
        zh_match = re.search(r"中国\s*(\d+)", text)
        if zh_match:
            size = zh_match.group(1)
    if not size:
        wide_match = re.search(r"(?<!\d)(\d{2})\s*wide\b", text, re.IGNORECASE)
        if wide_match:
            size = wide_match.group(1)
    if not size:
        label_match = re.search(r"标签尺[寸码][：:]\s*[^\d]*(\d+)", text)
        if label_match:
            size = label_match.group(1)
    return color, size


def parse_order_records(source_path, image_dir, reporter):
    """解析导出订单包裹表，返回记录列表。
    表结构：产品图片(嵌入) | 产品规格(中文) | 本地备注
    记录标记 sourceType='semi_managed_order'，build时SKC列填'（半托管订单）'。
    """
    if not source_path or not os.path.exists(source_path):
        return []
    wb = openpyxl.load_workbook(source_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    header_idx = find_header_row(rows, ["产品规格"])
    data_start = header_idx + 1 if header_idx >= 0 else 0
    headers = rows[header_idx] if header_idx >= 0 else []
    spec_col = pick_column(headers, ["产品规格", "规格"], 1) if headers else 1
    images_by_row = extract_row_images(source_path, image_dir, 0, -1)
    records = []
    for offset in range(data_start, len(rows)):
        row = rows[offset]
        if not any(clean_text(c) for c in row):
            continue
        spec_text = clean_text(row[spec_col]) if spec_col < len(row) else ""
        if not spec_text:
            continue
        source_row = offset + 1
        color, size = parse_order_spec(spec_text)
        rec = {
            "sourceRow": source_row,
            "size": size,
            "color": color,
            "itemNumber": "",
            "cost": None,
            "quantity": 1,
            "reporter": reporter,
            "rawSkuCode": "",
            "skcNumber": "",
            "skcId": "",
            "productAttributes": "",
            "productAttributeMatch": "not_checked",
            "groupKey": f"order-row:{source_row}",
            "imagePath": None,
            "mainImagePath": None,
            "sourceType": "semi_managed_order",
            "orderSpec": spec_text,
        }
        row_imgs = images_by_row.get(source_row, {})
        rec["imagePath"] = row_imgs.get("sku")
        attach_source(rec, source_path, ws.title)
        records.append(rec)
    return records

# ============ 提取嵌入图片 ============
def extract_row_images(source_path, image_dir, sku_img_col=-1, main_img_col=-1):
    """提取备货表嵌入图片，按锚点列区分 SKU 图和链接主图。

    sku_img_col/main_img_col 为表头列号(0-indexed, 与 drawing 锚点 col 一致)。
    返回 {excel_row: {"sku": path|None, "main": path|None}}。
    - SKU 图输出 row-{N}.ext；链接主图输出 row-{N}-main.ext
    - 旧备货表只有一列图片时(无法识别主图列)，全部作为 SKU 图，保证向后兼容
    """
    os.makedirs(image_dir, exist_ok=True)
    images_by_row = {}

    def classify(anchor_col):
        if main_img_col >= 0 and anchor_col == main_img_col:
            return "main"
        if sku_img_col >= 0 and anchor_col == sku_img_col:
            return "sku"
        # 列无法对应时：没有主图列则一律当 SKU 图；否则丢弃无关列
        if main_img_col < 0:
            return "sku"
        return None

    with zipfile.ZipFile(source_path) as z:
        names = z.namelist()
        drawing_paths = sorted(n for n in names if re.match(r"^xl/drawings/drawing\d+\.xml$", n, re.I))
        for drawing_path in drawing_paths:
            drawing_xml = z.read(drawing_path).decode("utf-8", errors="ignore")
            rels_path = os.path.join(os.path.dirname(drawing_path), "_rels", os.path.basename(drawing_path) + ".rels")
            rels_xml = z.read(rels_path).decode("utf-8", errors="ignore") if rels_path in names else ""
            rels = {}
            # Relationship 标签内 Id/Target 属性顺序不固定(WPS与openpyxl相反)，分别提取
            for rel_m in re.finditer(r"<Relationship\b[^>]*?/?>", rels_xml):
                tag = rel_m.group(0)
                id_m = re.search(r'Id="([^"]+)"', tag)
                tgt_m = re.search(r'Target="([^"]+)"', tag)
                if not id_m or not tgt_m:
                    continue
                target = tgt_m.group(1)
                if target.startswith("/"):
                    # 以 / 开头表示相对 zip 根的绝对路径
                    target = target.lstrip("/")
                else:
                    target = os.path.normpath(os.path.join(os.path.dirname(drawing_path), target))
                rels[id_m.group(1)] = target.replace("\\", "/")

            # 同时兼容 xdr: 前缀(WPS导出)和默认命名空间无前缀(openpyxl)两种写法
            anchor_re = re.compile(r"<(?:xdr:)?(oneCellAnchor|twoCellAnchor)>([\s\S]*?)</(?:xdr:)?\1>")
            for m in anchor_re.finditer(drawing_xml):
                body = m.group(2)
                row_match = re.search(r"<(?:xdr:)?row>(\d+)</(?:xdr:)?row>", body)
                col_match = re.search(r"<(?:xdr:)?col>(\d+)</(?:xdr:)?col>", body)
                embed_match = re.search(r'<a:blip\b[^>]*r:embed="([^"]+)"', body)
                if not row_match or not embed_match:
                    continue
                anchor_col = int(col_match.group(1)) if col_match else -1
                kind = classify(anchor_col)
                if kind is None:
                    continue
                excel_row = int(row_match.group(1)) + 1
                entry = images_by_row.setdefault(excel_row, {"sku": None, "main": None})
                if entry[kind]:
                    continue
                media_path = rels.get(embed_match.group(1))
                if not media_path or media_path not in names:
                    continue
                ext = os.path.splitext(media_path)[1].lower() or ".jpeg"
                suffix = "-main" if kind == "main" else ""
                out_path = os.path.join(image_dir, f"row-{excel_row}{suffix}{ext}")
                with open(out_path, "wb") as f:
                    f.write(z.read(media_path))
                entry[kind] = out_path
    return images_by_row

def make_small_copy(src_path, small_dir, max_size=1000):
    """把图片缩放到最长边≤max_size，输出到 small_dir；失败时回退原图路径。"""
    if not src_path or not os.path.exists(src_path):
        return None
    os.makedirs(small_dir, exist_ok=True)
    out_path = os.path.join(small_dir, os.path.basename(src_path))
    if os.path.exists(out_path) and os.path.getmtime(out_path) >= os.path.getmtime(src_path):
        return out_path
    try:
        from PIL import Image
        with Image.open(src_path) as im:
            im = im.convert("RGB")
            if max(im.size) > max_size:
                im.thumbnail((max_size, max_size), Image.LANCZOS)
            im.save(out_path, "JPEG", quality=90)
        return out_path
    except Exception:
        return src_path

# ============ 商品属性索引 ============
def extract_material_attributes(value):
    return normalize_material_attributes(clean_text(value))

def load_product_attribute_index(directory):
    """索引商品属性库。兼容两种结构:
    - 新库"商品在售详细数据*": 列含 SKC_ID(11位长数字,主键)/SKC货号/商品属性(多行文本)/SKU货号,
      SKC级信息只在该SKC首行出现, 其下SKU行留空(自动向下填充)。
    - 旧库"TEMU商品详细信息": 以SKC货号为主。
    索引键: SKC_ID长数字(优先) + SKC货号变体 + SKU货号变体(兜底)。
    """
    by_identifier = {}
    details = {"directoryPath": directory, "fileCount": 0, "indexedRowCount": 0,
               "skcIdCount": 0, "skippedFiles": []}
    if not directory or not os.path.isdir(directory):
        return by_identifier, details
    files = sorted(f for f in os.listdir(directory) if f.lower().endswith(".xlsx") and not f.startswith("~$"))
    details["fileCount"] = len(files)
    for fname in files:
        fpath = os.path.join(directory, fname)
        try:
            wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            rows = list(ws.iter_rows(values_only=True))
            header_idx = find_header_row(rows, ["商品属性"])
            if header_idx < 0:
                details["skippedFiles"].append(f"{fname}：缺少商品属性列")
                wb.close()
                continue
            headers = rows[header_idx]
            cols = {
                "attributes": pick_column(headers, ["商品属性"], -1),
                "skcId": pick_column(headers, ["SKC_ID", "SKC-ID", "SKC编号", "SKC"], -1),
                "skcCode": pick_column(headers, ["SKC货号"], -1),
                "skuCode": pick_column(headers, ["SKU货号"], -1),
            }
            if cols["skcId"] < 0 and cols["skcCode"] < 0:
                details["skippedFiles"].append(f"{fname}：缺少SKC_ID或SKC货号列")
                wb.close()
                continue
            # 向下填充（新库SKC级字段仅首行有值）
            cur_id, cur_code, cur_materials = "", "", ""
            seen_id_in_file = set()
            for ri in range(header_idx + 1, len(rows)):
                row = rows[ri]
                if not any(clean_text(c) for c in row):
                    continue
                row_id = long_id_text(row[cols["skcId"]]) if cols["skcId"] >= 0 else ""
                row_code = clean_text(row[cols["skcCode"]]) if cols["skcCode"] >= 0 else ""
                row_materials = extract_material_attributes(row[cols["attributes"]]) if cols["attributes"] >= 0 else ""
                if row_id:
                    cur_id = row_id
                if row_code:
                    cur_code = row_code
                if row_id or row_code:
                    if not row_materials:
                        cur_materials = ""
                if row_materials:
                    cur_materials = row_materials
                if not cur_materials:
                    continue
                entry = {"materials": cur_materials, "sourceFile": fname, "sourceRow": ri + 1}
                identifiers = set()
                if cur_id and re.match(r"^\d{6,}$", cur_id):
                    identifiers.add(cur_id)  # SKC_ID长数字主键，原样保留
                identifiers.update(identifier_variants(cur_code))  # SKC货号
                if cols["skuCode"] >= 0:
                    identifiers.update(identifier_variants(row[cols["skuCode"]]))  # SKU货号兜底
                for ident in identifiers:
                    if re.match(r"^\d{2}$", ident):
                        continue
                    existing = by_identifier.setdefault(ident, [])
                    if not any(e["materials"] == entry["materials"] and e["sourceFile"] == entry["sourceFile"] and e["sourceRow"] == entry["sourceRow"] for e in existing):
                        existing.append(entry)
                if identifiers:
                    details["indexedRowCount"] += 1
                    if cur_id and cur_id not in seen_id_in_file:
                        seen_id_in_file.add(cur_id)
                        details["skcIdCount"] += 1
            wb.close()
        except Exception as e:
            details["skippedFiles"].append(f"{fname}：读取失败 {e}")
    return by_identifier, details

def resolve_product_attributes(record, attr_index):
    # 有SKC_ID时只精确匹配；仅备货表缺SKC_ID才按货号兜底。
    skc_id = long_id_text(record.get("skcId"))
    tried = []
    id_matches = []
    if skc_id and re.match(r"^\d{6,}$", skc_id):
        tried.append(f"SKC_ID={skc_id}")
        id_matches = attr_index.get(skc_id, [])
    id_distinct = list(dict.fromkeys(e["materials"] for e in id_matches if e["materials"]))
    if len(id_distinct) == 1:
        return {"materials": id_distinct[0], "state": "matched", "matchBy": "skcId"}
    if len(id_distinct) > 1:
        return {"materials": "", "state": "ambiguous", "matchBy": "skcId"}
    if skc_id:
        return {"materials": "", "state": "not_found", "matchBy": "skcId"}
    # 兜底：SKC货号
    identifiers = set()
    identifiers.update(identifier_variants(record.get("skcNumber", "")))
    identifiers.update(identifier_variants(record.get("rawSkuCode", "")))
    matches = []
    for ident in identifiers:
        matches.extend(attr_index.get(ident, []))
    distinct = list(dict.fromkeys(e["materials"] for e in matches if e["materials"]))
    if len(distinct) == 1:
        return {"materials": distinct[0], "state": "matched", "matchBy": "skcCode"}
    if len(distinct) > 1:
        return {"materials": "", "state": "ambiguous", "matchBy": "skcCode"}
    return {"materials": "", "state": "not_found", "matchBy": ""}

# ============ 主流程 ============
def main():
    # 1. 读取备货表
    wb = openpyxl.load_workbook(SOURCE_XLSX, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    source_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    header_idx = find_header_row(source_rows, ["SKU规格", "下单数量"])
    if header_idx < 0:
        print("ERROR: 找不到 SKU规格 和 下单数量 表头", file=sys.stderr)
        sys.exit(1)
    headers = source_rows[header_idx]
    cols = {
        "quantity": pick_column(headers, ["下单数量"], 4),
        "spec": pick_column(headers, ["SKU规格"], 6),
        "skuCode": pick_column(headers, ["SKU货号"], 7),
        "skcCode": pick_column(headers, ["SKC货号"], -1),
        "skcId": pick_column(headers, ["SKC_ID", "SKC-ID", "SKC编号", "SKC"], -1),
        "skuImage": pick_column(headers, ["SKU图片", "SKU图"], -1),
        "mainImage": pick_column(headers, ["链接主图", "商品主图", "主图", "链接图片", "产品图片"], -1),
    }
    print(f"备货表表头行={header_idx+1}, 列映射: {cols}")

    # 2. 提取图片（SKU图 + 链接主图）
    images_dir = os.path.join(WORK_DIR, "images")
    small_dir = os.path.join(WORK_DIR, "images-small")
    main_small_dir = os.path.join(WORK_DIR, "images-main-small")
    images_by_row = extract_row_images(SOURCE_XLSX, images_dir, cols["skuImage"], cols["mainImage"])
    sku_img_count = sum(1 for v in images_by_row.values() if v.get("sku"))
    main_img_count = sum(1 for v in images_by_row.values() if v.get("main"))
    print(f"提取图片: SKU图 {sku_img_count} 张, 链接主图 {main_img_count} 张")

    # 3. 解析记录
    records = []
    for offset in range(header_idx + 1, len(source_rows)):
        row = source_rows[offset]
        if not any(clean_text(c) for c in row):
            continue
        if not clean_text(row[cols["spec"]]) and not clean_text(row[cols["skuCode"]]):
            continue
        source_row = offset + 1
        rec = parse_record(row, source_row, cols, REPORTER)
        row_imgs = images_by_row.get(source_row, {})
        rec["imagePath"] = row_imgs.get("sku")
        rec["mainImagePath"] = row_imgs.get("main")
        # 预生成≤1000的搜图用小图
        rec["smallImagePath"] = make_small_copy(rec["imagePath"], small_dir)
        rec["smallMainImagePath"] = make_small_copy(rec["mainImagePath"], main_small_dir)
        attach_source(rec, SOURCE_XLSX, ws.title)
        records.append(rec)

    # 3b. 解析订单包裹表（半托管订单，可选）
    order_file = ARGS.order_file
    order_records = []
    if order_file and os.path.exists(order_file):
        order_images_dir = os.path.join(WORK_DIR, "images", "orders")
        order_records = parse_order_records(order_file, order_images_dir, REPORTER)
        for rec in order_records:
            rec["smallImagePath"] = make_small_copy(rec["imagePath"], os.path.join(small_dir,"orders"))
            rec["smallMainImagePath"] = None
            records.append(rec)
        print(f"订单包裹表: 解析 {len(order_records)} 条半托管订单记录")
        for r in order_records if ARGS.verbose else []:
            print(f"  订单行{r['sourceRow']}: 颜色={r['color']}, 尺码={r['size']}, 规格='{r['orderSpec'][:40]}', 图={'有' if r['imagePath'] else '无'}")

    print(f"解析记录: {len(records)} 条 (备货表 {len(records)-len(order_records)} + 订单包裹 {len(order_records)})")
    for r in records if ARGS.verbose else []:
        print(f"  行{r['sourceRow']}: 货号={r['itemNumber']}, 颜色={r['color']}, 尺码={r['size']}, 成本={r['cost']}, 数量={r['quantity']}, SKU图={'有' if r['imagePath'] else '无'}, 主图={'有' if r['mainImagePath'] else '无'}")

    # 4. 商品属性（带缓存：源目录文件未变则直接加载）
    attr_sig = attribute_signature(PRODUCT_ATTR_DIR)
    attr_cache = load_cached(os.path.join(CACHE_DIR, "product-attributes.json"), attr_sig)
    if attr_cache is not None:
        attr_index = attr_cache["index"]
        attr_details = attr_cache["details"]
        print(f"\n商品属性索引: 缓存命中 ({attr_details['indexedRowCount']} 行, {attr_details['fileCount']} 文件)")
    else:
        attr_index, attr_details = load_product_attribute_index(PRODUCT_ATTR_DIR)
        save_cached(os.path.join(CACHE_DIR, "product-attributes.json"), attr_sig, {"index": attr_index, "details": attr_details})
        print(f"\n商品属性索引: 重建并缓存 ({attr_details['fileCount']} 文件, {attr_details['indexedRowCount']} 行, 跳过: {attr_details['skippedFiles']})")
    attr_summary = {"matched": 0, "ambiguous": 0, "notFound": 0}
    match_by = {"skcId": 0, "skcCode": 0}
    for rec in records:
        resolved = resolve_product_attributes(rec, attr_index)
        rec["productAttributes"] = resolved["materials"]
        rec["productAttributeMatch"] = resolved["state"]
        rec["productAttributeMatchBy"] = resolved.get("matchBy", "")
        attr_summary[resolved["state"] if resolved["state"] != "not_found" else "notFound"] += 1
        if resolved["state"] == "matched" and resolved.get("matchBy"):
            match_by[resolved["matchBy"]] = match_by.get(resolved["matchBy"], 0) + 1
    print(f"属性匹配: {attr_summary}; 命中方式 {match_by}")

    # 时效获取已取消；兼容旧输出字段，不读取原表或旧缓存。
    by_host = {}

    # 6. 归并搜图组
    groups_by_key = {}
    for rec in records:
        if rec["groupKey"] not in groups_by_key:
            groups_by_key[rec["groupKey"]] = {
                "groupKey": rec["groupKey"],
                "itemNumber": rec["itemNumber"],
                "colors": [],
                "sourceRows": [],
                "recordIds": [],
                "imagePath": None,
                "mainImagePath": None,
                "smallImagePath": None,
                "smallMainImagePath": None,
                "candidates": [],
            }
        g = groups_by_key[rec["groupKey"]]
        g["sourceRows"].append(rec["sourceRow"])
        g["recordIds"].append(record_key(rec))
        if rec["color"] and rec["color"] not in g["colors"]:
            g["colors"].append(rec["color"])
        if not g["imagePath"] and rec["imagePath"]:
            g["imagePath"] = rec["imagePath"]
        if not g["mainImagePath"] and rec["mainImagePath"]:
            g["mainImagePath"] = rec["mainImagePath"]
        if not g["smallImagePath"] and rec.get("smallImagePath"):
            g["smallImagePath"] = rec["smallImagePath"]
        if not g["smallMainImagePath"] and rec.get("smallMainImagePath"):
            g["smallMainImagePath"] = rec["smallMainImagePath"]
    groups = list(groups_by_key.values())
    print(f"搜图组: {len(groups)} 组")
    for g in groups if ARGS.verbose else []:
        print(f"  {g['groupKey']}: 货号={g['itemNumber']}, 颜色={g['colors']}, 源行={g['sourceRows']}, SKU图={'有' if g['imagePath'] else '无'}, 主图={'有' if g['mainImagePath'] else '无'}")

    # 7. 输出 JSON
    def write_json(path, data):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.write("\n")

    write_json(os.path.join(WORK_DIR, "records.json"), records)
    memory_result = lookup_memory(records, ARGS.memory, MEMORY_CACHE)
    save_json(os.path.join(WORK_DIR,"memory-hits.json"), memory_result)
    print(json.dumps({"memoryLookup":memory_result["summary"], **({"warning":memory_result["warning"]} if "warning" in memory_result else {})},ensure_ascii=False))
    search_groups = [{k: v for k, v in g.items() if k != "candidates"} for g in groups]
    write_json(os.path.join(WORK_DIR, "search-groups.json"), search_groups)
    write_json(os.path.join(WORK_DIR, "merchant-stats.json"), {"status": "disabled", "sourcePath": None, "byHost": by_host})
    write_json(os.path.join(WORK_DIR, "candidates.json"), groups)
    manifest = {
        "sourcePath": SOURCE_XLSX,
        "reporter": REPORTER,
        "statsPath": None,
        "merchantStatsStatus": "disabled",
        "productAttributesDir": PRODUCT_ATTR_DIR,
        "recordCount": len(records),
        "quantityTotal": sum(r["quantity"] for r in records),
        "searchGroupCount": len(groups),
        "extractedImageCount": sum(1 for r in records if r["imagePath"]),
        "extractedMainImageCount": sum(1 for r in records if r["mainImagePath"]),
        "indexedMerchantCount": len(by_host),
        "productAttributeFiles": attr_details["fileCount"],
        "productAttributeIndexRows": attr_details["indexedRowCount"],
        "productAttributeSummary": attr_summary,
        "productAttributeSkippedFiles": attr_details["skippedFiles"],
        "memoryPath": os.path.abspath(ARGS.memory),
        "memorySummary": memory_result["summary"],
    }
    write_json(os.path.join(WORK_DIR, "manifest.json"), manifest)

    print("\n===== prepare 完成 =====")
    print(json.dumps({
        "command": "prepare",
        "reporter": REPORTER,
        "workDir": WORK_DIR,
        "recordCount": len(records),
        "quantityTotal": sum(r["quantity"] for r in records),
        "searchGroupCount": len(groups),
        "extractedImageCount": sum(1 for r in records if r["imagePath"]),
        "extractedMainImageCount": sum(1 for r in records if r["mainImagePath"]),
        "indexedMerchantCount": len(by_host),
        "productAttributeSummary": attr_summary,
    }, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    main()
