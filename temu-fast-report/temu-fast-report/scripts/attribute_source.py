"""Locate the current attribute library and identify its derived cache."""
from datetime import date
import json
from pathlib import Path
import re
from local_paths import attribute_dir


# Change the parser version whenever normalized material values change.
MATERIAL_PARSER_VERSION = "2026-10-04-v3"
ATTRIBUTE_CACHE_VERSION = "2026-10-06-source-v1"
LIBRARY_PREFIX = "商品在售详细数据"


def attribute_files(directory):
    root = Path(directory)
    return sorted(p for p in root.iterdir()
                  if p.is_file() and p.suffix.lower() == ".xlsx"
                  and not p.name.startswith("~$"))


def _library_date(path):
    """Prefer the export date in the name over a later copy/edit time."""
    text = path.name[len(LIBRARY_PREFIX):]
    match = re.search(r"(?:(\d{4})年)?(\d{1,2})月(\d{1,2})", text)
    year = date.today().year
    try:
        if match:
            return date(int(match[1] or year), int(match[2]), int(match[3])).toordinal()
        compact = re.search(r"(?<!\d)(\d{8}|\d{4}|\d{3})(?!\d)", text)
        if compact:
            value = compact[1]
            if len(value) == 8:
                year, value = int(value[:4]), value[4:]
            return date(year, int(value[:-2]), int(value[-2:])).toordinal()
    except ValueError:
        pass
    return 0


def default_attribute_parents():
    return [attribute_dir().parent]


def resolve_attr_dir(path=None, parents=None):
    """Honor an explicit existing library; otherwise locate the latest export."""
    if path and Path(path).is_dir():
        return str(Path(path).resolve())
    if path is None and parents is None:
        return str(attribute_dir().resolve())
    roots = list(default_attribute_parents() if parents is None else parents)
    if path:
        roots.append(Path(path).parent)
    candidates = set()
    for root in dict.fromkeys(Path(p).resolve() for p in roots):
        if not root.is_dir():
            continue
        for child in root.iterdir():
            if child.name.startswith(LIBRARY_PREFIX) and child.is_dir() and attribute_files(child):
                candidates.add(child.resolve())
    if candidates:
        latest = max(candidates, key=lambda p: (_library_date(p), p.stat().st_mtime_ns, str(p)))
        return str(latest)
    return str(Path(path).resolve()) if path else ""


def attribute_signature(directory):
    """Include the source directory, parser version and every source file stamp."""
    if not directory or not Path(directory).is_dir():
        return None
    root = Path(directory).resolve()
    files = []
    for path in attribute_files(root):
        stat = path.stat()
        files.append([path.name, stat.st_mtime_ns, stat.st_size])
    return json.dumps([ATTRIBUTE_CACHE_VERSION, MATERIAL_PARSER_VERSION, str(root), files],
                      ensure_ascii=False, separators=(",", ":"))
