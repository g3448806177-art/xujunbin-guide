"""Normalize TEMU shoe material attributes into the report's fixed five-line layout."""

import re


MATERIAL_FIELDS = (
    ("鞋面材料", ("鞋面材料",)),
    ("鞋面材料（副）", ("鞋面材料（副）", "鞋面（副）", "副鞋面")),
    ("内里材料", ("内里材料",)),
    ("鞋底材料", ("鞋底材料",)),
    ("鞋垫材料", ("鞋垫材料",)),
)
_INPUT_LABELS = tuple(alias for _label, aliases in MATERIAL_FIELDS for alias in aliases)
_LABELS = "|".join(re.escape(label) for label in sorted(_INPUT_LABELS, key=len, reverse=True))


def _clean_material_value(value):
    text = str(value or "").strip()
    text = re.sub(r"^[\s\-—–•·]+", "", text)
    text = re.sub(r"[\s\-—–•·]+$", "", text)
    # Attribute sources occasionally leave a delimiter after the value
    # (for example ``鞋面材料：织物:``).  The report format keeps only the
    # material value, so trim both list punctuation and trailing colons.
    text = text.strip("，,；;：:")
    # A nested label such as “鞋面材料：PU（副鞋面：织物）” leaves one
    # unmatched bracket on each captured value. Keep balanced brackets that
    # are part of a material name, such as “超细纤维（超纤）”.
    text = re.sub(r"[（(]\s*$", "", text)
    if text.endswith(("）", ")")) and not any(char in text for char in "（("):
        text = text[:-1].rstrip()
    return text


def _material_parts(value):
    """Split a material list while preserving source order and removing duplicates."""
    parts = []
    for part in re.split(r"[、，,；;]+", _clean_material_value(value)):
        # Some exports use a colon around each list item, for example
        # ``织物:;PU:``.  Clean every item after splitting so neither colon
        # becomes part of the material name nor hides the second material.
        part = _clean_material_value(part)
        if part and part not in parts:
            parts.append(part)
    return parts


def normalize_material_attributes(value):
    """Return five lines; normalize multiple main-upper materials without merging the secondary field."""
    found = {}
    if isinstance(value, dict):
        for label, aliases in MATERIAL_FIELDS:
            for alias in aliases:
                material = _clean_material_value(value.get(alias))
                if material:
                    found[label] = material
                    break
    else:
        text = str(value or "").strip()
        if text:
            # A semicolon can separate material values (``织物:;PU:``), so it
            # cannot by itself end a field.  Stop at a following known field,
            # a line break, or end of text. This preserves every main-upper
            # material while retaining the five-field extraction boundary.
            pattern_tail = rf"(?=(?:{_LABELS})\s*[：:]|[\r\n]|$)"
            for label, aliases in MATERIAL_FIELDS:
                for alias in aliases:
                    match = re.search(
                        rf"{re.escape(alias)}\s*[：:]\s*(.*?){pattern_tail}",
                        text,
                        flags=re.DOTALL,
                    )
                    if match:
                        material = _clean_material_value(match.group(1))
                        if material:
                            found[label] = material
                            break
    if not found:
        return ""
    output = (
        ("鞋面材料", "+".join(_material_parts(found.get("鞋面材料", "")))),
        ("鞋面材料（副）", "、".join(_material_parts(found.get("鞋面材料（副）", "")))),
        ("内里材料", "、".join(_material_parts(found.get("内里材料", "")))),
        ("鞋底材料", "、".join(_material_parts(found.get("鞋底材料", "")))),
        ("鞋垫材料", "、".join(_material_parts(found.get("鞋垫材料", "")))),
    )
    return "\n".join(f"{label}：{material}" for label, material in output)
