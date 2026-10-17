"""Per-user settings written by the Windows installer: desktop paths and the reporter name."""
import json
import re
from pathlib import Path


CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "settings.json"
REPORTER_MAX_LENGTH = 20
# The reporter name is part of the report file name, so Windows file-name characters are rejected.
_INVALID_REPORTER_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def _config():
    try:
        value = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
        return value if isinstance(value, dict) else {}
    except (OSError, ValueError):
        return {}


def desktop_path():
    configured = _config().get("desktop")
    return Path(configured) if configured else Path.home() / "Desktop"


def memory_path():
    configured = _config().get("memory")
    return Path(configured) if configured else desktop_path() / "报单记忆.xlsx"


def attribute_dir():
    configured = _config().get("attribute_dir")
    return Path(configured) if configured else desktop_path() / "店铺属性库"


def reporter_name():
    """报单名由使用者在安装时自定义；未设置时返回空字符串。"""
    value = _config().get("reporter")
    return value.strip() if isinstance(value, str) else ""


def check_reporter(value):
    """校验报单名：必填、不超过20个字、不含文件名禁用字符。返回去掉首尾空格后的名字。"""
    text = str(value or "").strip()
    if not text:
        raise ValueError(
            f"未设置报单名：请重新运行安装程序设置报单名，或在命令中加 --reporter <报单名>。"
            f"（配置文件：{CONFIG_PATH}）"
        )
    if len(text) > REPORTER_MAX_LENGTH:
        raise ValueError(f"报单名不能超过{REPORTER_MAX_LENGTH}个字：{text}")
    if _INVALID_REPORTER_CHARS.search(text):
        raise ValueError(f'报单名不能包含 \\ / : * ? " < > | 等文件名禁用字符：{text}')
    return text
