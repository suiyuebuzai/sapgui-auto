import re
import json
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional

VKEY_MAP = {0: "ENTER", 3: "F3", 8: "F8", 11: "F11", 12: "F12", 26: "CTRL+S"}

NOISE_PATTERNS = [
    r'\.setFocus\b',
    r'\.caretPosition\s*=',
    r'\.maximize\b',
    r'\.iconName\s*=',
]

BOILERPLATE_STARTS = (
    "If Not IsObject",
    "Set SapGuiAuto",
    "Set application",
    "Set connection",
    "Set session",
    "Set WScript",
    "End If",
)


def _is_boilerplate(line: str) -> bool:
    return any(line.startswith(p) for p in BOILERPLATE_STARTS)


def _is_noise(line: str) -> bool:
    return any(re.search(p, line) for p in NOISE_PATTERNS)


def _parse_line(line: str, seq: int) -> Optional[Dict[str, Any]]:
    # 打开事务码
    m = re.match(
        r'session\.findById\("wnd\[0\]/tbar\[0\]/okcd"\)\.text\s*=\s*"/n([^"]+)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "open_tcode",
                "params": {"tcode": m.group(1)}, "raw_line": line}

    # 填写字段
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.text\s*=\s*"([^"]*)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "fill_field",
                "params": {"field": m.group(1), "value": m.group(2)}, "raw_line": line}

    # 按功能键
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.sendVKey\s+(\d+)',
        line, re.IGNORECASE,
    )
    if m:
        key = VKEY_MAP.get(int(m.group(2)), f"VKey{m.group(2)}")
        return {"seq": seq, "action": "press_key",
                "params": {"key": key}, "raw_line": line}

    # 点击按钮
    m = re.match(r'session\.findById\("([^"]+)"\)\.press\b', line, re.IGNORECASE)
    if m:
        return {"seq": seq, "action": "click_button",
                "params": {"button_id": m.group(1)}, "raw_line": line}

    # 未识别的 findById 行，保留原始内容
    if "session.findbyid" in line.lower():
        return {"seq": seq, "action": "unknown", "params": {}, "raw_line": line}

    return None


def parse_vbs(content: str) -> List[Dict[str, Any]]:
    """将 VBScript 字符串解析为结构化步骤列表。"""
    steps, seq = [], 1
    for line in content.splitlines():
        line = line.strip()
        if not line or line.startswith("'"):
            continue
        if _is_boilerplate(line):
            continue
        if _is_noise(line):
            continue
        step = _parse_line(line, seq)
        if step:
            steps.append(step)
            seq += 1
    return steps


def parse_vbs_file(path: str) -> Dict[str, Any]:
    """读取 .vbs 文件并返回结构化结果。"""
    content = Path(path).read_text(encoding="utf-8-sig")
    return {
        "raw_file": str(path),
        "parsed_at": datetime.now().isoformat(),
        "steps": parse_vbs(content),
    }


def save_steps(data: Dict[str, Any], output_path: str = ".sap_record/steps.json"):
    """将解析结果写入 JSON 文件。"""
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
