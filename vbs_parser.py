import re
import json
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional

# SAP GUI Virtual Key 完整映射
VKEY_MAP = {
    0: "ENTER",
    1: "F1",    2: "F2",    3: "F3",    4: "F4",
    5: "F5",    6: "F6",    7: "F7",    8: "F8",
    9: "F9",   10: "F10",  11: "F11",  12: "F12",
    24: "CTRL+F12",
    25: "CTRL+SHIFT+F1",
    26: "CTRL+S",
    33: "CTRL+SHIFT+F9",
    34: "CTRL+SHIFT+F10",
    70: "CTRL+E",
    71: "SHIFT+F1",  72: "SHIFT+F2",  73: "SHIFT+F3",
    74: "SHIFT+F4",  75: "SHIFT+F5",  76: "SHIFT+F6",
    77: "SHIFT+F7",  78: "SHIFT+F8",  79: "SHIFT+F9",
    80: "SHIFT+F10", 81: "SHIFT+F11", 82: "SHIFT+F12",
    83: "CTRL+F1",   84: "CTRL+F2",   85: "CTRL+F3",
    86: "CTRL+F4",   87: "CTRL+F5",   88: "CTRL+F6",
    89: "CTRL+F7",   90: "CTRL+F8",   91: "CTRL+F9",
    92: "CTRL+F10",  93: "CTRL+F11",
}

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
    """解析单行 VBScript，返回结构化步骤或 None。"""

    # 打开事务码：okcd.text = "/nXXXX"
    m = re.match(
        r'session\.findById\("wnd\[0\]/tbar\[0\]/okcd"\)\.text\s*=\s*"(/n)?([^"]+)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "open_tcode",
                "params": {"tcode": m.group(2)}, "raw_line": line}

    # 填写字段：.text = "value"
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.text\s*=\s*"([^"]*)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "fill_field",
                "params": {"field": m.group(1), "value": m.group(2)}, "raw_line": line}

    # 按功能键：.sendVKey N
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.sendVKey\s+(\d+)',
        line, re.IGNORECASE,
    )
    if m:
        key = VKEY_MAP.get(int(m.group(2)), f"VKey{m.group(2)}")
        return {"seq": seq, "action": "press_key",
                "params": {"key": key}, "raw_line": line}

    # 点击按钮：.press
    m = re.match(r'session\.findById\("([^"]+)"\)\.press\b', line, re.IGNORECASE)
    if m:
        return {"seq": seq, "action": "click_button",
                "params": {"button_id": m.group(1)}, "raw_line": line}

    # 选择 ComboBox / DropDown：.key = "value"
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.key\s*=\s*"([^"]*)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "select_combo",
                "params": {"field": m.group(1), "key": m.group(2)}, "raw_line": line}

    # 勾选/取消 CheckBox：.selected = true/false
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.selected\s*=\s*(true|false|-1|0)',
        line, re.IGNORECASE,
    )
    if m:
        checked = m.group(2).lower() in ("true", "-1")
        return {"seq": seq, "action": "set_checkbox",
                "params": {"field": m.group(1), "checked": checked}, "raw_line": line}

    # 选择 RadioButton：.select
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.select\b(?!\w)',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "select_radio",
                "params": {"field": m.group(1)}, "raw_line": line}

    # 双击：.doubleClick / .doubleClickItem / .doubleClickNode
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.(doubleClick(?:Item|Node)?)\b(.*)',
        line, re.IGNORECASE,
    )
    if m:
        params = {"field": m.group(1)}
        # doubleClickItem row, col / doubleClickNode "key"
        args = m.group(3).strip()
        if args:
            params["args"] = args
        return {"seq": seq, "action": "double_click",
                "params": params, "raw_line": line}

    # GridView 选中行：.selectedRows = "0" / "0,1,2"
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.selectedRows\s*=\s*"([^"]*)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "select_rows",
                "params": {"grid_id": m.group(1), "rows": m.group(2)}, "raw_line": line}

    # GridView 设置当前单元格：.setCurrentCell row, "col"
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.setCurrentCell\s+(\d+)\s*,\s*"([^"]*)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "set_current_cell",
                "params": {"grid_id": m.group(1), "row": int(m.group(2)),
                           "column": m.group(3)}, "raw_line": line}

    # GridView 点击当前单元格：.clickCurrentCell
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.clickCurrentCell\b',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "click_current_cell",
                "params": {"grid_id": m.group(1)}, "raw_line": line}

    # GridView 修改单元格：.modifyCell row, "col", "value"
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.modifyCell\s+(\d+)\s*,\s*"([^"]*)"\s*,\s*"([^"]*)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "modify_cell",
                "params": {"grid_id": m.group(1), "row": int(m.group(2)),
                           "column": m.group(3), "value": m.group(4)}, "raw_line": line}

    # GridView 工具栏按钮：.pressToolbarButton "btn" / .pressToolbarContextButton "btn"
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.pressToolbar(?:Context)?Button\s+"([^"]*)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "press_toolbar_button",
                "params": {"grid_id": m.group(1), "button": m.group(2)}, "raw_line": line}

    # GridView 选择列：.selectColumn "col"
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.selectColumn\s+"([^"]*)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "select_column",
                "params": {"grid_id": m.group(1), "column": m.group(2)}, "raw_line": line}

    # 右键菜单：.contextMenu
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.contextMenu\b',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "context_menu",
                "params": {"field": m.group(1)}, "raw_line": line}

    # 右键菜单项：.selectContextMenuItem "item" / .selectContextMenuItemByText "text"
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.selectContextMenuItem(?:ByText)?\s+"([^"]*)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "select_context_menu_item",
                "params": {"field": m.group(1), "item": m.group(2)}, "raw_line": line}

    # 滚动条：.verticalScrollbar.position = N / .horizontalScrollbar.position = N
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.(vertical|horizontal)Scrollbar\.position\s*=\s*(\d+)',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "scroll",
                "params": {"field": m.group(1), "direction": m.group(2).lower(),
                           "position": int(m.group(3))}, "raw_line": line}

    # Tree 操作：.expandNode / .collapseNode / .selectNode / .doubleClickNode
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.(expandNode|collapseNode|selectNode)\s+"([^"]*)"',
        line, re.IGNORECASE,
    )
    if m:
        action_map = {"expandnode": "expand_node", "collapsenode": "collapse_node",
                      "selectnode": "select_node"}
        return {"seq": seq, "action": action_map[m.group(2).lower()],
                "params": {"tree_id": m.group(1), "node_key": m.group(3)}, "raw_line": line}

    # Tab 选择：tabpXXXX.select (tabstrip)
    m = re.match(
        r'session\.findById\("([^"]+)"\)\.select\b',
        line, re.IGNORECASE,
    )
    if m and "/tabp" in m.group(1).lower():
        return {"seq": seq, "action": "select_tab",
                "params": {"tab_id": m.group(1)}, "raw_line": line}

    # StartTransaction（另一种打开事务码写法）
    m = re.match(
        r'session\.StartTransaction\s+"([^"]+)"',
        line, re.IGNORECASE,
    )
    if m:
        return {"seq": seq, "action": "open_tcode",
                "params": {"tcode": m.group(1)}, "raw_line": line}

    # 未识别的 findById / session 行，保留原始内容
    if "session.findbyid" in line.lower() or "session." in line.lower():
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
    """读取 .vbs 文件并返回结构化结果。支持 UTF-16 LE（SAP 默认）和 UTF-8。"""
    raw = Path(path).read_bytes()
    if raw[:2] in (b'\xff\xfe', b'\xfe\xff'):
        content = raw.decode("utf-16")
    else:
        content = raw.decode("utf-8-sig")
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
