# SAPGUI Auto 实现计划

> **面向执行者：** 请使用 superpowers:subagent-driven-development 或 superpowers:executing-plans 逐任务执行本计划。步骤用 `- [ ]` 跟踪进度。

**目标：** 从零构建 SAPGUI Auto 完整系统——SAP 执行层、录制管道（解析/AI增强/导出）、MCP Server、Web Chat。

**架构概要：** 三层模型。执行层（`sap_robot.py`）封装 SAP COM 原子操作；能力库（`sap_capabilities.yaml`）存储业务流程；录制管道（4个模块）生成能力库；接口层有两条路：MCP Server（与 Claude Desktop 集成）和 Web Chat（`llm_agent.py` + `web_server.py`，浏览器访问，Function Calling 驱动）。

**技术栈：** Python 3.8+、pywin32、pyyaml、anthropic SDK（claude-haiku-4-5-20251001 / claude-sonnet-4-6）、mcp>=1.0.0、fastapi、uvicorn、pytest

---

## 文件清单

| 文件 | 操作 | 职责 |
|---|---|---|
| `sap_robot.py` | 新建 | SAP GUI 原子操作封装，含 execute_sequence |
| `vbs_parser.py` | 新建 | VBScript → steps.json |
| `doc_exporter.py` | 新建 | enriched.json → YAML 追加 + MD 生成 |
| `step_enricher.py` | 新建 | LLM 描述生成 + 交互确认 → enriched.json |
| `sap_recorder.py` | 新建 | CLI 入口，串联各阶段 |
| `sap_mcp_server.py` | 新建 | MCP Server，将能力库暴露为 AI 工具 |
| `llm_agent.py` | 新建 | Anthropic Function Calling 会话引擎 |
| `web_server.py` | 新建 | FastAPI Web Chat 服务（含内联 HTML） |
| `sap_capabilities.yaml` | 新建 | 初始能力库（含登录SAP能力） |
| `sap_connections.yaml` | 新建 | SAP 系统连接凭据配置（不入版本控制） |
| `requirements.txt` | 新建 | 依赖清单 |
| `tests/fixtures/sample.vbs` | 新建 | 测试用 SAP 录制 fixture |
| `tests/test_vbs_parser.py` | 新建 | vbs_parser 单元测试 |
| `tests/test_doc_exporter.py` | 新建 | doc_exporter 单元测试 |
| `tests/test_step_enricher.py` | 新建 | step_enricher 单元测试（mock LLM） |
| `tests/test_llm_agent.py` | 新建 | llm_agent 单元测试（mock LLM） |

---

## 任务一：项目初始化

**文件：**
- 新建：`requirements.txt`
- 新建：`.gitignore`
- 新建：`sap_capabilities.yaml`
- 新建：`sap_connections.yaml`
- 新建：`tests/fixtures/sample.vbs`

- [ ] **步骤 1：安装依赖**

```bash
pip install pywin32 pyyaml anthropic "mcp>=1.0.0" fastapi uvicorn pytest
```

预期：所有包安装成功

- [ ] **步骤 2：新建 `requirements.txt`**

```
pywin32
pyyaml
anthropic
mcp>=1.0.0
fastapi
uvicorn
pytest
```

- [ ] **步骤 3：新建 `.gitignore`**

```
.sap_record/
sap_connections.yaml
__pycache__/
*.pyc
.pytest_cache/
*.egg-info/
```

- [ ] **步骤 4：新建初始能力库 `sap_capabilities.yaml`（含登录能力）**

```yaml
capabilities:
  - name: 登录SAP
    description: 登录指定 SAP 系统
    tcode: ""
    parameters:
      - name: system_name
        description: SAP 系统连接名称（对应 SAP Logon Pad 中的条目名）
        required: true
      - name: client
        description: 集团号
        required: true
      - name: user
        description: 用户名
        required: true
      - name: password
        description: 密码
        required: true
      - name: language
        description: 登录语言
        required: false
    steps:
      - action: open_connection
        params:
          system_name_from: system_name
        description: 打开 SAP 系统连接
      - action: fill_field
        params:
          field: "wnd[0]/usr/txtRSYST-MANDT"
          value_from: client
        description: 填写集团号
      - action: fill_field
        params:
          field: "wnd[0]/usr/txtRSYST-BNAME"
          value_from: user
        description: 填写用户名
      - action: fill_field
        params:
          field: "wnd[0]/usr/pwdRSYST-BCODE"
          value_from: password
        description: 填写密码
      - action: fill_field
        params:
          field: "wnd[0]/usr/txtRSYST-LANGU"
          value_from: language
        description: 填写登录语言
      - action: press_key
        params:
          key: ENTER
        description: 确认登录
```

- [ ] **步骤 5：新建 `sap_connections.yaml`（凭据模板）**

```yaml
# SAP 系统连接凭据（个人本地使用，不入版本控制）
connections:
  - name: DEV           # 对应 SAP Logon Pad 中的连接条目名
    client: "100"
    user: "YOUR_USER"
    password: "YOUR_PASSWORD"
    language: "ZH"

  - name: PRD
    client: "200"
    user: "YOUR_USER"
    password: "YOUR_PASSWORD"
    language: "ZH"

default: DEV            # 缺省连接
```

- [ ] **步骤 6：新建测试 fixture `tests/fixtures/sample.vbs`**

创建目录 `tests/fixtures/`，写入以下内容：

```vbscript
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session     = connection.Children(0)
End If
If Not IsObject(WScript) Then
   Set WScript = CreateObject("WScript.Shell")
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = "1000001"
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").setFocus
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 7
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
```

该 fixture 解析后应产生 5 个步骤：open_tcode、press_key、fill_field、press_key、click_button（maximize/setFocus/caretPosition 均为噪音被丢弃）。

---

## 任务二：新建 `sap_robot.py`（SAP 执行层）

**文件：**
- 新建：`sap_robot.py`

本模块封装 SAP GUI Scripting API，提供原子操作方法。依赖 win32com，无法在没有 SAP GUI 的环境下运行，因此本任务不写单元测试，在集成阶段人工验证。

- [ ] **步骤 1：新建 `sap_robot.py`**

```python
import win32com.client
import time
from typing import List, Dict, Any, Optional


class SAPRobot:
    """SAP GUI Scripting API 封装，提供原子操作方法。"""

    def __init__(self, system_name: str = None, timeout: int = 10):
        self.timeout = timeout
        self.session = None
        self._connect(system_name)

    def _connect(self, system_name: str = None) -> bool:
        """
        连接 SAP GUI。优先复用已有会话，没有则自动登录。
        - 已有会话 → 直接复用
        - 没有会话 → 读取 sap_connections.yaml → 调用 open_connection + 填写登录信息
        """
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine

            # 已有连接，直接复用
            if application.Children.Count > 0:
                connection = application.Children(0)
                if connection.Children.Count > 0:
                    self.session = connection.Children(0)
                    return True

            # 没有会话，走自动登录
            if system_name is None:
                system_name = self._get_default_system()
            return self._login(system_name)
        except Exception as e:
            print(f"连接 SAP GUI 失败：{e}")
            print("请确保 SAP GUI 已打开，且已启用 Scripting")
            return False

    def _get_default_system(self) -> str:
        """从 sap_connections.yaml 读取默认系统名。"""
        import yaml
        from pathlib import Path
        config_path = Path("sap_connections.yaml")
        if not config_path.exists():
            raise FileNotFoundError(
                "sap_connections.yaml 不存在，请创建凭据配置文件后重试"
            )
        config = yaml.safe_load(config_path.read_text(encoding="utf-8"))
        return config.get("default", "")

    def _login(self, system_name: str) -> bool:
        """读取凭据并登录指定 SAP 系统。"""
        import yaml
        from pathlib import Path
        config_path = Path("sap_connections.yaml")
        if not config_path.exists():
            print("sap_connections.yaml 不存在，无法自动登录")
            return False
        config = yaml.safe_load(config_path.read_text(encoding="utf-8"))
        conn = next(
            (c for c in config.get("connections", []) if c["name"] == system_name),
            None,
        )
        if not conn:
            names = [c["name"] for c in config.get("connections", [])]
            print(f"系统 '{system_name}' 未在配置中找到。可用系统：{names}")
            return False

        # 执行登录
        if not self.open_connection(system_name):
            return False
        self.set_text("wnd[0]/usr/txtRSYST-MANDT", conn.get("client", ""))
        self.set_text("wnd[0]/usr/txtRSYST-BNAME", conn.get("user", ""))
        self.set_text("wnd[0]/usr/pwdRSYST-BCODE", conn.get("password", ""))
        if conn.get("language"):
            self.set_text("wnd[0]/usr/txtRSYST-LANGU", conn["language"])
        self.press_key("ENTER")
        time.sleep(2)

        # 检查登录结果
        status = self.get_status()
        if status and ("错误" in status or "Error" in status or "invalid" in status.lower()):
            print(f"登录失败：{status}")
            return False
        return True

    def open_connection(self, system_name: str) -> bool:
        """通过 SAP Logon Pad 打开系统连接。"""
        try:
            application = win32com.client.GetObject("SAPGUI").GetScriptingEngine
            connection = application.OpenConnection(system_name, True)
            self.session = connection.Children(0)
            time.sleep(2)
            return True
        except Exception as e:
            print(f"打开连接 '{system_name}' 失败：{e}")
            return False

    def open_tcode(self, tcode: str) -> bool:
        """打开事务码。"""
        try:
            okcd = self.session.findById("wnd[0]/tbar[0]/okcd")
            okcd.Text = tcode
            self.session.findById("wnd[0]").SendVKey(0)
            time.sleep(1)
            return True
        except Exception as e:
            print(f"打开事务码 {tcode} 失败：{e}")
            return False

    def set_text(self, field_id: str, value: str) -> bool:
        """填写文本字段。"""
        try:
            field = self.session.findById(field_id)
            field.Text = str(value)
            return True
        except Exception as e:
            print(f"填写字段 {field_id} 失败：{e}")
            return False

    def press_key(self, key: str) -> bool:
        """按功能键。key 取值：ENTER / F3 / F8 / F12 / CTRL+S"""
        key_map = {"ENTER": 0, "F3": 3, "F8": 8, "F12": 12, "CTRL+S": 26}
        try:
            vkey = key_map.get(key.upper(), 0)
            self.session.findById("wnd[0]").SendVKey(vkey)
            time.sleep(0.5)
            return True
        except Exception as e:
            print(f"按键 {key} 失败：{e}")
            return False

    def click_button(self, button_id: str) -> bool:
        """点击按钮。"""
        try:
            self.session.findById(button_id).Press()
            return True
        except Exception as e:
            print(f"点击按钮 {button_id} 失败：{e}")
            return False

    def get_table_data(self, table_id: str, max_rows: int = 100) -> List[List[str]]:
        """
        读取 GridView 表格数据，返回行列二维数组。
        列名通过 grid.ColumnOrder 获取（SAP GridView.GetCellValue 要求列名字符串，非整数索引）。
        """
        try:
            grid = self.session.findById(table_id)
            rows = min(grid.RowCount, max_rows)
            col_names = list(grid.ColumnOrder)
            result = []
            for row in range(rows):
                row_data = []
                for col_name in col_names:
                    try:
                        row_data.append(grid.GetCellValue(row, col_name))
                    except Exception:
                        row_data.append("")
                result.append(row_data)
            return result
        except Exception as e:
            print(f"读取表格 {table_id} 失败：{e}")
            return []

    def get_status(self) -> str:
        """读取状态栏消息。"""
        try:
            return self.session.findById("wnd[0]/sbar").Text
        except Exception:
            return ""

    def wait_for_screen(self, screen_id: str = "wnd[0]", timeout: Optional[int] = None) -> bool:
        """等待屏幕加载完成。"""
        deadline = time.time() + (timeout or self.timeout)
        while time.time() < deadline:
            try:
                self.session.findById(screen_id)
                return True
            except Exception:
                time.sleep(0.5)
        return False

    def execute_sequence(
        self,
        steps: List[Dict[str, Any]],
        arguments: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """
        批量执行步骤序列。
        arguments 用于在运行时替换 value_from / tcode_from 参数。
        """
        if arguments is None:
            arguments = {}
        results: Dict[str, Any] = {"success": True, "messages": [], "outputs": {}}

        for step in steps:
            action = step.get("action")
            raw_params = dict(step.get("params", {}))

            # 将 value_from / tcode_from / system_name_from 替换为实际值
            params: Dict[str, Any] = {}
            for key, value in raw_params.items():
                if key == "value_from":
                    params["value"] = arguments.get(value, "")
                elif key == "tcode_from":
                    params["tcode"] = arguments.get(value, "")
                elif key == "system_name_from":
                    params["system_name"] = arguments.get(value, "")
                else:
                    params[key] = value

            success = False
            if action == "open_connection":
                success = self.open_connection(params.get("system_name", ""))
            elif action == "open_tcode":
                success = self.open_tcode(params.get("tcode", ""))
            elif action == "fill_field":
                success = self.set_text(params.get("field", ""), params.get("value", ""))
            elif action == "press_key":
                success = self.press_key(params.get("key", "ENTER"))
            elif action == "click_button":
                success = self.click_button(params.get("button_id", ""))
            elif action == "read_table":
                data = self.get_table_data(params.get("table_id", ""))
                results["outputs"][params.get("output_name", "table_data")] = data
                success = True
            elif action == "get_status":
                results["outputs"][params.get("output_name", "status")] = self.get_status()
                success = True
            elif action == "wait":
                time.sleep(params.get("seconds", 1))
                success = True
            else:
                results["messages"].append(f"未知 action：{action}")

            if not success and params.get("critical", False):
                results["success"] = False
                results["messages"].append(f"关键步骤失败：{action}")
                break

        return results
```

- [ ] **步骤 2：验证文件语法无误**

```bash
python -c "import ast; ast.parse(open('sap_robot.py').read()); print('语法正确')"
```

预期：`语法正确`

---

## 任务三：新建 `vbs_parser.py` + 单元测试

**文件：**
- 新建：`vbs_parser.py`
- 新建：`tests/test_vbs_parser.py`

- [ ] **步骤 1：先写失败测试 `tests/test_vbs_parser.py`**

```python
from pathlib import Path
import pytest
from vbs_parser import parse_vbs, parse_vbs_file

SAMPLE_VBS = Path("tests/fixtures/sample.vbs").read_text(encoding="utf-8-sig")


def test_识别打开事务码():
    steps = parse_vbs('session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA03"')
    assert len(steps) == 1
    assert steps[0]["action"] == "open_tcode"
    assert steps[0]["params"]["tcode"] == "VA03"


def test_识别填写字段():
    steps = parse_vbs('session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = "1000001"')
    assert steps[0]["action"] == "fill_field"
    assert steps[0]["params"]["field"] == "wnd[0]/usr/ctxtVBAK-VBELN"
    assert steps[0]["params"]["value"] == "1000001"


def test_识别回车():
    steps = parse_vbs('session.findById("wnd[0]").sendVKey 0')
    assert steps[0]["action"] == "press_key"
    assert steps[0]["params"]["key"] == "ENTER"


def test_识别F3():
    steps = parse_vbs('session.findById("wnd[0]").sendVKey 3')
    assert steps[0]["params"]["key"] == "F3"


def test_识别点击按钮():
    steps = parse_vbs('session.findById("wnd[0]/tbar[1]/btn[8]").press')
    assert steps[0]["action"] == "click_button"
    assert steps[0]["params"]["button_id"] == "wnd[0]/tbar[1]/btn[8]"


def test_噪音行被丢弃():
    noise = "\n".join([
        'session.findById("wnd[0]").maximize',
        'session.findById("wnd[0]/usr/ctxtVBAK-VBELN").setFocus',
        'session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 7',
    ])
    assert parse_vbs(noise) == []


def test_样板头被跳过():
    boilerplate = (
        "If Not IsObject(application) Then\n"
        '   Set SapGuiAuto = GetObject("SAPGUI")\n'
        "   Set application = SapGuiAuto.GetScriptingEngine\n"
        "End If\n"
    )
    assert parse_vbs(boilerplate) == []


def test_fixture_产生5个步骤():
    steps = parse_vbs(SAMPLE_VBS)
    assert len(steps) == 5
    assert [s["action"] for s in steps] == [
        "open_tcode", "press_key", "fill_field", "press_key", "click_button"
    ]


def test_seq编号连续():
    steps = parse_vbs(SAMPLE_VBS)
    assert [s["seq"] for s in steps] == [1, 2, 3, 4, 5]


def test_未识别findById行保留为unknown():
    line = 'session.findById("wnd[0]/usr/subSUB").unknownMethod'
    steps = parse_vbs(line)
    assert steps[0]["action"] == "unknown"
    assert steps[0]["raw_line"] == line


def test_parse_vbs_file读取文件(tmp_path):
    vbs = tmp_path / "t.vbs"
    vbs.write_text('session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM03"\n', encoding="utf-8")
    result = parse_vbs_file(str(vbs))
    assert result["steps"][0]["params"]["tcode"] == "MM03"
    assert "parsed_at" in result
```

- [ ] **步骤 2：确认测试失败**

```bash
pytest tests/test_vbs_parser.py -v
```

预期：`ImportError: No module named 'vbs_parser'`

- [ ] **步骤 3：新建 `vbs_parser.py`**

```python
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
    if "session.findById" in line.lower():
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
```

- [ ] **步骤 4：运行测试，确认全部通过**

```bash
pytest tests/test_vbs_parser.py -v
```

预期：11 项全部 PASS

---

## 任务四：新建 `doc_exporter.py` + 单元测试

**文件：**
- 新建：`doc_exporter.py`
- 新建：`tests/test_doc_exporter.py`

- [ ] **步骤 1：先写失败测试 `tests/test_doc_exporter.py`**

```python
import yaml
import pytest
from pathlib import Path
from doc_exporter import export_yaml, export_markdown

ENRICHED = {
    "capability_name": "查询销售订单",
    "capability_description": "根据订单号在 VA03 中查询销售订单详细信息",
    "tcode": "VA03",
    "created_at": "2026-05-28T10:30:00",
    "parameters": [{"name": "order_id", "description": "销售订单号", "required": True}],
    "steps": [
        {"seq": 1, "action": "open_tcode",
         "params": {"tcode": "VA03"}, "description": "打开事务码 VA03"},
        {"seq": 2, "action": "fill_field",
         "params": {"field": "wnd[0]/usr/ctxtVBAK-VBELN", "value_from": "order_id"},
         "description": "填写销售订单号"},
        {"seq": 3, "action": "press_key",
         "params": {"key": "ENTER"}, "description": "按回车确认"},
    ],
}


def test_导出yaml创建文件(tmp_path):
    ok, msg = export_yaml(ENRICHED, str(tmp_path / "caps.yaml"))
    assert ok is True and msg == "ok"


def test_导出yaml内容正确(tmp_path):
    path = str(tmp_path / "caps.yaml")
    export_yaml(ENRICHED, path)
    config = yaml.safe_load(open(path, encoding="utf-8"))
    cap = config["capabilities"][0]
    assert cap["name"] == "查询销售订单"
    assert cap["tcode"] == "VA03"
    assert len(cap["steps"]) == 3
    assert len(cap["parameters"]) == 1


def test_导出yaml同名返回duplicate(tmp_path):
    path = str(tmp_path / "caps.yaml")
    export_yaml(ENRICHED, path)
    ok, msg = export_yaml(ENRICHED, path)
    assert ok is False and msg == "duplicate"


def test_导出yaml保留已有条目(tmp_path):
    path = str(tmp_path / "caps.yaml")
    existing = {"capabilities": [{"name": "已有能力", "description": "x", "steps": []}]}
    import yaml as _yaml
    with open(path, "w", encoding="utf-8") as f:
        _yaml.dump(existing, f, allow_unicode=True)
    export_yaml(ENRICHED, path)
    config = yaml.safe_load(open(path, encoding="utf-8"))
    names = [c["name"] for c in config["capabilities"]]
    assert "已有能力" in names
    assert "查询销售订单" in names


def test_导出markdown创建文件(tmp_path):
    md_path = export_markdown(ENRICHED, str(tmp_path))
    assert Path(md_path).exists()


def test_导出markdown文件名正确(tmp_path):
    md_path = export_markdown(ENRICHED, str(tmp_path))
    assert Path(md_path).name == "查询销售订单.md"


def test_导出markdown包含关键内容(tmp_path):
    content = Path(export_markdown(ENRICHED, str(tmp_path))).read_text(encoding="utf-8")
    assert "查询销售订单" in content
    assert "VA03" in content
    assert "order_id" in content
    assert "wnd[0]/usr/ctxtVBAK-VBELN" in content
    assert "填写销售订单号" in content
```

- [ ] **步骤 2：确认测试失败**

```bash
pytest tests/test_doc_exporter.py -v
```

预期：`ImportError: No module named 'doc_exporter'`

- [ ] **步骤 3：新建 `doc_exporter.py`**

```python
import yaml
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Tuple


def export_yaml(
    enriched: Dict[str, Any],
    yaml_path: str = "sap_capabilities.yaml",
) -> Tuple[bool, str]:
    """
    将能力条目追加写入 YAML 能力库。
    返回 (True, "ok") 成功；(False, "duplicate") 表示同名能力已存在。
    """
    path = Path(yaml_path)
    config = yaml.safe_load(path.read_text(encoding="utf-8")) if path.exists() else {}
    config.setdefault("capabilities", [])

    name = enriched["capability_name"]
    if any(c.get("name") == name for c in config["capabilities"]):
        return False, "duplicate"

    config["capabilities"].append({
        "name": name,
        "description": enriched["capability_description"],
        "tcode": enriched.get("tcode", ""),
        "created_at": enriched.get("created_at", datetime.now().isoformat()),
        "parameters": enriched.get("parameters", []),
        "steps": [
            {"action": s["action"], "params": s["params"], "description": s.get("description", "")}
            for s in enriched["steps"]
        ],
    })
    with open(path, "w", encoding="utf-8") as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False)
    return True, "ok"


def export_markdown(enriched: Dict[str, Any], docs_dir: str = "docs/sop") -> str:
    """生成 SOP Markdown 文档，返回文件路径。"""
    Path(docs_dir).mkdir(parents=True, exist_ok=True)
    name = enriched["capability_name"]

    params = enriched.get("parameters", [])
    param_rows = "\n".join(
        f"| {p['name']} | {p.get('description', '')} | {'是' if p.get('required') else '否'} |"
        for p in params
    ) or "| — | — | — |"

    step_rows = []
    for s in enriched["steps"]:
        p = s.get("params", {})
        if s["action"] == "fill_field":
            elem = f"`{p.get('field', '')}`"
        elif s["action"] == "click_button":
            elem = f"`{p.get('button_id', '')}`"
        else:
            elem = "—"
        step_rows.append(f"| {s['seq']} | {s['action']} | {elem} | {s.get('description', '')} |")

    content = (
        f"# {name}\n\n"
        f"> {enriched['capability_description']}\n\n"
        f"- **事务码**：{enriched.get('tcode', '—')}\n"
        f"- **录制时间**：{enriched.get('created_at', '')[:10]}\n\n"
        f"## 参数\n\n"
        f"| 参数名 | 说明 | 必填 |\n|---|---|---|\n{param_rows}\n\n"
        f"## 操作步骤\n\n"
        f"| # | 操作 | 元素路径 | 说明 |\n|---|---|---|---|\n"
        + "\n".join(step_rows)
        + "\n\n## 原始录制文件\n\n- 来源：`.sap_record/session.vbs`\n"
    )
    path = Path(docs_dir) / f"{name}.md"
    path.write_text(content, encoding="utf-8")
    return str(path)
```

- [ ] **步骤 4：运行测试，确认全部通过**

```bash
pytest tests/test_doc_exporter.py -v
```

预期：7 项全部 PASS

---

## 任务五：新建 `step_enricher.py` + 单元测试

**文件：**
- 新建：`step_enricher.py`
- 新建：`tests/test_step_enricher.py`

- [ ] **步骤 1：先写失败测试 `tests/test_step_enricher.py`**

```python
import json
import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock
from step_enricher import suggest_param_name, run_interactive

STEPS_DATA = {
    "raw_file": "session.vbs",
    "parsed_at": "2026-05-28T10:00:00",
    "steps": [
        {"seq": 1, "action": "open_tcode", "params": {"tcode": "VA03"}, "raw_line": ""},
        {"seq": 2, "action": "fill_field",
         "params": {"field": "wnd[0]/usr/ctxtVBAK-VBELN", "value": "1000001"}, "raw_line": ""},
        {"seq": 3, "action": "press_key", "params": {"key": "ENTER"}, "raw_line": ""},
    ],
}


def _msg(text):
    m = MagicMock()
    m.content = [MagicMock(text=text)]
    return m


def test_推断参数名使用LLM结果():
    with patch("step_enricher._get_client") as mock:
        mock.return_value.messages.create.return_value = _msg("order_id")
        assert suggest_param_name("wnd[0]/usr/ctxtVBAK-VBELN") == "order_id"


def test_推断参数名LLM无效时回退():
    with patch("step_enricher._get_client") as mock:
        mock.return_value.messages.create.return_value = _msg("无法确定参数名称，请手动输入")
        # 回退：取路径最后段 "-" 后半部分小写
        assert suggest_param_name("wnd[0]/usr/ctxtVBAK-VBELN") == "vbeln"


def test_交互流程写入enriched_json(tmp_path):
    steps_path = str(tmp_path / "steps.json")
    out_path = str(tmp_path / "enriched.json")
    Path(steps_path).write_text(json.dumps(STEPS_DATA, ensure_ascii=False), encoding="utf-8")

    meta_resp = json.dumps({"name": "查询订单", "description": "查询销售订单信息"})
    inputs = ["y", "y", "y", "n", "y"]  # 确认meta、确认3步、第2步不参数化

    with patch("step_enricher._get_client") as mock, \
         patch("builtins.input", side_effect=inputs):
        mock.return_value.messages.create.side_effect = [
            _msg(meta_resp),
            _msg("打开销售订单事务码"),
            _msg("填写销售订单号"),
            _msg("按回车确认"),
        ]
        run_interactive(steps_path, out_path)

    result = json.loads(Path(out_path).read_text(encoding="utf-8"))
    assert result["capability_name"] == "查询订单"
    assert len(result["steps"]) == 3
    assert result["parameters"] == []
    assert result["steps"][1]["params"]["value"] == "1000001"


def test_参数化流程正确生成value_from(tmp_path):
    steps_path = str(tmp_path / "steps.json")
    out_path = str(tmp_path / "enriched.json")
    Path(steps_path).write_text(json.dumps(STEPS_DATA, ensure_ascii=False), encoding="utf-8")

    meta_resp = json.dumps({"name": "查询订单", "description": "查询销售订单信息"})
    inputs = ["y", "y", "y", "y", "", "销售订单号", "y"]
    # 确认meta、确认步骤1、确认步骤2描述、选择参数化、回车接受默认名、填写说明、确认步骤3

    with patch("step_enricher._get_client") as mock, \
         patch("builtins.input", side_effect=inputs):
        mock.return_value.messages.create.side_effect = [
            _msg(meta_resp),
            _msg("打开销售订单事务码"),
            _msg("填写销售订单号"),
            _msg("order_id"),   # suggest_param_name
            _msg("按回车确认"),
        ]
        run_interactive(steps_path, out_path)

    result = json.loads(Path(out_path).read_text(encoding="utf-8"))
    fill = result["steps"][1]
    assert "value_from" in fill["params"]
    assert fill["params"]["value_from"] == "order_id"
    assert "value" not in fill["params"]
    assert result["parameters"][0]["name"] == "order_id"
```

- [ ] **步骤 2：确认测试失败**

```bash
pytest tests/test_step_enricher.py -v
```

预期：`ImportError: No module named 'step_enricher'`

- [ ] **步骤 3：新建 `step_enricher.py`**

```python
import json
import anthropic
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List, Optional

_client: Optional[anthropic.Anthropic] = None


def _get_client() -> anthropic.Anthropic:
    global _client
    if _client is None:
        _client = anthropic.Anthropic()
    return _client


def suggest_capability_meta(steps: List[Dict]) -> Dict[str, str]:
    """批量请求：根据所有步骤建议能力名称和描述。"""
    summary = "\n".join(
        f"- {s['action']}: {json.dumps(s['params'], ensure_ascii=False)}" for s in steps
    )
    msg = _get_client().messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=200,
        messages=[{"role": "user", "content": (
            f"你是 SAP 业务专家。根据以下操作步骤，建议一个能力名称（5-10字）和描述（20-40字）。"
            f"以 JSON 格式输出：{{\"name\": \"...\", \"description\": \"...\"}}\n\n步骤：\n{summary}"
        )}],
    )
    try:
        return json.loads(msg.content[0].text)
    except Exception:
        return {"name": "未命名能力", "description": ""}


def suggest_step_description(step: Dict, tcode: str) -> str:
    """单步请求：用一句中文描述操作，≤20字。"""
    msg = _get_client().messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=50,
        messages=[{"role": "user", "content": (
            f"用一句中文描述这个 SAP 操作，不超过20字：\n"
            f"操作类型: {step['action']}\n"
            f"参数: {json.dumps(step['params'], ensure_ascii=False)}\n"
            f"当前事务码: {tcode}\n只输出描述文字。"
        )}],
    )
    return msg.content[0].text.strip()


def suggest_param_name(field_path: str) -> str:
    """根据 SAP 字段路径推断 snake_case 参数名。LLM 失败时回退到路径末段。"""
    msg = _get_client().messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=20,
        messages=[{"role": "user", "content": (
            f"SAP 字段路径 '{field_path}' 对应的英文参数名是什么？"
            f"只输出参数名（snake_case，无其他内容）。"
        )}],
    )
    name = msg.content[0].text.strip().lower().replace("-", "_").replace(" ", "_")
    if not name or len(name) > 30 or " " in name:
        last = field_path.split("/")[-1]
        name = last.split("-")[-1].lower() if "-" in last else last.lower()
    return name


def run_interactive(
    steps_path: str = ".sap_record/steps.json",
    output_path: str = ".sap_record/enriched.json",
):
    """交互式 AI 增强会话，引导用户逐步确认每个步骤的业务描述和参数化。"""
    with open(steps_path, encoding="utf-8") as f:
        data = json.load(f)

    steps = data["steps"]
    tcode = next(
        (s["params"].get("tcode", "") for s in steps if s["action"] == "open_tcode"), ""
    )

    print(f"\n{'=' * 50}\nSAP 能力录制 - AI 增强阶段\n共 {len(steps)} 个步骤\n{'=' * 50}")

    # 确认能力元信息
    meta = suggest_capability_meta(steps)
    print(f"\nAI 建议能力名称：{meta['name']}\nAI 建议描述：{meta['description']}")
    if input("[y] 确认  [m] 修改  > ").strip().lower() == "m":
        meta["name"] = input("能力名称：").strip() or meta["name"]
        meta["description"] = input("能力描述：").strip() or meta["description"]

    parameters: List[Dict] = []
    enriched_steps: List[Dict] = []

    for step in steps:
        print(f"\n{'─' * 40}")
        print(f"步骤 {step['seq']}/{len(steps)}  {step['action']}  →  "
              f"{json.dumps(step['params'], ensure_ascii=False)}")

        if step["action"] == "unknown":
            print("⚠ 未识别步骤，跳过")
            enriched_steps.append({**step, "description": "（未识别）"})
            continue

        desc = suggest_step_description(step, tcode)
        print(f"AI 描述：{desc}")
        choice = input("[y] 确认  [m] 修改  [s] 跳过  > ").strip().lower()
        if choice == "s":
            continue
        if choice == "m":
            desc = input("描述：").strip() or desc

        new_params = dict(step["params"])
        if step["action"] == "fill_field" and step["params"].get("value"):
            if input(f'\n值 "{step["params"]["value"]}" 是否设为运行时参数？[y] 是  [n] 否  > ').strip().lower() == "y":
                suggested = suggest_param_name(step["params"].get("field", ""))
                param_name = input(f"参数名（回车默认 {suggested}）> ").strip() or suggested
                new_params = {"field": step["params"]["field"], "value_from": param_name}
                if not any(p["name"] == param_name for p in parameters):
                    param_desc = input(f"参数 '{param_name}' 的说明：").strip() or param_name
                    parameters.append({"name": param_name, "description": param_desc, "required": True})

        enriched_steps.append({"seq": step["seq"], "action": step["action"],
                                "params": new_params, "description": desc})

    result = {
        "capability_name": meta["name"],
        "capability_description": meta["description"],
        "tcode": tcode,
        "created_at": datetime.now().isoformat(),
        "parameters": parameters,
        "steps": enriched_steps,
    }
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n{'=' * 50}\n增强完成，已写入 {output_path}")
```

- [ ] **步骤 4：运行测试，确认全部通过**

```bash
pytest tests/test_step_enricher.py -v
```

预期：4 项全部 PASS

---

## 任务六：新建 `sap_recorder.py`（CLI 入口）

**文件：**
- 新建：`sap_recorder.py`

- [ ] **步骤 1：新建 `sap_recorder.py`**

```python
import argparse
import json
import sys
from pathlib import Path

SAP_RECORD_DIR = ".sap_record"
SESSION_VBS    = f"{SAP_RECORD_DIR}/session.vbs"
STEPS_JSON     = f"{SAP_RECORD_DIR}/steps.json"
ENRICHED_JSON  = f"{SAP_RECORD_DIR}/enriched.json"


def _require(path: str, hint: str):
    if not Path(path).exists():
        print(f"错误：{path} 不存在。{hint}")
        sys.exit(1)


def cmd_start(_):
    try:
        import win32com.client
        utils = win32com.client.GetObject("SAPGUI").Utils
        Path(SAP_RECORD_DIR).mkdir(exist_ok=True)
        utils.Record(str(Path(SESSION_VBS).resolve()))
        print("录制已开启，请在 SAP 中完成操作...")
        print(f"完成后运行：python sap_recorder.py stop")
    except Exception as e:
        print(f"启动录制失败：{e}")
        print("请确保 SAP GUI 已打开并启用了 Scripting（sapgui/user_scripting=TRUE）")
        sys.exit(1)


def cmd_stop(_):
    try:
        win32com.client.GetObject("SAPGUI").Utils.Stop()
        print(f"录制停止，文件已保存至 {SESSION_VBS}")
    except Exception as e:
        print(f"停止录制失败：{e}")
        sys.exit(1)


def cmd_parse(args):
    from vbs_parser import parse_vbs_file, save_steps
    vbs = args.file or SESSION_VBS
    _require(vbs, "请提供 .vbs 文件路径，或先运行 start/stop")
    data = parse_vbs_file(vbs)
    save_steps(data, STEPS_JSON)
    print(f"解析完成，共 {len(data['steps'])} 个步骤，已写入 {STEPS_JSON}")


def cmd_enrich(_):
    from step_enricher import run_interactive
    _require(STEPS_JSON, "请先运行 parse")
    run_interactive(STEPS_JSON, ENRICHED_JSON)


def cmd_export(_):
    from doc_exporter import export_yaml, export_markdown
    _require(ENRICHED_JSON, "请先运行 enrich")
    with open(ENRICHED_JSON, encoding="utf-8") as f:
        enriched = json.load(f)

    name = enriched["capability_name"]
    ok, msg = export_yaml(enriched)
    if not ok and msg == "duplicate":
        choice = input(f"\n能力 '{name}' 已存在。[o] 覆盖  [r] 重命名  [c] 取消  > ").strip().lower()
        if choice == "o":
            from config_manager import ConfigManager
            ConfigManager().remove_capability(name)
            export_yaml(enriched)
        elif choice == "r":
            enriched["capability_name"] = input("新名称：").strip()
            export_yaml(enriched)
        else:
            print("已取消")
            return

    md = export_markdown(enriched)
    print(f"\n导出完成：\n  YAML → sap_capabilities.yaml\n  文档 → {md}")


def cmd_pipeline(args):
    cmd_start(args)
    input("\n操作完成后按 Enter 键停止录制 > ")
    cmd_stop(args)
    class _A: file = None
    cmd_parse(_A())
    cmd_enrich(args)
    cmd_export(args)


def main():
    p = argparse.ArgumentParser(description="SAP GUI 操作录制工具")
    sub = p.add_subparsers(dest="cmd")
    sub.add_parser("start")
    sub.add_parser("stop")
    pp = sub.add_parser("parse")
    pp.add_argument("file", nargs="?", default=None)
    sub.add_parser("enrich")
    sub.add_parser("export")
    sub.add_parser("pipeline")
    args = p.parse_args()
    {
        "start": cmd_start, "stop": cmd_stop, "parse": cmd_parse,
        "enrich": cmd_enrich, "export": cmd_export, "pipeline": cmd_pipeline,
    }.get(args.cmd, lambda _: p.print_help())(args)


if __name__ == "__main__":
    main()
```

- [ ] **步骤 2：用 fixture 冒烟测试 parse 命令**

```bash
python sap_recorder.py parse tests/fixtures/sample.vbs
```

预期输出：
```
解析完成，共 5 个步骤，已写入 .sap_record/steps.json
```

并验证 `.sap_record/steps.json` 中 5 个步骤的 action 分别为：
`open_tcode, press_key, fill_field, press_key, click_button`

---

## 任务七：新建 `sap_mcp_server.py`

**文件：**
- 新建：`sap_mcp_server.py`

本模块在 stdio 模式下与 Claude Desktop 集成，simple 模式下用于命令行测试。

- [ ] **步骤 1：新建 `sap_mcp_server.py`**

```python
import asyncio
import json
import yaml
import sys
from pathlib import Path
from typing import Any, Dict, List

try:
    from mcp.server import Server
    from mcp.server.models import InitializationOptions
    import mcp.server.stdio
    import mcp.types as types
    MCP_AVAILABLE = True
except ImportError:
    MCP_AVAILABLE = False
    print("警告：MCP SDK 未安装，仅 simple 模式可用。运行：pip install mcp", file=sys.stderr)

from sap_robot import SAPRobot


class SAPMCPServer:
    def __init__(self, config_path: str = "sap_capabilities.yaml"):
        self.config_path = Path(config_path)
        self.capabilities = self._load()
        self._robot: SAPRobot = None

    def _load(self) -> Dict:
        if not self.config_path.exists():
            return {"capabilities": []}
        return yaml.safe_load(self.config_path.read_text(encoding="utf-8")) or {"capabilities": []}

    def _robot_instance(self) -> SAPRobot:
        if self._robot is None:
            self._robot = SAPRobot()
        return self._robot

    def _execute(self, cap_name: str, arguments: Dict[str, Any]) -> str:
        cap = next((c for c in self.capabilities.get("capabilities", [])
                    if c.get("name") == cap_name), None)
        if not cap:
            return f"错误：未找到能力 '{cap_name}'"
        result = self._robot_instance().execute_sequence(cap.get("steps", []), arguments)
        if result["success"]:
            out = result["outputs"]
            return f"执行成功\n{json.dumps(out, ensure_ascii=False, indent=2)}" if out else "执行成功"
        return f"执行失败：{'; '.join(result['messages'])}"

    def _tools(self) -> List[Dict]:
        tools = []
        for cap in self.capabilities.get("capabilities", []):
            props = {p["name"]: {"type": "string", "description": p.get("description", "")}
                     for p in cap.get("parameters", [])}
            required = [p["name"] for p in cap.get("parameters", []) if p.get("required")]
            tools.append({
                "name": cap["name"],
                "description": cap["description"],
                "inputSchema": {"type": "object", "properties": props, "required": required},
            })
        return tools

    async def run_stdio(self):
        if not MCP_AVAILABLE:
            print("MCP SDK 不可用，请运行：pip install mcp")
            return
        server = Server("sap-mcp-server")

        @server.list_tools()
        async def list_tools() -> List[types.Tool]:
            return [
                types.Tool(name=t["name"], description=t["description"],
                           inputSchema=t["inputSchema"])
                for t in self._tools()
            ]

        @server.call_tool()
        async def call_tool(name: str, arguments: dict) -> List[types.TextContent]:
            return [types.TextContent(type="text", text=self._execute(name, arguments))]

        async with mcp.server.stdio.stdio_server() as (r, w):
            await server.run(r, w, InitializationOptions(
                server_name="sap-mcp-server", server_version="1.0.0"
            ))

    def run_simple(self):
        print("=" * 50)
        print("SAP MCP Server - 命令行测试模式")
        tools = self._tools()
        for i, t in enumerate(tools, 1):
            print(f"  {i}. {t['name']}：{t['description']}")
        print("=" * 50)
        print("用法：能力名称 参数1=值1 参数2=值2")
        print("输入 list 查看能力列表，exit 退出\n")
        while True:
            try:
                cmd = input("> ").strip()
                if cmd.lower() == "exit":
                    break
                if cmd.lower() == "list":
                    for t in tools:
                        print(f"\n{t['name']}\n  描述：{t['description']}")
                    continue
                parts = cmd.split(maxsplit=1)
                cap_name = parts[0]
                args = {}
                if len(parts) > 1:
                    for pair in parts[1].split():
                        if "=" in pair:
                            k, v = pair.split("=", 1)
                            args[k] = v
                print(self._execute(cap_name, args))
            except KeyboardInterrupt:
                break
            except Exception as e:
                print(f"错误：{e}")


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--mode", choices=["stdio", "simple"], default="simple")
    p.add_argument("--config", default="sap_capabilities.yaml")
    args = p.parse_args()
    server = SAPMCPServer(args.config)
    if args.mode == "stdio" and MCP_AVAILABLE:
        asyncio.run(server.run_stdio())
    else:
        server.run_simple()
```

- [ ] **步骤 2：验证语法无误**

```bash
python -c "import ast; ast.parse(open('sap_mcp_server.py').read()); print('语法正确')"
```

预期：`语法正确`

---

## 任务八：全量测试 + 冒烟验证

- [ ] **步骤 1：运行全部单元测试**

```bash
pytest tests/ -v --tb=short
```

预期：全部 PASS，无失败

- [ ] **步骤 2：完整 parse → export 管道冒烟**

```bash
# 解析 fixture
python sap_recorder.py parse tests/fixtures/sample.vbs

# 查看 steps.json 内容确认正确
python -c "
import json
data = json.load(open('.sap_record/steps.json', encoding='utf-8'))
for s in data['steps']:
    print(s['seq'], s['action'], s['params'])
"
```

预期输出：
```
1 open_tcode {'tcode': 'VA03'}
2 press_key {'key': 'ENTER'}
3 fill_field {'field': 'wnd[0]/usr/ctxtVBAK-VBELN', 'value': '1000001'}
4 press_key {'key': 'ENTER'}
5 click_button {'button_id': 'wnd[0]/tbar[1]/btn[8]'}
```

- [ ] **步骤 3：验证 simple 模式 MCP Server 启动**

```bash
python sap_mcp_server.py --mode=simple
```

预期：启动成功，显示能力列表（初始为空），输入 `exit` 退出

- [ ] **步骤 4：最终提交**

```bash
git add .
git commit -m "feat: 完成 SAPGUI Auto 初版实现（执行层、录制管道、MCP Server）"
```

---

---

## 任务九：新建 `llm_agent.py` + 单元测试

**文件：**
- 新建：`llm_agent.py`
- 新建：`tests/test_llm_agent.py`

- [ ] **步骤 1：先写失败测试 `tests/test_llm_agent.py`**

```python
import json
import yaml
import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock
from llm_agent import load_tools, chat

YAML_CONTENT = """
capabilities:
  - name: 查询销售订单
    description: 根据订单号在 VA03 中查询销售订单详细信息
    tcode: VA03
    parameters:
      - name: order_id
        description: 销售订单号
        required: true
    steps:
      - action: open_tcode
        params:
          tcode: VA03
      - action: fill_field
        params:
          field: wnd[0]/usr/ctxtVBAK-VBELN
          value_from: order_id
      - action: press_key
        params:
          key: ENTER
      - action: get_status
        params:
          output_name: result
"""


def test_load_tools_格式正确(tmp_path):
    yaml_file = tmp_path / "caps.yaml"
    yaml_file.write_text(YAML_CONTENT, encoding="utf-8")
    tools = load_tools(str(yaml_file))
    assert len(tools) == 1
    assert tools[0]["name"] == "查询销售订单"
    assert "input_schema" in tools[0]
    assert tools[0]["input_schema"]["properties"]["order_id"]["type"] == "string"
    assert tools[0]["input_schema"]["required"] == ["order_id"]


def test_load_tools_空文件返回空列表(tmp_path):
    yaml_file = tmp_path / "caps.yaml"
    yaml_file.write_text("capabilities: []", encoding="utf-8")
    assert load_tools(str(yaml_file)) == []


def test_chat_直接文本回复(tmp_path):
    yaml_file = tmp_path / "caps.yaml"
    yaml_file.write_text("capabilities: []", encoding="utf-8")

    text_block = MagicMock()
    text_block.type = "text"
    text_block.text = "你好，我是SAP助手"
    mock_resp = MagicMock()
    mock_resp.stop_reason = "end_turn"
    mock_resp.content = [text_block]

    with patch("llm_agent._get_client") as mock_client, \
         patch("llm_agent.SAPRobot"):
        mock_client.return_value.messages.create.return_value = mock_resp
        reply, history = chat("你好", [], str(yaml_file))

    assert reply == "你好，我是SAP助手"
    assert len(history) == 2  # user message + assistant message


def test_chat_工具调用后返回结果(tmp_path):
    yaml_file = tmp_path / "caps.yaml"
    yaml_file.write_text(YAML_CONTENT, encoding="utf-8")

    # 第一次响应：tool_use
    tool_block = MagicMock()
    tool_block.type = "tool_use"
    tool_block.name = "查询销售订单"
    tool_block.id = "tool_123"
    tool_block.input = {"order_id": "1000001"}
    resp1 = MagicMock()
    resp1.stop_reason = "tool_use"
    resp1.content = [tool_block]

    # 第二次响应：文本
    text_block = MagicMock()
    text_block.type = "text"
    text_block.text = "订单 1000001 已发货"
    resp2 = MagicMock()
    resp2.stop_reason = "end_turn"
    resp2.content = [text_block]

    mock_robot = MagicMock()
    mock_robot.execute_sequence.return_value = {
        "success": True,
        "outputs": {"result": "订单已发货"},
        "messages": [],
    }

    with patch("llm_agent._get_client") as mock_client, \
         patch("llm_agent.SAPRobot", return_value=mock_robot):
        mock_client.return_value.messages.create.side_effect = [resp1, resp2]
        reply, history = chat("查询订单1000001", [], str(yaml_file))

    assert reply == "订单 1000001 已发货"
    mock_robot.execute_sequence.assert_called_once()
    # history: user + assistant(tool_use) + user(tool_result) + assistant(text)
    assert len(history) == 4


def test_chat_保留传入历史(tmp_path):
    yaml_file = tmp_path / "caps.yaml"
    yaml_file.write_text("capabilities: []", encoding="utf-8")

    prior_history = [
        {"role": "user", "content": "你好"},
        {"role": "assistant", "content": [MagicMock(type="text", text="你好！")]},
    ]
    text_block = MagicMock()
    text_block.type = "text"
    text_block.text = "本次回复"
    mock_resp = MagicMock()
    mock_resp.stop_reason = "end_turn"
    mock_resp.content = [text_block]

    with patch("llm_agent._get_client") as mock_client, \
         patch("llm_agent.SAPRobot"):
        mock_client.return_value.messages.create.return_value = mock_resp
        reply, history = chat("继续", prior_history, str(yaml_file))

    # 调用时 messages 应包含完整历史 + 新消息
    call_args = mock_client.return_value.messages.create.call_args
    msgs = call_args[1]["messages"]
    assert msgs[0] == prior_history[0]
    assert msgs[1] == prior_history[1]
    assert msgs[2] == {"role": "user", "content": "继续"}
```

- [ ] **步骤 2：确认测试失败**

```bash
pytest tests/test_llm_agent.py -v
```

预期：`ImportError: No module named 'llm_agent'`

- [ ] **步骤 3：新建 `llm_agent.py`**

```python
import json
import yaml
import anthropic
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

_client: Optional[anthropic.Anthropic] = None


def _get_client() -> anthropic.Anthropic:
    global _client
    if _client is None:
        _client = anthropic.Anthropic()
    return _client


def load_tools(config_path: str = "sap_capabilities.yaml") -> List[Dict[str, Any]]:
    """将 sap_capabilities.yaml 中的能力转换为 Anthropic tools 格式。"""
    config = yaml.safe_load(Path(config_path).read_text(encoding="utf-8")) or {}
    tools = []
    for cap in config.get("capabilities", []):
        props = {
            p["name"]: {"type": "string", "description": p.get("description", "")}
            for p in cap.get("parameters", [])
        }
        required = [p["name"] for p in cap.get("parameters", []) if p.get("required")]
        tools.append({
            "name": cap["name"],
            "description": cap["description"],
            "input_schema": {
                "type": "object",
                "properties": props,
                "required": required,
            },
        })
    return tools


def chat(
    user_message: str,
    history: List[Dict[str, Any]],
    config_path: str = "sap_capabilities.yaml",
) -> Tuple[str, List[Dict[str, Any]]]:
    """
    单轮对话。接受用户消息和历史，返回 (回复文本, 更新后的完整历史)。
    内部驱动 Function Calling 循环直到 stop_reason 为 end_turn。
    """
    from sap_robot import SAPRobot
    tools = load_tools(config_path)
    robot = SAPRobot()
    messages = list(history) + [{"role": "user", "content": user_message}]

    while True:
        resp = _get_client().messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            tools=tools,
            messages=messages,
        )
        messages.append({"role": "assistant", "content": resp.content})

        if resp.stop_reason == "tool_use":
            config = yaml.safe_load(Path(config_path).read_text(encoding="utf-8")) or {}
            tool_results = []
            for block in resp.content:
                if block.type != "tool_use":
                    continue
                cap = next(
                    (c for c in config.get("capabilities", []) if c["name"] == block.name),
                    None,
                )
                if cap:
                    result = robot.execute_sequence(cap["steps"], dict(block.input))
                    if result["success"]:
                        content = (
                            json.dumps(result["outputs"], ensure_ascii=False)
                            if result["outputs"]
                            else "执行成功"
                        )
                    else:
                        content = f"执行失败：{'; '.join(result['messages'])}"
                else:
                    content = f"未找到能力 '{block.name}'"
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": content,
                })
            messages.append({"role": "user", "content": tool_results})
        else:
            reply = next(
                (b.text for b in resp.content if hasattr(b, "text")),
                "（无回复）",
            )
            return reply, messages
```

- [ ] **步骤 4：运行测试，确认全部通过**

```bash
pytest tests/test_llm_agent.py -v
```

预期：5 项全部 PASS

---

## 任务十：新建 `web_server.py`（Web Chat）

**文件：**
- 新建：`web_server.py`

本模块不写自动化单元测试（需要网络 + SAP），在步骤 2 用 curl 手动冒烟。

- [ ] **步骤 1：新建 `web_server.py`**

```python
import uvicorn
from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from typing import List, Dict, Any

from llm_agent import chat as agent_chat

app = FastAPI(title="SAP Web Chat")

_HTML = """<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="UTF-8">
  <title>SAP 助手</title>
  <style>
    body { font-family: sans-serif; max-width: 800px; margin: 40px auto; padding: 0 20px; }
    #messages { height: 420px; overflow-y: auto; border: 1px solid #ddd; padding: 12px;
                margin-bottom: 10px; border-radius: 4px; background: #fafafa; }
    .user { text-align: right; color: #0066cc; margin: 6px 0; }
    .assistant { text-align: left; color: #333; margin: 6px 0; white-space: pre-wrap; }
    #input-row { display: flex; gap: 8px; }
    #msg { flex: 1; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
    button { padding: 8px 18px; cursor: pointer; border-radius: 4px;
             background: #0066cc; color: white; border: none; }
    button:hover { background: #0052a3; }
  </style>
</head>
<body>
  <h2>SAP 助手</h2>
  <div id="messages"></div>
  <div id="input-row">
    <input id="msg" placeholder="输入指令，例如：查询订单 1000001"
           onkeydown="if(event.key==='Enter')send()">
    <button onclick="send()">发送</button>
  </div>
  <script>
    let history = [];
    async function send() {
      const input = document.getElementById('msg');
      const text = input.value.trim();
      if (!text) return;
      input.value = '';
      append('user', text);
      const res = await fetch('/chat', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({message: text, history})
      });
      const data = await res.json();
      history = data.history;
      append('assistant', data.reply);
    }
    function append(role, text) {
      const div = document.createElement('div');
      div.className = role;
      div.textContent = (role === 'user' ? '你：' : 'SAP 助手：') + text;
      document.getElementById('messages').appendChild(div);
      document.getElementById('messages').scrollTop = 9999;
    }
  </script>
</body>
</html>"""


class ChatRequest(BaseModel):
    message: str
    history: List[Dict[str, Any]] = []


class ChatResponse(BaseModel):
    reply: str
    history: List[Dict[str, Any]]


@app.get("/", response_class=HTMLResponse)
async def index():
    return _HTML


@app.post("/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    reply, updated_history = agent_chat(req.message, req.history)
    return ChatResponse(reply=reply, history=updated_history)


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
```

- [ ] **步骤 2：验证语法无误**

```bash
python -c "import ast; ast.parse(open('web_server.py').read()); print('语法正确')"
```

预期：`语法正确`

- [ ] **步骤 3：启动服务（不连接 SAP，仅测试路由）**

新开一个终端，启动服务：

```bash
python web_server.py
```

预期：`Uvicorn running on http://0.0.0.0:8000`

在另一个终端验证 HTML 页面可访问：

```bash
curl -s http://localhost:8000/ | head -5
```

预期输出包含 `<!DOCTYPE html>`

验证 `/chat` 端点结构（不实际调用 SAP，用空 capabilities）：

先临时将 `sap_capabilities.yaml` 内容改为 `capabilities: []`，然后：

```bash
curl -s -X POST http://localhost:8000/chat \
  -H "Content-Type: application/json" \
  -d "{\"message\": \"你好\", \"history\": []}" | python -m json.tool
```

预期：返回包含 `reply` 和 `history` 字段的 JSON

用 `Ctrl+C` 停止服务。

- [ ] **步骤 4：运行全量测试**

```bash
pytest tests/ -v --tb=short
```

预期：全部 PASS（含新增的 test_llm_agent.py）

- [ ] **步骤 5：最终提交**

```bash
git add llm_agent.py web_server.py tests/test_llm_agent.py requirements.txt
git commit -m "feat: 新增 Web Chat 功能（llm_agent.py + web_server.py）"
```

---

## Claude Desktop 集成（手动步骤，非代码）

完成以上任务后，按以下步骤接入 Claude Desktop：

1. 用录制管道录制至少一个真实能力（需要 SAP GUI 环境）：
   ```bash
   python sap_recorder.py pipeline
   ```

2. 编辑 Claude Desktop 配置文件 `%APPDATA%\Claude\claude_desktop_config.json`：
   ```json
   {
     "mcpServers": {
       "my-sap": {
         "command": "python",
         "args": ["C:/完整路径/sap_mcp_server.py", "--mode=stdio"]
       }
     }
   }
   ```

3. 重启 Claude Desktop

4. 在对话中测试：*"帮我查一下订单 1000001"*

## Web Chat 使用（手动步骤）

1. 确保 SAP GUI 已打开并登录，且 `sap_capabilities.yaml` 中有已录制的能力

2. 启动 Web Chat 服务：
   ```bash
   python web_server.py
   ```

3. 浏览器打开 `http://localhost:8000`

4. 在输入框中输入自然语言指令，例如：*查询订单 1000001*
