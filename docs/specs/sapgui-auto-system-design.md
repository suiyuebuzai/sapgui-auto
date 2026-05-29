# SAPGUI Auto 整体系统设计

- **日期**：2026-05-28
- **状态**：设计确认

---

## 一、问题与目标

### 核心问题

SAP GUI 承载了大量企业业务操作，但存在三个痛点：

1. 操作依赖人工手动点击，无法自动化执行
2. 业务流程只存在操作人脑中，无法沉淀复用
3. 非技术人员无法描述操作步骤，技术人员看不懂业务含义

### 目标

构建一套工具链，实现两件事：

- **能力沉淀**：用户操作一遍 SAP，系统自动生成可执行配置 + 可阅读文档
- **自然语言驱动**：业务人员说"查询订单 1000001"，系统自动完成 SAP 操作并返回结果

### 成功标准

- 录制一个业务能力的全流程不超过 5 分钟
- 录制产出的 YAML 配置可直接被执行引擎运行
- 录制产出的 Markdown 文档业务人员能直接读懂
- Claude Desktop 可通过自然语言调用任意已录制能力

---

## 二、整体架构

### 三层模型

```
┌─────────────────────────────────────────────────────────────┐
│  第一层：自然语言接口层                                         │
│                                                              │
│  方式 A（Claude Desktop）                                     │
│  用户 → Claude Desktop → MCP Server                          │
│                                                              │
│  方式 B（Web Chat）                                           │
│  用户 → 浏览器 → FastAPI → llm_agent.py (Function Calling)   │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│  第二层：能力库层                                              │
│                                                              │
│   sap_capabilities.yaml                                      │
│   ├── 查询销售订单  (VA03)                                    │
│   ├── 创建物料凭证  (MIGO)                                    │
│   └── ...（持续积累）                                         │
│                                                              │
│   录制管道（录制 → 解析 → AI 增强 → 导出）                     │
└─────────────────────────────────────────────────────────────┘
                           ↓
┌─────────────────────────────────────────────────────────────┐
│  第三层：SAP 执行层                                            │
│                                                              │
│   sap_robot.py → win32com → SAP GUI Scripting API → SAP     │
└─────────────────────────────────────────────────────────────┘
```

### 两条主线

**主线 A：能力积累（离线，一次性录制）**

```
用户在 SAP GUI 中操作一遍业务流程
        ↓ SAP 内置录制器（通过 COM 控制）
   .sap_record/session.vbs（VBScript 原始录制输出）
        ↓ vbs_parser.py（解析）
   .sap_record/steps.json（结构化步骤）
        ↓ step_enricher.py + LLM（AI 增强 + 用户交互确认）
   .sap_record/enriched.json（含业务描述和参数定义）
        ↓ doc_exporter.py（导出）
   ┌───────────────────────────┐
   ↓                           ↓
sap_capabilities.yaml       docs/sop/{能力名}.md
（供执行层调用）              （供业务人员阅读）
```

**主线 B1：Claude Desktop（MCP Server，在线）**

```
用户自然语言输入（Claude Desktop）
        ↓ MCP 协议
   sap_mcp_server.py（解析 YAML，注册为 MCP 工具）
        ↓ 工具调用（含运行时参数）
   sap_robot.py（执行原子操作序列）
        ↓ win32com → SAP GUI Scripting API
   SAP GUI 完成操作
        ↓
   结果返回用户
```

**主线 B2：Web Chat（Function Calling，在线）**

```
用户在浏览器中输入（http://localhost:8000）
        ↓ HTTP POST /chat
   web_server.py → llm_agent.py
        ↓ Anthropic API（Function Calling 工具循环）
   Claude 决策 → 调用能力（sap_capabilities.yaml 中的条目）
        ↓ sap_robot.execute_sequence()
   SAP GUI 完成操作
        ↓
   结果文本返回 → 浏览器显示
```

---

## 三、组件详细设计

### 3.1 SAP 执行层：`sap_robot.py`

SAP GUI Scripting API 的 Python 封装，对外提供原子操作方法。

**连接方式（COM）：**

```python
sap_gui_auto = win32com.client.GetObject("SAPGUI")
application  = sap_gui_auto.GetScriptingEngine
connection   = application.Children(0)   # 第一个 SAP 连接
session      = connection.Children(0)    # 第一个会话窗口
```

**连接策略（优先复用已有会话）：**

```python
def _connect(self, system_name=None):
    """
    1. 尝试获取已有会话 → 成功则直接复用，无需登录
    2. 没有已有会话 → 调用"登录SAP"能力自动登录
    """
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine

    # 已有连接，直接复用
    if application.Children.Count > 0:
        connection = application.Children(0)
        if connection.Children.Count > 0:
            self.session = connection.Children(0)
            return True

    # 没有会话，走自动登录（需提供 system_name 或使用默认系统）
    if system_name is None:
        system_name = self._get_default_system()
    return self._login(system_name)
```

**原子方法：**

| 方法 | 说明 |
|---|---|
| `open_connection(system_name)` | 通过 SAP Logon Pad 打开系统连接 |
| `open_tcode(tcode)` | 打开事务码 |
| `set_text(field_id, value)` | 填写文本字段 |
| `press_key(key)` | 按功能键（ENTER / F3 / F8 / F12 / CTRL+S） |
| `click_button(button_id)` | 点击按钮 |
| `get_table_data(table_id)` | 读取表格，返回行列二维数组 |
| `get_status()` | 读取状态栏消息 |
| `wait_for_screen(screen_id)` | 等待屏幕加载完成 |
| `execute_sequence(steps, arguments)` | 批量执行步骤序列，支持运行时参数替换 |

**`open_connection` 方法说明：**

通过 SAP GUI Scripting API 的 `OpenConnection` 方法打开系统连接：

```python
def open_connection(self, system_name: str) -> bool:
    """通过 SAP Logon Pad 打开系统连接。"""
    application = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    connection = application.OpenConnection(system_name, True)  # True=同步
    self.session = connection.Children(0)
    return True
```

`system_name` 对应 SAP Logon Pad 中配置的连接条目名称。

**`execute_sequence` 参数替换规则：**

YAML 中 `value_from: order_id` 表示该字段的值在运行时从 `arguments["order_id"]` 取得。
执行时将 `value_from` 键替换为 `value` 键，值从 `arguments` 中查找。

```
params: {field: "wnd[0]/...", value_from: "order_id"}
arguments: {order_id: "1000001"}
→ 实际执行: set_text("wnd[0]/...", "1000001")
```

**读取表格数据注意事项：**

SAP GridView 的 `GetCellValue(row, colName)` 第二参数是列名字符串，不是列索引整数。
需先通过 `grid.ColumnOrder` 获取列名列表，再按名读取。

---

### 3.2 能力库：`sap_capabilities.yaml`

存储所有已沉淀的 SAP 业务能力。每个能力条目结构如下：

```yaml
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

- name: 查询销售订单
  description: 根据订单号在 VA03 中查询销售订单详细信息
  tcode: VA03
  created_at: "2026-05-28T10:30:00"
  parameters:
    - name: order_id
      description: 销售订单号
      required: true
  steps:
    - action: open_tcode
      params:
        tcode: VA03
      description: 打开销售订单显示事务码 VA03
    - action: fill_field
      params:
        field: wnd[0]/usr/ctxtVBAK-VBELN
        value_from: order_id       # 运行时从参数取值
      description: 填写销售订单号
    - action: press_key
      params:
        key: ENTER
      description: 按回车进入订单详情
    - action: get_status
      params:
        output_name: result
      description: 读取操作结果
```

**支持的 action 类型：**

| action | 关键参数 | 说明 |
|---|---|---|
| `open_connection` | `system_name` 或 `system_name_from` | 通过 SAP Logon Pad 打开系统连接 |
| `open_tcode` | `tcode` | 打开事务码 |
| `fill_field` | `field` + `value` 或 `value_from` | 填写字段 |
| `press_key` | `key` | 按键 |
| `click_button` | `button_id` | 点击按钮 |
| `read_table` | `table_id`, `output_name` | 读取表格到输出 |
| `get_status` | `output_name` | 读取状态栏到输出 |
| `wait` | `seconds` | 等待 |

**登录凭据与连接策略：**

登录能力运行时所需的凭据从 `sap_connections.yaml` 读取（该文件不入版本控制）：

```yaml
# sap_connections.yaml — 存储 SAP 系统连接凭据（本地个人使用）
connections:
  - name: DEV           # 对应 SAP Logon Pad 中的连接条目名
    client: "100"
    user: "YANGLI"
    password: "xxx"
    language: "ZH"

  - name: PRD
    client: "200"
    user: "YANGLI"
    password: "yyy"
    language: "ZH"

default: DEV            # 缺省连接
```

**`_connect()` 逻辑：**

1. 尝试获取已有 SAP 会话 → 成功则直接复用，无需登录
2. 没有已有会话 → 读取 `sap_connections.yaml` 中对应系统凭据 → 调用"登录SAP"能力自动登录

调用方式：
- `SAPRobot()` — 有会话就复用，没有就用 `default` 系统登录
- `SAPRobot(system_name="PRD")` — 指定登录哪个系统

---

### 3.3 录制管道

用户操作一遍 SAP，系统自动生成可执行配置 + 可阅读文档。由四个模块串联实现。

#### 管道阶段总览

```
【录制阶段】
  全自动：python sap_recorder.py start
             ↓ COM: SAPGuiAuto.Utils.Record(path)
          用户在 SAP 中正常操作
             ↓ COM: SAPGuiAuto.Utils.Stop()
  半自动：用户手动录制并保存 .vbs 文件
          python sap_recorder.py parse my_recording.vbs

【解析阶段】vbs_parser.py
  .sap_record/session.vbs → .sap_record/steps.json

【增强阶段】step_enricher.py + LLM
  .sap_record/steps.json → .sap_record/enriched.json
  （交互式终端：AI 建议 → 用户逐步确认）

【导出阶段】doc_exporter.py
  .sap_record/enriched.json
        ├→ sap_capabilities.yaml（追加）
        └→ docs/sop/{名称}.md（新建）
```

每个阶段独立可重跑，中间产物保存在 `.sap_record/` 临时目录。

#### 录制控制：`sap_recorder.py`

SAP GUI 的内置录制器通过 COM 接口 `SAPGuiAuto.Utils` 暴露控制方法，可由 Python 直接驱动：

```python
sap_gui_auto = win32com.client.GetObject("SAPGUI")
utils = sap_gui_auto.Utils
utils.Record(r"C:\绝对路径\.sap_record\session.vbs")   # 开始录制
utils.Stop()                                            # 停止录制
```

录制器启动后，用户在 SAP GUI 中的所有操作会被自动写入指定的 `.vbs` 文件。

**CLI 命令：**

| 命令 | 说明 |
|---|---|
| `start` | 通过 COM 开启 SAP 内置录制器 |
| `stop` | 停止录制，session.vbs 自动写入 |
| `parse [file.vbs]` | 解析 VBScript（省略路径则用 session.vbs） |
| `enrich` | AI 增强 + 终端交互确认 |
| `export` | 导出 YAML + Markdown |
| `pipeline` | 一键全流程（start → 等待回车 → stop → parse → enrich → export） |

#### VBScript 解析：`vbs_parser.py`

SAP 录制输出的 VBScript 样例：

```vbscript
' 样板头（忽略）
If Not IsObject(application) Then ...

' 有效操作行
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA03"        ← 打开事务码
session.findById("wnd[0]").sendVKey 0                           ← 按回车
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = "1000001" ← 填字段
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").setFocus         ← 噪音，忽略
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 7 ← 噪音，忽略
session.findById("wnd[0]").sendVKey 0                           ← 按回车
session.findById("wnd[0]/tbar[1]/btn[8]").press                ← 点击按钮
```

**解析规则：**

| VBScript 模式 | 输出 action | 说明 |
|---|---|---|
| `okcd").text = "/nXXXX"` | `open_tcode` | 打开事务码（特殊 fill_field） |
| `findById("...").text = "value"` | `fill_field` | 填写字段 |
| `sendVKey 0` | `press_key: ENTER` | |
| `sendVKey 3/8/12/26` | `press_key: F3/F8/F12/CTRL+S` | |
| `.press` | `click_button` | 点击按钮 |
| `.setFocus` / `.caretPosition` / `.maximize` | 丢弃 | 噪音 |
| 其他 `findById` 用法 | `unknown` | 保留原始行，不静默丢弃 |
| 样板头（`If Not IsObject` 等） | 跳过 | |

`unknown` 步骤保留到 enrich 阶段，由用户人工决策，不自动丢弃。

#### AI 增强：`step_enricher.py`

**交互流程：**

```
加载 steps.json
    ↓
LLM 批量分析所有步骤 → 建议整体能力名称和描述 → 用户确认或修改
    ↓
逐步处理每个 step：
  LLM 生成业务描述（≤20字）
  用户：[y] 确认 / [m] 修改 / [s] 跳过
  若为 fill_field 且有固定值：
    询问是否参数化
    若是：LLM 推断参数名，用户确认，提升到 parameters 列表
    ↓
写入 enriched.json
```

**终端交互示例：**

```
AI 建议能力名称：查询销售订单
AI 建议描述：根据订单号在 VA03 中查询销售订单详细信息
[y] 确认  [m] 修改  > y

步骤 2/5  fill_field → wnd[0]/usr/ctxtVBAK-VBELN = "1000001"
AI 描述：填写销售订单号
[y] 确认  [m] 修改  [s] 跳过  > y

值 "1000001" 是否设为运行时参数？[y] 是  [n] 否  > y
参数名（回车默认 order_id）> [回车]
参数 'order_id' 的说明：销售订单号
```

**LLM 调用设计：**

| 调用时机 | 模型 | 说明 |
|---|---|---|
| 启动时，批量分析所有步骤 | claude-haiku-4-5-20251001 | 生成能力名称和整体描述 |
| 每步一次 | claude-haiku-4-5-20251001 | ≤20字业务描述 |
| fill_field 参数化时 | claude-haiku-4-5-20251001 | 从字段路径推断 snake_case 参数名 |

**参数化逻辑：**

`value: "1000001"` 在用户确认参数化后：
- `params` 中 `value` 键替换为 `value_from: order_id`
- `order_id` 提升到顶层 `parameters` 列表
- 同名参数不重复添加

#### 双格式导出：`doc_exporter.py`

**输出一：追加到 `sap_capabilities.yaml`**
- 从 `enriched.json` 映射为能力库格式，追加写入
- 若同名能力已存在：提示 `[o] 覆盖 / [r] 重命名 / [c] 取消`

**输出二：`docs/sop/{能力名}.md`**

```markdown
# 查询销售订单

> 根据订单号在 VA03 中查询销售订单详细信息

- **事务码**：VA03
- **录制时间**：2026-05-28

## 参数

| 参数名 | 说明 | 必填 |
|---|---|---|
| order_id | 销售订单号 | 是 |

## 操作步骤

| # | 操作 | 元素路径 | 说明 |
|---|---|---|---|
| 1 | open_tcode | — | 打开销售订单显示事务码 VA03 |
| 2 | fill_field | `wnd[0]/usr/ctxtVBAK-VBELN` | 填写销售订单号 |
| 3 | press_key ENTER | — | 按回车进入订单详情 |

## 原始录制文件

- 来源：`.sap_record/session.vbs`
```

---

### 3.4 MCP Server：`sap_mcp_server.py`

将 `sap_capabilities.yaml` 中的每个能力动态注册为 MCP 工具，供 Claude 调用。

**工作流程：**

1. 启动时读取 `sap_capabilities.yaml`
2. 每个 capability 注册为一个 MCP Tool，`inputSchema` 从 `parameters` 字段生成
3. Claude 调用工具时，接收参数并调用 `sap_robot.execute_sequence(steps, arguments)`
4. 将执行结果返回给 Claude

**与 Claude Desktop 集成：**

在 `%APPDATA%\Claude\claude_desktop_config.json` 中添加：

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

**运行模式：**

| 模式 | 命令参数 | 用途 |
|---|---|---|
| stdio | `--mode=stdio` | 与 Claude Desktop 集成（正式使用） |
| simple | `--mode=simple` | 本地命令行测试验证 |

---

### 3.5 Web Chat：`llm_agent.py` + `web_server.py`

为不使用 Claude Desktop 的用户提供本地浏览器聊天界面，通过 Anthropic Function Calling 驱动 SAP 操作。

**与 MCP Server 的区别：**

| | MCP Server | Web Chat |
|---|---|---|
| 对话循环控制 | Claude Desktop 内部控制 | `llm_agent.py` 自己控制 |
| 使用方式 | 需要安装 Claude Desktop | 任意浏览器打开 |
| 适用场景 | 个人单机使用 | 团队内部共享 |

#### `llm_agent.py`：Function Calling 会话引擎

**职责：**

- 读取 `sap_capabilities.yaml`，将每个能力转换为 Anthropic tools 格式
- 驱动 Function Calling 对话循环：接收用户消息 → 调用 Claude API → 若有 `tool_use` 则执行并将结果反馈给 Claude → 循环直到获得文本回复

**核心函数签名：**

```python
def load_tools(config_path: str) -> List[Dict]:
    """读取 YAML，转换为 Anthropic tools 格式（含 input_schema）"""

def chat(user_message: str, history: List, config_path: str) -> Tuple[str, List]:
    """单轮对话入口。返回 (回复文本, 更新后的完整历史)。"""
```

**Function Calling 循环逻辑：**

```
用户消息 + 历史
    ↓ Claude API（携带 tools）
stop_reason == "tool_use"？
    ↓ 是
    执行所有 tool_use block → sap_robot.execute_sequence()
    将 tool_result 追加到消息历史
    ↓ 再次调用 Claude API
stop_reason == "end_turn"
    ↓
提取文本块 → 返回给调用方
```

#### `web_server.py`：FastAPI 聊天服务

**端点：**

| 端点 | 方法 | 说明 |
|---|---|---|
| `/` | GET | 返回内联 HTML 聊天页面 |
| `/chat` | POST | 接收 `{message, history}` → 返回 `{reply, history}` |

**设计约定：**

- HTML 聊天页面内联在 `web_server.py` 中，无需额外文件
- 消息历史由前端持有并随每次请求上传，服务端完全无状态
- 默认监听 `0.0.0.0:8000`

**启动方式：**

```bash
python web_server.py
# 浏览器打开 http://localhost:8000
```

---

## 四、中间数据格式

### `steps.json`（解析产物）

```json
{
  "raw_file": ".sap_record/session.vbs",
  "parsed_at": "2026-05-28T10:00:00",
  "steps": [
    {
      "seq": 1,
      "action": "open_tcode",
      "params": { "tcode": "VA03" },
      "raw_line": "session.findById(\"wnd[0]/tbar[0]/okcd\").text = \"/nVA03\""
    },
    {
      "seq": 2,
      "action": "fill_field",
      "params": { "field": "wnd[0]/usr/ctxtVBAK-VBELN", "value": "1000001" },
      "raw_line": "session.findById(\"wnd[0]/usr/ctxtVBAK-VBELN\").text = \"1000001\""
    }
  ]
}
```

### `enriched.json`（AI 增强产物）

```json
{
  "capability_name": "查询销售订单",
  "capability_description": "根据订单号在 VA03 中查询销售订单详细信息",
  "tcode": "VA03",
  "created_at": "2026-05-28T10:30:00",
  "parameters": [
    { "name": "order_id", "description": "销售订单号", "required": true }
  ],
  "steps": [
    {
      "seq": 1,
      "action": "open_tcode",
      "params": { "tcode": "VA03" },
      "description": "打开销售订单显示事务码 VA03"
    },
    {
      "seq": 2,
      "action": "fill_field",
      "params": { "field": "wnd[0]/usr/ctxtVBAK-VBELN", "value_from": "order_id" },
      "description": "填写销售订单号"
    }
  ]
}
```

---

## 五、错误处理

| 场景 | 处理方式 |
|---|---|
| SAP GUI 未运行（start 命令） | 捕获 COM 异常，提示用户开启 SAP，退出 |
| `sap_connections.yaml` 不存在（自动登录时） | 提示创建凭据配置文件，退出 |
| 指定的 `system_name` 未在配置中找到 | 列出可用系统名，提示用户选择 |
| 登录失败（密码错误等） | 读取状态栏错误消息，返回具体原因 |
| SAP Scripting 未启用 | 捕获 COM 异常，提示联系 Basis 管理员设置 `sapgui/user_scripting=TRUE` |
| `.vbs` 文件不存在（parse 命令） | 提示文件路径不存在，退出 |
| `steps.json` 不存在（enrich 命令） | 提示先运行 parse，退出 |
| LLM 调用失败 | 跳过 AI 描述，直接进入用户手动输入模式 |
| 同名能力已存在（export） | 交互询问：覆盖 / 重命名 / 取消 |
| unknown 步骤 | 终端显示提示，不自动跳过，由用户决定 |

---

## 六、文件结构

```
sap_mcp_project/
│
├── sap_robot.py              SAP GUI 原子操作封装（执行层）
│
├── sap_capabilities.yaml     能力配置库（持续积累）
├── sap_connections.yaml      SAP 系统连接凭据（不入版本控制）
├── config_manager.py         能力库增删改查工具
│
├── sap_recorder.py           录制管道 CLI 入口
├── vbs_parser.py             VBScript 解析器
├── step_enricher.py          AI 增强 + 交互确认
├── doc_exporter.py           双格式导出
│
├── sap_mcp_server.py         MCP Server
├── llm_agent.py              Function Calling 会话引擎
├── web_server.py             Web Chat FastAPI 服务（含内联 HTML）
│
├── .sap_record/              临时目录（不入版本控制）
│   ├── session.vbs
│   ├── steps.json
│   └── enriched.json
│
├── docs/
│   ├── specs/                设计文档
│   ├── plans/                实现计划
│   └── sop/                  录制产出的 SOP 文档
│
├── tests/
│   └── fixtures/
│       └── sample.vbs
│
├── requirements.txt
└── run_server.bat
```

---

## 七、依赖清单

```
pywin32      SAP GUI COM 接口（Windows 专属）
pyyaml       YAML 读写
anthropic    LLM 调用（AI 增强 + Function Calling）
mcp>=1.0.0   MCP Server SDK
fastapi      Web Chat 服务框架
uvicorn      ASGI 服务器（运行 FastAPI）
pytest       单元测试
```

---

## 八、实现顺序

```
第一步：新建 sap_robot.py（执行层，包含完整原子方法）
         ↓
第二步：新建 vbs_parser.py（VBScript 解析，纯函数，可独立测试）
         ↓
第三步：新建 doc_exporter.py（YAML + MD 导出，纯函数，可独立测试）
         ↓
第四步：新建 step_enricher.py（AI 增强 + 交互）
         ↓
第五步：新建 sap_recorder.py（CLI 入口，串联前四步）
         ↓
第六步：端到端验证（用 fixture 跑通 parse → enrich → export）
         ↓
第七步：新建 sap_mcp_server.py（MCP 集成）
         ↓
第八步：新建 llm_agent.py（Function Calling 会话引擎）
         ↓
第九步：新建 web_server.py（Web Chat 服务）
         ↓
第十步：集成测试（Claude Desktop + Web Chat 双路径验证）
```
