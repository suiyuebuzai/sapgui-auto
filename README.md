# SAPGUI Auto

将 SAP GUI 上的业务操作能力沉淀为标准化文档，并通过 AI 实现自然语言驱动 SAP 自动化操作。

## 核心功能

- **操作录制**：通过 SAP GUI Scripting API 录制用户操作，自动转化为结构化步骤
- **AI 增强**：LLM 自动生成业务描述、推断参数名，交互式确认后导出为能力库
- **MCP Server**：将能力库暴露为 AI 工具，与 Claude Desktop 直接集成
- **Web Chat**：浏览器聊天界面，自然语言驱动 SAP 操作（Anthropic Function Calling）

## 架构

```
用户自然语言输入
      ↓
AI 接口层（二选一）
  方式 A：Claude Desktop → MCP Server（sap_mcp_server.py）
  方式 B：浏览器 → Web Chat → llm_agent.py（Function Calling）
      ↓
执行层（sap_robot.py，调用 SAP GUI Scripting API）
      ↓
能力库（sap_capabilities.yaml，录制管道持续积累）
```

## 环境要求

- Windows 系统（SAP GUI Scripting API 依赖 COM）
- SAP GUI 7.40+（推荐 7.70+）
- SAP 系统参数 `sapgui/user_scripting = TRUE`
- Python 3.8+
- Anthropic API Key

## 安装

```bash
pip install -r requirements.txt
```

## 快速开始

### 1. 配置 SAP 连接

复制凭据模板并填写：

```bash
cp sap_connections.yaml.example sap_connections.yaml
# 编辑 sap_connections.yaml，填入你的 SAP 系统信息
```

### 2. 录制能力（完整管道）

```bash
python sap_recorder.py pipeline
```

或分步执行：

```bash
python sap_recorder.py start      # 开始录制
# 在 SAP GUI 中完成操作...
python sap_recorder.py stop       # 停止录制
python sap_recorder.py parse      # 解析 VBS → steps.json
python sap_recorder.py enrich     # AI 增强 + 交互确认
python sap_recorder.py export     # 导出到能力库 + 生成文档
```

### 3. 使用 Web Chat

```bash
python web_server.py
# 浏览器打开 http://localhost:8000
```

### 4. 接入 Claude Desktop（MCP）

编辑 `%APPDATA%\Claude\claude_desktop_config.json`：

```json
{
  "mcpServers": {
    "my-sap": {
      "command": "python",
      "args": ["C:/你的路径/sap_mcp_server.py", "--mode=stdio"]
    }
  }
}
```

重启 Claude Desktop 后即可用自然语言操作 SAP。

## 项目结构

```
├── sap_robot.py          # SAP GUI 原子操作封装
├── vbs_parser.py         # VBScript 录制文件解析
├── step_enricher.py      # LLM 描述生成 + 交互确认
├── doc_exporter.py       # 导出 YAML 能力库 + MD 文档
├── sap_recorder.py       # CLI 入口，串联录制管道
├── sap_mcp_server.py     # MCP Server（Claude Desktop 集成）
├── llm_agent.py          # Anthropic Function Calling 会话引擎
├── web_server.py         # FastAPI Web Chat 服务
├── sap_capabilities.yaml # 能力库
├── sap_connections.yaml  # SAP 连接凭据（不入版本控制）
├── tests/                # 单元测试（27 项）
└── docs/                 # 设计文档 + 实现计划
```

## 测试

```bash
pytest tests/ -v
```

## License

MIT
