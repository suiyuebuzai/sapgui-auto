# SAPGUI Auto

## 项目目标

帮助用户将 SAP GUI 上的业务处理能力沉淀为标准化文档，并通过 AI 实现自然语言驱动 SAP 自动化操作。

## 核心想法

1. **能力沉淀**：将用户在 SAP GUI 上的操作录制下来，转化为标准化的业务流程文档（MD / YAML / JSON）
2. **底层驱动**：基于 SAP GUI Scripting API（官方 COM 接口），用 Python 封装常用原子操作
3. **元素寻址**：SAP 使用唯一字符串路径定位界面元素，如 `wnd[0]/usr/txtRSYST-MANDT`
4. **AI 转换引擎**：构建"自然语言 → SAP 操作"的转换链路，形成可复用的能力库
5. **MCP 集成**：将能力库通过 MCP Server 暴露为 AI 可调用工具，与 Claude Desktop 直接集成
6. **Web Chat**：通过 Anthropic Function Calling 提供浏览器聊天界面，无需 Claude Desktop

## 整体架构（三层模型）

```
用户自然语言输入
      ↓
AI 接口层（两种方式二选一）
  方式 A：Claude Desktop → MCP Server（sap_mcp_server.py）
  方式 B：浏览器 → Web Chat → llm_agent.py（Function Calling）
      ↓
执行层（sap_robot.py，调用 SAP GUI Scripting API）
      ↓
能力库（sap_capabilities.yaml，录制管道持续积累）
```

## 技术环境要求

- Windows 系统（SAP GUI Scripting API 依赖 COM）
- SAP GUI 7.40+（推荐 7.70+）
- SAP 系统参数 `sapgui/user_scripting = TRUE`（需 Basis 管理员开启）
- Python 3.8+
- Anthropic API Key（AI 增强 + Function Calling 均需要）

## 设计文档

- 整体系统设计（含录制管道）：`docs/specs/sapgui-auto-system-design.md`

## 实现计划

- `docs/plans/sapgui-auto-implementation.md`
