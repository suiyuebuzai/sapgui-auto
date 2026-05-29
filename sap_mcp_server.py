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
