import json
import yaml
import anthropic
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

from sap_robot import SAPRobot

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
