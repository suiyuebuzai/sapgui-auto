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
