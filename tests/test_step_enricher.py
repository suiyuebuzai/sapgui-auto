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
