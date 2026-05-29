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
