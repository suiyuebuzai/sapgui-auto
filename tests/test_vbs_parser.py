from pathlib import Path
import pytest
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
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
