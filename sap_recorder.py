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
        import win32com.client
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
        if choice == "r":
            enriched["capability_name"] = input("新名称：").strip()
            export_yaml(enriched)
        elif choice != "c":
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
