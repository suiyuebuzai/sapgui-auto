import yaml
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Tuple


def export_yaml(
    enriched: Dict[str, Any],
    yaml_path: str = "sap_capabilities.yaml",
) -> Tuple[bool, str]:
    """
    将能力条目追加写入 YAML 能力库。
    返回 (True, "ok") 成功；(False, "duplicate") 表示同名能力已存在。
    """
    path = Path(yaml_path)
    config = yaml.safe_load(path.read_text(encoding="utf-8")) if path.exists() else {}
    config.setdefault("capabilities", [])

    name = enriched["capability_name"]
    if any(c.get("name") == name for c in config["capabilities"]):
        return False, "duplicate"

    config["capabilities"].append({
        "name": name,
        "description": enriched["capability_description"],
        "tcode": enriched.get("tcode", ""),
        "created_at": enriched.get("created_at", datetime.now().isoformat()),
        "parameters": enriched.get("parameters", []),
        "steps": [
            {"action": s["action"], "params": s["params"], "description": s.get("description", "")}
            for s in enriched["steps"]
        ],
    })
    with open(path, "w", encoding="utf-8") as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False)
    return True, "ok"


def export_markdown(enriched: Dict[str, Any], docs_dir: str = "docs/sop") -> str:
    """生成 SOP Markdown 文档，返回文件路径。"""
    Path(docs_dir).mkdir(parents=True, exist_ok=True)
    name = enriched["capability_name"]

    params = enriched.get("parameters", [])
    param_rows = "\n".join(
        f"| {p['name']} | {p.get('description', '')} | {'是' if p.get('required') else '否'} |"
        for p in params
    ) or "| — | — | — |"

    step_rows = []
    for s in enriched["steps"]:
        p = s.get("params", {})
        if s["action"] == "fill_field":
            elem = f"`{p.get('field', '')}`"
        elif s["action"] == "click_button":
            elem = f"`{p.get('button_id', '')}`"
        else:
            elem = "—"
        step_rows.append(f"| {s['seq']} | {s['action']} | {elem} | {s.get('description', '')} |")

    content = (
        f"# {name}\n\n"
        f"> {enriched['capability_description']}\n\n"
        f"- **事务码**：{enriched.get('tcode', '—')}\n"
        f"- **录制时间**：{enriched.get('created_at', '')[:10]}\n\n"
        f"## 参数\n\n"
        f"| 参数名 | 说明 | 必填 |\n|---|---|---|\n{param_rows}\n\n"
        f"## 操作步骤\n\n"
        f"| # | 操作 | 元素路径 | 说明 |\n|---|---|---|---|\n"
        + "\n".join(step_rows)
        + "\n\n## 原始录制文件\n\n- 来源：`.sap_record/session.vbs`\n"
    )
    path = Path(docs_dir) / f"{name}.md"
    path.write_text(content, encoding="utf-8")
    return str(path)
