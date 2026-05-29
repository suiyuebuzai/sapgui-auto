import json
import anthropic
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List, Optional

_client: Optional[anthropic.Anthropic] = None


def _get_client() -> anthropic.Anthropic:
    global _client
    if _client is None:
        import os
        api_key = os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")
        base_url = os.environ.get("ANTHROPIC_BASE_URL")
        kwargs = {"api_key": api_key}
        if base_url:
            kwargs["base_url"] = base_url
        _client = anthropic.Anthropic(**kwargs)
    return _client


def _get_model() -> str:
    import os
    return os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")


def suggest_capability_meta(steps: List[Dict]) -> Dict[str, str]:
    """批量请求：根据所有步骤建议能力名称和描述。"""
    summary = "\n".join(
        f"- {s['action']}: {json.dumps(s['params'], ensure_ascii=False)}" for s in steps
    )
    msg = _get_client().messages.create(
        model=_get_model(),
        max_tokens=200,
        messages=[{"role": "user", "content": (
            f"你是 SAP 业务专家。根据以下操作步骤，建议一个能力名称（5-10字）和描述（20-40字）。"
            f"以 JSON 格式输出：{{\"name\": \"...\", \"description\": \"...\"}}\n\n步骤：\n{summary}"
        )}],
    )
    try:
        return json.loads(msg.content[0].text)
    except Exception:
        return {"name": "未命名能力", "description": ""}


def suggest_step_description(step: Dict, tcode: str) -> str:
    """单步请求：用一句中文描述操作，≤20字。"""
    msg = _get_client().messages.create(
        model=_get_model(),
        max_tokens=50,
        messages=[{"role": "user", "content": (
            f"用一句中文描述这个 SAP 操作，不超过20字：\n"
            f"操作类型: {step['action']}\n"
            f"参数: {json.dumps(step['params'], ensure_ascii=False)}\n"
            f"当前事务码: {tcode}\n只输出描述文字。"
        )}],
    )
    return msg.content[0].text.strip()


def suggest_param_name(field_path: str) -> str:
    """根据 SAP 字段路径推断 snake_case 参数名。LLM 失败时回退到路径末段。"""
    msg = _get_client().messages.create(
        model=_get_model(),
        max_tokens=20,
        messages=[{"role": "user", "content": (
            f"SAP 字段路径 '{field_path}' 对应的英文参数名是什么？"
            f"只输出参数名（snake_case，无其他内容）。"
        )}],
    )
    import re
    name = msg.content[0].text.strip().lower().replace("-", "_").replace(" ", "_")
    if not name or len(name) > 30 or not re.match(r'^[a-z][a-z0-9_]*$', name):
        last = field_path.split("/")[-1]
        name = last.split("-")[-1].lower() if "-" in last else last.lower()
    return name


def run_interactive(
    steps_path: str = ".sap_record/steps.json",
    output_path: str = ".sap_record/enriched.json",
):
    """交互式 AI 增强会话，引导用户逐步确认每个步骤的业务描述和参数化。"""
    with open(steps_path, encoding="utf-8") as f:
        data = json.load(f)

    steps = data["steps"]
    tcode = next(
        (s["params"].get("tcode", "") for s in steps if s["action"] == "open_tcode"), ""
    )

    print(f"\n{'=' * 50}\nSAP 能力录制 - AI 增强阶段\n共 {len(steps)} 个步骤\n{'=' * 50}")

    # 确认能力元信息
    meta = suggest_capability_meta(steps)
    print(f"\nAI 建议能力名称：{meta['name']}\nAI 建议描述：{meta['description']}")
    if input("[y] 确认  [m] 修改  > ").strip().lower() == "m":
        meta["name"] = input("能力名称：").strip() or meta["name"]
        meta["description"] = input("能力描述：").strip() or meta["description"]

    parameters: List[Dict] = []
    enriched_steps: List[Dict] = []

    for step in steps:
        print(f"\n{'─' * 40}")
        print(f"步骤 {step['seq']}/{len(steps)}  {step['action']}  →  "
              f"{json.dumps(step['params'], ensure_ascii=False)}")

        if step["action"] == "unknown":
            print("⚠ 未识别步骤，跳过")
            enriched_steps.append({**step, "description": "（未识别）"})
            continue

        desc = suggest_step_description(step, tcode)
        print(f"AI 描述：{desc}")
        choice = input("[y] 确认  [m] 修改  [s] 跳过  > ").strip().lower()
        if choice == "s":
            continue
        if choice == "m":
            desc = input("描述：").strip() or desc

        new_params = dict(step["params"])
        if step["action"] == "fill_field" and step["params"].get("value"):
            if input(f'\n值 "{step["params"]["value"]}" 是否设为运行时参数？[y] 是  [n] 否  > ').strip().lower() == "y":
                suggested = suggest_param_name(step["params"].get("field", ""))
                param_name = input(f"参数名（回车默认 {suggested}）> ").strip() or suggested
                new_params = {"field": step["params"]["field"], "value_from": param_name}
                if not any(p["name"] == param_name for p in parameters):
                    param_desc = input(f"参数 '{param_name}' 的说明：").strip() or param_name
                    parameters.append({"name": param_name, "description": param_desc, "required": True})

        enriched_steps.append({"seq": step["seq"], "action": step["action"],
                                "params": new_params, "description": desc})

    result = {
        "capability_name": meta["name"],
        "capability_description": meta["description"],
        "tcode": tcode,
        "created_at": datetime.now().isoformat(),
        "parameters": parameters,
        "steps": enriched_steps,
    }
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"\n{'=' * 50}\n增强完成，已写入 {output_path}")
