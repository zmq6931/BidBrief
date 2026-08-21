"""DeepSeek (OpenAI-compatible API) extraction client."""
import json
import time

import requests

# Category names are Chinese on purpose: they are data keys shared with the
# prompt, the pipeline, and the exported reports.
CATEGORIES = ["资质要求", "商务要求", "技术要求", "时间节点", "废标项", "评分标准"]

# Safety cap per category per chunk to guard against runaway model output.
MAX_ITEMS_PER_CATEGORY = 30


class CancelledError(Exception):
    """Raised when the user requested a stop."""


class ApiAuthError(Exception):
    """Auth-type failure (bad key / no quota); retrying other files is pointless."""


# Prompt sent to the LLM, in Chinese on purpose: it drives extraction of
# Chinese-language tender documents and returns the Chinese category keys above.
SYSTEM_PROMPT = """你是资深的招投标文件分析专家。任务：从用户给出的一段招标文件内容中，抽取以下六类关键信息，供投标前核对：

1. 资质要求：对投标人的资格、资质证书、业绩、人员、信用、财务等要求
2. 商务要求：工期/交货期/服务期、最高限价（控制价）、投标保证金、付款方式、质保期等
3. 技术要求：技术参数、规格标准、实施方案、服务要求等
4. 时间节点：投标截止、开标、澄清/答疑截止、保证金缴纳截止、现场踏勘等具体时间安排
5. 废标项：导致投标无效/废标/否决投标的条款（一票否决项）
6. 评分标准：评分项、分值、评分办法

严格要求：
- 只抽取文中明确出现的内容，禁止编造、推测或补充常识
- 每条必须附"原文摘录"：从原文逐字复制的片段（不超过60个字），用于定位页码，必须与原文完全一致
- 金额、日期、数量、工期等数字必须与原文完全一致
- 本段没有某类内容时，对应数组返回空数组
- 单类条目不超过15条，只保留重要的

只输出如下 JSON，不要输出任何其他内容：
{
  "资质要求": [{"内容": "", "原文摘录": ""}],
  "商务要求": [{"内容": "", "原文摘录": ""}],
  "技术要求": [{"内容": "", "原文摘录": ""}],
  "时间节点": [{"事项": "", "时间": "", "原文摘录": ""}],
  "废标项": [{"内容": "", "原文摘录": ""}],
  "评分标准": [{"评分项": "", "分值": "", "说明": "", "原文摘录": ""}]
}"""


def _build_system_prompt(custom):
    """Base prompt plus the user's custom focus instructions, if any.

    The custom text may narrow the scope (e.g. subcontract-only bids); the JSON
    schema must stay unchanged so downstream parsing keeps working.
    """
    if not custom:
        return SYSTEM_PROMPT
    return (
        SYSTEM_PROMPT
        + "\n\n【用户补充要求（优先级高于以上默认要求）】\n"
        + custom.strip()
        + "\n\n若补充要求限定了抽取范围，范围之外的内容可以不抽取；但输出 JSON 结构保持不变，没有命中的类别返回空数组。"
    )


def extract_chunk(cfg, chunk, cancel=None):
    """Extract key points from one text chunk; returns {category: [item dict]}."""
    sp, ep = chunk["start_page"], chunk["end_page"]
    if sp is not None:
        loc = f"第 {sp}-{ep} 页" if sp != ep else f"第 {sp} 页"
        user = f"【招标文件片段】{loc}\n\n{chunk['text']}"
    else:
        user = f"【招标文件片段】（无页码信息）\n\n{chunk['text']}"

    content = _chat(cfg, [
        {"role": "system", "content": _build_system_prompt(cfg.get("custom_prompt", ""))},
        {"role": "user", "content": user},
    ], cancel)
    return _parse_json(content)


def _chat(cfg, messages, cancel=None):
    """Call the chat-completions endpoint with retry on transient failures."""
    url = cfg["base_url"].rstrip("/") + "/chat/completions"
    headers = {
        "Authorization": f"Bearer {cfg['api_key']}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": cfg["model"],
        "messages": messages,
        "temperature": cfg.get("temperature", 0.1),
        "max_tokens": cfg.get("max_tokens", 8192),
        "response_format": {"type": "json_object"},
        "stream": False,
    }
    last_err = None
    for attempt in range(max(1, cfg.get("max_retries", 3))):
        if cancel is not None and cancel.is_set():
            raise CancelledError()
        try:
            r = requests.post(url, headers=headers, json=payload,
                              timeout=cfg.get("request_timeout", 180))
            if r.status_code == 200:
                return r.json()["choices"][0]["message"]["content"]
            body = r.text[:300]
            if r.status_code in (401, 403):
                raise ApiAuthError(f"API 认证失败({r.status_code})，请检查 api_key：{body}")
            if r.status_code in (400, 404, 422):
                raise ApiAuthError(f"API 请求被拒绝({r.status_code})，请检查 base_url/model：{body}")
            last_err = f"HTTP {r.status_code}: {body}"
        except (requests.Timeout, requests.ConnectionError) as e:
            last_err = f"网络错误: {e}"
        if attempt < cfg.get("max_retries", 3) - 1:
            if cancel is not None and cancel.is_set():
                raise CancelledError()
            wait = cfg.get("retry_wait", 5) * (attempt + 1)
            time.sleep(wait)
    raise RuntimeError(f"API 调用失败（重试后仍失败）：{last_err}")


def _parse_json(content):
    """Parse model output robustly (tolerates code fences and stray text)."""
    s = (content or "").strip()
    if s.startswith("```"):
        s = s.strip("`")
        if s.lower().startswith("json"):
            s = s[4:]
        s = s.strip()
    try:
        data = json.loads(s)
    except json.JSONDecodeError:
        lo, hi = s.find("{"), s.rfind("}")
        if lo < 0 or hi <= lo:
            return {c: [] for c in CATEGORIES}
        try:
            data = json.loads(s[lo : hi + 1])
        except json.JSONDecodeError:
            return {c: [] for c in CATEGORIES}

    out = {}
    for cat in CATEGORIES:
        items = data.get(cat) if isinstance(data, dict) else None
        clean = []
        if isinstance(items, list):
            for it in items:
                if isinstance(it, dict):
                    clean.append({
                        "内容": str(it.get("内容", "")).strip(),
                        "事项": str(it.get("事项", "")).strip(),
                        "时间": str(it.get("时间", "")).strip(),
                        "评分项": str(it.get("评分项", "")).strip(),
                        "分值": str(it.get("分值", "")).strip(),
                        "说明": str(it.get("说明", "")).strip(),
                        "原文摘录": str(it.get("原文摘录", "")).strip(),
                    })
        out[cat] = clean[:MAX_ITEMS_PER_CATEGORY]
    return out
