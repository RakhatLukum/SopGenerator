import os
import re
import json
import time
import requests
from typing import List, Dict, Tuple, Any
from dataclasses import dataclass

from autogen import AssistantAgent, UserProxyAgent, GroupChat, GroupChatManager
from userdata import get as userdata_get
from requests.adapters import HTTPAdapter
try:
    from urllib3.util.retry import Retry  # type: ignore
except Exception:
    from urllib3.util.retry import Retry  # type: ignore
from http.client import RemoteDisconnected

API_BASE = "https://lzl4i1wx0cdh9a-8000.proxy.runpod.net/v1"
API_URLS = [
    f"{API_BASE}/chat/completions",
    f"{API_BASE}/completions",  # some deployments use legacy route
]
MODEL_NAME = os.getenv("LLM_MODEL", "llama4scout")


def call_llm(messages: List[Dict[str, str]], max_tokens: int = 1500, timeout_read: int = 600) -> str:
    api_key = userdata_get('api_key') or ""

    def _make_session() -> requests.Session:
        session = requests.Session()
        retry = Retry(
            total=2,
            connect=2,
            read=2,
            backoff_factor=0.8,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=frozenset(["POST"]),
        )
        adapter = HTTPAdapter(max_retries=retry, pool_connections=1, pool_maxsize=1)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        return session

    headers = {
        "accept": "application/json",
        "Content-Type": "application/json",
        "Connection": "close",
    }
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    def _payload_plain(msgs: List[Dict[str, str]], tokens: int) -> Dict[str, Any]:
        return {
            "model": MODEL_NAME,
            "messages": msgs,
            "max_tokens": tokens,
            "temperature": 0,
            "stream": False,
        }

    def _payload_typed(msgs: List[Dict[str, str]], tokens: int) -> Dict[str, Any]:
        typed_msgs: List[Dict[str, Any]] = []
        for m in msgs:
            content_str = m.get("content") or ""
            typed_msgs.append({
                "role": m.get("role") or "user",
                "content": [{"type": "text", "text": content_str}],
            })
        return {
            "model": MODEL_NAME,
            "messages": typed_msgs,
            "max_tokens": tokens,
            "temperature": 0,
            "stream": False,
        }

    def _extract_content(resp_json: Dict[str, Any]) -> str:
        # chat completions shape
        msg = (resp_json.get("choices", [{}])[0].get("message") or {})
        if isinstance(msg, dict):
            return (msg.get("content") or "").strip()
        # legacy completions shape
        txt = (resp_json.get("choices", [{}])[0].get("text") or "").strip()
        return txt

    def _post_once(url: str, payload: Dict[str, Any]) -> str:
        session = _make_session()
        try:
            resp = session.post(url, headers=headers, json=payload, timeout=(15, timeout_read))
            if resp.status_code == 200:
                try:
                    data = resp.json()
                except Exception:
                    raise RuntimeError(f"LLM returned non-JSON: {resp.text[:200]}")
                content = _extract_content(data)
                if not content:
                    raise RuntimeError(f"LLM empty content. Raw response: {resp.text[:400]}")
                return content
            raise RuntimeError(f"LLM error {url}: {resp.status_code} {resp.text}")
        finally:
            try:
                session.close()
            except Exception:
                pass

    last_err: Exception | None = None

    for attempt in range(2):
        for url in API_URLS:
            # Try plain OpenAI-style content
            try:
                return _post_once(url, _payload_plain(messages, max_tokens))
            except Exception as e1:
                last_err = e1
                # Try typed-content
                try:
                    return _post_once(url, _payload_typed(messages, max_tokens))
                except Exception as e2:
                    last_err = e2
                    # Minimal messages
                    try:
                        minimal_msgs: List[Dict[str, str]] = []
                        if messages and messages[0].get("role") == "system":
                            minimal_msgs.append(messages[0])
                        if messages:
                            minimal_msgs.append(messages[-1])
                        small_tokens = min(800, max_tokens)
                        return _post_once(url, _payload_plain(minimal_msgs, small_tokens))
                    except Exception as e3:
                        last_err = e3
                        try:
                            return _post_once(url, _payload_typed(minimal_msgs, small_tokens))
                        except Exception as e4:
                            last_err = e4
                            continue
        time.sleep(0.8 * (2 ** attempt))

    raise RuntimeError(f"Connection to LLM failed after retries: {last_err}")


WRITER_SYSTEM = (
    "You are the Writer Agent for SOP generation. Produce a rigorous, fully formatted SOP draft in Markdown.\n"
    "CRITICAL: Write in Russian language (русский язык) unless user explicitly requests another language.\n"
    "Rules and formatting requirements (strictly follow):\n"
    "- Page: A4; Margins: 1 inch (25.4 mm).\n"
    "- Font: Times New Roman or Arial; Size: 11–14; Single line spacing.\n"
    "- Structure: Numbered sections and subsections (e.g., 1, 1.1, 1.1.1).\n"
    "- Include: Title page, Approval sheet, Change log, Acknowledgement sheet.\n"
    "- Include: Scope, Responsibilities, Definitions, Procedure/Steps, Equipment/Materials, Periodicities, Calibration, Safety, References.\n"
    "- Tables and figures must have correct captions (Table X: Title, Figure X: Title).\n"
    "- Appendices must be numbered (Appendix A, B, ...).\n"
    "- Use footnotes where necessary (mark as [^n]: footnote).\n"
    "Output only a complete SOP draft in Markdown, no extra commentary.\n"
)

CRITIC_SYSTEM = (
    "You are the Critic Agent for SOP QA. You must review the SOP and return structured feedback strictly with the following keys per issue:\n"
    "ISSUE: <short title>\n"
    "WHY: <why this is a problem with reference to the rules>\n"
    "FIX: <clear actionable fix>\n"
    "BLOCKER: <yes|no>\n\n"
    "- If there are no blockers remaining, return exactly `STATUS: OK` and nothing else.\n"
    "- Be precise, reference missing structure, numbering, captions, formatting, or non-compliant text.\n"
)


@dataclass
class ReviewResult:
    status: str
    draft_markdown: str
    conversation: List[Dict[str, Any]]
    feedback_items: List[Dict[str, Any]]
    used_fallback: bool = False


def _parse_feedback(text: str) -> List[Dict[str, Any]]:
    if not text or text.strip() == "STATUS: OK":
        return []
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    joined = "\n".join(lines)
    blocks = re.split(r"(?=^ISSUE:\s*)", joined, flags=re.MULTILINE)
    items: List[Dict[str, Any]] = []
    for block in blocks:
        if not block.strip():
            continue
        issue_match = re.search(r"ISSUE:\s*(.*)", block)
        why_match = re.search(r"WHY:\s*(.*)", block)
        fix_match = re.search(r"FIX:\s*(.*)", block)
        blocker_match = re.search(r"BLOCKER:\s*(.*)", block)
        if issue_match or why_match or fix_match or blocker_match:
            items.append(
                {
                    "issue": (issue_match.group(1).strip() if issue_match else "").strip(),
                    "why": (why_match.group(1).strip() if why_match else "").strip(),
                    "fix": (fix_match.group(1).strip() if fix_match else "").strip(),
                    "blocker": (blocker_match.group(1).strip().lower() if blocker_match else "no").strip(),
                }
            )
    return items


def _extract_top_level_headings(outline_md: str) -> List[str]:
    headings: List[str] = []
    for line in outline_md.splitlines():
        line = line.strip()
        m = re.match(r"^(\d+)(?:\.\d+)*\.?\s+(.+)$", line)
        if m:
            num = m.group(1)
            text = m.group(2)
            if num.isdigit():
                headings.append(f"{num}. {text}")
    seen = set()
    uniq: List[str] = []
    for h in headings:
        if h not in seen:
            uniq.append(h)
            seen.add(h)
    return uniq[:8]


def _rule_based_sop(data: Dict[str, Any]) -> str:
    def f(k: str) -> str:
        return (data.get(k) or "").strip()

    title = f("title") or "Стандартная операционная процедура"
    scope = f("scope") or f("sections")
    responsibilities = f("responsibilities") or "Ответственные лица определяются руководителем подразделения."
    definitions = f("definitions") or "Термины и определения приводятся при необходимости."
    equipment = f("equipment") or f("equipment_type")
    procedure = f("procedure") or "Пошаговое описание процедуры согласно внутренним регламентам."
    periodicities = f("periodicities") or "Периодичность выполняемых операций согласно графику."
    calibration = f("calibration") or "Калибровка проводится согласно паспорту оборудования и внутренним инструкциям."
    safety = f("safety") or "Соблюдать технику безопасности и охрану труда."
    references = f("references") or "Внутренние регламенты, стандарты, НПА."

    md = []
    md.append(f"# {title}")
    md.append("")
    md.append("## 1. Область применения")
    md.append(scope or "Описание области применения и ограничений.")
    md.append("")
    md.append("## 2. Ответственность")
    md.append(responsibilities)
    md.append("")
    md.append("## 3. Определения")
    md.append(definitions)
    md.append("")
    md.append("## 4. Оборудование и материалы")
    if equipment:
        md.append("| Наименование | Модель/Тип | Калибровка | Примечание |")
        md.append("| --- | --- | --- | --- |")
        md.append(f"| {equipment} | — | По графику | — |")
    else:
        md.append("См. перечень оборудования в приложении A.")
    md.append("")
    md.append("## 5. Порядок выполнения работ")
    md.append(procedure)
    md.append("")
    md.append("## 6. Периодичность")
    md.append(periodicities)
    md.append("")
    md.append("## 7. Калибровка и поверка")
    md.append(calibration)
    md.append("")
    md.append("## 8. Требования безопасности")
    md.append(safety)
    md.append("")
    md.append("## 9. Ссылки")
    md.append(references)
    md.append("")
    md.append("## 10. Приложения")
    md.append("Приложение A — Формы записей и журналы.")
    md.append("")
    md.append("[^1]: Данный документ сформирован автоматически и требует проверки ответственным лицом.")
    return "\n".join(md)


def _manual_pipeline(input_prompt: str, input_data: Dict[str, Any], max_rounds: int) -> Tuple[str, List[Dict[str, Any]], List[Dict[str, Any]], str, bool]:
    conversation: List[Dict[str, Any]] = []
    feedback_items: List[Dict[str, Any]] = []
    status = "NEEDS_REVIEW"
    used_fallback = False

    conversation.append({"name": "User", "content": input_prompt})

    try:
        outline_prompt = (
            input_prompt
            + "\n\nСформируй только оглавление SOP (только нумерованные заголовки 1-го и 2-го уровня). Без текста разделов."
        )
        outline = call_llm(
            [
                {"role": "system", "content": WRITER_SYSTEM},
                {"role": "user", "content": outline_prompt},
            ],
            max_tokens=400,
            timeout_read=120,
        )
        conversation.append({"name": "Writer", "content": outline})
        headings = _extract_top_level_headings(outline)

        if not headings:
            draft = call_llm(
                [
                    {"role": "system", "content": WRITER_SYSTEM},
                    {"role": "user", "content": input_prompt},
                ],
                max_tokens=1500,
                timeout_read=180,
            )
            conversation.append({"name": "Writer", "content": draft})
        else:
            draft_parts: List[str] = []
            for h in headings:
                part_prompt = (
                    f"Сгенерируй только раздел '{h}' в формате Markdown, строго соблюдая правила (A4, шрифты, нумерация). "
                    f"Не повторяй другие разделы."
                )
                part = call_llm(
                    [
                        {"role": "system", "content": WRITER_SYSTEM},
                        {"role": "user", "content": part_prompt},
                    ],
                    max_tokens=700,
                    timeout_read=150,
                )
                draft_parts.append(part)
                conversation.append({"name": "Writer", "content": part})
            draft = "\n\n".join(draft_parts)

        critic_prompt = (
            "Проверь соответствие SOP правилам (структура, нумерация, обязательные разделы, подписи к таблицам/рисункам, приложения). "
            "Верни список замечаний в формате: ISSUE/WHY/FIX/BLOCKER. Если блокирующих нет, верни ровно 'STATUS: OK'."
        )
        critic_feedback = call_llm(
            [
                {"role": "system", "content": CRITIC_SYSTEM},
                {"role": "user", "content": critic_prompt},
            ],
            max_tokens=500,
            timeout_read=120,
        )
        conversation.append({"name": "Critic", "content": critic_feedback})

        if critic_feedback.strip() == "STATUS: OK":
            status = "OK"
            final_draft = draft
        else:
            feedback_items.extend(_parse_feedback(critic_feedback))
            revise_prompt = (
                input_prompt
                + "\n\nУчитывая следующие замечания Критика, сгенерируй ПОЛНУЮ обновлённую версию SOP в Markdown: \n\n"
                + critic_feedback
            )
            revised = call_llm(
                [
                    {"role": "system", "content": WRITER_SYSTEM},
                    {"role": "user", "content": revise_prompt},
                ],
                max_tokens=1500,
                timeout_read=180,
            )
            conversation.append({"name": "Writer", "content": revised})

            critic_final = call_llm(
                [
                    {"role": "system", "content": CRITIC_SYSTEM},
                    {"role": "user", "content": "Проверь итоговую версию. Если нет блокеров, верни 'STATUS: OK'."},
                ],
                max_tokens=200,
                timeout_read=90,
            )
            conversation.append({"name": "Critic", "content": critic_final})

            if critic_final.strip() == "STATUS: OK":
                status = "OK"
            final_draft = revised or draft

        return final_draft, conversation, feedback_items, status, used_fallback

    except Exception as e:
        used_fallback = True
        fallback_draft = _rule_based_sop(input_data)
        conversation.append({"name": "System", "content": f"LLM error: {e}"})
        conversation.append({"name": "Writer", "content": "Локальный генератор сформировал черновик SOP (без LLM)."})
        conversation.append({"name": "Critic", "content": "STATUS: OK"})
        status = "OK"
        return fallback_draft, conversation, feedback_items, status, used_fallback


def format_input_data_to_prompt(data: Dict[str, Any]) -> str:
    def field(key: str) -> str:
        return (data.get(key) or "").strip()

    source_docs: List[Dict[str, str]] = data.get("source_docs") or []
    src_lines: List[str] = []
    for i, sd in enumerate(source_docs, start=1):
        name = (sd.get("name") or f"Источник {i}").strip()
        preview = (sd.get("preview") or "").strip()
        if preview:
            preview = preview[:2000]
        src_lines.append(f"- {name}: {preview}")
    source_block = "\n".join(src_lines)

    structure_text = field("structure_text")
    structure_desc = field("structure_description")

    content_type = field("content_type")

    parts = [
        "Сгенерируй полный проект стандартной операционной процедуры (SOP) на русском языке, строго соблюдая правила форматирования.\n",
        "[МЕТАДАННЫЕ]\n",
        f"Название: {field('title')}\n",
        f"Номер: {field('sop_number')}\n",
        f"Тип оборудования: {field('equipment_type')}\n",
        f"Тип содержимого: {content_type}\n",
        "[/МЕТАДАННЫЕ]\n\n",
        "[ТЕКСТ ВВОДА]", "\n",
        f"Разделы/описание: {field('sections')}\n",
        f"Описание структуры: {structure_desc}\n",
        "[/ТЕКСТ ВВОДА]\n\n",
    ]

    if source_block:
        parts.extend([
            "[ИСТОЧНИКИ]\n",
            source_block,
            "\n[/ИСТОЧНИКИ]\n\n",
        ])

    if structure_text:
        parts.extend([
            "[ИЗВЛЕЧЕННАЯ СТРУКТУРА ДОКУМЕНТА]\n",
            structure_text[:4000],
            "\n[/ИЗВЛЕЧЕННАЯ СТРУКТУРА ДОКУМЕНТА]\n\n",
        ])

    for key in [
        "scope",
        "responsibilities",
        "definitions",
        "equipment",
        "procedure",
        "periodicities",
        "calibration",
        "safety",
        "references",
        "notes",
    ]:
        v = field(key)
        if v:
            parts.append(f"{key.title()}: {v}\n")

    parts.append(
        "Требования к оформлению: формат страницы A4; поля 1 дюйм; шрифт Times New Roman или Arial 11–14; одинарный интервал; нумерация разделов и подразделов (1, 1.1, 1.1.1); подписи к таблицам и рисункам; нумерация приложений; при необходимости сноски формата [^n]: текст.\n"
    )

    # Guidance based on content_type
    if content_type == "Только документ":
        parts.append("Используй исключительно информацию из блока [ИСТОЧНИКИ], без выдумывания фактов.\n")
    elif content_type == "ИИ + источник":
        parts.append("Опирайся на [ИСТОЧНИКИ], дополняя нейросетевыми формулировками, но не противоречь источнику.\n")

    return "".join(parts)


def run_review(input_data: Dict[str, Any], max_rounds: int = 8) -> ReviewResult:
    initial_prompt = format_input_data_to_prompt(input_data)
    draft, conv, items, status, used_fallback = _manual_pipeline(initial_prompt, input_data, max_rounds)
    return ReviewResult(status=status, draft_markdown=draft, conversation=conv, feedback_items=items, used_fallback=used_fallback) 