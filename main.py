import os
from datetime import datetime
from typing import Dict, Any, List

import streamlit as st

from agents import run_review
from utils import ensure_versions_dir, save_version_docx, compute_unified_diff, extract_uploaded_text, format_size, sanitize_markdown

APP_DIR = os.path.dirname(os.path.abspath(__file__))
VERSIONS_DIR = ensure_versions_dir(APP_DIR)


def _init_state() -> None:
    if "versions" not in st.session_state:
        st.session_state["versions"] = []  # list of dicts with content_md, status, feedback, conversation, file_path
    if "current_idx" not in st.session_state:
        st.session_state["current_idx"] = None


def _current_version() -> Dict[str, Any] | None:
    idx = st.session_state.get("current_idx")
    if idx is None:
        return None
    versions: List[Dict[str, Any]] = st.session_state.get("versions", [])
    if 0 <= idx < len(versions):
        return versions[idx]
    return None


def _add_version(entry: Dict[str, Any]) -> int:
    versions: List[Dict[str, Any]] = st.session_state["versions"]
    versions.append(entry)
    st.session_state["current_idx"] = len(versions) - 1
    return st.session_state["current_idx"]


def main():
    st.set_page_config(page_title="SOP Автор-Критик", layout="wide")
    _init_state()

    st.title("SOP Автор-Критик (AutoGen + Streamlit)")

    left, right = st.columns(2)

    with left:
        st.text_input("Название СОП", key="title")
        st.text_input("Номер СОП", key="sop_number")
        st.text_input("Тип оборудования", key="equipment_type")
        st.text_area(
            "Разделы",
            placeholder=(
                "Опиши правила работы с оборудованием, технические характеристики, а также информацию из файла.\n"
                "Кратко расскажи, для чего применяется и зачем нужен."
            ),
            key="sections",
            height=140,
        )
        content_type = st.selectbox(
            "Тип содержимого",
            [
                "ИИ + источник",
                "Только ИИ",
                "Только документ",
            ],
            index=0,
        )
        st.text_area(
            "Описание структуры документа (опционально)",
            key="structure_description",
            placeholder="Распиши каждый раздел развёрнуто и подробно",
            height=120,
        )
    with right:
        st.caption("Исходный документ(ы)")
        source_files = st.file_uploader(
            "Загрузите исходный документ(ы)",
            type=["pdf", "docx", "txt", "xlsx", "xls", "csv"],
            accept_multiple_files=True,
        )
        st.caption("Документ со структурой (опционально)")
        struct_file = st.file_uploader(
            "Загрузите документ со структурой",
            type=["pdf", "docx", "txt", "xlsx", "xls", "csv"],
            accept_multiple_files=False,
        )

        structure_text_preview = ""
        if struct_file is not None:
            full_text, preview = extract_uploaded_text(struct_file)
            structure_text_preview = preview
            st.write(struct_file.name)
            st.code(preview or "Предпросмотр недоступен для данного формата.")

        source_docs: List[Dict[str, str]] = []
        if source_files:
            st.write("Загруженные источники:")
            for f in source_files:
                _, preview = extract_uploaded_text(f)
                source_docs.append({"name": f.name, "preview": preview})
                size = 0
                try:
                    size = len(f.getvalue())
                except Exception:
                    pass
                st.markdown(f"- {f.name} ({format_size(size) if size else ''})")

    col1, col2 = st.columns([1, 1])
    with col1:
        run_clicked = st.button("Запустить проверку", type="primary")
    with col2:
        export_clicked = st.button("Экспорт в DOCX")
        download_placeholder = st.empty()

    if run_clicked:
        if content_type == "Только документ" and not source_docs:
            st.error("Вы выбрали 'Только документ', но не загрузили исходный документ.")
            return

        # Build context blocks
        input_data: Dict[str, Any] = {
            "title": st.session_state.get("title") or "",
            "sop_number": st.session_state.get("sop_number") or "",
            "equipment_type": st.session_state.get("equipment_type") or "",
            "sections": st.session_state.get("sections") or "",
            "content_type": content_type,
            "structure_description": st.session_state.get("structure_description") or "",
            "structure_text": structure_text_preview,
            "source_docs": source_docs,
        }
        try:
            with st.spinner("Запуск многодельного обзора..."):
                result = run_review(input_data=input_data, max_rounds=8)
            entry = {
                "timestamp": datetime.utcnow().isoformat(),
                "title": input_data["title"] or "Без названия",
                "status": result.status,
                "content_md": sanitize_markdown(result.draft_markdown),
                "feedback_items": result.feedback_items,
                "conversation": result.conversation,
                "file_path": None,
            }
            _add_version(entry)
            if getattr(result, 'used_fallback', False):
                st.warning("LLM недоступен — сформирован локальный черновик SOP. Проверьте и дополните при необходимости.")
            st.success(f"Проверка завершена. Статус: {result.status}")
        except Exception as e:
            st.error(f"Ошибка при обращении к LLM: {e}")

    current = _current_version()

    if export_clicked:
        if not current:
            st.error("Нет текущего черновика. Сначала запустите проверку.")
        elif not current.get("content_md"):
            st.error("Текущий черновик пуст.")
        else:
            from export import export_docx_bytes
            label = "approved" if (current.get("status") == "OK") else "draft"
            metadata = {"title": current.get("title", "Без названия")}
            data = export_docx_bytes(current.get("content_md") or "", metadata)
            filename = f"SOP_{metadata['title']}_{label}.docx".replace(" ", "_")
            with download_placeholder:
                st.download_button(
                    label="Скачать DOCX",
                    data=data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            st.success("Файл готов к скачиванию.")

    tabs = st.tabs(["Черновик", "Диалог", "Обратная связь", "Сравнение версий"])

    with tabs[0]:
        st.subheader("Текущий черновик SOP")
        if current and current.get("content_md"):
            st.markdown(current["content_md"])
            st.caption(f"Статус: {current.get('status')} | Экспорт: {'Да' if current.get('file_path') else 'Нет'}")
        else:
            st.info("Черновик отсутствует. Нажмите 'Запустить проверку'.")

    with tabs[1]:
        st.subheader("Диалог: Автор ↔ Критик")
        if current and current.get("conversation"):
            for msg in current["conversation"]:
                name = msg.get("name") or msg.get("role") or "User"
                content = msg.get("content") or ""
                st.markdown(f"**{name}:**")
                st.write(content)
                st.divider()
        else:
            st.info("Диалогов пока нет.")

    with tabs[2]:
        st.subheader("Обратная связь Критика")
        if current:
            if current.get("status") == "OK":
                st.success("STATUS: OK. Блокирующих замечаний нет.")
            items = current.get("feedback_items") or []
            if items:
                st.dataframe(items, use_container_width=True)
            elif current and current.get("status") != "OK":
                st.info("Структурированных замечаний не найдено.")
        else:
            st.info("Запустите проверку, чтобы увидеть замечания.")

    with tabs[3]:
        st.subheader("Сравнение версий")
        versions: List[Dict[str, Any]] = st.session_state.get("versions", [])
        if len(versions) < 2:
            st.info("Для сравнения нужно минимум две версии.")
        else:
            labels = [f"{i+1}: {v['title']} ({v['status']})" for i, v in enumerate(versions)]
            idx_a = st.selectbox("Из версии", list(range(len(versions))), format_func=lambda i: labels[i], key="cmp_a")
            idx_b = st.selectbox("В версию", list(range(len(versions))), format_func=lambda i: labels[i], key="cmp_b")
            do_compare = st.button("Сравнить версии")
            if do_compare:
                a = versions[idx_a]
                b = versions[idx_b]
                diff = compute_unified_diff(a.get("content_md") or "", b.get("content_md") or "", a_label=f"v{idx_a+1}", b_label=f"v{idx_b+1}")
                if diff.strip():
                    st.code(diff)
                else:
                    st.info("Различий нет.")


if __name__ == "__main__":
    main() 