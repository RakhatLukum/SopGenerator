import os
import re
from datetime import datetime
from typing import Dict, List, Tuple, Optional
from difflib import unified_diff

from export import export_docx


def ensure_versions_dir(base_dir: Optional[str] = None) -> str:
    base_dir = base_dir or os.path.dirname(os.path.abspath(__file__))
    versions_dir = os.path.join(base_dir, "versions")
    os.makedirs(versions_dir, exist_ok=True)
    return versions_dir


def _scan_existing_indices(versions_dir: str) -> List[int]:
    indices: List[int] = []
    for name in os.listdir(versions_dir):
        m = re.match(r"^(\d{3})_.*\.docx$", name)
        if m:
            try:
                indices.append(int(m.group(1)))
            except Exception:
                pass
    return sorted(indices)


def next_index(versions_dir: str) -> int:
    indices = _scan_existing_indices(versions_dir)
    if not indices:
        return 1
    return max(indices) + 1


def save_version_docx(markdown_content: str, metadata: Dict, label: str = "draft", base_dir: Optional[str] = None) -> str:
    versions_dir = ensure_versions_dir(base_dir)
    idx = next_index(versions_dir)
    filename = f"{idx:03d}_{label}.docx"
    output_path = os.path.join(versions_dir, filename)

    # Attach version info into metadata
    _meta = dict(metadata or {})
    _meta.setdefault("version", idx)
    _meta.setdefault("author", "Writer Agent")

    export_docx(markdown_content, _meta, output_path)
    return output_path


def list_saved_versions(base_dir: Optional[str] = None) -> List[Tuple[int, str]]:
    versions_dir = ensure_versions_dir(base_dir)
    results: List[Tuple[int, str]] = []
    for name in os.listdir(versions_dir):
        m = re.match(r"^(\d{3})_.*\.docx$", name)
        if m:
            results.append((int(m.group(1)), os.path.join(versions_dir, name)))
    results.sort(key=lambda x: x[0])
    return results


def compute_unified_diff(a_text: str, b_text: str, a_label: str = "A", b_label: str = "B") -> str:
    a_lines = (a_text or "").splitlines(keepends=True)
    b_lines = (b_text or "").splitlines(keepends=True)
    diff = unified_diff(a_lines, b_lines, fromfile=a_label, tofile=b_label)
    return "".join(diff)


def sanitize_markdown(md: str) -> str:
    if not md:
        return md
    # Remove common page-break divs and any standalone HTML tags
    cleaned = re.sub(r"<div[^>]*>\s*</div>", "", md, flags=re.IGNORECASE)
    cleaned = re.sub(r"<br\s*/?>", "\n", cleaned, flags=re.IGNORECASE)
    # Strip any remaining HTML tags
    cleaned = re.sub(r"<[^>]+>", "", cleaned)

    # Remove LLM front-matter/meta lines
    lines = cleaned.splitlines()
    meta_patterns = [
        r"^\s*\*\*?СТАНДАРТНАЯ ОПЕРАЦИОННАЯ ПРОЦЕДУРА\*\*?\s*$",
        r"^\s*\*\*?Номер:\*\*?.*$",
        r"^\s*\*\*?Тип оборудования:\*\*?.*$",
        r"^\s*\*\*?Тип содержимого:\*\*?.*$",
        r"^\s*\*\*?Страница:\*\*?.*$",
        r"^\s*\*\*?Шрифт:\*\*?.*$",
        r"^\s*Номер:\s*.*$",
        r"^\s*Тип оборудования:\s*.*$",
        r"^\s*Тип содержимого:\s*.*$",
        r"^\s*Страница:\s*.*$",
        r"^\s*Шрифт:\s*.*$",
    ]
    meta_regex = [re.compile(p, re.IGNORECASE) for p in meta_patterns]

    kept: List[str] = []
    prev_was_meta = False
    for idx, line in enumerate(lines):
        is_meta = any(rx.search(line) for rx in meta_regex)
        # Skip setext underline immediately following meta/title
        if not is_meta and prev_was_meta and re.match(r"^[=-]{3,}\s*$", line):
            prev_was_meta = False
            continue
        if is_meta:
            prev_was_meta = True
            continue
        prev_was_meta = False
        kept.append(line)

    cleaned = "\n".join(kept)

    # Normalize multiple blank lines
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned


def _read_all_bytes(uploaded_file) -> bytes:
    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    data = uploaded_file.read()
    try:
        uploaded_file.seek(0)
    except Exception:
        pass
    return data


def _pdf_text_with_fallback(uploaded_file) -> str:
    text = ""
    try:
        import pdfplumber  # type: ignore
        with pdfplumber.open(uploaded_file) as pdf:
            pages = [page.extract_text() or "" for page in pdf.pages]
            text = "\n".join(pages)
    except Exception:
        try:
            from pypdf import PdfReader  # type: ignore
            uploaded_file.seek(0)
            reader = PdfReader(uploaded_file)
            pages = []
            for p in reader.pages:
                try:
                    pages.append(p.extract_text() or "")
                except Exception:
                    pages.append("")
            text = "\n".join(pages)
        except Exception:
            text = ""
    finally:
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
    return text


def extract_uploaded_text(uploaded_file) -> Tuple[str, str]:
    """Return (full_text, preview). Supports .docx, .pdf, .txt/.csv, and .xlsx/.xls.
    Preview is truncated to ~1500 chars.
    """
    name = (getattr(uploaded_file, "name", "file") or "file").lower()
    text = ""

    if name.endswith(".docx"):
        try:
            from docx import Document  # type: ignore
            doc = Document(uploaded_file)
            text = "\n".join(p.text for p in doc.paragraphs)
        except Exception:
            text = ""
        finally:
            try:
                uploaded_file.seek(0)
            except Exception:
                pass
    elif name.endswith(".pdf"):
        text = _pdf_text_with_fallback(uploaded_file)
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        try:
            import openpyxl  # type: ignore
            data = _read_all_bytes(uploaded_file)
            from io import BytesIO
            wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
            parts: List[str] = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    cells = [str(c) if c is not None else "" for c in row]
                    if any(cells):
                        parts.append("\t".join(cells))
            text = "\n".join(parts)
        except Exception:
            text = ""
    elif name.endswith(".txt") or name.endswith(".csv"):
        try:
            data = _read_all_bytes(uploaded_file)
            text = data.decode(errors="ignore")
        except Exception:
            text = ""
    else:
        text = ""

    preview_len = 1500
    preview = text[:preview_len]
    return text, preview


def format_size(num_bytes: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if num_bytes < 1024.0:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.1f} TB" 