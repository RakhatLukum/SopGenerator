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


def extract_uploaded_text(uploaded_file) -> Tuple[str, str]:
    """Return (full_text, preview). Supports .docx and .txt/.csv; others return empty text.
    The preview is truncated to ~1500 characters for UI display and prompt context.
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
    elif name.endswith(".txt") or name.endswith(".csv"):
        try:
            data = _read_all_bytes(uploaded_file)
            text = data.decode(errors="ignore")
        except Exception:
            text = ""
    else:
        text = ""  # PDF/XLSX not supported without extra deps

    preview_len = 1500
    preview = text[:preview_len]
    return text, preview


def format_size(num_bytes: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if num_bytes < 1024.0:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.1f} TB" 