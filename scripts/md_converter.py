"""Markdown 转换工具 V3.1 — 将 PDF/DOCX/TXT 文献转换为 Markdown 格式。

V3.1 变更:
  - 写后验证重试使用 _retry1 后缀（明确命名规则）

V3 功能:
  - 结构化内容模板（标题/作者/年份/来源 + 正文）
  - 写后强制验证（检查文件存在且大小 > 0）
  - 空文件检测（0 KB / < 1 KB）

使用方式:
  python md_converter.py convert <目录>           # 转换目录下所有文献
  python md_converter.py convert <文件> --structured  # 结构化转换单个文件
  python md_converter.py check <md目录>           # 检测空文件
"""

import os
import re
import sys
import argparse
from pathlib import Path
from typing import List, Tuple, Dict, Optional

# ── Constants ───────────────────────────────────────────────────
PAPER_EXTS = {'.pdf', '.docx', '.doc', '.txt', '.epub', '.mobi'}
EMPTY_THRESHOLD_KB = 0      # 0 KB → 完全失败
SMALL_THRESHOLD_KB = 1      # < 1 KB → 高度疑似失败


# ── Ensure unique filename ──────────────────────────────────────
def ensure_unique(dest_dir: str, filename: str) -> str:
    """Ensure target filename is unique by appending incrementing number."""
    base, ext = os.path.splitext(filename)
    dest = os.path.join(dest_dir, filename)
    if not os.path.exists(dest):
        return dest
    counter = 1
    while True:
        new_name = f"{base}_{counter}{ext}"
        new_dest = os.path.join(dest_dir, new_name)
        if not os.path.exists(new_dest):
            return new_dest
        counter += 1


# ── PDF → Markdown ──────────────────────────────────────────────
def pdf_to_markdown(filepath: str) -> Tuple[str, bool, str]:
    """Extract text from PDF and return as Markdown-formatted string."""
    text_parts = []
    try:
        try:
            import fitz
            doc = fitz.open(filepath)
            for page in doc:
                t = page.get_text()
                if t:
                    text_parts.append(t.strip())
            doc.close()
            return '\n\n'.join(text_parts), True, ""
        except ImportError:
            pass

        try:
            import PyPDF2
            reader = PyPDF2.PdfReader(filepath)
            for page in reader.pages:
                t = page.extract_text()
                if t:
                    text_parts.append(t.strip())
            return '\n\n'.join(text_parts), True, ""
        except ImportError:
            pass

        return "", False, "No PDF library available (install PyMuPDF or PyPDF2)"
    except Exception as e:
        return "", False, f"PDF extraction failed: {e}"


# ── DOCX → Markdown ─────────────────────────────────────────────
def docx_to_markdown(filepath: str) -> Tuple[str, bool, str]:
    """Extract text from DOCX and return as Markdown-formatted string."""
    try:
        try:
            from docx import Document
            doc = Document(filepath)
            lines = []
            for para in doc.paragraphs:
                text = para.text.strip()
                if not text:
                    lines.append('')
                    continue
                style = para.style.name if para.style else ''
                if 'Heading 1' in style or 'heading 1' in style.lower():
                    lines.append(f'# {text}')
                elif 'Heading 2' in style or 'heading 2' in style.lower():
                    lines.append(f'## {text}')
                elif 'Heading 3' in style or 'heading 3' in style.lower():
                    lines.append(f'### {text}')
                else:
                    lines.append(text)
            return '\n\n'.join(lines), True, ""
        except ImportError:
            return "", False, "python-docx not installed"
    except Exception as e:
        return "", False, f"DOCX extraction failed: {e}"


# ── TXT → Markdown ──────────────────────────────────────────────
def txt_to_markdown(filepath: str) -> Tuple[str, bool, str]:
    """Read plain text file as Markdown."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read().strip(), True, ""
    except UnicodeDecodeError:
        try:
            with open(filepath, 'r', encoding='gbk') as f:
                return f.read().strip(), True, ""
        except Exception as e:
            return "", False, f"TXT encoding error: {e}"
    except Exception as e:
        return "", False, f"TXT read failed: {e}"


# ── Structured MD template ──────────────────────────────────────
def _build_structured_md(raw_text: str, filename: str,
                         metadata: Optional[Dict[str, str]] = None) -> str:
    """Build a structured Markdown file with header metadata.

    Args:
        raw_text: Extracted text content
        filename: Source filename (for fallback title)
        metadata: Optional dict with 'title', 'author', 'year', 'source' keys

    Returns:
        Formatted Markdown string
    """
    if metadata is None:
        metadata = {}

    base_name = os.path.splitext(filename)[0]
    lines = []

    # Title
    title = metadata.get('title', '') or base_name.replace('_', ' ').replace('-', ' ')
    lines.append(f'# {title}')
    lines.append('')

    # Metadata section
    author = metadata.get('author', '')
    year = metadata.get('year', '')
    source = metadata.get('source', '')

    if author:
        lines.append(f'**Author:** {author}')
    if year:
        lines.append(f'**Year:** {year}')
    if source:
        lines.append(f'**Source:** {source}')

    if author or year or source:
        lines.append('')
        lines.append('---')
        lines.append('')

    # Body
    clean_text = raw_text
    if clean_text.startswith('# '):
        first_line_end = clean_text.find('\n')
        if first_line_end > 0:
            first_line = clean_text[:first_line_end].strip()
            if filename.lower().replace('_', ' ').replace('-', ' ') in first_line.lower():
                clean_text = clean_text[first_line_end + 1:].lstrip()

    lines.append(clean_text)
    return '\n'.join(lines)


# ── V3.1: Write with verify ─────────────────────────────────────
def _write_with_verify(filepath: str, content: str, max_retries: int = 1) -> Tuple[bool, str]:
    """Write content to file and verify it was written correctly.

    V3.1: On failure, retry with _retry1 suffix.

    Returns (success, note).
    """
    dir_path = os.path.dirname(filepath)
    if dir_path:
        os.makedirs(dir_path, exist_ok=True)

    # Attempt 1
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)

        # Verify
        if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
            return True, "OK"
    except Exception as e:
        pass  # Fall to retry

    # V3.1: Retry with _retry1 suffix
    if max_retries > 0:
        base, ext = os.path.splitext(filepath)
        retry_path = f"{base}_retry1{ext}"
        try:
            with open(retry_path, 'w', encoding='utf-8') as f:
                f.write(content)
            if os.path.exists(retry_path) and os.path.getsize(retry_path) > 0:
                return True, f"Retried as {os.path.basename(retry_path)}"
        except Exception as e:
            return False, f"Retry failed: {e}"

    return False, "Write verification failed"


# ── Single file conversion ──────────────────────────────────────
def convert_file(filepath: str, output_dir: str = None, dry_run: bool = False,
                 structured: bool = False, metadata: Optional[Dict] = None) -> Dict:
    """Convert a single file to Markdown.

    Returns: {"source": ..., "output": ..., "status": "ok|skip|error|preview", "note": ...}
    """
    if output_dir is None:
        output_dir = os.path.dirname(filepath)

    ext = Path(filepath).suffix.lower()
    base_name = os.path.splitext(os.path.basename(filepath))[0]
    md_filename = f"{base_name}.md"
    md_path = ensure_unique(output_dir, md_filename)

    if dry_run:
        return {"source": filepath, "output": md_path, "status": "preview", "note": "dry-run"}

    # Convert based on type
    if ext == '.pdf':
        md_text, ok, err = pdf_to_markdown(filepath)
    elif ext == '.docx':
        md_text, ok, err = docx_to_markdown(filepath)
    elif ext == '.doc':
        md_text, ok, err = docx_to_markdown(filepath)
        if not ok:
            return {"source": filepath, "output": "", "status": "skip",
                    "note": f"Cannot parse .doc: {err}"}
    elif ext == '.txt':
        md_text, ok, err = txt_to_markdown(filepath)
    elif ext in ('.epub', '.mobi'):
        return {"source": filepath, "output": "", "status": "skip",
                "note": f"Cannot convert {ext} (requires specialized tools)"}
    else:
        return {"source": filepath, "output": "", "status": "skip",
                "note": f"Unsupported format: {ext}"}

    if not ok:
        return {"source": filepath, "output": "", "status": "skip", "note": err}

    # Apply structured template if requested
    if structured:
        md_text = _build_structured_md(md_text, os.path.basename(filepath), metadata)

    # Write with V3.1 verify+retry
    written, note = _write_with_verify(md_path, md_text)

    if written:
        return {"source": filepath, "output": md_path, "status": "ok", "note": note}
    else:
        return {"source": filepath, "output": "", "status": "error", "note": note}


# ── Directory conversion ────────────────────────────────────────
def convert_directory(target_dir: str, dry_run: bool = False,
                      structured: bool = False) -> Dict:
    """Convert all papers in a directory (excluding Duplicates).

    Returns: {"converted": [...], "skipped": [...], "errors": [...]}
    """
    result = {"converted": [], "skipped": [], "errors": []}

    for root, dirs, files in os.walk(target_dir):
        dirs[:] = [d for d in dirs if d != 'Duplicates']
        dirs[:] = [d for d in dirs if not d.startswith('.')]

        for f in files:
            ext = Path(f).suffix.lower()
            if ext not in PAPER_EXTS:
                continue

            fpath = os.path.join(root, f)
            res = convert_file(fpath, output_dir=root, dry_run=dry_run,
                               structured=structured)

            rel_source = os.path.relpath(fpath, target_dir)
            if res["status"] == "ok":
                result["converted"].append({
                    "source": rel_source,
                    "output": os.path.relpath(res["output"], target_dir),
                    "note": res.get("note", "")
                })
            elif res["status"] == "skip":
                result["skipped"].append({
                    "source": rel_source,
                    "note": res["note"]
                })
            else:
                result["errors"].append({
                    "source": rel_source,
                    "note": res["note"]
                })

    return result


# ── Empty file detection ────────────────────────────────────────
def check_empty_files(md_dir: str) -> Dict:
    """Check all .md files in directory for empty/small files.

    Returns: {"ok": count, "empty": [(name, size)], "small": [(name, size)]}
    """
    result = {"ok": 0, "empty": [], "small": []}

    for root, dirs, files in os.walk(md_dir):
        for f in files:
            if not f.endswith('.md'):
                continue
            fpath = os.path.join(root, f)
            size_bytes = os.path.getsize(fpath)
            size_kb = size_bytes / 1024
            rel_path = os.path.relpath(fpath, md_dir)

            if size_bytes == 0:
                result["empty"].append((rel_path, round(size_kb, 2)))
            elif size_kb < SMALL_THRESHOLD_KB:
                result["small"].append((rel_path, round(size_kb, 2)))
            else:
                result["ok"] += 1

    return result


# ── Check all three MD folders ──────────────────────────────────
def check_all_md_folders(md_base: str) -> List[Dict]:
    """Check md_all, md_ByYear, md_ByTopic folders."""
    results = []
    for folder in ['md_all', 'md_ByYear', 'md_ByTopic']:
        # Find the actual folder (may have parent prefix)
        found = None
        for root, dirs, _ in os.walk(md_base):
            for d in dirs:
                if d.endswith(folder):
                    found = os.path.join(root, d)
                    break
            if found:
                break

        if found and os.path.isdir(found):
            check = check_empty_files(found)
            check['folder'] = folder
            results.append(check)

    return results


# ── CLI ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Markdown 转换工具 V3.1")
    sub = parser.add_subparsers(dest="cmd")

    convert_p = sub.add_parser("convert", help="转换文献为 Markdown")
    convert_p.add_argument("path", help="文件或目录路径")
    convert_p.add_argument("--dry-run", action="store_true", help="仅预览")
    convert_p.add_argument("--structured", action="store_true",
                           help="使用结构化模板（标题/作者/年份/来源）")

    check_p = sub.add_parser("check", help="检测空 MD 文件")
    check_p.add_argument("dir", help="MD 文件夹路径")

    args = parser.parse_args()

    if args.cmd == "convert":
        if os.path.isdir(args.path):
            res = convert_directory(args.path, dry_run=args.dry_run,
                                    structured=args.structured)
            print(f"\n## Converted ({len(res['converted'])})")
            for item in res["converted"]:
                note_str = f" [{item['note']}]" if item.get('note') and item['note'] != 'OK' else ''
                print(f"  {item['source']} -> {item['output']}{note_str}")
            print(f"\n## Skipped ({len(res['skipped'])})")
            for item in res["skipped"]:
                print(f"  {item['source']}: {item['note']}")
            if res["errors"]:
                print(f"\n## Errors ({len(res['errors'])})")
                for item in res["errors"]:
                    print(f"  {item['source']}: {item['note']}")
        else:
            res = convert_file(args.path, dry_run=args.dry_run,
                               structured=args.structured)
            print(f"  {res['status']}: {res.get('output', '')} {res.get('note', '')}")

    elif args.cmd == "check":
        res = check_empty_files(args.dir)
        print(f"OK: {res['ok']}")
        if res["empty"]:
            print(f"\n## EMPTY (0 KB) - {len(res['empty'])} files")
            for fname, size in res["empty"]:
                print(f"  {fname}: {size} KB")
        if res["small"]:
            print(f"\n## SMALL (< 1 KB) - {len(res['small'])} files")
            for fname, size in res["small"]:
                print(f"  {fname}: {size} KB")

        if res["empty"] or res["small"]:
            print("\n--- WARNING ---")
            print("The following MD files are empty or very small.")
            print("Original files may be scanned PDFs or image-only documents.")
            print("Consider using OCR tools to reprocess them.")
