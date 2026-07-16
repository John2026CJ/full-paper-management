#!/usr/bin/env python3
"""Markdown 转换工具 V2 — 将 PDF/DOCX/TXT 文献转换为 Markdown 格式。

任务四专用脚本：
  Step 2: 执行格式转换（PDF/DOCX/TXT → MD）
  Step 5: 空文件检测与提醒

功能:
  - PDF → MD: 提取文本层，保留段落结构
  - DOCX → MD: 提取段落和标题层级
  - TXT → MD: 直接复制内容
  - 空文件检测: 0 KB 和 < 1 KB 文件标记
  - 不删除原文件，不覆盖已有 MD

使用方式:
  python md_converter.py convert <目录>           # 转换目录下所有文献
  python md_converter.py convert <文件>           # 转换单个文件
  python md_converter.py check <md目录>           # 检测空文件
"""

import os, sys, re, shutil
from pathlib import Path

# ── 依赖检测 ──────────────────────────────────────────────────
HAS_PYPDF = False
HAS_DOCX = False

try:
    import PyPDF2
    HAS_PYPDF = True
except ImportError:
    pass

try:
    import docx
    HAS_DOCX = True
except ImportError:
    pass

PAPER_EXTS = {'.pdf', '.docx', '.doc', '.txt', '.epub', '.mobi'}


def ensure_unique(dest_dir: str, filename: str) -> str:
    """确保目标文件名唯一。"""
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


# ── PDF → Markdown ────────────────────────────────────────────

def pdf_to_markdown(filepath: str) -> tuple:
    """将 PDF 转换为 Markdown。

    Returns:
        (markdown_text, success_bool, error_msg)
    """
    if not HAS_PYPDF:
        return ("", False, "PyPDF2 not installed")

    try:
        reader = PyPDF2.PdfReader(filepath)
        md_parts = []

        for i, page in enumerate(reader.pages):
            try:
                text = page.extract_text()
                if text:
                    # 清理和格式化文本
                    lines = text.split('\n')
                    cleaned_lines = []
                    for line in lines:
                        line = line.strip()
                        if line:
                            cleaned_lines.append(line)

                    page_text = '\n\n'.join(cleaned_lines)
                    md_parts.append(page_text)
            except Exception:
                continue

        if not md_parts:
            return ("", False, "No text extracted (likely scanned PDF)")

        md_text = '\n\n---\n\n'.join(md_parts)

        # 添加标题
        filename = os.path.splitext(os.path.basename(filepath))[0]
        md_text = f"# {filename}\n\n{md_text}"

        return (md_text, True, "")

    except Exception as e:
        return ("", False, str(e))


# ── DOCX → Markdown ───────────────────────────────────────────

def docx_to_markdown(filepath: str) -> tuple:
    """将 DOCX 转换为 Markdown。"""
    if not HAS_DOCX:
        return ("", False, "python-docx not installed")

    try:
        doc = docx.Document(filepath)
        md_parts = []

        filename = os.path.splitext(os.path.basename(filepath))[0]
        md_parts.append(f"# {filename}\n")

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue

            style_name = para.style.name.lower() if para.style else ""

            # 根据样式添加 Markdown 格式
            if "heading 1" in style_name or "title" in style_name:
                md_parts.append(f"\n## {text}\n")
            elif "heading 2" in style_name:
                md_parts.append(f"\n### {text}\n")
            elif "heading 3" in style_name:
                md_parts.append(f"\n#### {text}\n")
            elif "heading 4" in style_name:
                md_parts.append(f"\n##### {text}\n")
            elif "list" in style_name:
                md_parts.append(f"- {text}")
            elif "quote" in style_name:
                md_parts.append(f"> {text}")
            else:
                md_parts.append(text)

        md_text = '\n\n'.join(md_parts)

        if len(md_text.strip()) < 50:
            return ("", False, "Document too short or empty")

        return (md_text, True, "")

    except Exception as e:
        return ("", False, str(e))


# ── TXT → Markdown ────────────────────────────────────────────

def txt_to_markdown(filepath: str) -> tuple:
    """将 TXT 转换为 Markdown。"""
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()

        if not text.strip():
            return ("", False, "Empty file")

        filename = os.path.splitext(os.path.basename(filepath))[0]
        md_text = f"# {filename}\n\n{text}"

        return (md_text, True, "")

    except Exception as e:
        return ("", False, str(e))


# ── 统一转换入口 ──────────────────────────────────────────────

def convert_file(filepath: str, output_dir: str = None, dry_run: bool = False) -> dict:
    """转换单个文件为 Markdown。

    Args:
        filepath: 源文件路径
        output_dir: 输出目录（默认为源文件所在目录）
        dry_run: 仅预览

    Returns:
        {"source": filepath, "output": path, "status": "ok|skip|error", "note": "..."}
    """
    if output_dir is None:
        output_dir = os.path.dirname(filepath)

    ext = Path(filepath).suffix.lower()
    base_name = os.path.splitext(os.path.basename(filepath))[0]
    md_filename = f"{base_name}.md"

    # 不覆盖已有 MD
    md_path = ensure_unique(output_dir, md_filename)

    if dry_run:
        return {"source": filepath, "output": md_path, "status": "preview", "note": "dry-run"}

    # 根据类型转换
    if ext == '.pdf':
        md_text, ok, err = pdf_to_markdown(filepath)
    elif ext == '.docx':
        md_text, ok, err = docx_to_markdown(filepath)
    elif ext == '.doc':
        # .doc 格式尝试用 docx 读取（可能失败）
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

    # 写入 MD 文件
    try:
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(md_text)
    except Exception as e:
        return {"source": filepath, "output": "", "status": "error", "note": str(e)}

    return {"source": filepath, "output": md_path, "status": "ok", "note": ""}


def convert_directory(target_dir: str, dry_run: bool = False) -> dict:
    """转换目录下所有文献文件（排除 Duplicates）。

    Returns:
        {"converted": [...], "skipped": [...], "errors": [...]}
    """
    result = {"converted": [], "skipped": [], "errors": []}

    for root, dirs, files in os.walk(target_dir):
        # 严格排除 Duplicates
        dirs[:] = [d for d in dirs if d != 'Duplicates']
        # 排除隐藏目录
        dirs[:] = [d for d in dirs if not d.startswith('.')]

        for f in files:
            ext = Path(f).suffix.lower()
            if ext not in PAPER_EXTS:
                continue

            fpath = os.path.join(root, f)
            res = convert_file(fpath, output_dir=root, dry_run=dry_run)

            if res["status"] == "ok":
                result["converted"].append({
                    "source": os.path.relpath(fpath, target_dir),
                    "output": os.path.relpath(res["output"], target_dir)
                })
            elif res["status"] == "skip":
                result["skipped"].append({
                    "source": os.path.relpath(fpath, target_dir),
                    "note": res["note"]
                })
            else:
                result["errors"].append({
                    "source": os.path.relpath(fpath, target_dir),
                    "note": res["note"]
                })

    return result


# ── 空文件检测 ────────────────────────────────────────────────

def check_empty_files(md_dir: str) -> dict:
    """检测 Markdown 文件夹中的空文件和小文件。

    判定标准:
        0 KB   → 转换失败（完全空文件）
        < 1 KB → 高度疑似转换失败
        >= 1 KB → 正常

    Returns:
        {"empty": [(name, size_kb)], "small": [(name, size_kb)], "ok": int}
    """
    result = {"empty": [], "small": [], "ok": 0}

    for root, dirs, files in os.walk(md_dir):
        for f in files:
            if not f.lower().endswith('.md'):
                continue

            fpath = os.path.join(root, f)
            try:
                size_bytes = os.path.getsize(fpath)
                size_kb = round(size_bytes / 1024, 2)
            except Exception:
                continue

            if size_bytes == 0:
                result["empty"].append((f, size_kb))
            elif size_kb < 1:
                result["small"].append((f, size_kb))
            else:
                result["ok"] += 1

    return result


# ── CLI ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Markdown 转换工具 V2")
    sub = parser.add_subparsers(dest="cmd")

    convert_p = sub.add_parser("convert", help="转换文献为 Markdown")
    convert_p.add_argument("path", help="文件或目录路径")
    convert_p.add_argument("--dry-run", action="store_true", help="仅预览")

    check_p = sub.add_parser("check", help="检测空 MD 文件")
    check_p.add_argument("dir", help="_Com_md 文件夹路径")

    args = parser.parse_args()

    if args.cmd == "convert":
        if os.path.isdir(args.path):
            res = convert_directory(args.path, dry_run=args.dry_run)
            print(f"\n## Converted ({len(res['converted'])})")
            for item in res["converted"]:
                print(f"  {item['source']} -> {item['output']}")
            print(f"\n## Skipped ({len(res['skipped'])})")
            for item in res["skipped"]:
                print(f"  {item['source']}: {item['note']}")
            if res["errors"]:
                print(f"\n## Errors ({len(res['errors'])})")
                for item in res["errors"]:
                    print(f"  {item['source']}: {item['note']}")
        else:
            res = convert_file(args.path, dry_run=args.dry_run)
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
