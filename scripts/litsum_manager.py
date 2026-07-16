#!/usr/bin/env python3
"""LitsSum.xlsx 管理工具 V2 — 创建、读取、更新、备份文献综述 Excel 表。

V2 新增:
  - 列 E: 关键词（最多 5 个，分号分隔）
  - 去重时关键词数量比较
  - APA 自检集成
  - 排序时处理 5 列数据

Sheet: LitsSum
Columns: A=出版年份 B=文件名 C=完整引用(APA) D=原始文件夹 E=关键词
"""

import os, sys, json, shutil, re
from datetime import datetime
from pathlib import Path

HAS_OPENPYXL = False
try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    pass

# ── 常量 ────────────────────────────────────────────────────────
SHEET_NAME = "LitsSum"
HEADERS = ["出版年份", "文件名", "完整引用(APA)", "原始文件夹", "关键词"]
COL_YEAR, COL_FNAME, COL_APA, COL_FOLDER, COL_KEYWORDS = 0, 1, 2, 3, 4
NUM_COLS = 5


def ensure_openpyxl():
    if not HAS_OPENPYXL:
        raise RuntimeError("需要安装 openpyxl: pip install openpyxl")


def _header_style(ws):
    """设置表头样式。"""
    header_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="1a3a5c", end_color="1a3a5c", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    for col in range(1, NUM_COLS + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border


def create_litsum(filepath: str) -> str:
    """创建新的 LitsSum.xlsx（含 5 列）。"""
    ensure_openpyxl()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    for i, h in enumerate(HEADERS, 1):
        ws.cell(row=1, column=i, value=h)
    _header_style(ws)
    # 设置列宽
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 70
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 50
    wb.save(filepath)
    return filepath


def backup_litsum(filepath: str) -> str:
    """备份 LitsSum.xlsx。"""
    if not os.path.exists(filepath):
        return ""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    parent = os.path.dirname(filepath)
    backup_name = f"LitsSum_backup_{ts}.xlsx"
    backup_path = os.path.join(parent, backup_name)
    shutil.copy2(filepath, backup_path)
    return backup_path


def read_litsum(filepath: str) -> list:
    """读取 LitsSum.xlsx 中所有数据行，返回 {header: value} 列表。"""
    ensure_openpyxl()
    if not os.path.exists(filepath):
        return []
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(c is not None for c in row):
            row_data = {}
            for i, h in enumerate(HEADERS):
                if i < len(row):
                    row_data[h] = str(row[i]) if row[i] is not None else ""
                else:
                    row_data[h] = ""
            rows.append(row_data)
    wb.close()
    return rows


def find_existing_by_filename(rows: list, filename: str) -> int:
    """在已有数据中按文件名查找行索引，未找到返回 -1。"""
    for i, row in enumerate(rows):
        if row[HEADERS[COL_FNAME]] == filename:
            return i
    return -1


def count_apa_fields(apa_str: str) -> int:
    """统计 APA 字符串中非空字段数（用于比较完整度）。

    检查: 作者+1, 年份+1, 期刊+1, 页码+1, DOI+1
    """
    if not apa_str:
        return 0
    count = 0
    # 作者: 开头有非括号内容
    if re.match(r'^[A-Za-z\u4e00-\u9fff]', apa_str):
        count += 1
    # 年份: (YYYY)
    if re.search(r'\(\d{4}\)', apa_str):
        count += 1
    # 期刊: *斜体* 或中文期刊名
    if '*' in apa_str or re.search(r'[\u4e00-\u9fff]{2,}', apa_str[20:] if len(apa_str) > 20 else ""):
        count += 1
    # 页码: 数字-数字
    if re.search(r'\d+\s*[-–]\s*\d+', apa_str):
        count += 1
    # DOI
    if 'doi.org' in apa_str.lower():
        count += 1
    return count


def count_keywords(keywords_str: str) -> int:
    """统计关键词数量。"""
    if not keywords_str or not keywords_str.strip():
        return 0
    return len([k for k in re.split(r'[;；]', keywords_str) if k.strip()])


def update_litsum(filepath: str, new_entries: list, dry_run: bool = False) -> dict:
    """更新 LitsSum.xlsx。

    new_entries: [
        {
            "filename": "Smith2020.pdf",
            "year": "2020",
            "apa": "Smith, J. D. (2020). ...",
            "folder": "Papers",
            "keywords": "Deep Learning; NLP; Transformer"
        },
        ...
    ]

    返回:
        {"appended": [...], "updated": [...], "skipped": [...], "errors": [...]}
    """
    ensure_openpyxl()

    result = {"appended": [], "updated": [], "skipped": [], "errors": []}

    if not os.path.exists(filepath):
        create_litsum(filepath)

    if dry_run:
        existing = read_litsum(filepath)
        for entry in new_entries:
            fname = entry.get("filename", "")
            idx = find_existing_by_filename(existing, fname)
            if idx < 0:
                result["appended"].append(fname)
            else:
                new_fields = count_apa_fields(entry.get("apa", ""))
                old_fields = count_apa_fields(existing[idx].get(HEADERS[COL_APA], ""))
                if new_fields > old_fields:
                    result["updated"].append(fname)
                else:
                    result["skipped"].append(fname)
        return result

    # 实际写入
    wb = openpyxl.load_workbook(filepath)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    existing = read_litsum(filepath)

    for entry in new_entries:
        try:
            fname = entry.get("filename", "")
            year = entry.get("year", "Unknown")
            apa = entry.get("apa", "")
            folder = entry.get("folder", "")
            keywords = entry.get("keywords", "")

            idx = find_existing_by_filename(existing, fname)
            if idx < 0:
                # 新文献：追加
                next_row = ws.max_row + 1
                ws.cell(row=next_row, column=1, value=year)
                ws.cell(row=next_row, column=2, value=fname)
                ws.cell(row=next_row, column=3, value=apa)
                ws.cell(row=next_row, column=4, value=folder)
                ws.cell(row=next_row, column=5, value=keywords)
                existing.append({
                    HEADERS[COL_YEAR]: year,
                    HEADERS[COL_FNAME]: fname,
                    HEADERS[COL_APA]: apa,
                    HEADERS[COL_FOLDER]: folder,
                    HEADERS[COL_KEYWORDS]: keywords,
                })
                result["appended"].append(fname)
            else:
                new_fields = count_apa_fields(apa)
                old_fields = count_apa_fields(existing[idx].get(HEADERS[COL_APA], ""))
                if new_fields > old_fields:
                    # 更新 APA 列（列 C），文件名列不动
                    ws.cell(row=idx + 2, column=3, value=apa)
                    existing[idx][HEADERS[COL_APA]] = apa
                    result["updated"].append(fname)
                else:
                    result["skipped"].append(fname)

                # 关键词列：保留数量更多的一方；数量相同则保留新提取的
                new_kw_count = count_keywords(keywords)
                old_kw_count = count_keywords(existing[idx].get(HEADERS[COL_KEYWORDS], ""))
                if new_kw_count >= old_kw_count and new_kw_count > 0:
                    ws.cell(row=idx + 2, column=5, value=keywords)
                    existing[idx][HEADERS[COL_KEYWORDS]] = keywords

        except Exception as e:
            result["errors"].append({"file": fname, "error": str(e)})

    wb.save(filepath)
    wb.close()
    return result


def sort_by_year(filepath: str) -> str:
    """对 LitsSum 表按出版年份升序排列（Unknown 排最后）。"""
    ensure_openpyxl()
    if not os.path.exists(filepath):
        return ""

    wb = openpyxl.load_workbook(filepath)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    # 读取所有数据行
    data_rows = []
    max_row = ws.max_row
    for row_idx in range(2, max_row + 1):
        vals = []
        for col in range(1, NUM_COLS + 1):
            vals.append(ws.cell(row=row_idx, column=col).value)
        data_rows.append(vals)

    # 排序：数字年份在前，"Unknown" 在后
    def sort_key(row):
        y = str(row[0]) if row[0] else ""
        if y == "Unknown" or y == "":
            return (1, "")
        try:
            return (0, int(y))
        except ValueError:
            return (1, y)

    data_rows.sort(key=sort_key)

    # 清除旧数据（保留表头）
    for row_idx in range(max_row, 1, -1):
        for col in range(1, NUM_COLS + 1):
            ws.cell(row=row_idx, column=col).value = None

    # 写回
    for i, row in enumerate(data_rows):
        for j, val in enumerate(row):
            ws.cell(row=i + 2, column=j + 1, value=val)

    wb.save(filepath)
    wb.close()
    return filepath


def build_year_map(filepath: str) -> dict:
    """从 LitsSum.xlsx 中提取 {文件名: 年份} 映射。"""
    rows = read_litsum(filepath)
    year_map = {}
    for row in rows:
        fname = row.get(HEADERS[COL_FNAME], "")
        year = row.get(HEADERS[COL_YEAR], "")
        if fname:
            # 同时映射带扩展名和不带扩展名的版本
            year_map[fname] = year
            base_name = os.path.splitext(fname)[0]
            year_map[base_name] = year
    return year_map


# ── CLI ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="LitsSum.xlsx 管理工具 V2")
    sub = parser.add_subparsers(dest="cmd")

    create_p = sub.add_parser("create", help="创建新的 LitsSum.xlsx")
    create_p.add_argument("path", help="输出路径")

    read_p = sub.add_parser("read", help="读取 LitsSum.xlsx")
    read_p.add_argument("path", help="文件路径")

    backup_p = sub.add_parser("backup", help="备份 LitsSum.xlsx")
    backup_p.add_argument("path", help="文件路径")

    update_p = sub.add_parser("update", help="更新 LitsSum.xlsx")
    update_p.add_argument("path", help="文件路径")
    update_p.add_argument("--entries", help="JSON 格式的新条目文件路径")
    update_p.add_argument("--dry-run", action="store_true", help="仅预览")

    sort_p = sub.add_parser("sort", help="排序 LitsSum.xlsx")
    sort_p.add_argument("path", help="文件路径")

    yearmap_p = sub.add_parser("yearmap", help="提取年份映射")
    yearmap_p.add_argument("path", help="文件路径")

    args = parser.parse_args()

    if args.cmd == "create":
        create_litsum(args.path)
        print(f"Created: {args.path}")

    elif args.cmd == "read":
        rows = read_litsum(args.path)
        for i, r in enumerate(rows):
            kw = r.get(HEADERS[COL_KEYWORDS], "")
            print(f"{i+2}: {r[HEADERS[COL_YEAR]]} | {r[HEADERS[COL_FNAME]]} | {r[HEADERS[COL_APA]][:60]}... | {r[HEADERS[COL_FOLDER]]} | {kw}")
        print(f"\nTotal: {len(rows)} records")

    elif args.cmd == "backup":
        bp = backup_litsum(args.path)
        print(f"Backup: {bp}")

    elif args.cmd == "update":
        entries = []
        if args.entries:
            with open(args.entries, 'r', encoding='utf-8') as f:
                entries = json.load(f)
        result = update_litsum(args.path, entries, dry_run=args.dry_run)
        for k, v in result.items():
            if v:
                print(f"{k}: {len(v)} - {v}")

    elif args.cmd == "sort":
        sort_by_year(args.path)
        print(f"Sorted: {args.path}")

    elif args.cmd == "yearmap":
        ym = build_year_map(args.path)
        print(json.dumps(ym, ensure_ascii=False, indent=2))
