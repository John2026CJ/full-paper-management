#!/usr/bin/env python3
"""文献文件操作工具 V2 — 收集、重命名（含冲突降级）、去重、年代/主题分类。

V2 新增:
  - 重命名冲突降级（姓氏 → 姓氏+名字全拼）
  - 按主题分类（classify --mode topic）
  - MD 转换文件分类（classify_md）

子命令:
  collect     - 递归收集文献到 Combined 文件夹（任务一）
  rename      - 标准化重命名，含冲突降级（任务二）
  dedup       - 重复文献检测（任务三）
  classify    - 按年代/主题分类（任务六）
  classify_md - 将 MD 文件分类到 _Com_md 文件夹（任务四 Step 4）
"""

import os, sys, json, re, shutil, hashlib
from pathlib import Path
from collections import defaultdict

# ── 文献支持的文件类型 ─────────────────────────────────────────
PAPER_EXTS = {'.pdf', '.docx', '.doc', '.txt', '.epub', '.mobi'}

# ── 排除的文件/文件夹 ──────────────────────────────────────────
EXCLUDE_FILES = {'README.md', 'README.txt', 'readme.md', 'readme.txt',
                 '.DS_Store', 'Thumbs.db', 'desktop.ini'}
EXCLUDE_DIRS = {'.git', '__pycache__', 'node_modules', '.venv', 'venv', 'Duplicates'}
EXCLUDE_EXT = {'.py', '.js', '.html', '.css', '.json', '.xml', '.yaml', '.yml',
               '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico',
               '.exe', '.dll', '.so', '.dylib', '.zip', '.rar', '.7z', '.tar', '.gz',
               '.xlsx', '.xls', '.pptx', '.ppt', '.csv', '.md'}


def is_paper(filepath: str) -> bool:
    """判断文件是否为文献文件。"""
    name = os.path.basename(filepath)
    if name in EXCLUDE_FILES or name.startswith('.') or name.startswith('~'):
        return False
    ext = Path(filepath).suffix.lower()
    return ext in PAPER_EXTS


def safe_filename(name: str) -> str:
    """清理文件名：移除特殊字符、空格。"""
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    name = re.sub(r'[\s_]+', '_', name)
    name = name.strip('_')
    return name


def ensure_unique(dest_dir: str, filename: str) -> str:
    """确保目标文件名唯一，冲突时添加 _数字 后缀。"""
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


# ── 收集文件 ────────────────────────────────────────────────────

def collect_files(source_dir: str, dest_dir: str, dry_run: bool = False) -> dict:
    """递归收集所有文献文件到目标文件夹。"""
    result = {"copied": [], "skipped": [], "errors": []}
    parent_name = os.path.basename(os.path.abspath(source_dir))

    os.makedirs(dest_dir, exist_ok=True)

    for root, dirs, files in os.walk(source_dir):
        dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS and not d.startswith('.')]

        for f in files:
            src = os.path.join(root, f)

            if not is_paper(src):
                result["skipped"].append(f"{src} (non-paper)")
                continue

            dest = ensure_unique(dest_dir, safe_filename(f))

            if not dry_run:
                try:
                    shutil.copy2(src, dest)
                except Exception as e:
                    result["errors"].append(f"{src}: {e}")
                    continue

            rel_src = os.path.relpath(src, source_dir)
            result["copied"].append((rel_src, os.path.relpath(dest, dest_dir)))

    return result


# ── 标准化重命名（V2: 含冲突降级） ──────────────────────────────

def rename_files(target_dir: str, dry_run: bool = False) -> dict:
    """对目标文件夹中的文献文件标准化重命名。

    V2 命名规则:
    - 基础格式: <第一作者姓氏><年份>[区分后缀].<扩展名>
    - 姓氏与年份间严禁空格、下划线等分隔符
    - 冲突降级: 同姓氏同年份不同文献 → 退化为"姓氏+名字全拼"
    - 非英文姓名: 中文姓名用拼音全名

    返回: {"renamed": [(old, new), ...], "skipped": [...], "degraded": [...]}
    """
    result = {"renamed": [], "skipped": [], "degraded": [], "conflicts": {}}

    sys.path.insert(0, os.path.dirname(__file__))
    from metadata_utils import extract_metadata, build_target_filename, PAPER_EXTS

    files = [f for f in os.listdir(target_dir)
             if os.path.isfile(os.path.join(target_dir, f))
             and Path(f).suffix.lower() in PAPER_EXTS]

    # 第一遍：提取所有文件的元数据
    all_metas = []
    file_meta_pairs = []
    for f in files:
        fpath = os.path.join(target_dir, f)
        meta = extract_metadata(fpath)
        file_meta_pairs.append((f, fpath, meta))
        all_metas.append(meta)

    # 检测同姓氏同年份冲突
    surname_year_map = defaultdict(list)
    for f, fpath, meta in file_meta_pairs:
        surname = meta.get("first_author_surname", "")
        year = meta.get("year", "")
        fullname = meta.get("first_author_fullname", "")
        if surname and year:
            key = f"{surname}_{year}"
            surname_year_map[key].append((f, fullname, meta))

    # 标记需要降级的文件
    degrade_set = set()
    for key, entries in surname_year_map.items():
        if len(entries) > 1:
            # 检查是否真的是不同文献（不同 fullname）
            fullnames = set(e[1] for e in entries if e[1])
            if len(fullnames) > 1:
                for e in entries:
                    degrade_set.add(e[0])

    # 第二遍：构建目标文件名
    existing_names = set(files)  # 当前目录已有文件名
    rename_plan = []

    for f, fpath, meta in file_meta_pairs:
        if not meta.get("first_author_surname") or not meta.get("year"):
            result["skipped"].append((f, "cannot extract author or year"))
            continue

        ext = Path(f).suffix.lower()
        target = build_target_filename(meta, ext, existing_names, all_metas)

        # 确保扩展名正确
        if not target.lower().endswith(ext):
            target = os.path.splitext(target)[0] + ext

        if target == f:
            result["skipped"].append((f, "already conforms"))
            continue

        # 冲突处理：不同文献同名
        if target in existing_names and target != f:
            base, e = os.path.splitext(target)
            for letter in 'abcdefghijklmnopqrstuvwxy':
                alt = f"{base}{letter}{e}"
                if alt not in existing_names:
                    target = alt
                    break
            else:
                target = os.path.basename(ensure_unique(target_dir, target))

        is_degraded = f in degrade_set
        rename_plan.append((f, target, is_degraded))
        existing_names.discard(f)
        existing_names.add(target)

    # 执行重命名
    for old, new, is_degraded in rename_plan:
        if not dry_run:
            try:
                old_path = os.path.join(target_dir, old)
                new_path = os.path.join(target_dir, new)
                os.rename(old_path, new_path)
            except Exception as e:
                result["skipped"].append((old, str(e)))
                continue

        result["renamed"].append((old, new))
        if is_degraded:
            result["degraded"].append((old, new))

    return result


# ── 重复检测 ───────────────────────────────────────────────────

def detect_duplicates(target_dir: str) -> dict:
    """检测目标文件夹中的重复文献。"""
    sys.path.insert(0, os.path.dirname(__file__))
    from metadata_utils import extract_metadata, file_md5, PAPER_EXTS

    papers = []
    for f in os.listdir(target_dir):
        fpath = os.path.join(target_dir, f)
        if os.path.isfile(fpath) and Path(f).suffix.lower() in PAPER_EXTS:
            meta = extract_metadata(fpath)
            papers.append({"filename": f, "filepath": fpath, "meta": meta})

    groups = defaultdict(list)

    for i, p1 in enumerate(papers):
        assigned = False
        for j, p2 in enumerate(papers):
            if j <= i:
                continue
            m1, m2 = p1["meta"], p2["meta"]

            # DOI 相同
            if m1["doi"] and m2["doi"] and m1["doi"] == m2["doi"]:
                groups[f"doi_{m1['doi']}"].extend([p1, p2])
                assigned = True
                continue

            # 作者+年份+标题前40字符
            if (m1["first_author_surname"] == m2["first_author_surname"]
                    and m1["year"] == m2["year"]
                    and m1["title"][:40] == m2["title"][:40]
                    and m1["title"]):
                key = f"author_{m1['first_author_surname']}_{m1['year']}"
                groups[key].extend([p1, p2])
                assigned = True
                continue

            # MD5 相同
            if m1["md5"] and m2["md5"] and m1["md5"] == m2["md5"]:
                groups[f"md5_{m1['md5']}"].extend([p1, p2])
                assigned = True
                continue

            # 文件名编辑距离 ≤ 3 且作者+年份一致
            f1 = os.path.splitext(p1["filename"])[0]
            f2 = os.path.splitext(p2["filename"])[0]
            if len(f1) > 5 and len(f2) > 5 and edit_distance(f1, f2) <= 3:
                if m1["first_author_surname"] == m2["first_author_surname"] and m1["year"] == m2["year"]:
                    groups[f"fname_{f1}"].extend([p1, p2])

        if not assigned and p1 not in [x for g in groups.values() for x in g]:
            groups[f"unique_{p1['filename']}"].append(p1)

    result_groups = []
    seen = set()
    for key, group in groups.items():
        unique_group = []
        for p in group:
            if p["filename"] not in seen:
                unique_group.append(p)
                seen.add(p["filename"])

        if len(unique_group) <= 1:
            continue

        best = select_best(unique_group)
        dups = [p for p in unique_group if p["filename"] != best["filename"]]
        result_groups.append({
            "kept": best["filename"],
            "duplicates": [d["filename"] for d in dups],
            "reason": f"score: {best['meta']['completeness_score']}/10, pages: {best['meta']['page_count']}"
        })

    return {"groups": result_groups}


def select_best(papers: list) -> dict:
    """从一组重复论文中选择最优文件。

    优先级: 信息完整度 → 文字清晰度(页数) → 文件体积 → 格式
    """
    def sort_key(p):
        m = p["meta"]
        score = m.get("completeness_score", 0)
        pages = m.get("page_count", 0)
        size = os.path.getsize(p["filepath"]) if os.path.exists(p["filepath"]) else 0
        ext = Path(p["filename"]).suffix.lower()
        ext_rank = {'.pdf': 0, '.docx': 1, '.doc': 2, '.txt': 3, '.epub': 4, '.mobi': 5}
        return (score, pages, -size, ext_rank.get(ext, 9))

    papers.sort(key=sort_key, reverse=True)
    return papers[0]


def move_duplicates(target_dir: str, groups: list, dry_run: bool = False) -> dict:
    """将冗余文件移动到 Duplicates 文件夹。"""
    dup_dir = os.path.join(target_dir, "Duplicates")
    result = {"moved": [], "errors": []}

    if groups:
        os.makedirs(dup_dir, exist_ok=True)

    for group in groups:
        for dup in group["duplicates"]:
            src = os.path.join(target_dir, dup)
            if not os.path.exists(src):
                continue
            dst = ensure_unique(dup_dir, dup)
            if not dry_run:
                try:
                    shutil.move(src, dst)
                except Exception as e:
                    result["errors"].append(f"{dup}: {e}")
                    continue
            result["moved"].append((dup, os.path.relpath(dst, target_dir)))

    return result


def edit_distance(s1: str, s2: str) -> int:
    """计算两个字符串的编辑距离。"""
    if len(s1) < len(s2):
        return edit_distance(s2, s1)
    if len(s2) == 0:
        return len(s1)
    prev = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        cur = [i + 1]
        for j, c2 in enumerate(s2):
            cur.append(min(cur[j] + 1, prev[j + 1] + 1, prev[j] + (c1 != c2)))
        prev = cur
    return prev[-1]


# ── 年代分类 ──────────────────────────────────────────────────

YEAR_RANGES = [
    (0, 1989, "1900-1989"),
    (1990, 1999, "1990-1999"),
    (2000, 2009, "2000-2009"),
    (2010, 2019, "2010-2019"),
    (2020, 2029, "2020-2029"),
]


def classify_by_year(target_dir: str, year_map: dict = None, dry_run: bool = False) -> dict:
    """按年代将文献文件分类到子文件夹。"""
    sys.path.insert(0, os.path.dirname(__file__))
    from metadata_utils import extract_metadata, PAPER_EXTS

    result = {"classified": {}, "errors": []}

    for _, _, folder in YEAR_RANGES:
        result["classified"][folder] = []
    result["classified"]["UnknownYear"] = []

    files = [f for f in os.listdir(target_dir)
             if os.path.isfile(os.path.join(target_dir, f))
             and Path(f).suffix.lower() in PAPER_EXTS
             and f != "Duplicates"]

    if not year_map:
        year_map = {}

    for f in files:
        fpath = os.path.join(target_dir, f)

        year_str = year_map.get(f, "") or year_map.get(os.path.splitext(f)[0], "")
        if not year_str:
            meta = extract_metadata(fpath)
            year_str = meta.get("year", "")

        year = None
        try:
            year = int(year_str)
        except (ValueError, TypeError):
            pass

        target_folder = "UnknownYear"
        if year:
            for lo, hi, folder in YEAR_RANGES:
                if lo <= year <= hi:
                    target_folder = folder
                    break

        folder_path = os.path.join(target_dir, target_folder)
        os.makedirs(folder_path, exist_ok=True)
        dst = ensure_unique(folder_path, f)

        if not dry_run:
            try:
                shutil.move(fpath, dst)
            except Exception as e:
                result["errors"].append(f"{f}: {e}")
                continue

        result["classified"][target_folder].append(f)

    return result


# ── 主题分类（V2 新增） ────────────────────────────────────────

def classify_by_topic(target_dir: str, topics: list, topic_map: dict = None,
                       dry_run: bool = False) -> dict:
    """按主题将文献文件分类到子文件夹。

    Args:
        target_dir: Combined 文件夹路径
        topics: 主题列表，如 ["机器学习", "自然语言处理", "计算机视觉"]
        topic_map: {filename: topic} 可选的预设映射
        dry_run: 仅预览

    Returns:
        {"classified": {topic_folder: [files]}, "unclassified": [files], "errors": [...]}
    """
    sys.path.insert(0, os.path.dirname(__file__))
    from metadata_utils import extract_metadata, PAPER_EXTS

    result = {"classified": {}, "unclassified": [], "errors": []}

    # 初始化主题文件夹
    for i, topic in enumerate(topics):
        folder_name = safe_filename(f"{i+1:02d}_{topic}")
        result["classified"][folder_name] = []
    result["classified"]["UnClassified"] = []

    files = [f for f in os.listdir(target_dir)
             if os.path.isfile(os.path.join(target_dir, f))
             and Path(f).suffix.lower() in PAPER_EXTS]

    if not topic_map:
        topic_map = {}

    for f in files:
        fpath = os.path.join(target_dir, f)

        # 检查预设映射
        assigned_topic = topic_map.get(f, "") or topic_map.get(os.path.splitext(f)[0], "")

        if not assigned_topic:
            # 从元数据推断主题
            meta = extract_metadata(fpath)
            text = meta.get("text_preview", "") + " " + meta.get("title", "")
            keywords = meta.get("keywords", [])

            # 简单关键词匹配
            best_match = ""
            best_score = 0
            for topic in topics:
                score = 0
                topic_lower = topic.lower()
                if topic_lower in text.lower():
                    score += 3
                for kw in keywords:
                    if topic_lower in kw.lower() or kw.lower() in topic_lower:
                        score += 2
                if score > best_score:
                    best_score = score
                    best_match = topic

            if best_score > 0:
                assigned_topic = best_match

        # 确定目标文件夹
        if assigned_topic and assigned_topic in topics:
            idx = topics.index(assigned_topic)
            folder_name = safe_filename(f"{idx+1:02d}_{assigned_topic}")
        else:
            folder_name = "UnClassified"

        folder_path = os.path.join(target_dir, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        dst = ensure_unique(folder_path, f)

        if not dry_run:
            try:
                shutil.move(fpath, dst)
            except Exception as e:
                result["errors"].append(f"{f}: {e}")
                continue

        result["classified"][folder_name].append(f)

    return result


# ── MD 文件分类（任务四 Step 4） ───────────────────────────────

def classify_md_files(source_dir: str, md_dir: str, year_map: dict = None,
                       dry_run: bool = False) -> dict:
    """将 Markdown 文件按年份分类到 _Com_md 文件夹。

    Args:
        source_dir: Combined 文件夹（MD 文件所在）
        md_dir: _Com_md 目标文件夹
        year_map: {filename: year} 年份映射
        dry_run: 仅预览
    """
    result = {"classified": {}, "errors": []}

    for _, _, folder in YEAR_RANGES:
        result["classified"][folder] = []
    result["classified"]["UnknownYear"] = []

    # 收集所有 .md 文件（排除 Duplicates）
    md_files = []
    for root, dirs, files in os.walk(source_dir):
        dirs[:] = [d for d in dirs if d != 'Duplicates']
        for f in files:
            if f.lower().endswith('.md'):
                md_files.append(os.path.join(root, f))

    if not year_map:
        year_map = {}

    os.makedirs(md_dir, exist_ok=True)

    for fpath in md_files:
        fname = os.path.basename(fpath)
        base_name = os.path.splitext(fname)[0]

        # 获取年份
        year_str = year_map.get(base_name, "") or year_map.get(fname, "")
        if not year_str:
            # 尝试从文件名提取年份
            m = re.search(r'(\d{4})', base_name)
            if m:
                year_str = m.group(1)

        year = None
        try:
            year = int(year_str)
        except (ValueError, TypeError):
            pass

        target_folder = "UnknownYear"
        if year:
            for lo, hi, folder in YEAR_RANGES:
                if lo <= year <= hi:
                    target_folder = folder
                    break

        folder_path = os.path.join(md_dir, target_folder)
        os.makedirs(folder_path, exist_ok=True)
        dst = ensure_unique(folder_path, fname)

        if not dry_run:
            try:
                shutil.move(fpath, dst)
            except Exception as e:
                result["errors"].append(f"{fname}: {e}")
                continue

        result["classified"][target_folder].append(fname)

    return result


# ── 空文件检测（任务四 Step 5） ────────────────────────────────

def check_empty_md(md_dir: str) -> dict:
    """检测 _Com_md 文件夹中的空文件和小文件。

    Returns:
        {"empty": [(filename, size_kb)], "small": [(filename, size_kb)], "ok": count}
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
    parser = argparse.ArgumentParser(description="文献文件操作工具 V2")
    sub = parser.add_subparsers(dest="cmd")

    collect_p = sub.add_parser("collect", help="收集文献到 Combined 文件夹")
    collect_p.add_argument("source", help="源目录")
    collect_p.add_argument("--dest", help="目标目录")
    collect_p.add_argument("--dry-run", action="store_true")

    rename_p = sub.add_parser("rename", help="标准化重命名（含冲突降级）")
    rename_p.add_argument("dir", help="目标目录")
    rename_p.add_argument("--dry-run", action="store_true")

    dedup_p = sub.add_parser("dedup", help="检测并整合重复文献")
    dedup_p.add_argument("dir", help="目标目录")
    dedup_p.add_argument("--move", action="store_true")
    dedup_p.add_argument("--dry-run", action="store_true")

    classify_p = sub.add_parser("classify", help="按年代/主题分类文献")
    classify_p.add_argument("dir", help="目标目录")
    classify_p.add_argument("--mode", choices=["year", "topic"], default="year")
    classify_p.add_argument("--year-map", help="JSON 年份映射文件")
    classify_p.add_argument("--topics", help="JSON 主题列表文件")
    classify_p.add_argument("--topic-map", help="JSON 主题映射文件")
    classify_p.add_argument("--dry-run", action="store_true")

    classify_md_p = sub.add_parser("classify_md", help="将 MD 文件分类到 _Com_md")
    classify_md_p.add_argument("source", help="Combined 文件夹")
    classify_md_p.add_argument("--md-dir", help="_Com_md 目标文件夹")
    classify_md_p.add_argument("--year-map", help="JSON 年份映射文件")
    classify_md_p.add_argument("--dry-run", action="store_true")

    check_md_p = sub.add_parser("check_md", help="检测空 MD 文件")
    check_md_p.add_argument("dir", help="_Com_md 文件夹")

    args = parser.parse_args()

    if args.cmd == "collect":
        parent = os.path.basename(os.path.abspath(args.source))
        dest = args.dest or f"{parent}_Combined"
        if not os.path.isabs(dest):
            dest = os.path.join(os.path.dirname(args.source), dest)

        res = collect_files(args.source, dest, dry_run=args.dry_run)
        for k, v in res.items():
            if v:
                print(f"\n## {k} ({len(v)})")
                for item in v:
                    print(f"  {item}")

    elif args.cmd == "rename":
        res = rename_files(args.dir, dry_run=args.dry_run)
        for k, v in res.items():
            if v:
                print(f"\n## {k} ({len(v)})")
                if k == "renamed":
                    for old, new in v:
                        print(f"  {old} -> {new}")
                elif k == "degraded":
                    print("  (conflict degraded to fullname)")
                    for old, new in v:
                        print(f"  {old} -> {new}")
                elif k == "skipped":
                    for fname, reason in v:
                        print(f"  {fname}: {reason}")

    elif args.cmd == "dedup":
        res = detect_duplicates(args.dir)
        print(json.dumps(res, ensure_ascii=False, indent=2))
        if args.move:
            move_res = move_duplicates(args.dir, res["groups"], dry_run=args.dry_run)
            for k, v in move_res.items():
                if v:
                    print(f"\n## {k} ({len(v)})")
                    for item in v:
                        print(f"  {item}")

    elif args.cmd == "classify":
        if args.mode == "year":
            ym = {}
            if args.year_map and os.path.exists(args.year_map):
                with open(args.year_map, 'r', encoding='utf-8') as f:
                    ym = json.load(f)
            res = classify_by_year(args.dir, year_map=ym, dry_run=args.dry_run)
        else:
            topics = []
            if args.topics and os.path.exists(args.topics):
                with open(args.topics, 'r', encoding='utf-8') as f:
                    topics = json.load(f)
            tm = {}
            if args.topic_map and os.path.exists(args.topic_map):
                with open(args.topic_map, 'r', encoding='utf-8') as f:
                    tm = json.load(f)
            res = classify_by_topic(args.dir, topics, topic_map=tm, dry_run=args.dry_run)

        total = 0
        for folder, files in res.get("classified", {}).items():
            print(f"  {folder}: {len(files)}")
            total += len(files)
        print(f"  Total: {total}")

    elif args.cmd == "classify_md":
        parent = os.path.basename(os.path.abspath(args.source))
        parent = parent.replace("_Combined", "")
        md_dir = args.md_dir or f"{parent}_Com_md"
        if not os.path.isabs(md_dir):
            md_dir = os.path.join(os.path.dirname(args.source), md_dir)

        ym = {}
        if args.year_map and os.path.exists(args.year_map):
            with open(args.year_map, 'r', encoding='utf-8') as f:
                ym = json.load(f)

        res = classify_md_files(args.source, md_dir, year_map=ym, dry_run=args.dry_run)
        total = 0
        for folder, files in res["classified"].items():
            print(f"  {folder}: {len(files)}")
            total += len(files)
        print(f"  Total: {total}")

    elif args.cmd == "check_md":
        res = check_empty_md(args.dir)
        print(f"OK: {res['ok']}")
        if res["empty"]:
            print(f"\n## EMPTY ({len(res['empty'])})")
            for fname, size in res["empty"]:
                print(f"  {fname}: {size} KB")
        if res["small"]:
            print(f"\n## SMALL ({len(res['small'])})")
            for fname, size in res["small"]:
                print(f"  {fname}: {size} KB")
