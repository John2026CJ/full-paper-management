# Full Paper Management V2 — 全流程文献管理技能
# Full Paper Management V2 — Complete Literature Management Skill

> 对文件夹中的学术文献执行七项标准化处理：整合、重命名、去重、Markdown 转换、LitsSum 文献表生成（含关键词）、分类整理、增量入库。
> Perform seven standardized tasks on academic papers in any folder: collect, rename, deduplicate, convert to Markdown, generate LitsSum literature spreadsheet (with keywords), classify, and incremental ingestion.

## 核心特性

- **七项任务，完整流水线** — 从散乱文献到结构化知识库，一步到位
- **先预览后执行** — 每步操作均支持 `--dry-run`，确认无误再执行
- **不删除任何原始文件** — 所有操作为复制/重命名/移动，绝不删除
- **幂等可重复** — 多次执行不会产生重复数据
- **自动备份** — 操作 LitsSum.xlsx 前自动创建时间戳备份

## Key Features

- **Seven tasks, complete pipeline** — From scattered papers to a structured knowledge base in one pass
- **Preview before execution** — Every step supports `--dry-run` for safe preview before committing
- **Never deletes original files** — All operations are copy/rename/move, never delete
- **Idempotent** — Repeated executions produce no duplicate data
- **Auto-backup** — LitsSum.xlsx is automatically backed up with timestamp before any modification
- **What's new in V2** — Markdown conversion, keyword extraction, field-by-field APA validation, topic-based classification, empty file detection

---

## 快速开始
## Quick Start


### 1. 安装依赖
### 1. Install Dependencies

```bash
pip install PyPDF2 python-docx openpyxl requests
```

### 2. 触发技能
### 2. Trigger the Skill

在 WorkBuddy 中输入以下任意短语即可触发：
Use any of the following phrases in WorkBuddy:

- `文献管理` / `整理文献` / `全流程文献管理`
- `文献重命名` / `去重` / `生成文献表`
- `分类文献` / `新增文献入库`
- `process papers` / `manage literature` / `organize papers`

### 3. 首次运行
### 3. First Run

```
Task 1 -> Task 2 -> Task 3 -> Task 4 -> Task 5 -> Task 6
整合     重命名     去重      MD转换    LitsSum    分类
Collect   Rename    Dedup     MD Convert LitsSum   Classify
```

### 4. 增量运行（有新文献加入时）
### 4. Incremental Run (when new papers are added)

```
Task 7: Step 0 -> 1 -> 2 -> 3 -> 4 -> 5 -> 6 -> 7
        发现     重命名   预整合   识别    去重   更新表  分类   汇总
        Discover  Rename  Pre-     Identify Dedup Update Classify Summarize
                 collect  merge
```

---

## 七项任务详解
## The Seven Tasks

### 任务一：文件整合 (Collect)


将分散在各级子文件夹中的文献汇聚到统一目录 `XXX_Combined`。


```bash
python scripts/file_ops.py collect 源目录 [--dest 目标目录] [--dry-run]
```

- 递归扫描所有子文件夹
- 同名冲突自动添加 `_数字` 后缀

### Task 1: Collect Papers

Gather papers scattered across subfolders into a unified directory `XXX_Combined`.

```bash
python scripts/file_ops.py collect <source_dir> [--dest <dest_dir>] [--dry-run]
```

- Recursively scans all subfolders
- Auto-appends `_number` suffix on name conflicts


### 任务二：文件重命名 (Rename)


将文件名标准化为 `第一作者姓氏年份[后缀].扩展名` 格式。

**命名规则：**
- 基础格式：`Smith2020.pdf`
- 冲突降级：同姓同年不同人 → `ZhangYuan2020.pdf`、`ZhangFei2020.pdf`
- 非英文姓名：冲突时用拼音全名 `WangXiaoming2023.pdf`

**信息提取优先级：** PDF 元数据 > 内容解析 > 文件名推断 > DOI 解析

```bash
python scripts/file_ops.py rename 目标目录 [--dry-run]
```

### Task 2: Rename Papers

Standardize filenames to `FirstAuthorLastNameYear[suffix].extension` format.

**Naming Rules:**
- Basic format: `Smith2020.pdf`
- Conflict fallback: Same last name + same year but different authors → `ZhangYuan2020.pdf`, `ZhangFei2020.pdf`
- Non-English names: Uses pinyin full name on conflict → `WangXiaoming2023.pdf`

**Information Extraction Priority:** PDF metadata > Content parsing > Filename inference > DOI lookup

```bash
python scripts/file_ops.py rename <target_dir> [--dry-run]
```



### 任务三：重复文献判别与去重 (Dedup)


检测重复文献，每组保留最优版本，冗余文件移入 `Duplicates` 文件夹。

**重复判定标准（满足任一）：**
1. DOI 完全相同
2. 作者+年份+标题前 40 字符一致
3. MD5 哈希相同
4. 文件名编辑距离 ≤ 3 且作者+年份一致

**最优版本选择：** 信息完整性打分（满分 10）> 页数 > 文件大小 > 格式优先级

```bash
python scripts/file_ops.py dedup 目标目录 [--move] [--dry-run]
```

### Task 3: Deduplication

Detect duplicate papers, keep the best version per group, move redundants to `Duplicates` folder.

**Duplicate Criteria (any one):**
1. Identical DOI
2. Same author + year + first 40 chars of title
3. Identical MD5 hash
4. Filename edit distance ≤ 3 AND same author + year

**Best Version Selection:** Information completeness score (max 10) > Page count > File size > Format priority

```bash
python scripts/file_ops.py dedup <target_dir> [--move] [--dry-run]
```



### 任务四：文献格式转换为 Markdown (MD Convert) — V2 新增

将 Combined 文件夹中的文献转换为 Markdown，分类到 `_Com_md` 文件夹，并检测空文件。

```bash
# 转换
python scripts/md_converter.py convert Combined目录 [--dry-run]

# 分类到 _Com_md
python scripts/file_ops.py classify_md Combined目录 [--md-dir Com_md目录] [--year-map year_map.json] [--dry-run]

# 空文件检测
python scripts/md_converter.py check Com_md目录
```

**空文件判定：**

| 文件大小 | 判定 | 处理 |
|----------|------|------|
| 0 KB | 转换失败 | 提醒用户 |
| < 1 KB | 高度疑似失败 | 提醒用户 |
| ≥ 1 KB | 正常 | 不额外处理 |


### Task 4: Convert to Markdown (V2 New)

Convert papers in the Combined folder to Markdown, classify into `_Com_md` folder, and detect empty files.

```bash
# Convert
python scripts/md_converter.py convert <combined_dir> [--dry-run]

# Classify into _Com_md
python scripts/file_ops.py classify_md <combined_dir> [--md-dir <com_md_dir>] [--year-map year_map.json] [--dry-run]

# Empty file detection
python scripts/md_converter.py check <com_md_dir>
```

**Empty File Thresholds:**

| File Size | Verdict | Action |
|-----------|---------|--------|
| 0 KB | Conversion failed | Alert user |
| < 1 KB | Likely failed | Alert user |
| ≥ 1 KB | Normal | No action |




### 任务五：生成/更新 LitsSum 文献表 (LitsSum)

在 Combined 文件夹中创建或更新 `LitsSum.xlsx`。

**表结构（V2 新增列 E）：**

| 列 | 表头 | 说明 |
|----|------|------|
| A | 出版年份 | 四位数字，无法提取则 Unknown |
| B | 文件名 | 去除扩展名后的纯文件名 |
| C | 完整引用(APA) | APA 格式文献引用 |
| D | 原始文件夹 | 文件最初来源文件夹名 |
| E | 关键词 | 最多 5 个，分号分隔 — **V2 新增** |

**关键词提取来源：** DOI 解析 > 文件内容词频 > 标题术语

**APA 格式自检（每条生成后强制执行 8 项检查）：**
1. 作者格式（姓, 首字母缩写.）
2. 年份括号+句点
3. 标题仅首字母大写
4. 期刊名斜体
5. 卷号斜体、期号不斜体
6. 页码格式
7. DOI 以 `https://doi.org/` 开头
8. 无多余标点

```bash
python scripts/metadata_utils.py --batch --apa 目标目录 -o metadata_batch.json
python scripts/litsum_manager.py update LitsSum路径 --entries metadata_batch.json [--dry-run]
python scripts/litsum_manager.py sort LitsSum路径
```

### Task 5: Generate / Update LitsSum Spreadsheet

Create or update `LitsSum.xlsx` in the Combined folder.

**Table Structure (V2 adds Column E):**

| Column | Header | Description |
|--------|--------|-------------|
| A | Year | Four-digit year; "Unknown" if unextractable |
| B | Filename | Filename without extension |
| C | Full Citation (APA) | APA-formatted citation |
| D | Source Folder | Original folder name |
| E | Keywords | Up to 5 keywords, semicolon-separated — **V2 New** |

**Keyword Extraction Sources:** DOI lookup > File content word frequency > Title terms

**APA Self-Check (8 mandatory checks after each citation):**
1. Author format (Last, F. M.)
2. Year in parentheses with period
3. Title in sentence case
4. Journal name in italics
5. Volume in italics, issue not italicized
6. Page range format
7. DOI starts with `https://doi.org/`
8. No extra punctuation

```bash
python scripts/metadata_utils.py --batch --apa <target_dir> -o metadata_batch.json
python scripts/litsum_manager.py update <litsum_path> --entries metadata_batch.json [--dry-run]
python scripts/litsum_manager.py sort <litsum_path>
```


### 任务六：文献分类整理（按年份 / 按主题）— V2 增强主题分类

**模式 A：按年份分类**（默认）

| 年份范围 | 文件夹名 |
|----------|----------|
| ≤ 1989 | 1900-1989 |
| 1990-1999 | 1990-1999 |
| 2000-2009 | 2000-2009 |
| 2010-2019 | 2010-2019 |
| 2020-2029 | 2020-2029 |
| 无法识别 | UnknownYear |

**模式 B：按主题分类** — V2 新增

- 主题来源优先级：用户提供 > LitsSum 推断 > 文件名推断 > 请用户提供
- 文件夹命名：`01_机器学习`、`02_自然语言处理`
- 无法归类放入 `UnClassified`

```bash
# 按年份
python scripts/file_ops.py classify 目标目录 --mode year --year-map year_map.json [--dry-run]

# 按主题
python scripts/file_ops.py classify 目标目录 --mode topic --topics topics.json [--topic-map topic_map.json] [--dry-run]
```

### Task 6: Classify Papers (By Year / By Topic) — V2 Enhanced

**Mode A: By Year (default)**

| Year Range | Folder Name |
|------------|-------------|
| ≤ 1989 | 1900-1989 |
| 1990-1999 | 1990-1999 |
| 2000-2009 | 2000-2009 |
| 2010-2019 | 2010-2019 |
| 2020-2029 | 2020-2029 |
| Unknown | UnknownYear |

**Mode B: By Topic** — V2 New

- Topic source priority: User-provided > LitsSum inference > Filename inference > Ask user
- Folder naming: `01_Machine_Learning`, `02_NLP`
- Unclassifiable papers go to `UnClassified`

```bash
# By year
python scripts/file_ops.py classify <target_dir> --mode year --year-map year_map.json [--dry-run]

# By topic
python scripts/file_ops.py classify <target_dir> --mode topic --topics topics.json [--topic-map topic_map.json] [--dry-run]
```



### 任务七：新增文献增量处理 (Incremental)

当用户添加了新的文献文件夹后，执行增量入库，不影响已有数据。

```
Step 0: 发现新增源文件夹
Step 1: 原地重命名（含冲突降级）
Step 2: 复制到 Combined（保留源文件）
Step 3: 识别新增文献（对比 LitsSum 列 B）
Step 4: 重复检测（新增间 + 新增与库存间）
Step 5: 更新 LitsSum（列 D 填源文件夹，列 E 填关键词）
Step 6: 分类到年代/主题文件夹
Step 7: 排序 + 汇总统计表
```

### Task 7: Incremental Ingestion

Process newly added paper folders without affecting existing data.

```
Step 0: Discover new source folders
Step 1: Rename in-place (with conflict fallback)
Step 2: Copy to Combined (preserve source files)
Step 3: Identify new papers (compare against LitsSum Column B)
Step 4: Deduplication (new vs. new + new vs. existing)
Step 5: Update LitsSum (Column D = source folder, Column E = keywords)
Step 6: Classify into year/topic folders
Step 7: Sort + summary statistics table
```


---

## 脚本工具一览
## Script Reference


| 脚本 | 功能 | 主要命令 |
|------|------|----------|
| `scripts/metadata_utils.py` | 元数据提取、DOI 解析、APA 构建、关键词提取、冲突降级 | `--batch --apa` |
| `scripts/file_ops.py` | 文件收集、重命名、去重、年代/主题分类、MD 分类、空文件检测 | `collect` / `rename` / `dedup` / `classify` / `classify_md` |
| `scripts/litsum_manager.py` | LitsSum.xlsx 创建、读取、更新、备份、排序、年份映射 | `backup` / `update` / `sort` / `yearmap` |
| `scripts/md_converter.py` | PDF/DOCX/TXT 转 Markdown、空文件检测 | `convert` / `check` |

| Script | Function | Key Commands |
|--------|----------|--------------|
| `scripts/metadata_utils.py` | Metadata extraction, DOI parsing, APA construction, keyword extraction, conflict fallback | `--batch --apa` |
| `scripts/file_ops.py` | File collection, renaming, dedup, year/topic classification, MD classification, empty file detection | `collect` / `rename` / `dedup` / `classify` / `classify_md` |
| `scripts/litsum_manager.py` | LitsSum.xlsx create, read, update, backup, sort, year mapping | `backup` / `update` / `sort` / `yearmap` |
| `scripts/md_converter.py` | PDF/DOCX/TXT to Markdown conversion, empty file detection | `convert` / `check` |



---

## 支持的文献格式
## Supported File Formats

| 格式 | 扩展名 |
|------|--------|
| PDF | .pdf |
| Word | .docx, .doc |
| 纯文本 | .txt |
| 电子书 | .epub, .mobi |

| Format | Extensions |
|--------|------------|
| PDF | .pdf |
| Word | .docx, .doc |
| Plain Text | .txt |
| E-book | .epub, .mobi |



**排除项：** 隐藏文件、系统文件、代码文件、图片、README、Excel、PPT 等非文献文件。
**Excluded:** Hidden files, system files, code files, images, README, Excel, PPT, and other non-literature files.


---

## 全局规则
## Global Rules

1. **先预览后执行** — 每项任务先 `--dry-run` 预览，确认后再正式执行
2. **不删除原始文件** — 所有操作为复制/重命名/移动
3. **五状态日志** — 成功、跳过、冲突已处理、已移动、已更新
4. **幂等性** — 支持重复执行，不产生重复数据
5. **自动备份** — 操作 LitsSum.xlsx 前自动备份为 `LitsSum_backup_YYYYMMDD_HHmmss.xlsx`
6. **错误不中断** — 单个文件失败记录原因后跳过，不影响整体流程

1. **Preview before execution** — Always run with `--dry-run` first, confirm, then execute
2. **Never delete originals** — All operations are copy/rename/move
3. **Five status log** — Success, Skipped, Conflict resolved, Moved, Updated
4. **Idempotency** — Repeated execution produces no duplicate data
5. **Auto-backup** — LitsSum.xlsx backed up as `LitsSum_backup_YYYYMMDD_HHmmss.xlsx` before modification
6. **Error resilience** — Single file failure is logged and skipped; overall pipeline continues

---

## V2 vs V1 变更摘要

| 变更项 | V1 | V2 |
|--------|----|----|
| 任务数量 | 6 项 | **7 项**（新增 MD 转换） |
| LitsSum 列 | A-D（4列） | A-E（**新增关键词列**） |
| APA 格式 | 基础规则 | **逐字段规范 + 8 项强制自检** |
| 分类模式 | 仅按年份 | 按年份 / **按主题** |
| 增量处理 | 基本流程 | **多源文件夹 + 重命名先于整合** |
| 空文件检测 | 无 | **0 KB / < 1 KB / ≥ 1 KB 三级判定** |
| 冲突降级 | 无 | **同姓同年自动降级为全名** |

## V2 vs V1 Change Summary

| Change | V1 | V2 |
|--------|----|----|
| Task count | 6 tasks | **7 tasks** (added MD conversion) |
| LitsSum columns | A-D (4 columns) | A-E (**added keywords column**) |
| APA format | Basic rules | **Field-by-field spec + 8 mandatory self-checks** |
| Classification | Year only | Year / **Topic** |
| Incremental | Basic flow | **Multi-source folders + rename-before-merge** |
| Empty file detection | None | **0 KB / < 1 KB / ≥ 1 KB three-tier detection** |
| Name conflict fallback | None | **Same last name + same year auto-fallback to full name** |

---

## 目录结构示例

```
Papers/                          # 用户文献根目录
  ├─ subfolder1/                 # 散乱子文件夹
  ├─ subfolder2/
  ├─ Papers_Combined/            # 任务一：整合后
  │   ├─ Smith2020.pdf
  │   ├─ ZhangYuan2023.pdf
  │   ├─ 2020-2029/              # 任务六：分类后
  │   │   ├─ Smith2020.pdf
  │   │   └─ ZhangYuan2023.pdf
  │   ├─ 2010-2019/
  │   ├─ LitsSum.xlsx            # 任务五：文献表
  │   ├─ Duplicates/             # 任务三：冗余文件
  │   └─ Papers_Com_md/          # 任务四：Markdown 版本
  │       ├─ 2020-2029/
  │       │   ├─ Smith2020.md
  │       │   └─ ZhangYuan2023.md
  │       └─ UnknownYear/
  └─ new_papers/                 # 新增文献（任务七增量处理）
```

## Directory Structure Example

```
Papers/                          # User's paper root directory
  ├─ subfolder1/                 # Scattered subfolders
  ├─ subfolder2/
  ├─ Papers_Combined/            # Task 1: After collection
  │   ├─ Smith2020.pdf
  │   ├─ ZhangYuan2023.pdf
  │   ├─ 2020-2029/              # Task 6: After classification
  │   │   ├─ Smith2020.pdf
  │   │   └─ ZhangYuan2023.pdf
  │   ├─ 2010-2019/
  │   ├─ LitsSum.xlsx            # Task 5: Literature spreadsheet
  │   ├─ Duplicates/             # Task 3: Redundant files
  │   └─ Papers_Com_md/          # Task 4: Markdown versions
  │       ├─ 2020-2029/
  │       │   ├─ Smith2020.md
  │       │   └─ ZhangYuan2023.md
  │       └─ UnknownYear/
  └─ new_papers/                 # New papers (Task 7 incremental)
```

---

## 技术依赖
## Dependencies

| 依赖 | 用途 |
|------|------|
| PyPDF2 | PDF 元数据提取与文本解析 |
| python-docx | DOCX 文件文本提取 |
| openpyxl | LitsSum.xlsx 读写 |
| requests | Crossref DOI API 调用 |



| Dependency | Purpose |
|------------|---------|
| PyPDF2 | PDF metadata extraction and text parsing |
| python-docx | DOCX text extraction |
| openpyxl | LitsSum.xlsx read/write |
| requests | Crossref DOI API calls |

---

## 许可证
## License

本技能由 WorkBuddy 技能创建器生成，可自由使用和修改。
Generated by WorkBuddy Skill Creator. Free to use and modify.

