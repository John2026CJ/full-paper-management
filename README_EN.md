# full-paper-management V3.1

Complete literature management skill — executes seven standardized tasks on literature files within a folder, achieving fully automated library management from file consolidation to Markdown conversion.

## V3.1 Changelog

| # | Change | Detail |
|---|--------|--------|
| 1 | **Topic Word Cap** | Max 10 topic words per 100 papers; exceed → reuse existing or classify as "Unclassified" |
| 2 | **User Confirmation** | Task 5 topic classification now requires user confirmation after inference |
| 3 | **MD Retry Suffix** | Failed writes retried with `_retry1` suffix |
| 4 | **Enhanced Exclusion** | Task 7 incremental adds prefix-match exclusion for system folders |

## Seven Tasks

| Task | Description | Operation |
|------|-------------|-----------|
| **Task 1** | File consolidation → `Combined` folder | Copy |
| **Task 2** | File renaming (`AuthorYYYY.ext` format) | Rename |
| **Task 3** | Duplicate detection & redundancy → `Duplicates` | Move |
| **Task 4** | Generate/update `LitsSum.xlsx` (6 columns, incl. subjects) | Write |
| **Task 5** | Dual classification (`ByYear` + `ByTopic`, **with user confirmation**) | Copy |
| **Task 6** | Markdown conversion + 3-folder output (`md_all`/`md_ByYear`/`md_ByTopic`) | Convert+Copy |
| **Task 7** | Incremental additions (multi-source discovery, rename, dedup, update) | Mixed |

## File Structure

```
full-paper-management/
├── SKILL.md                     # Skill specification (7 tasks + 9 global rules)
├── README.md                    # Chinese README
├── README_EN.md                 # This file
└── scripts/
    ├── metadata_utils.py         # Metadata extraction, DOI resolution, APA, subject terms
    ├── file_ops.py               # File collection, rename, dedup, dual classification, MD classification
    ├── litsum_manager.py         # LitsSum.xlsx CRUD (6 columns)
    └── md_converter.py           # PDF/DOCX/TXT → MD conversion + structured template
```

## LitsSum.xlsx Schema

| Col | Header | Description |
|-----|--------|-------------|
| A | Year | 4-digit, Unknown last |
| B | Filename | Without extension |
| C | APA Citation | APA 7th format |
| D | Source Folder | Original folder name |
| E | Keywords | Semicolon-separated, max 5 |
| F | Subject Terms | 1-2 words, V3.1 cap: 10/100 papers |

## Global Rules

1. **Dry-Run First** — Preview all write operations before execution
2. **Granular Logging** — Status codes: ✅ ⏭ ⚠️ 📦 📁 🔄 ❌
3. **File Safety** — Never delete original files
4. **Robust Error Handling** — Single-step and global fault tolerance
5. **Strict Task Order** — First run: 1→6, incremental: Task 7
6. **Backup & Recovery** — Auto-backup LitsSum.xlsx before modifications
7. **Idempotency** — Repeated execution produces no duplicate data
8. **Resource Limits** — Parse first 3000 chars, DOI timeout 10s
9. **Unified Increment** — Conflict filenames increment from `_1`

## Dependencies

```bash
pip install openpyxl          # Excel I/O (required)
pip install PyMuPDF           # PDF text extraction (recommended)
# or
pip install PyPDF2            # PDF text extraction (alternative)
pip install python-docx       # DOCX text extraction
```

## Usage

Load the skill, then say "文献管理" or "manage literature" in conversation. Follow the dry-run confirmation prompts for each task.
