# -*- coding: utf-8 -*-
"""
批量提取 Office 文档文本（docx / pptx / xlsx）

扫描指定文件夹，按内容相关性挑选重要文件，提取正文与表格内容，
输出成 .txt，并生成一份提取索引。

相比第一版的修正：
  1) 修 pptx 表格解析 bug（python-pptx 的 table.rows 不支持切片）
  2) 选文件改为「内容优先 + 时间次之」：方案/文本/说明书/汇报/策划 优先，
     结算清单/报价/模板/登记表 降权
  3) 文件数量多的项目自动下钻一层子目录，按子目录分别提取

只读源文件，不改动任何原始文档。
"""

import os
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path

# ---------------- 配置 ----------------
ROOTS = [
    Path("./input"),        # 要扫描的文件夹
]
OUT_ROOT = Path("./output")

PER_PROJECT = 4
MAX_FILE_MB = 80
CHAR_LIMIT = 12000
MAX_FILES_TOTAL = 220
DRILL_THRESHOLD = 25      # 项目内 Office 文件超过这个数，就下钻一层子目录

SUFFIX = {".docx", ".pptx", ".xlsx"}
SKIP_DIR = {"node_modules", "__pycache__", ".git", "$RECYCLE.BIN"}

GOOD_WORDS = ("文本", "方案", "说明书", "规划", "汇报", "策划", "设计",
              "可行性", "研究", "报告", "成果", "图集", "总规", "控规")
BAD_WORDS = ("结算", "清单", "报价", "发票", "模板", "指引", "登记表",
             "评分表", "问卷", "委托书", "合同", "考勤", "费用", "付款", "目录")

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass


def safe_name(s: str, limit: int = 80) -> str:
    s = re.sub(r'[<>:"/\\|?*\r\n\t]', "_", s).strip(" .")
    return (s[:limit] or "未命名")


def collect(root: Path):
    found = []
    for dp, dns, fns in os.walk(root):
        dns[:] = [d for d in dns if d not in SKIP_DIR and not d.startswith("~$")]
        for fn in fns:
            if fn.startswith("~$"):
                continue
            p = Path(dp) / fn
            if p.suffix.lower() not in SUFFIX:
                continue
            try:
                st = p.stat()
            except OSError:
                continue
            if st.st_size > MAX_FILE_MB * 1024 * 1024 or st.st_size == 0:
                continue
            found.append((st.st_mtime, p))
    return found


def score(item):
    """内容优先打分：相关词加分，无关词重罚，其次看时间"""
    mtime, p = item
    name = p.name
    s = 0.0
    for w in GOOD_WORDS:
        if w in name:
            s += 2.0
            break
    for w in BAD_WORDS:
        if w in name:
            s -= 4.0
            break
    if p.suffix.lower() == ".xlsx":
        s -= 1.0          # 表格不产生成段文字
    if "~$" in name:
        s -= 10.0
    return (s, mtime)


def read_docx(p: Path) -> str:
    from docx import Document
    d = Document(str(p))
    out = [para.text.strip() for para in d.paragraphs if para.text.strip()]
    for ti, tb in enumerate(d.tables, 1):
        out.append(f"\n[表{ti}]")
        for row in list(tb.rows)[:40]:
            cells = [c.text.strip().replace("\n", " ") for c in row.cells]
            out.append(" | ".join(cells))
    return "\n".join(out)


def read_pptx(p: Path) -> str:
    from pptx import Presentation
    prs = Presentation(str(p))
    out = []
    for i, slide in enumerate(prs.slides, 1):
        buf = []
        for sh in slide.shapes:
            try:
                if sh.has_text_frame and sh.text_frame.text.strip():
                    buf.append(sh.text_frame.text.strip().replace("\n", " / "))
                if getattr(sh, "has_table", False) and sh.has_table:
                    # 注意：table.rows 不支持切片，必须先用 list() 包一层
                    for row in list(sh.table.rows)[:15]:
                        buf.append(" | ".join(c.text.strip().replace("\n", " ")
                                              for c in list(row.cells)))
            except Exception:
                continue
        if buf:
            out.append(f"[第{i}页] " + " ｜ ".join(buf))
    return "\n".join(out)


def read_xlsx(p: Path) -> str:
    import openpyxl
    wb = openpyxl.load_workbook(str(p), read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets[:6]:
        out.append(f"[工作表] {ws.title}")
        for ri, row in enumerate(ws.iter_rows(values_only=True)):
            if ri >= 30:
                break
            vals = ["" if v is None else str(v) for v in list(row)[:20]]
            if any(v.strip() for v in vals):
                out.append(" | ".join(vals))
    wb.close()
    return "\n".join(out)


READERS = {".docx": read_docx, ".pptx": read_pptx, ".xlsx": read_xlsx}


def units(root: Path):
    """产出 (显示名, 目录, 文件列表)"""
    projects = sorted([p for p in root.iterdir() if p.is_dir()],
                      key=lambda p: p.stat().st_mtime, reverse=True)
    for proj in projects:
        files = collect(proj)
        if not files:
            yield proj.name, proj, []
            continue
        if len(files) >= DRILL_THRESHOLD:
            subs = sorted([d for d in proj.iterdir() if d.is_dir()],
                          key=lambda d: d.stat().st_mtime, reverse=True)
            for sub in subs:
                sfiles = collect(sub)
                if len(sfiles) >= 3:
                    yield f"{proj.name}／{sub.name}", sub, sfiles
        yield proj.name, proj, files


def main():
    OUT_ROOT.mkdir(parents=True, exist_ok=True)
    index = ["# 文档文本提取索引", ""]
    done = 0

    for root in ROOTS:
        if not root.exists():
            index.append(f"- ⚠️ 路径不存在：{root}")
            continue
        institute = root.name
        index.append(f"\n## {institute}")
        print(f"\n=== {institute} ===")

        for disp, proj, files in units(root):
            if not files:
                index.append(f"- {disp}：无可用 Office 文件")
                continue
            picked = sorted(files, key=score, reverse=True)[:PER_PROJECT]
            index.append(f"\n### {disp}")
            target = OUT_ROOT / safe_name(institute) / safe_name(disp)
            target.mkdir(parents=True, exist_ok=True)

            for mtime, f in picked:
                if done >= MAX_FILES_TOTAL:
                    index.append("  - （已达总上限，停止）")
                    break
                rel = str(f.relative_to(root))
                try:
                    text = READERS[f.suffix.lower()](f)
                except Exception as e:
                    index.append(f"  - ❌ {rel} —— {e.__class__.__name__}: {e}")
                    continue
                text = (text or "").strip()
                if not text:
                    index.append(f"  - ⚠️ {rel} —— 无文本内容")
                    continue
                text = text[:CHAR_LIMIT]
                outfile = target / (safe_name(f.name, 60) + ".txt")
                header = (f"源文件：{rel}\n修改时间：{datetime.fromtimestamp(mtime):%Y-%m-%d %H:%M}\n"
                          f"{'-' * 60}\n")
                outfile.write_text(header + text, encoding="utf-8")
                done += 1
                index.append(f"  - ✅ {rel}（{datetime.fromtimestamp(mtime):%Y-%m-%d}，"
                             f"{len(text)} 字）→ {outfile.name}")
                print(f"  ✓ {disp} / {f.name}")

    (OUT_ROOT / "index.md").write_text("\n".join(index), encoding="utf-8")
    print(f"\n完成：共导出 {done} 个文件 → {OUT_ROOT}")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(1)
