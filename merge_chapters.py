"""
将 byChapterMDcn/ 下的章节 MD 文件按顺序拼接成一个合并 MD 文件。
  - 按 文件名前缀数字排序（00_前言, 01_单例, ..., 23_解释器, 24_附录）
  - 输出文件名: 23个设计模式_VBScript版_YYYYMMDD_HHMMSS.md
  - 输出位置: 项目根目录（与本脚本同级）
用法:  python merge_chapters.py
"""
import os
import glob
import re
from datetime import datetime

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
CHAPTER_DIR = os.path.join(BASE_DIR, "byChapterMDcn")

# 合并文件命名前缀
OUTPUT_PREFIX = "23个设计模式_VBScript版"

def list_chapters():
    """列出 byChapterMDcn/ 下所有 .md 文件，按文件名前缀数字升序排序"""
    mds = glob.glob(os.path.join(CHAPTER_DIR, "*.md"))
    if not mds:
        return []

    def sort_key(path):
        name = os.path.basename(path)
        m = re.match(r"^(\d+)_", name)
        return int(m.group(1)) if m else 999

    return sorted(mds, key=sort_key)

def merge():
    chapters = list_chapters()
    if not chapters:
        print(f"[ERROR] No .md files found in {CHAPTER_DIR}")
        return None

    parts = []
    for path in chapters:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read().rstrip()
        parts.append(content)
        print(f"  [MERGE] {os.path.basename(path)}")

    # 章节之间用空行分隔（各文件末尾通常已含 --- 分隔符）
    merged = "\n\n".join(parts) + "\n"

    # 生成时间戳文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"{OUTPUT_PREFIX}_{timestamp}.md"
    out_path = os.path.join(BASE_DIR, out_name)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(merged)

    print(f"\n[MERGE] Done. {len(chapters)} chapters merged.")
    print(f"[MERGE] Output: {out_path}")
    return out_path

def main():
    print(f"[MERGE] Reading chapters from {CHAPTER_DIR}")
    merge()

if __name__ == "__main__":
    main()
