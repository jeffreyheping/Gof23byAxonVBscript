"""
从 Markdown 提取 VBScript 设计模式代码
  - 传统 VBScript 版 -> classicASPcode/  (.vbs)
  - Axon VBScript 版  -> axonASPcode/    (.asp, 自动包裹 <% %>)
  - 按章节拆分 MD 文件 -> byChapterMDcn/  (25个文件)
用法:  python extract_code.py [md_file]
"""
import os, re, sys, glob

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
CLASSIC_DIR = os.path.join(BASE_DIR, "classicASPcode")
AXON_DIR    = os.path.join(BASE_DIR, "axonASPcode")
ASPPY_DIR   = os.path.join(BASE_DIR, "aspPycode")
CHAPTER_DIR = os.path.join(BASE_DIR, "byChapterMDcn")

# 章节标题里中文数字 -> 阿拉伯数字
CN_NUM = {"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,
          "九":9,"十":10,"十一":11,"十二":12,"十三":13,"十四":14,"十五":15,
          "十六":16,"十七":17,"十八":18,"十九":19,"二十":20,"二十一":21,
          "二十二":22,"二十三":23}

def find_md():
    """找设计模式 md（排除 test_report.md）"""
    mds = glob.glob(os.path.join(BASE_DIR, "*.md"))
    cands = [m for m in mds if "report" not in os.path.basename(m).lower()]
    return cands[0] if cands else (mds[0] if mds else None)

def extract(md_path):
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    # 匹配每章内的两个代码块（传统版 / Axon 版）
    # 用占位符迭代：先抓所有 ## 第N章 块的位置
    chapter_re = re.compile(r"^##\s*第(\d+|[一二三四五六七八九十]+)章\s*([^（(\n]+)", re.M)
    block_re   = re.compile(
        r"###\s*(传统\s*VBScript\s*版|Axon\s*VBScript\s*版[^\n]*)\s*\n+```vbscript\s*\n(.*?)```",
        re.S,
    )

    results = []
    for cm in chapter_re.finditer(content):
        ch_text = cm.group(1)
        ch_name = cm.group(2).strip()
        try:
            ch_num = int(ch_text)
        except ValueError:
            try:
                ch_num = CN_NUM[ch_text]
            except KeyError:
                continue
        # 章节标题
        slug = re.sub(r"[^\w\u4e00-\u9fa5]+", "", ch_name)  # 模式名做文件名
        ch_slug = f"{ch_num:02d}_{slug}"

        # 在本章范围内（下一个 ## 之前）找代码块
        start = cm.end()
        nxt = chapter_re.search(content, start)
        end  = nxt.start() if nxt else len(content)
        region = content[start:end]

        for bm in block_re.finditer(region):
            label = bm.group(1).strip()
            code  = bm.group(2).strip()
            if label.startswith("传统"):
                engine = "ClassicASP"
                ext    = ".vbs"
            else:
                engine = "AxonASP"
                ext    = ".asp"
            results.append({
                "type":    engine,
                "filename": f"{ch_slug}{ext}",
                "title":   ch_name,
                "chapter": ch_num,
                "content": code,
            })
    return results

def save(items):
    os.makedirs(CLASSIC_DIR, exist_ok=True)
    os.makedirs(AXON_DIR,    exist_ok=True)
    os.makedirs(ASPPY_DIR,   exist_ok=True)

    # 先清空三个目录里的旧文件
    for d in (CLASSIC_DIR, AXON_DIR, ASPPY_DIR):
        for f in os.listdir(d):
            fp = os.path.join(d, f)
            if os.path.isfile(fp):
                os.remove(fp)

    saved = []
    for item in items:
        target = CLASSIC_DIR if item["type"] == "ClassicASP" else AXON_DIR
        path   = os.path.join(target, item["filename"])
        content = item["content"]
        # AxonASP 没 <% %> 就包一层
        if item["type"] == "AxonASP" and "<%" not in content:
            content = "<%\n" + content + "\n%>"
        # ClassicASP 在 cscript 下运行，没有 Response 对象
        # 顶部初始化 Response（只占1行，尽量不影响行号），底部放类定义
        if item["type"] == "ClassicASP":
            content = (
                "Dim Response: Set Response = New ResponseStub\n"
                "' -- inject: ResponseStub class below user code --\n"
                + content + "\n\n"
                "Class ResponseStub\n"
                "    Public Sub Write(s)\n"
                "        WScript.Echo s\n"
                "    End Sub\n"
                "End Class\n"
            )
        # cscript/wscript 读 .vbs 用系统 ANSI 代码页，中文会乱码
        # axonasp-cli 读 .asp 用 UTF-8
        # 所以 ClassicASP 写成 GBK(系统代码页)，AxonASP 写成 UTF-8
        if item["type"] == "ClassicASP":
            data = content.encode("gbk", errors="replace")
        else:
            data = content.encode("utf-8")
        with open(path, "wb") as f:
            f.write(data)
        saved.append({**item, "path": path})
        print(f"  {item['type']:10s} | ch{item['chapter']:02d} | {item['filename']}")

        # ASPPY: 用传统版代码（不注入 ResponseStub），包裹 <% %>，UTF-8 编码
        if item["type"] == "ClassicASP":
            asppy_content = item["content"]
            if "<%" not in asppy_content:
                asppy_content = "<%\n" + asppy_content + "\n%>"
            asppy_path = os.path.join(ASPPY_DIR, item["filename"].replace(".vbs", ".asp"))
            with open(asppy_path, "wb") as f:
                f.write(asppy_content.encode("utf-8"))
            print(f"  ASPPY      | ch{item['chapter']:02d} | {os.path.basename(asppy_path)}")
    return saved

def split_chapters(md_path):
    """将 MD 文件按章节拆分为 25 个文件（头+尾+23个模式）"""
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    os.makedirs(CHAPTER_DIR, exist_ok=True)
    
    # 清空目录中的旧文件
    for f in os.listdir(CHAPTER_DIR):
        fp = os.path.join(CHAPTER_DIR, f)
        if os.path.isfile(fp):
            os.remove(fp)
    
    # 找所有章节标题（## 开头）
    chapter_re = re.compile(r"^##\s+", re.M)
    matches = list(chapter_re.finditer(content))
    
    if not matches:
        print("[SPLIT] No chapters found")
        return
    
    # 头部：第一个 ## 之前的内容
    head_end = matches[0].start()
    head_content = content[:head_end].strip()
    if head_content:
        head_path = os.path.join(CHAPTER_DIR, "00_前言.md")
        with open(head_path, "w", encoding="utf-8") as f:
            f.write(head_content)
        print(f"  [SPLIT] 00_前言.md")
    
    # 找出设计模式章节（第1-23章）和附录章节
    design_chapters = []
    appendix_start = len(content)
    
    for i, m in enumerate(matches):
        line_end = content.find("\n", m.start())
        line_text = content[m.start():line_end].strip()
        
        # 判断是否为设计模式章节（第N章格式）
        if re.match(r"^##\s*第\d+章", line_text) or re.match(r"^##\s*第[一二三四五六七八九十]+章", line_text):
            design_chapters.append(m)
        # 判断是否为附录开头
        elif line_text.startswith("## 附："):
            appendix_start = m.start()
            break
    
    # 处理设计模式章节（只取前23个）
    for i, m in enumerate(design_chapters[:23]):
        start = m.start()
        # 找下一个章节的位置
        if i + 1 < len(design_chapters):
            next_start = design_chapters[i + 1].start()
        else:
            next_start = appendix_start
        
        chapter_content = content[start:next_start].strip()
        
        # 提取章节标题作为文件名
        line_end = content.find("\n", start)
        title_line = content[start:line_end].strip()
        title = re.sub(r"^##\s*", "", title_line)
        
        # 提取章节号和模式名
        ch_match = re.match(r"第(\d+|[一二三四五六七八九十]+)章\s*(.+)", title)
        if ch_match:
            ch_text = ch_match.group(1)
            ch_name = ch_match.group(2).strip()
            try:
                ch_num = int(ch_text)
            except ValueError:
                ch_num = CN_NUM.get(ch_text, i + 1)
            slug = re.sub(r"[^\w\u4e00-\u9fa5]+", "", ch_name)
            filename = f"{ch_num:02d}_{slug}.md"
        else:
            filename = f"{i+1:02d}_{title[:20]}.md"
        
        path = os.path.join(CHAPTER_DIR, filename)
        with open(path, "w", encoding="utf-8") as f:
            f.write(chapter_content)
        print(f"  [SPLIT] {filename}")
    
    # 尾部：附录部分（合并所有附录）
    if appendix_start < len(content):
        tail_content = content[appendix_start:].strip()
        tail_path = os.path.join(CHAPTER_DIR, "24_附录.md")
        with open(tail_path, "w", encoding="utf-8") as f:
            f.write(tail_content)
        print(f"  [SPLIT] 24_附录.md")
    
    print(f"[SPLIT] Done. Saved to {CHAPTER_DIR}")

def main():
    md_file = sys.argv[1] if len(sys.argv) > 1 else find_md()
    if not md_file or not os.path.isfile(md_file):
        print(f"[ERROR] MD file not found: {md_file}")
        sys.exit(1)

    print(f"[EXTRACT] Reading {md_file}")
    items = extract(md_file)
    c = sum(1 for i in items if i["type"] == "ClassicASP")
    a = sum(1 for i in items if i["type"] == "AxonASP")
    print(f"[EXTRACT] Found {len(items)} code blocks (ClassicASP={c}, AxonASP={a})")
    save(items)
    
    # 按章节拆分 MD 文件
    print(f"[SPLIT] Splitting by chapter...")
    split_chapters(md_file)
    
    print(f"[EXTRACT] Done.")

if __name__ == "__main__":
    main()