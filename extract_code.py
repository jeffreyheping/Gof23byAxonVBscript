"""
从 Markdown 提取 VBScript 设计模式代码
  - 传统 VBScript 版 -> classicASPcode/  (.vbs)
  - Axon VBScript 版  -> axonAspModernCode/    (.asp, 自动包裹 <% %>)
  - VB.NET 版         -> vbNetcode/      (.vb)
  - 传统版复用        -> aspPycode/      (.asp)  供 ASPPY 运行
                       -> axonAspClassicCode/ (.asp)  供 AxonASP 跑传统语法
  - 按章节拆分 MD 文件 -> byChapterMDcn/  (25个文件)
用法:  python extract_code.py [md_file]
"""
import os, re, sys, glob

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
CLASSIC_DIR = os.path.join(BASE_DIR, "classicASPcode")
AXON_MODERN_DIR = os.path.join(BASE_DIR, "axonAspModernCode")
ASPPY_DIR   = os.path.join(BASE_DIR, "aspPycode")
VBNET_DIR   = os.path.join(BASE_DIR, "vbNetcode")
AXON_CLASSIC_DIR = os.path.join(BASE_DIR, "axonAspClassicCode")
CHAPTER_DIR = os.path.join(BASE_DIR, "byChapterMDcn")


# ── VB.NET 后处理辅助 ──────────────────────────────────────────
_VBNET_TYPE_OPEN_RE = re.compile(
    r"^\s*"
    r"(?:Public\s+|Private\s+|Protected\s+|Friend\s+|Protected\s+Friend\s+)*"
    r"(?:MustInherit\s+|NotInheritable\s+|Partial\s+|Shadows\s+)*"
    r"(?:Class|Structure|Enum|Module|Namespace)\b",
    re.I,
)
_VBNET_TYPE_CLOSE_RE = re.compile(
    r"^\s*"
    r"(?:Public\s+|Private\s+|Protected\s+|Friend\s+)*"
    r"End\s+(?:Class|Structure|Enum|Module|Namespace)\b",
    re.I,
)
_VBNET_INTERFACE_OPEN_RE = re.compile(
    r"^\s*"
    r"(?:Public\s+|Private\s+|Protected\s+|Friend\s+)*"
    r"Interface\b",
    re.I,
)
_VBNET_INTERFACE_CLOSE_RE = re.compile(
    r"^\s*"
    r"(?:Public\s+|Private\s+|Protected\s+|Friend\s+)*"
    r"End\s+Interface\b",
    re.I,
)
_VBNET_INH_IMPL_RE = re.compile(r"^\s*(Inherits|Implements)\b", re.I)


def _vbnet_wrap_module_and_main(code: str, module_name: str) -> str:
    """将一段原生的"教科书"VB.NET 代码（类定义 + 顶层演示语句）包成可编译的 Module。

    分类策略（按行扫描 + 嵌套栈）：
      - 类型头（Public/MustInherit Class/Structure/...）→ 打开嵌套，归 outer
      - 类型尾（End Class/...）→ 关闭嵌套，归 outer
      - Inherits/Implements 必须紧跟在类头后面 → outer
      - 其它非空行：
         * 若在类型体内（type_depth>0）→ outer（成员方法、字段、嵌套类）
         * 否则 → inner（演示语句：Dim/赋值/Console.WriteLine...）
      - 用户写的 Option/Imports 归 outer，注入时会去重前置
    """
    code_lines = code.splitlines()

    # 收集用户写的 Option/Imports（可选，可能没有）
    user_opt_imports_lines = []
    user_body_lines = []
    for raw_line in code_lines:
        s = raw_line.strip()
        if s.startswith("'") or s == "":
            # 开头的注释/空行也当作 opt/imp 区域的一部分
            if not user_body_lines:
                user_opt_imports_lines.append(raw_line)
            else:
                user_body_lines.append(raw_line)
            continue
        if re.match(r"^Option\s+", s, re.I) or re.match(r"^Imports\s+", s, re.I):
            # 前面如果没有 body 行，还属于前置
            if not user_body_lines:
                user_opt_imports_lines.append(raw_line)
                continue
        user_body_lines.append(raw_line)

    outer_lines = []  # 类型定义与类型体内的内容（放 Module 级别）
    inner_lines = []  # Sub Main 里的演示语句
    type_depth = 0
    for raw_line in user_body_lines:
        stripped = raw_line.strip()
        if not stripped:
            (outer_lines if type_depth > 0 else inner_lines).append(raw_line)
            continue
        if stripped.startswith("'"):
            # 注释：按是否在类型体内归属
            (outer_lines if type_depth > 0 else inner_lines).append(raw_line)
            continue
        depth_change = 0
        is_outer = False
        if _VBNET_TYPE_OPEN_RE.match(raw_line):
            depth_change = +1
            is_outer = True
        elif _VBNET_TYPE_CLOSE_RE.match(raw_line):
            depth_change = -1
            is_outer = True
        elif _VBNET_INTERFACE_OPEN_RE.match(raw_line):
            depth_change = +1
            is_outer = True
        elif _VBNET_INTERFACE_CLOSE_RE.match(raw_line):
            depth_change = -1
            is_outer = True
        elif _VBNET_INH_IMPL_RE.match(raw_line):
            is_outer = True
        if type_depth > 0 or depth_change != 0 or is_outer:
            outer_lines.append(raw_line)
        else:
            inner_lines.append(raw_line)
        type_depth += depth_change

    # 注入的 Option/Imports（放在 Module 声明之后、第一行之前，无缩进或 4 空格缩进都行；
    # 为兼容 vbc /langversion 默认，这里保持缩进 4 空格，但要放在类型定义之前）
    injected_header_lines = [
        "Option Strict On",
        "Option Explicit On",
        "Imports System",
        "Imports System.Collections.Generic",
        "Imports System.Collections",
        "Imports System.Linq",
    ]
    # 去重：用户写过的 Imports/Option 不再重复注入
    def _norm(s):
        return re.sub(r"\s+", " ", s.strip().lower())
    user_norm = {_norm(l) for l in user_opt_imports_lines if l.strip() and not l.strip().startswith("'")}
    final_header = []
    for l in injected_header_lines:
        if _norm(l) not in user_norm:
            final_header.append(l)

    # 用户写过的 Option/Imports（如果是在文件开头写的）追加到注入 header 后面
    for l in user_opt_imports_lines:
        s = l.strip()
        if not s or s.startswith("'"):
            continue
        if re.match(r"^(Option|Imports)\s+", s, re.I):
            final_header.append(s)

    # 组装：
    #   Option / Imports 必须在任何 namespace/module/class 声明之前（file-level）
    #   所以结构：
    #     Option Strict Off
    #     Option Explicit On
    #     Imports ...
    #     Module X
    #         Class ...
    #         Sub Main() ...
    #     End Module
    IND4  = "    "
    IND8  = IND4 + IND4

    out_lines = []
    # file-level Option/Imports（不加缩进，放在 Module 之前）
    for l in final_header:
        out_lines.append(l)
    out_lines.append(f"Module {module_name}")
    # 用户类型定义（Module 级，缩进 4）
    for l in outer_lines:
        out_lines.append((IND4 + l) if l else "")
    # Sub Main 入口（缩进 4）
    out_lines.append(IND4 + "Sub Main()")
    for l in inner_lines:
        out_lines.append((IND8 + l) if l else "")
    out_lines.append(IND4 + "End Sub")
    out_lines.append("End Module")
    out_lines.append("")
    return "\r\n".join(out_lines)


# 章节标题里中文数字 -> 阿拉伯数字
CN_NUM = {"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,
          "九":9,"十":10,"十一":11,"十二":12,"十三":13,"十四":14,"十五":15,
          "十六":16,"十七":17,"十八":18,"十九":19,"二十":20,"二十一":21,
          "二十二":22,"二十三":23}

def find_md():
    """找设计模式总 md：只认 23个设计模式*.md，按文件名时间戳取最新"""
    mds = glob.glob(os.path.join(BASE_DIR, "23个设计模式*.md"))
    if not mds:
        return None
    return max(mds, key=os.path.basename)   # 文件名含时间戳，字典序=时间序

def extract(md_path):
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    # 匹配每章内的三个代码块（传统版 / Axon 版 / VB.NET 版）
    # 用占位符迭代：先抓所有 ## 第N章 块的位置
    chapter_re = re.compile(r"^##\s*第(\d+|[一二三四五六七八九十]+)章\s*([^（(\n]+)", re.M)
    block_re   = re.compile(
        r"###\s*(传统\s*VBScript\s*版|Axon\s*VBScript\s*版[^\n]*|VB\.NET\s*版[^\n]*)\s*\n(?:.*?\n)*?```(vbscript|vba|vbnet)\s*\n(.*?)```",
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
            lang  = bm.group(2).strip()
            code  = bm.group(3).strip()
            # 统一换行符：保证所有引擎都得到 LF 分隔，后处理一致
            code = code.replace("\r\n", "\n").replace("\r", "\n")
            if label.startswith("传统"):
                engine = "ClassicASP"
                ext    = ".vbs"
            elif label.startswith("Axon"):
                engine = "AxonASP"
                ext    = ".asp"
            elif label.startswith("VB.NET"):
                engine = "VBNET"
                ext    = ".vb"
            else:
                continue
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
    os.makedirs(AXON_MODERN_DIR, exist_ok=True)
    os.makedirs(ASPPY_DIR,   exist_ok=True)
    os.makedirs(VBNET_DIR,   exist_ok=True)
    os.makedirs(AXON_CLASSIC_DIR, exist_ok=True)

    # 先清空目录里的旧文件
    for d in (CLASSIC_DIR, AXON_MODERN_DIR, ASPPY_DIR, VBNET_DIR, AXON_CLASSIC_DIR):
        for f in os.listdir(d):
            fp = os.path.join(d, f)
            if os.path.isfile(fp):
                os.remove(fp)

    saved = []
    for item in items:
        if item["type"] == "ClassicASP":
            target = CLASSIC_DIR
        elif item["type"] == "AxonASP":
            target = AXON_MODERN_DIR
        else:
            target = VBNET_DIR
        path   = os.path.join(target, item["filename"])
        content = item["content"]

        if item["type"] == "VBNET":
            # VB.NET: 用 .NET Framework vbc.exe 编译成控制台 EXE，UTF-8 编码
            # 需要在用户代码外包裹一个 Module + Sub Main（作为控制台入口）
            module_name = f"Ch{item['chapter']:02d}Module"
            has_entry = ("Sub Main" in content) or ("Module " in content) or ("Namespace " in content)
            if not has_entry:
                content = _vbnet_wrap_module_and_main(content, module_name)
            data = content.encode("utf-8-sig")
        elif item["type"] == "AxonASP":
            # AxonASP 没 <% %> 就包一层，注入 Option Explicit（VBScript 语法：不带 On）
            if "<%" not in content:
                content = "<%\nOption Explicit\n" + content + "\n%>"
            else:
                # 已有 <% %>，在开头的 <% 后面紧跟 Option Explicit
                content = content.replace("<%", "<%\nOption Explicit\n", 1)
            data = content.encode("utf-8")
        else:
            # ClassicASP: 注入 Option Explicit + ResponseStub，GBK 编码
            content = (
                "Option Explicit\n"
                "Dim Response: Set Response = New ResponseStub\n"
                "' -- inject: ResponseStub class below user code --\n"
                + content + "\n\n"
                "Class ResponseStub\n"
                "    Public Sub Write(s)\n"
                "        WScript.Echo s\n"
                "    End Sub\n"
                "End Class\n"
            )
            data = content.encode("gbk", errors="replace")

        with open(path, "wb") as f:
            f.write(data)
        saved.append({**item, "path": path})
        print(f"  {item['type']:10s} | ch{item['chapter']:02d} | {item['filename']}")

        # ASPPY / AxonASP-Classic: 用传统版代码（不注入 ResponseStub），
        # 包裹 <% %>，UTF-8 编码；注入 Option Explicit（VBScript 语法不带 On）
        if item["type"] == "ClassicASP":
            wrapped = item["content"]
            if "<%" not in wrapped:
                wrapped = "<%\nOption Explicit\n" + wrapped + "\n%>"
            else:
                wrapped = wrapped.replace("<%", "<%\nOption Explicit\n", 1)
            asp_name = item["filename"].replace(".vbs", ".asp")
            for out_dir, tag in ((ASPPY_DIR, "ASPPY"), (AXON_CLASSIC_DIR, "AxonClassic")):
                out_path = os.path.join(out_dir, asp_name)
                with open(out_path, "wb") as f:
                    f.write(wrapped.encode("utf-8"))
                print(f"  {tag:10s} | ch{item['chapter']:02d} | {asp_name}")
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
    
    # 先找出设计模式章节（第N章）和附录章节
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

    # 头部：第一个"第N章"之前的内容（含前言各节）
    if design_chapters:
        head_end = design_chapters[0].start()
    else:
        head_end = matches[0].start()
    head_content = content[:head_end].strip()
    if head_content:
        head_path = os.path.join(CHAPTER_DIR, "00_前言.md")
        with open(head_path, "w", encoding="utf-8") as f:
            f.write(head_content)
        print(f"  [SPLIT] 00_前言.md")
    
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
    v = sum(1 for i in items if i["type"] == "VBNET")
    print(f"[EXTRACT] Found {len(items)} code blocks (ClassicASP={c}, AxonASP={a}, VBNET={v})")
    save(items)
    
    # 按章节拆分 MD 文件
    print(f"[SPLIT] Splitting by chapter...")
    split_chapters(md_file)
    
    print(f"[EXTRACT] Done.")

if __name__ == "__main__":
    main()