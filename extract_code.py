"""
从 Markdown 提取 VBScript 设计模式代码
  - 传统 VBScript 版 -> classicASPcode/  (.vbs)
  - Axon VBScript 版  -> axonASPcode/    (.asp, 自动包裹 <% %>)
  - VB.NET 版         -> vbNetcode/      (.vb)
  - 按章节拆分 MD 文件 -> byChapterMDcn/  (25个文件)
用法:  python extract_code.py [md_file]
"""
import os, re, sys, glob

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
CLASSIC_DIR = os.path.join(BASE_DIR, "classicASPcode")
AXON_DIR    = os.path.join(BASE_DIR, "axonASPcode")
ASPPY_DIR   = os.path.join(BASE_DIR, "aspPycode")
VBNET_DIR   = os.path.join(BASE_DIR, "vbNetcode")
CHAPTER_DIR = os.path.join(BASE_DIR, "byChapterMDcn")


def _fix_vbs_trailing_comment_parens(code):
    """修复 VBScript 中 `Response.Write(X   ' 注释)` 这种把 `)` 写进注释的问题。

    模式：一行里有个未闭合的 `(`，行尾 `' ... )` 注释里包含了本应在括号外的 `)`。
    修复策略：对每行
      1. 检查该行是否以 `' 注释)` 这种模式结尾，并且注释体内含有 ')'
      2. 如果该行代码部分的 '(' 数多于 ')'，将注释体里多余的 ')' 数，
         逐个从注释最左 ')' 位置移除，依次补到代码末尾。
    """
    lines = code.split("\n")
    new_lines = []
    for line in lines:
        # 找代码/注释分界：第一个不在字符串里的单引号（简化：只看第一个 '）
        # 这个简化在真实演示代码里足够用，因为演示代码一般不在 Response.Write 参数里写含 ' 的字面量
        comment_start = line.find("'")
        if comment_start < 0:
            new_lines.append(line)
            continue
        code_part    = line[:comment_start]
        comment_part = line[comment_start + 1:]
        opens  = code_part.count("(")
        closes = code_part.count(")")
        deficit = opens - closes
        if deficit <= 0:
            new_lines.append(line)
            continue
        # 从注释里找 deficit 个 ')'，逐个挪到 code_part 尾
        # 取最靠左的 deficit 个 ')'
        moving_closes = min(deficit, comment_part.count(")"))
        if moving_closes <= 0:
            new_lines.append(line)
            continue
        # 从 comment_part 中删除前 moving_closes 个 ')'
        fixed_comment = list(comment_part)
        removed = 0
        i = 0
        while removed < moving_closes and i < len(fixed_comment):
            if fixed_comment[i] == ")":
                fixed_comment.pop(i)
                removed += 1
            else:
                i += 1
        # 补到 code_part 尾，保留原缩进
        new_code = code_part.rstrip() + (")" * moving_closes) + (" " * max(0, len(code_part) - len(code_part.rstrip())))
        new_lines.append(new_code + "'" + "".join(fixed_comment))
    return "\n".join(new_lines)


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


def _vbnet_fix_readonly_auto_and_default(code: str) -> str:
    """修补 MD 中常见 VB.NET 语法简写在 vbc.exe 中报错的场景：

    1) C# 风格的 ReadOnly 自动属性：
         Public ReadOnly Property Name As String
         Public Sub New(...)\n            Me.Name = value\n        End Sub
       → 转为带 Get 私有字段的完整属性。
       因为 vbc.exe（.NET Framework 4.x）默认 /langversion 不支持 ReadOnly 自动属性
       （需要在构造函数里赋值），会报 BC30126。

    2) Default(Public ReadOnly Property Item(...) As T) 这种括号包裹的语法
       → Default Public ReadOnly Property Item(...) As T
    """
    # 先处理 2）：Default(...) 变成 Default ...（去掉最外层包裹括号）
    # 匹配：行首可选空格 Default(...) → Default ...
    # 还要处理：因为上面行的缩进 + Default(...) 被替换后，行末可能有个多余的 )
    # 策略：先整段把 `Default(Public ... As T)` 全局替换再逐行二次处理
    code = re.sub(
        r"Default\((?P<body>(?:Public|Private|Protected|Friend|Protected\s+Friend)\s+"
        r"(?:ReadOnly\s+|WriteOnly\s+)?Property\s+Item(?:\([^)]*\))?\s*As\s+[^)\n]+?)\)",
        lambda m: "Default " + m.group("body"),
        code,
        flags=re.I,
    )
    # 清理因上面正则不完整残留的行尾 )（形如 "... As T)"）
    code = re.sub(
        r"^(\s*Default\s+Public\s+ReadOnly\s+Property\s+Item\([^)]*\)\s+As\s+[A-Za-z0-9_.,<> ]+)\)\s*$",
        r"\1",
        code,
        flags=re.M,
    )

    lines = code.split("\n")
    out = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        # 匹配：Public ReadOnly Property <Name> As <Type> （无 Get 块、且无行尾 = 初始化）
        m_prop = re.match(
            r"^(?P<access>Public|Private|Protected|Friend|Protected\s+Friend)\s+"
            r"ReadOnly\s+Property\s+(?P<name>[A-Za-z_][A-Za-z0-9_]*)\s+As\s+(?P<type>[A-Za-z_][A-Za-z0-9_.,<> ]+?)\s*$",
            stripped,
            re.I,
        )
        # 下一行若为 End Property 或其它属性，就不要转（已经是完整块）
        if m_prop and (i + 1 >= len(lines) or not lines[i + 1].strip().startswith("Get")):
            access = m_prop.group("access")
            pname  = m_prop.group("name")
            ptype  = m_prop.group("type").strip()
            field  = f"m_{pname}"
            indent = line[:len(line) - len(line.lstrip())]
            # 生成：私有字段 + 完整属性
            out.append(f"{indent}Private {field} As {ptype}")
            out.append(f"{indent}{access} ReadOnly Property {pname} As {ptype}")
            out.append(f"{indent}    Get")
            out.append(f"{indent}        Return {field}")
            out.append(f"{indent}    End Get")
            out.append(f"{indent}End Property")
            # 继续扫描，把后续构造函数中的 Me.<Name> = xxx → <field> = xxx
            # （简单处理：之后所有构造函数范围里对 Me.Name 的赋值都替换）
            i += 1
            continue
        # 替换构造函数 / 属性里的 Me.Name = xxx（前提：上面已经有这个 Name 的 ReadOnly 属性定义）
        # 这里全局替换所有 Me.<Name> = xxx → m_<Name> = xxx（Name 开头大写即可）
        # 为安全起见不全局，只针对上面匹配过的 pname：此处先放行，下面统一 pass 处理
        out.append(line)
        i += 1

    result = "\n".join(out)

    # 全局第二轮：若代码中出现 ReadOnly Property 定义的自动属性形式（见上面），
    # 同时出现 `Me.<PName> = xxx` 对 Me.PropertyName 赋值，统一替换成 m_<PName> = xxx。
    # 提取上面所有已改写的 ReadOnly 属性名（从 result 里找 m_* 字段 + ReadOnly Property）
    for m in re.finditer(r"Private\s+m_([A-Za-z_][A-Za-z0-9_]*)\s+As\s+.+?\n\s*(?:Public|Private|Protected|Friend)\s+ReadOnly\s+Property\s+\1\b",
                         result, flags=re.I):
        pname = m.group(1)
        result = re.sub(rf"Me\.{pname}\s*=",
                        lambda mm: f"m_{pname} =",
                        result,
                        flags=re.I)
    return result


def _vbnet_patch_syntax(code: str) -> str:
    """修补 MD 中 VB.NET 代码片段常见的语法瑕疵（在包 Module 之前）。"""
    lines = code.split("\n")
    out = []
    for line in lines:
        stripped = line.lstrip()
        # 1. 修复 m_Content(&= text) 这种括号吞掉空格问题 → m_Content &= text
        #    匹配：identifier(&= operand)
        fixed = re.sub(r"([A-Za-z_][A-Za-z0-9_]*)\(\s*(&=)\s*(.+?)\s*\)\s*$",
                       lambda m: f"{m.group(1)} {m.group(2)} {m.group(3)}",
                       line)
        # 2. 通用：标识符后面跟 (& 或 (+= 等）但少一个空格的模式
        #    例如 foo(& bar) → foo & bar
        fixed = re.sub(r"([A-Za-z_][A-Za-z0-9_]*)\(\s*([&+\-*/])\s*(?!=)",
                       lambda m: f"{m.group(1)} {m.group(2)} ",
                       fixed)
        # 2b. 复合赋值：i(+= 1) → i += 1；i(+ = 1) → i += 1；i(+ = 1) 空格变体
        fixed = re.sub(r"([A-Za-z_][A-Za-z0-9_]*)\(\s*([+\-*/&])\s*=\s*(.+?)\s*\)",
                       lambda m: f"{m.group(1)} {m.group(2)}= {m.group(3)}",
                       fixed)
        # 2c. 更宽松（允许 "+ =" 这种中间有空格，并且不强制以 ) 结尾）
        fixed = re.sub(r"([A-Za-z_][A-Za-z0-9_]*)\(\s*([+\-*/&])\s*=\s*(.+?)\)\s*$",
                       lambda m: f"{m.group(1)} {m.group(2)}= {m.group(3)}",
                       fixed)
        # 2d. 遗留："i + = 1)" 括号没包住整个的情况 —— "i + = 1)" → i += 1
        fixed = re.sub(r"([A-Za-z_][A-Za-z0-9_]*)\s+([+\-*/&])\s*\+?\s*=\s*([^\n)]+?)\)\s*$",
                       lambda m: f"{m.group(1)} {m.group(2)}= {m.group(3).strip()}",
                       fixed)
        # 3. 字符串插值 $"..." → String.Format("...", ...)
        #    简化处理：对每一行尝试，匹配 $"…{X}…{Y}…"
        while True:
            m = re.search(r'\$"([^"]*)"', fixed)
            if not m:
                break
            body = m.group(1)
            args = []
            idx = [0]
            def _sub_hole(hm):
                args.append(hm.group(1))
                r = "{" + str(idx[0]) + "}"
                idx[0] += 1
                return r
            fmt_body = re.sub(r"\{([^{}]+)\}", _sub_hole, body)
            if args:
                replaced = 'String.Format("' + fmt_body + '", ' + ", ".join(args) + ")"
            else:
                replaced = '"' + fmt_body + '"'
            fixed = fixed[:m.start()] + replaced + fixed[m.end():]
        out.append(fixed)
    return "\n".join(out)


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
    """找设计模式 md（排除 test_report.md）"""
    mds = glob.glob(os.path.join(BASE_DIR, "*.md"))
    cands = [m for m in mds if "report" not in os.path.basename(m).lower()]
    return cands[0] if cands else (mds[0] if mds else None)

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
            # VBScript 代码统一做尾随注释闭合括号修复
            if engine in ("ClassicASP", "AxonASP"):
                code = _fix_vbs_trailing_comment_parens(code)
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
    os.makedirs(VBNET_DIR,   exist_ok=True)

    # 先清空四个目录里的旧文件
    for d in (CLASSIC_DIR, AXON_DIR, ASPPY_DIR, VBNET_DIR):
        for f in os.listdir(d):
            fp = os.path.join(d, f)
            if os.path.isfile(fp):
                os.remove(fp)

    saved = []
    for item in items:
        if item["type"] == "ClassicASP":
            target = CLASSIC_DIR
        elif item["type"] == "AxonASP":
            target = AXON_DIR
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

        # ASPPY: 用传统版代码（不注入 ResponseStub），包裹 <% %>，UTF-8 编码；注入 Option Explicit（VBScript 语法不带 On）
        if item["type"] == "ClassicASP":
            asppy_content = item["content"]
            if "<%" not in asppy_content:
                asppy_content = "<%\nOption Explicit\n" + asppy_content + "\n%>"
            else:
                asppy_content = asppy_content.replace("<%", "<%\nOption Explicit\n", 1)
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