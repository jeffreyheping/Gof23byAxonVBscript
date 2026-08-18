# AxonASP Bug #3: Paren-less method call on array element silently drops the argument (no error)

> 以下内容可直接粘贴到 GitHub Issue。标题用下方建议标题。

**建议标题 / Suggested title:**

```
Paren-less method call on array element silently drops the argument (no error, wrong output)
```

---

## Body

Hi! I found a third variant of the "paren-less method call on array element" bug family I reported earlier (the compile-error one and the Object-required one). **This one is the most dangerous: it produces no error at all — the argument is silently dropped and becomes `Empty`.**

### Environment

- axonasp-cli.exe (latest local build, 2026-08-18)
- Windows 11

### Minimal reproduction

```asp
<%
Option Explicit
Class Obs
    Public Name
    Public Function Update(news)
        Response.Write(Name & " got: [" & news & "] TypeName=" & TypeName(news))
    End Function
End Class

Class Pub
    Private m_Obs()
    Private m_Count
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Obs(10)
    End Sub
    Public Function Add(o)
        Set m_Obs(m_Count) = o
        m_Count = m_Count + 1
    End Function
    Public Function Notify(news)
        Dim i
        For i = 0 To m_Count - 1
            m_Obs(i).Update news          ' paren-less call on array element
        Next
    End Function
End Class

Dim p, a
Set p = New Pub
Set a = New Obs
a.Name = "A"
p.Add(a)

' Control group: simple variable as caller, same paren-less style
Dim simple
Set simple = New Obs
simple.Name = "S"

Response.Write("== simple variable caller (control) ==")
simple.Update "BREAKING NEWS"
Response.Write("== array element caller (bug) ==")
p.Notify("BREAKING NEWS")
%>
```

### Actual output (axonasp-cli)

```
== simple variable caller (control) ==S got: [BREAKING NEWS] TypeName=String== array element caller (bug) ==A got: [] TypeName=Empty
```

### Expected output (matches cscript.exe)

```
== simple variable caller (control) ==
S got: [BREAKING NEWS] TypeName=String
== array element caller (bug) ==
A got: [BREAKING NEWS] TypeName=String
```

### Analysis

Same root cause family as the two bugs I reported earlier — the trigger is always **an array element as the caller + a paren-less method call**. Depending on the argument shape, it fails in three different ways:

| # | Argument shape | Behavior | Severity |
|---|---|---|---|
| 1 | expression (`indent & "  "`) | Compile error 800A03EA | low — fails fast |
| 2 | two arguments (`msg, fromUser`) | Runtime error 800A01A8 "Object required" | medium — errors at runtime |
| 3 | **single simple argument (`news`)** | **Silent: argument becomes `Empty`, no error** | **high — wrong results with no warning** |

Variant 3 is the scariest one: real-world code like the Observer pattern (`m_Observers(i).Update news`) keeps running and just produces empty/incorrect output, so it can go unnoticed in production.

As always, happy to test any fix — my full 23-pattern test suite covers this automatically.

---

## 附：中文备注（不用贴）

- 复现文件：`repro3_silent_arg_loss.asp`（保留在本仓库）
- 触发场景：ch14 观察者模式 `m_Observers(i).Update news`，测试框架显示 PASS（无报错）但输出为空
- 测试框架目前只判"有无报错"，未比对输出内容，故此 bug 长期漏网
