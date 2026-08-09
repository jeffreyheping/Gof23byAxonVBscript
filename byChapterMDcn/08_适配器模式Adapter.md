## 第8章 适配器模式（Adapter）

**核心思想**：把不兼容的接口转成目标接口。

**示例说明**：OldPrinter 只有 OldPrint 方法（接收字符串），新系统期望用 Print 方法（接收 Document 对象）。PrinterAdapter 在中间做转换，把 doc.Content 取出来传给 OldPrint。

### 传统 VBScript 版

```vbscript
' 旧类：只有 OldPrint 方法，接收字符串
Class OldPrinter
    ' 旧接口：直接打印字符串
    Public Function OldPrint(s)
        Response.Write("【旧打印机】" & s)

    End Function
End Class

' 适配器：把旧接口转换成新接口
Class PrinterAdapter
    Private m_OldPrinter

    ' 注入被适配的旧对象
    Public Function Init(oldPrinter)
        Set m_OldPrinter = oldPrinter
    End Function

    ' 新接口：接收 Document 对象，提取 Content 后转调旧接口
    Public Function Print(doc)
        m_OldPrinter.OldPrint(doc.Content)

    End Function
End Class

' 新系统的数据载体
Class Document
    Public Content
End Class

' 演示：用新接口 Print 调用旧打印机
Dim doc
Set doc = New Document
doc.Content = "Hello World"

Dim adapter
Set adapter = New PrinterAdapter
adapter.Init(New OldPrinter)

adapter.Print(doc)   ' 用新接口调用旧打印机

```

**传统 VBScript 版妥协说明**：
- **无目标接口**：`PrinterAdapter` 没有 `IPrinter` 接口可实现，"新接口"只是约定 `Print` 方法名。如果有多个适配器，无法保证接口一致。
- **无继承**：经典适配器可选择"类适配器"（多重继承 Adapter 同时继承 Target 和 Adaptee）或"对象适配器"（组合 Adaptee）。VBScript 无继承，只能用对象适配器（组合），无法用类适配器一次复用两边代码。

### Axon VBScript 版（支持 Implements）

```vba
' 目标接口
Class IPrinter
    Public Function Print(doc As Document)
    End Function
End Class

' 旧类
Class OldPrinter
    Public Function OldPrint(s As String)
        Response.Write("【旧打印机】" & s)

    End Function
End Class

' 新系统的数据载体
Class Document
    Public Content As String
End Class

' 适配器：实现目标接口，内部组合旧对象
Class PrinterAdapter
    Implements IPrinter
    Private m_OldPrinter As OldPrinter

    Public Function Init(oldPrinter As OldPrinter)
        Set m_OldPrinter = oldPrinter
    End Function

    Public Function IPrinter_Print(doc As Document)
        m_OldPrinter.OldPrint(doc.Content)

    End Function
End Class

' 演示：通过接口使用适配器
Dim doc As Document
Set doc = New Document
doc.Content = "Hello World"

Dim adapter As PrinterAdapter
Set adapter = New PrinterAdapter
adapter.Init(New OldPrinter)


Dim ip As IPrinter
Set ip = adapter
ip.Print(doc)

```

**Axon VBScript 版妥协说明**：
- `Implements` 接口机制解决了目标契约问题：`IPrinter` 强制适配器实现 `Print`，多个适配器共享同一接口签名，调用方通过 `IPrinter` 引用统一调用，无需关心具体适配器类型。
- 缺失语法点：**代码复用机制（继承）**。经典适配器有类适配器与对象适配器两种写法——类适配器通过多重继承同时继承 Target 和 Adaptee，可直接复用两边方法而无需手动转发；AxonASP 只能用对象适配器（组合 `OldPrinter`），每个适配方法都要手写一次转调代码，新增方法时样板代码线性增长。Go 同样无继承，但 Go 用 struct embedding（`PrinterAdapter struct { *OldPrinter }`）让嵌入方法自动提升到外层，只需重写 `Print` 即可，其余方法零成本透传。AxonASP 目前只能手动逐个委托。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有 `Interface` + `Implements` 完整语法，对象适配器写法与 Axon 版结构一一对应：`PrinterAdapter Implements IPrinter`，内部组合 `OldPrinter`，在 `Print` 中转调 `OldPrint`。

```vbnet
' ① 目标接口：新系统期望的契约
Public Interface IPrinter
    Function Print(doc As Document) As Object
End Interface

' ② 适配者（旧类）：接口不兼容，只有 OldPrint
Public Class OldPrinter
    Public Function OldPrint(s As String) As Object
        Console.WriteLine("【旧打印机】" & s)
    End Function
End Class

' ③ 新系统数据载体
Public Class Document
    Public Property Content As String
End Class

' ④ 适配器：Implements IPrinter，内部组合 OldPrinter，在 Print 中转调
Public Class PrinterAdapter
    Implements IPrinter

    Private ReadOnly m_OldPrinter As OldPrinter

    ' 带参构造：创建时即注入适配者，无需 Init 两步
    Public Sub New(oldPrinter As OldPrinter)
        m_OldPrinter = oldPrinter
    End Sub

    ' Implements IPrinter.Print，编译器强制签名一致
    Public Function Print(doc As Document) As Object Implements IPrinter.Print
        m_OldPrinter.OldPrint(doc.Content)
    End Function
End Class

' 演示：通过接口引用使用适配器
Dim doc As New Document With {.Content = "Hello World"}
Dim adapter As IPrinter = New PrinterAdapter(New OldPrinter())
adapter.Print(doc)   ' 【旧打印机】Hello World
```

**VB.NET 版说明**：
- **`Interface` + `Implements` 编译期契约**：`Implements IPrinter.Print` 让适配器与目标接口签名绑定，漏写或拼写错误直接编译报错。Axon 版 `IPrinter_Print` 靠命名约定，错误到运行时才暴露。
- **带参构造函数替代 Init**：`New PrinterAdapter(New OldPrinter())` 创建时即注入适配者，不存在"先 New 后 Init"的半初始化窗口。
- **无需 `Set`**：VB.NET 对象赋值直接用 `=`，不再区分 `Set`/`Let`。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 目标契约 | 方法名约定（靠自觉） | `Implements IPrinter` 接口约束 | `Interface` + `Implements` 编译期强制 |
| 适配方式 | 对象组合 | 对象组合 | 对象组合 |
| 适配者注入 | `Init` 两步（易忘） | `Init` 两步（易忘） | 带参构造 `New(OldPrinter)` 一步到位 |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
---