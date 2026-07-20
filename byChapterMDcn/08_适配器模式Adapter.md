## 第8章 适配器模式（Adapter）

**核心思想**：把不兼容的接口转成目标接口。

**示例说明**：OldPrinter 只有 OldPrint 方法（接收字符串），新系统期望用 Print 方法（接收 Document 对象）。PrinterAdapter 在中间做转换，把 doc.Content 取出来传给 OldPrint。

### 传统 VBScript 版

```vbscript
' 旧类：只有 OldPrint 方法，接收字符串
Class OldPrinter
    ' 旧接口：直接打印字符串
    Public Function OldPrint(s)
        Response.Write "【旧打印机】" & s
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
        m_OldPrinter.OldPrint doc.Content
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
adapter.Init New OldPrinter
adapter.Print doc   ' 用新接口调用旧打印机
```

**传统 VBScript 版妥协说明**：
- **无目标接口**：`PrinterAdapter` 没有 `IPrinter` 接口可实现，"新接口"只是约定 `Print` 方法名。如果有多个适配器，无法保证接口一致。

### Axon VBScript 版（支持 Implements）

```vbscript
' 目标接口
Class IPrinter
    Public Function Print(doc As Document)
    End Function
End Class

' 旧类
Class OldPrinter
    Public Function OldPrint(s As String)
        Response.Write "【旧打印机】" & s
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
        m_OldPrinter.OldPrint doc.Content
    End Function
End Class

' 演示：通过接口使用适配器
Dim doc As Document
Set doc = New Document
doc.Content = "Hello World"

Dim adapter As PrinterAdapter
Set adapter = New PrinterAdapter
adapter.Init New OldPrinter

Dim ip As IPrinter
Set ip = adapter
ip.Print doc
```

**Axon VBScript 版妥协说明**：
- 此模式在 AxonASP 中实现较为自然，接口机制解决了核心多态问题，无显著妥协。
---