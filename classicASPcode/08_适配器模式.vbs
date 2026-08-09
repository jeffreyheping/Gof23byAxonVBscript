Option Explicit
Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
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

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
