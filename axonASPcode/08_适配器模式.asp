<%
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
%>