Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch08Module
    Public Interface IPrinter
        Sub Print(doc As Document)
    End Interface
    Public Class OldPrinter
        Public Sub OldPrint(s As String)
            Console.WriteLine("【旧打印机】" & s)
        End Sub
    End Class
    Public Class Document
        Public Property Content As String
    End Class
    Public Class PrinterAdapter
        Implements IPrinter

        Private ReadOnly m_OldPrinter As OldPrinter

        ' 带参构造：创建时即注入适配者，无需 Init 两步
        Public Sub New(oldPrinter As OldPrinter)
            m_OldPrinter = oldPrinter
        End Sub

        ' Implements IPrinter.Print，编译器强制签名一致
        Public Sub Print(doc As Document) Implements IPrinter.Print
            m_OldPrinter.OldPrint(doc.Content)
        End Sub
    End Class
    Sub Main()

        ' ② 适配者（旧类）：接口不兼容，只有 OldPrint

        ' ③ 新系统数据载体

        ' ④ 适配器：Implements IPrinter，内部组合 OldPrinter，在 Print 中转调

        ' 演示：通过接口引用使用适配器
        Dim doc As New Document With {.Content = "Hello World"}
        Dim adapter As IPrinter = New PrinterAdapter(New OldPrinter())
        adapter.Print(doc)   ' 【旧打印机】Hello World
    End Sub
End Module
