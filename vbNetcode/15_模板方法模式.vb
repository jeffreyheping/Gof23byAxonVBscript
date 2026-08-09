Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch15Module
    Public MustInherit Class DataMiner
        ' ② 模板方法（骨架）：调用顺序固定，子类无需也无法改流程
        Public Function MineData(filePath As String) As Object
            Load(filePath)              ' 固定步骤1：加载
            Parse(filePath)             ' 可变步骤：解析（子类必须实现）
            Analyze()                   ' 固定步骤2：分析
            SendReport()                ' 固定步骤3：发送报告
        End Function

        ' 固定步骤：Private 子类不可见，彻底防止被覆盖
        Private Function Load(filePath As String) As Object
            Console.WriteLine($"加载文件：{filePath}")
        End Function

        Private Function Analyze() As Object
            Console.WriteLine("分析数据")
        End Function

        Private Function SendReport() As Object
            Console.WriteLine("发送报告")
        End Function

        ' ③ 可变步骤：MustOverride 强制所有子类必须实现，漏写直接编译报错
        Public MustOverride Function Parse(filePath As String) As Object
    End Class
    Public Class PDFMiner
        Inherits DataMiner

        Public Overrides Function Parse(filePath As String) As Object
            Console.WriteLine("【解析 PDF】")
        End Function
    End Class
    Public Class CSVMiner
        Inherits DataMiner

        Public Overrides Function Parse(filePath As String) As Object
            Console.WriteLine("【解析 CSV】")
        End Function
    End Class
    Sub Main()

        ' ④ 具体子类1：PDF 解析器

        ' ⑤ 具体子类2：CSV 解析器

        ' 演示：调用方只依赖抽象基类，完全不关心具体子类实现
        Dim pdf As DataMiner = New PDFMiner()
        pdf.MineData("data.pdf")    ' 加载文件 → 【解析 PDF】→ 分析数据 → 发送报告

        Dim csv As DataMiner = New CSVMiner()
        csv.MineData("data.csv")    ' 加载文件 → 【解析 CSV】→ 分析数据 → 发送报告
    End Sub
End Module
