<%
Option Explicit
' 数据挖掘器：定义算法骨架
Class DataMiner
    ' 模板方法：固定流程，parser 参数提供可变步骤
    Public Function MineData(filePath, parser)
        Load(filePath)

        parser.Parse
        Analyze
        SendReport
    End Function

    ' 公共步骤：加载文件
    Private Function Load(filePath)
        Response.Write("加载文件：" & filePath)

    End Function

    ' 公共步骤：分析数据
    Private Function Analyze
        Response.Write("分析数据")

    End Function

    ' 公共步骤：发送报告
    Private Function SendReport
        Response.Write("发送报告")

    End Function
End Class

' 具体子类：PDF 格式解析（组合方式模拟继承）
Class PDFMiner
    Private m_Miner

    Private Sub Class_Initialize
        Set m_Miner = New DataMiner
    End Sub

    Public Function MineData(filePath)
        m_Miner.MineData filePath, Me
    End Function

    ' 覆盖父类的 Parse 方法
    Public Function Parse
        Response.Write("【解析 PDF】")

    End Function
End Class

' 具体子类：CSV 格式解析（组合方式模拟继承）
Class CSVMiner
    Private m_Miner

    Private Sub Class_Initialize
        Set m_Miner = New DataMiner
    End Sub

    Public Function MineData(filePath)
        m_Miner.MineData filePath, Me
    End Function

    ' 覆盖父类的 Parse 方法
    Public Function Parse
        Response.Write("【解析 CSV】")

    End Function
End Class

' 演示：两个子类复用相同流程，只需实现 Parse
Dim pdf, csv
Set pdf = New PDFMiner
pdf.MineData("data.pdf")


Set csv = New CSVMiner
csv.MineData("data.csv")
%>