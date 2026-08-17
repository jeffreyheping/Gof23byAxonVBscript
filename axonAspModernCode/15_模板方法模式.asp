<%
Option Explicit
' 提取器接口（可变步骤）
Class IExtractor
    Public Function Extract(data As String) As String
    End Function
End Class

' 模板矿工接口（骨架）
Class ITemplateMiner
    Public Function SetExtractor(extractor As IExtractor)
    End Function
    Public Function MineData(filePath As String)
    End Function
End Class

' 具体提取器：PDF
Class PDFMiner
    Implements IExtractor
    Public Function IExtractor_Extract(data As String) As String
        Response.Write("【解析 PDF】")

    End Function
End Class

' 具体提取器：CSV
Class CSVMiner
    Implements IExtractor
    Public Function IExtractor_Extract(data As String) As String
        Response.Write("【解析 CSV】")

    End Function
End Class

' 模板矿工：持有提取器引用，执行固定流程
Class DataMiner
    Implements ITemplateMiner
    Private m_Extractor As IExtractor

    Public Function ITemplateMiner_SetExtractor(extractor As IExtractor)
        Set m_Extractor = extractor
    End Function

    Public Function ITemplateMiner_MineData(filePath As String)
        Response.Write("加载文件：" & filePath)

        m_Extractor.Extract(filePath)

        Response.Write("分析数据")

        Response.Write("发送报告")

    End Function
End Class

' 演示
Dim miner As ITemplateMiner
Set miner = New DataMiner
Dim pdfExtractor As IExtractor
Set pdfExtractor = New PDFMiner
miner.SetExtractor(pdfExtractor)

miner.MineData("data.pdf")


Dim csvExtractor As IExtractor
Set csvExtractor = New CSVMiner
miner.SetExtractor(csvExtractor)

miner.MineData("data.csv")
%>