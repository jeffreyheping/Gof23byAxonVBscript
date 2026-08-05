## 第15章 模板方法模式（Template Method）

**核心思想**：定义算法骨架，把可变步骤留给子类实现。

**示例说明**：DataMiner 定义了 MineData 的固定流程（加载→解析→分析→发送报告）。PDFMiner 和 CSVMiner 只需各自实现 Parse 方法，其他步骤复用父类逻辑。

### 传统 VBScript 版

```vbscript
' 数据挖掘器：定义算法骨架
Class DataMiner
    ' 模板方法：固定流程，parser 参数提供可变步骤
    Public Function MineData(filePath, parser)
        Load filePath
        parser.Parse
        Analyze
        SendReport
    End Function

    ' 公共步骤：加载文件
    Private Function Load(filePath)
        Response.Write "加载文件：" & filePath
    End Function

    ' 公共步骤：分析数据
    Private Function Analyze
        Response.Write "分析数据"
    End Function

    ' 公共步骤：发送报告
    Private Function SendReport
        Response.Write "发送报告"
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
        Response.Write "【解析 PDF】"
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
        Response.Write "【解析 CSV】"
    End Function
End Class

' 演示：两个子类复用相同流程，只需实现 Parse
Dim pdf, csv
Set pdf = New PDFMiner
pdf.MineData "data.pdf"

Set csv = New CSVMiner
csv.MineData "data.csv"
```

**传统 VBScript 版妥协说明**：
- **无继承、无 abstract**：VBScript 没有继承，"子类覆盖 Parse"只能靠同名方法覆盖（实际是两个独立类）。`DataMiner` 的 `MineData` 里调用的是传入的 `parser.Parse`，如果传入的对象没有 `Parse` 方法，运行时才报错——这不符合模板方法"必须实现"的语义。
- **无强制覆盖**：VBScript 无法标记某个方法为"必须覆盖"，开发者可能忘记在子类中实现 Parse。

### Axon VBScript 版（支持 Implements）

```vbscript
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
        Response.Write "【解析 PDF】"
    End Function
End Class

' 具体提取器：CSV
Class CSVMiner
    Implements IExtractor
    Public Function IExtractor_Extract(data As String) As String
        Response.Write "【解析 CSV】"
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
        Response.Write "加载文件：" & filePath
        m_Extractor.Extract filePath
        Response.Write "分析数据"
        Response.Write "发送报告"
    End Function
End Class

' 演示
Dim miner As ITemplateMiner
Set miner = New DataMiner
Dim pdfExtractor As IExtractor
Set pdfExtractor = New PDFMiner
miner.SetExtractor pdfExtractor
miner.MineData "data.pdf"

Dim csvExtractor As IExtractor
Set csvExtractor = New CSVMiner
miner.SetExtractor csvExtractor
miner.MineData "data.csv"
```

**Axon VBScript 版妥协说明**：
- 接口 + 组合实现了流程骨架与可变步骤的分离。`DataMiner` 持有 `IExtractor` 引用，在 `ITemplateMiner_MineData` 中直接调用 `m_Extractor.Extract` 即可自动路由到具体提取器实现，无需辅助方法。`ITemplateMiner` 接口增加了 `SetExtractor` 方法用于注入提取器。残留限制：缺失语法点：**无抽象方法强制覆盖机制**。Go 同样无继承、无 abstract 方法——Go 用接口 + 组合实现模板方法，与 AxonASP 当前做法一致。真正的痛点是 `SetExtractor` 必须在 `MineData` 之前调用，否则 `m_Extractor` 为 `Nothing` 运行时报错，编译期无法保证调用顺序。缺失语法点：**带参构造函数**——若有构造函数可在创建时注入 `IExtractor`，则不存在调用顺序问题。
---