## 第15章 模板方法模式（Template Method）

**核心思想**：定义算法骨架，把可变步骤留给子类实现。

**示例说明**：DataMiner 定义了 MineData 的固定流程（加载→解析→分析→发送报告）。PDFMiner 和 CSVMiner 只需各自实现 Parse 方法，其他步骤复用父类逻辑。

### 传统 VBScript 版

```vbscript
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

```

**传统 VBScript 版妥协说明**：
- **无继承、无 abstract**：VBScript 没有继承，"子类覆盖 Parse"只能靠同名方法覆盖（实际是两个独立类）。`DataMiner` 的 `MineData` 里调用的是传入的 `parser.Parse`，如果传入的对象没有 `Parse` 方法，运行时才报错——这不符合模板方法"必须实现"的语义。
- **无强制覆盖**：VBScript 无法标记某个方法为"必须覆盖"，开发者可能忘记在子类中实现 Parse。

### Axon VBScript 版（支持 Implements）

```vba
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

```

**Axon VBScript 版妥协说明**：
- 接口 + 组合实现了流程骨架与可变步骤的分离。`DataMiner` 持有 `IExtractor` 引用，在 `ITemplateMiner_MineData` 中直接调用 `m_Extractor.Extract` 即可自动路由到具体提取器实现，无需辅助方法。`ITemplateMiner` 接口增加了 `SetExtractor` 方法用于注入提取器。残留限制：缺失语法点：**无继承、无 abstract 方法强制覆盖机制**。经典模板方法的精髓是"抽象基类定义骨架（`MustInherit` + 非虚 `TemplateMethod`）+ 子类 `MustOverride` 实现可变步骤"，AxonASP 只能用接口+组合模拟，缺少编译期的"强制覆盖"约束（虽然 `IExtractor` 接口能约束 `Extract` 存在，但无法保证 `MineData` 调用顺序正确，`SetExtractor` 必须在 `MineData` 之前调用）。Go 同样无继承、无 abstract 方法——Go 用接口 + 组合实现模板方法，与 AxonASP 当前做法一致。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有 `MustInherit`（抽象类）+ `MustOverride`（抽象方法）+ `Overrides`（子类重写），可以写出教科书式的模板方法——抽象基类 `MustInherit` 定义算法骨架，`MustOverride` 强制子类实现可变步骤。

```vbnet
' ① 抽象基类：MustInherit 禁止直接 New，定义算法骨架
Public MustInherit Class DataMiner
    ' ② 模板方法（骨架）：调用顺序固定，子类无需也无法改流程
    Public Sub MineData(filePath As String)
        Load(filePath)              ' 固定步骤1：加载
        Parse(filePath)             ' 可变步骤：解析（子类必须实现）
        Analyze()                   ' 固定步骤2：分析
        SendReport()                ' 固定步骤3：发送报告
    End Sub

    ' 固定步骤：Private 子类不可见，彻底防止被覆盖
    Private Sub Load(filePath As String)
        Console.WriteLine($"加载文件：{filePath}")
    End Sub

    Private Sub Analyze()
        Console.WriteLine("分析数据")
    End Sub

    Private Sub SendReport()
        Console.WriteLine("发送报告")
    End Sub

    ' ③ 可变步骤：MustOverride 强制所有子类必须实现，漏写直接编译报错
    Public MustOverride Sub Parse(filePath As String)
End Class

' ④ 具体子类1：PDF 解析器
Public Class PDFMiner
    Inherits DataMiner

    Public Overrides Sub Parse(filePath As String)
        Console.WriteLine("【解析 PDF】")
    End Sub
End Class

' ⑤ 具体子类2：CSV 解析器
Public Class CSVMiner
    Inherits DataMiner

    Public Overrides Sub Parse(filePath As String)
        Console.WriteLine("【解析 CSV】")
    End Sub
End Class

' 演示：调用方只依赖抽象基类，完全不关心具体子类实现
Dim pdf As DataMiner = New PDFMiner()
pdf.MineData("data.pdf")    ' 加载文件 → 【解析 PDF】→ 分析数据 → 发送报告

Dim csv As DataMiner = New CSVMiner()
csv.MineData("data.csv")    ' 加载文件 → 【解析 CSV】→ 分析数据 → 发送报告
```

**VB.NET 版说明**：
- **`MustInherit` + `MustOverride` 编译期强制契约**：`MustInherit Class DataMiner` 禁止 `New DataMiner()`（只能实例化具体子类），`MustOverride Sub Parse` 强制所有子类必须实现 `Parse`，漏写直接编译报错。Axon 版虽用 `IExtractor` 接口约束 `Extract` 存在，但 `SetExtractor` + `MineData` 的调用顺序靠开发者自觉，编译期无法保证。
- **继承绑定骨架与步骤**：`MineData` 在基类定义固定调用顺序，子类只能 `Overrides` 可变步骤 `Parse`，无法修改整体流程。Axon 版用接口+组合模拟，实现类可随意改 `MineData` 内步骤顺序。
- **继承复用公共逻辑**：`Load`/`Analyze`/`SendReport` 等固定步骤在基类写一次，所有子类自动获得。Axon 版各提取器是平行类，无继承复用。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 骨架+步骤绑定 | 松散：`MineData(filePath, parser)` 手动传入 parser | 接口组合：`SetExtractor` + `MineData` 两步调用（顺序靠自觉） | 继承绑定：基类 `MineData` 调用 `MustOverride Parse`（编译期强制） |
| 步骤强制实现 | 方法名约定（漏写运行时报错） | `IExtractor` 接口约束 `Extract` 存在 | `MustOverride` 编译期报错（漏写子类无法通过编译） |
| 骨架防止篡改 | 无（任何人都可以改 `MineData` 流程） | 无（实现类可随意改 `MineData` 顺序） | 继承绑定，子类只能 `Overrides` 可变步骤 |
| 代码复用 | 无（PDFMiner/CSVMiner 各写一份组合样板） | 无（各提取器平行类，无继承） | 基类固定步骤自动传给所有子类 |
---