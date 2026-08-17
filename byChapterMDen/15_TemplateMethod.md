## Chapter 15: Template Method

**Core idea**: Define the skeleton of an algorithm, leaving variable steps for subclasses to implement.

**Example**: DataMiner defines a fixed mining pipeline (load → parse → analyze → send report). PDFMiner and CSVMiner only need to implement the Parse step; all other steps reuse the parent logic.

### Classic VBScript Version

```vbscript
' Data miner: defines algorithm skeleton
Class DataMiner
    ' Template method: fixed pipeline; parser argument provides the variable step
    Public Function MineData(filePath, parser)
        Load filePath
        parser.Parse
        Analyze
        SendReport
    End Function

    ' Common step: load file
    Private Function Load(filePath)
        Response.Write "Loading file: " & filePath
    End Function

    ' Common step: analyze data
    Private Function Analyze
        Response.Write "Analyzing data"
    End Function

    ' Common step: send report
    Private Function SendReport
        Response.Write "Sending report"
    End Function
End Class

' Concrete variant: PDF parsing (composition simulates inheritance)
Class PDFMiner
    Private m_Miner

    Private Sub Class_Initialize
        Set m_Miner = New DataMiner
    End Sub

    Public Function MineData(filePath)
        m_Miner.MineData filePath, Me
    End Function

    ' Override the Parse step
    Public Function Parse
        Response.Write "[Parsing PDF]"
    End Function
End Class

' Concrete variant: CSV parsing (composition simulates inheritance)
Class CSVMiner
    Private m_Miner

    Private Sub Class_Initialize
        Set m_Miner = New DataMiner
    End Sub

    Public Function MineData(filePath)
        m_Miner.MineData filePath, Me
    End Function

    ' Override the Parse step
    Public Function Parse
        Response.Write "[Parsing CSV]"
    End Function
End Class

' Demo: both variants reuse the same pipeline, only implementing Parse
Dim pdf, csv
Set pdf = New PDFMiner
pdf.MineData "data.pdf"

Set csv = New CSVMiner
csv.MineData "data.csv"
```

**Classic VBScript trade-offs**:
- **No inheritance, no abstract**: VBScript has no inheritance. "Subclass overrides Parse" is really just two independent classes with the same method name. `DataMiner.MineData` calls the passed-in `parser.Parse` — if the passed object lacks `Parse`, the error only surfaces at runtime. This doesn't match the template method's "must implement" semantics.
- **No forced override**: VBScript can't mark a method as "must override". Developers might forget to implement Parse in a subclass.

### Axon VBScript Version (supports Implements)

```vbscript
' Extractor interface (variable step)
Class IExtractor
    Public Function Extract(data As String) As String
    End Function
End Class

' Template miner interface (skeleton)
Class ITemplateMiner
    Public Function SetExtractor(extractor As IExtractor)
    End Function
    Public Function MineData(filePath As String)
    End Function
End Class

' Concrete extractor: PDF
Class PDFMiner
    Implements IExtractor
    Public Function IExtractor_Extract(data As String) As String
        Response.Write "[Parsing PDF]"
    End Function
End Class

' Concrete extractor: CSV
Class CSVMiner
    Implements IExtractor
    Public Function IExtractor_Extract(data As String) As String
        Response.Write "[Parsing CSV]"
    End Function
End Class

' Template miner: holds extractor reference, runs fixed pipeline
Class DataMiner
    Implements ITemplateMiner
    Private m_Extractor As IExtractor

    Public Function ITemplateMiner_SetExtractor(extractor As IExtractor)
        Set m_Extractor = extractor
    End Function

    Public Function ITemplateMiner_MineData(filePath As String)
        Response.Write "Loading file: " & filePath
        m_Extractor.Extract filePath
        Response.Write "Analyzing data"
        Response.Write "Sending report"
    End Function
End Class

' Demo
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

**Axon VBScript trade-offs**:
- Interface + composition achieves the separation of algorithm skeleton and variable steps. `DataMiner` holds an `IExtractor` reference, calls `m_Extractor.Extract` directly in `ITemplateMiner_MineData` — auto-dispatches to the concrete extractor without helper methods. `ITemplateMiner` interface adds `SetExtractor` method for injecting the extractor. Remaining gap: **no inheritance, no abstract method forced-override mechanism**. Classic Template Method's essence is "abstract base class defines skeleton (`MustInherit` + non-virtual `TemplateMethod`) + subclass `MustOverride` implements variable steps". AxonASP can only simulate with interface + composition, lacking compile-time "forced override" constraint (although `IExtractor` interface constrains `Extract` exists, it can't guarantee `MineData` call order is correct — `SetExtractor` must be called before `MineData`). Go also lacks inheritance and abstract methods — Go uses interface + composition for template method, same as AxonASP's current approach.

### VB.NET Version (syntactically complete baseline)

VB.NET has `MustInherit` (abstract class) + `MustOverride` (abstract method) + `Overrides` (subclass override), enabling textbook Template Method — abstract base class `MustInherit` defines algorithm skeleton, `MustOverride` forces subclasses to implement variable steps.

```vbnet
' ① Abstract base class: MustInherit prevents direct New, defines algorithm skeleton
Public MustInherit Class DataMiner
    ' ② Template method (skeleton): fixed call order, subclasses can't and don't need to change the flow
    Public Sub MineData(filePath As String)
        Load(filePath)              ' Fixed step 1: Load
        Parse(filePath)             ' Variable step: Parse (subclasses must implement)
        Analyze()                   ' Fixed step 2: Analyze
        SendReport()                ' Fixed step 3: Send report
    End Sub

    ' Fixed steps: Private, invisible to subclasses, completely prevents overriding
    Private Sub Load(filePath As String)
        Console.WriteLine($"Loading file: {filePath}")
    End Sub

    Private Sub Analyze()
        Console.WriteLine("Analyzing data")
    End Sub

    Private Sub SendReport()
        Console.WriteLine("Sending report")
    End Sub

    ' ③ Variable step: MustOverride forces all subclasses to implement, missing = compile error
    Public MustOverride Sub Parse(filePath As String)
End Class

' ④ Concrete subclass 1: PDF parser
Public Class PDFMiner
    Inherits DataMiner

    Public Overrides Sub Parse(filePath As String)
        Console.WriteLine("[Parsing PDF]")
    End Sub
End Class

' ⑤ Concrete subclass 2: CSV parser
Public Class CSVMiner
    Inherits DataMiner

    Public Overrides Sub Parse(filePath As String)
        Console.WriteLine("[Parsing CSV]")
    End Sub
End Class

' Demo: caller depends only on abstract base class, completely independent of concrete subclass
Dim pdf As DataMiner = New PDFMiner()
pdf.MineData("data.pdf")    ' Load file → [Parsing PDF] → Analyze data → Send report

Dim csv As DataMiner = New CSVMiner()
csv.MineData("data.csv")    ' Load file → [Parsing CSV] → Analyze data → Send report
```

**VB.NET version notes**:
- **`MustInherit` + `MustOverride` compile-time forced contract**: `MustInherit Class DataMiner` prevents `New DataMiner()` (only concrete subclasses can be instantiated), `MustOverride Sub Parse` forces all subclasses to implement `Parse`, missing implementation causes compile error. Axon version uses `IExtractor` interface to constrain `Extract` exists, but `SetExtractor` + `MineData` call order relies on developer discipline — compiler can't guarantee it.
- **Inheritance binds skeleton and steps**: `MineData` defines fixed call order in base class, subclasses can only `Overrides` the variable step `Parse`, can't modify the overall flow. Axon version simulates with interface + composition, implementation classes can freely change `MineData` step order.
- **Inheritance reuses common logic**: `Load`/`Analyze`/`SendReport` fixed steps written once in base class, all subclasses automatically get them. Axon version extractors are parallel classes, no inheritance reuse.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Skeleton + step binding | Loose: `MineData(filePath, parser)` manually passes parser | Interface composition: `SetExtractor` + `MineData` two-step call (order relies on discipline) | Inheritance binding: base class `MineData` calls `MustOverride Parse` (compile-time enforced) |
| Step forced implementation | Method name convention (missing = runtime error) | `IExtractor` interface constrains `Extract` exists | `MustOverride` compile error (missing = subclass won't compile) |
| Skeleton tamper prevention | None (anyone can change `MineData` flow) | None (implementation classes can freely change `MineData` order) | Inheritance binding, subclasses can only `Overrides` variable steps |
| Code reuse | None (PDFMiner/CSVMiner each write their own composition boilerplate) | None (parallel extractor classes, no inheritance) | Base class fixed steps automatically inherited by all subclasses |
---
