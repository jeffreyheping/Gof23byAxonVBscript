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
- Interface + composition achieves the separation of algorithm skeleton and variable steps. `DataMiner` holds an `IExtractor` reference and calls `m_Extractor.Extract` directly in `ITemplateMiner_MineData` — auto-dispatches to the concrete extractor. `ITemplateMiner` adds `SetExtractor` for injecting the extractor. Remaining gap: **No abstract method forced-override mechanism**. Go also lacks inheritance and abstract methods — Go uses interface + composition for template method, same as AxonASP's current approach. The real pain point is that `SetExtractor` must be called before `MineData`, otherwise `m_Extractor` is `Nothing` and you get a runtime error. The compiler can't enforce the call order. Missing syntax: **Parameterized constructor** — if a constructor could inject `IExtractor` at creation time, the call-order problem wouldn't exist.
---
