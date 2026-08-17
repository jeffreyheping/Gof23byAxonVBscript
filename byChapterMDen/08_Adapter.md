## Chapter 8: Adapter

**Core idea**: Convert an incompatible interface into the target interface.

**Example**: OldPrinter only has `OldPrint` (takes a string). The new system expects `Print` (takes a Document object). PrinterAdapter bridges the two — it extracts `doc.Content` and passes it to `OldPrint`.

### Classic VBScript Version

```vbscript
' Legacy class: only has OldPrint, takes a string
Class OldPrinter
    ' Old interface: print a string directly
    Public Function OldPrint(s)
        Response.Write "[Old Printer] " & s
    End Function
End Class

' Adapter: converts old interface to new interface
Class PrinterAdapter
    Private m_OldPrinter

    ' Inject the adapted legacy object
    Public Function Init(oldPrinter)
        Set m_OldPrinter = oldPrinter
    End Function

    ' New interface: takes a Document, extracts Content, delegates to old interface
    Public Function Print(doc)
        m_OldPrinter.OldPrint doc.Content
    End Function
End Class

' New system's data carrier
Class Document
    Public Content
End Class

' Demo: call the old printer through the new Print interface
Dim doc
Set doc = New Document
doc.Content = "Hello World"

Dim adapter
Set adapter = New PrinterAdapter
adapter.Init New OldPrinter
adapter.Print doc   ' Call old printer via new interface
```

**Classic VBScript trade-offs**:
- **No target interface**: `PrinterAdapter` has no `IPrinter` interface to implement. The "new interface" is just a convention on the `Print` method name. With multiple adapters, there's no guarantee of consistency.

### Axon VBScript Version (supports Implements)

```vbscript
' Target interface
Class IPrinter
    Public Function Print(doc As Document)
    End Function
End Class

' Legacy class
Class OldPrinter
    Public Function OldPrint(s As String)
        Response.Write "[Old Printer] " & s
    End Function
End Class

' New system's data carrier
Class Document
    Public Content As String
End Class

' Adapter: implements target interface, internally holds legacy object
Class PrinterAdapter
    Implements IPrinter
    Private m_OldPrinter As OldPrinter

    Public Function Init(oldPrinter As OldPrinter)
        Set m_OldPrinter = oldPrinter
    End Function

    Public Function IPrinter_Print(doc As Document)
        m_OldPrinter.OldPrint doc.Content
    End Function
End Class

' Demo: use adapter through interface
Dim doc As Document
Set doc = New Document
doc.Content = "Hello World"

Dim adapter As PrinterAdapter
Set adapter = New PrinterAdapter
adapter.Init New OldPrinter

Dim ip As IPrinter
Set ip = adapter
ip.Print doc
```

**Axon VBScript trade-offs**:
- `Implements` interface mechanism solves the target contract problem: `IPrinter` forces the adapter to implement `Print`, multiple adapters share the same interface signature, callers uniformly invoke via `IPrinter` reference without caring about specific adapter type.
- Missing syntax: **code reuse mechanism (inheritance)**. Classic Adapter has two variants — class adapter and object adapter. Class adapter uses multiple inheritance to inherit both Target and Adaptee, directly reusing methods from both sides without manual forwarding. AxonASP can only use object adapter (composing `OldPrinter`), every adaptation method requires hand-written forwarding code, boilerplate grows linearly when new methods are added. Go also lacks inheritance, but Go uses struct embedding (`PrinterAdapter struct { *OldPrinter }`) to promote embedded methods to the outer level — only `Print` needs rewriting, all other methods pass through at zero cost. AxonASP currently requires manual delegation for every method.

### VB.NET Version (syntactically complete baseline)

VB.NET has full `Interface` + `Implements` syntax. Object adapter structure corresponds one-to-one with Axon version: `PrinterAdapter Implements IPrinter`, internally composes `OldPrinter`, forwards `Print` to `OldPrint`.

```vbnet
' ① Target interface: new system's expected contract
Public Interface IPrinter
    Sub Print(doc As Document)
End Interface

' ② Adaptee (legacy class): incompatible interface, only has OldPrint
Public Class OldPrinter
    Public Sub OldPrint(s As String)
        Console.WriteLine("[Old Printer] " & s)
    End Sub
End Class

' ③ New system data carrier
Public Class Document
    Public Property Content As String
End Class

' ④ Adapter: Implements IPrinter, internally composes OldPrinter, forwards Print
Public Class PrinterAdapter
    Implements IPrinter

    Private ReadOnly m_OldPrinter As OldPrinter

    ' Parameterized constructor: inject adaptee at creation, no Init two-step
    Public Sub New(oldPrinter As OldPrinter)
        m_OldPrinter = oldPrinter
    End Sub

    ' Implements IPrinter.Print, compiler enforces matching signature
    Public Sub Print(doc As Document) Implements IPrinter.Print
        m_OldPrinter.OldPrint(doc.Content)
    End Sub
End Class

' Demo: use adapter via interface reference
Dim doc As New Document With {.Content = "Hello World"}
Dim adapter As IPrinter = New PrinterAdapter(New OldPrinter())
adapter.Print(doc)   ' [Old Printer] Hello World
```

**VB.NET version notes**:
- **`Interface` + `Implements` compile-time contract**: `Implements IPrinter.Print` binds the adapter to the target interface signature, missing or misspelled methods cause compile errors. Axon's `IPrinter_Print` relies on naming convention, errors only surface at runtime.
- **Parameterized constructor replaces Init**: `New PrinterAdapter(New OldPrinter())` injects adaptee at creation, no "New then Init" half-initialized window.
- **No `Set` needed**: VB.NET object assignment uses `=` directly, no `Set`/`Let` distinction.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Target contract | Method name convention (honor system) | `Implements IPrinter` interface constraint | `Interface` + `Implements` compile-time enforced |
| Adaptation approach | Object composition | Object composition | Object composition |
| Adaptee injection | `Init` two-step (easy to forget) | `Init` two-step (easy to forget) | Parameterized constructor `New(OldPrinter)` one-step |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |
---
