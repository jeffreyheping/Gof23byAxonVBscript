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
- This pattern maps naturally to AxonASP. The interface mechanism solves the core polymorphism problem with no significant trade-offs.
---
