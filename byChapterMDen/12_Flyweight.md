## Chapter 12: Flyweight

**Core idea**: Share fine-grained objects to reduce memory usage.

**Example**: A forest has many trees, but the "types" (name + color) are only a few. TreeFactory uses a Dictionary to cache TreeType — the same config is created only once, and multiple Tree instances share the same TreeType.

### Classic VBScript Version

```vbscript
' Flyweight: intrinsic tree properties (name, color), shared across trees
Class TreeType
    Public Name, Color

    ' Draw tree at given coordinates
    Public Function Draw(x, y)
        Response.Write "Draw " & Color & " " & Name & " at (" & x & "," & y & ")"
    End Function
End Class

' Flyweight factory: cache and reuse TreeType objects
Class TreeFactory
    Private m_Types   ' Dictionary: key → TreeType

    ' Constructor: create the dictionary
    Private Sub Class_Initialize
        Set m_Types = CreateObject("Scripting.Dictionary")
    End Sub

    ' Get or create TreeType: same args return the same object
    ' name: tree name, color: color
    ' Returns: shared TreeType instance
    Public Function GetTreeType(name, color)
        Dim key
        key = name & "|" & color
        If Not m_Types.Exists(key) Then
            Dim t
            Set t = New TreeType
            t.Name = name
            t.Color = color
            Set m_Types(key) = t
        End If
        Set GetTreeType = m_Types(key)
    End Function
End Class

' Demo: 3 trees share the same TreeType object
Dim factory, oakType, i
Set factory = New TreeFactory
Set oakType = factory.GetTreeType("Oak", "Green")

For i = 0 To 2
    oakType.Draw i, i * 2
Next
Response.Write "3 trees, but only 1 TreeType object"
```

**Classic VBScript trade-offs**:
This pattern maps naturally to VBScript. `Scripting.Dictionary` provides exactly the "cache objects by key" capability the flyweight factory needs — a perfect fit. The only annoyance is that Dictionary requires explicit `Set` when storing/retrieving objects.

### Axon VBScript Version (supports strong typing)

```vbscript
' Flyweight: intrinsic tree properties (name, color), shared across trees
Class TreeType
    Public Name As String
    Public Color As String

    ' Draw tree at given coordinates
    Public Function Draw(x As Long, y As Long)
        Response.Write "Draw " & Color & " " & Name & " at (" & x & "," & y & ")"
    End Function
End Class

' Flyweight factory: cache and reuse TreeType objects
Class TreeFactory
    ' Note: Scripting.Dictionary is a COM object; As can't annotate it, stays Variant
    Private m_Types

    ' Constructor: create the dictionary
    Private Sub Class_Initialize
        Set m_Types = CreateObject("Scripting.Dictionary")
    End Sub

    ' Get or create TreeType: same args return the same object
    ' name: tree name, color: color
    ' Returns: shared TreeType instance
    Public Function GetTreeType(name As String, color As String) As TreeType
        Dim key As String
        key = name & "|" & color
        If Not m_Types.Exists(key) Then
            Dim t As TreeType
            Set t = New TreeType
            t.Name = name
            t.Color = color
            Set m_Types(key) = t
        End If
        Set GetTreeType = m_Types(key)
    End Function
End Class

' Demo: 3 trees share the same TreeType object
Dim factory As TreeFactory
Dim oakType As TreeType
Dim i As Long
Set factory = New TreeFactory
Set oakType = factory.GetTreeType("Oak", "Green")

For i = 0 To 2
    oakType.Draw i, i * 2
Next
Response.Write "3 trees, but only 1 TreeType object"
```

**Axon VBScript trade-offs**:
- The classic version's `Scripting.Dictionary` for object caching remains the best approach. The strongly-typed version annotates `TreeType` fields, `GetTreeType` parameters and return type, and loop variables with types, letting the IDE auto-complete members after `oakType.`. The only field that can't be typed is `m_Types` — `Scripting.Dictionary` is a COM object and `As` can't annotate it. This is a COM interop limitation, unrelated to the Flyweight pattern itself.

---
