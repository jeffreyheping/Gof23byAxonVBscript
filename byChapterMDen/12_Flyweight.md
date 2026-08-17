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
- The classic version's `Scripting.Dictionary` for object caching remains the best approach (already idiomatic in classic). The strongly-typed version annotates `TreeType` fields, `GetTreeType` parameters and return type, and loop variables with types, letting the IDE auto-complete members after `oakType.`. The only field that can't be typed is `m_Types` — `Scripting.Dictionary` is a COM object and `As` doesn't support annotating it. This is a COM interop limitation, unrelated to the Flyweight Pattern itself.

### VB.NET Version (syntactically complete baseline)

Classic version is already idiomatic. VB.NET uses generic `Dictionary(Of String, TreeType)` instead of COM `Scripting.Dictionary`, with compile-time type safety, no `Set` needed for storage. Same scenario as Axon version: 3 trees sharing the same TreeType.

```vbnet
' ① Flyweight object: tree "type" (name, color), shared by multiple trees
Public Class TreeType
    Public ReadOnly Property Name As String
    Public ReadOnly Property Color As String

    Public Sub New(name As String, color As String)
        Me.Name = name
        Me.Color = color
    End Sub

    ' Draw tree at given coordinates (x, y are external state, passed by caller)
    Public Sub Draw(x As Long, y As Long)
        Console.WriteLine($"Drawing {Color}{Name} at ({x},{y})")
    End Sub
End Class

' ② Flyweight factory: cache and reuse TreeType
Public Class TreeFactory
    ' Generic dictionary: Key is always String, Value is always TreeType
    Private ReadOnly m_Types As New Dictionary(Of String, TreeType)()

    ' Get or create TreeType: same parameters return same object
    Public Function GetTreeType(name As String, color As String) As TreeType
        Dim key = name & "|" & color
        If Not m_Types.ContainsKey(key) Then
            m_Types(key) = New TreeType(name, color)
        End If
        Return m_Types(key)
    End Function
End Class

' Demo: 3 trees share the same TreeType object
Dim factory As New TreeFactory()
Dim oakType = factory.GetTreeType("Oak", "Green")

For i = 0 To 2
    oakType.Draw(i, i * 2)
Next
Console.WriteLine("3 trees, but only 1 TreeType object")
' Drawing Green Oak at (0,0)
' Drawing Green Oak at (1,2)
' Drawing Green Oak at (2,4)
```

**VB.NET version notes**:
- **Generic `Dictionary(Of String, TreeType)` compile-time type safety**: Key and Value types are locked in angle brackets, `m_Types(key)` returns strongly-typed `TreeType`, no `CType` conversion needed. COM `Scripting.Dictionary` Key/Value are both `Object`, type errors only surface at runtime.
- **`ReadOnly` property + constructor assignment**: `TreeType` Name/Color are immutable after creation, preventing shared objects from being accidentally modified. Axon/classic version `Public Name` is a mutable field.
- **No `Set` needed**: Object assignment uses `=` directly, no need for verbose `Set m_Types(key) = t` syntax.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Cache dictionary type | COM `Scripting.Dictionary` (all Object) | COM `Scripting.Dictionary` (can't annotate types) | Generic `Dictionary(Of String, TreeType)` (compile-time locked types) |
| Type safety | None, all Variant conversion | Strong typing outside, Dictionary inside still Object | Full-chain strong typing (no conversion needed) |
| Flyweight mutability | `Public Name` mutable field | Same (with `As String`) | `ReadOnly` property + constructor assignment, immutable |
| Object assignment | `Set m_Types(key) = t` | Same (COM limitation) | Direct `m_Types(key) = New TreeType(...)` |
---
