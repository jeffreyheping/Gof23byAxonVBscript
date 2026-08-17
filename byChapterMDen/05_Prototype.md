## Chapter 5: Prototype

**Core idea**: Create new objects by copying existing ones.

**Example**: Create a resume template, then Clone it to produce copies. Modifying the clone's skills array doesn't affect the original — proving it's a deep copy.

### Classic VBScript Version

```vbscript
' Resume class: has Name, Age, and a Skills array
Class MyResume
    Public Name, Age, Skills

    ' Shallow clone: create a new MyResume and copy fields one by one
    ' Array elements are copied by value; in this example Skills is a string array,
    ' so modifying the copy won't affect the original
    ' Returns: a new MyResume instance
    Public Function Clone
        Dim copy
        Set copy = New MyResume
        copy.Name = Me.Name
        copy.Age = Me.Age
        Dim ub, i, arr
        ub = UBound(Me.Skills)
        ReDim arr(ub)
        For i = 0 To ub
            arr(i) = Me.Skills(i)
        Next
        copy.Skills = arr
        Set Clone = copy
    End Function
End Class

' Demo: clone, then modify the copy — original is unaffected
Dim r1, r2
Set r1 = New MyResume
r1.Name = "Zhang San"
r1.Age = 25
r1.Skills = Array("VBScript", "HTML")

Set r2 = r1.Clone
r2.Name = "Li Si"
r2.Skills(0) = "JavaScript"

Response.Write r1.Name & " " & r1.Skills(0)   ' Zhang San VBScript
Response.Write r2.Name & " " & r2.Skills(0)   ' Li Si JavaScript
```

**Classic VBScript trade-offs**:
- **No built-in Clone**: VBScript has no `Clone()` method or serialization mechanism. You must copy fields manually. The more fields, the longer the Clone method — and it's easy to forget adding a new field to Clone.
- **No ICloneable interface**: There's no way to require all classes to implement Clone. It's up to the developer.

### Axon VBScript Version (supports Implements)

```vbscript
' Clone interface
Class ICloneable
    Public Function Clone As ICloneable
    End Function
End Class

' Resume class: implements ICloneable
Class MyResume
    Implements ICloneable
    Public Name As String
    Public Age As Integer
    Public Skills

    ' Deep clone: copy each field; copy array element by element
    Public Function ICloneable_Clone As ICloneable
        Dim copy As MyResume
        Set copy = New MyResume
        copy.Name = Me.Name
        copy.Age = Me.Age
        Dim ub As Integer, i As Integer, arr
        ub = UBound(Me.Skills)
        ReDim arr(ub)
        For i = 0 To ub
            arr(i) = Me.Skills(i)
        Next
        copy.Skills = arr
        Set ICloneable_Clone = copy
    End Function
End Class

' Demo: call Clone through interface reference
Dim r1 As MyResume
Dim r2 As ICloneable
Dim r2Copy As MyResume
Set r1 = New MyResume
r1.Name = "Zhang San"
r1.Age = 25
r1.Skills = Array("VBScript", "HTML")
Set r2 = r1.Clone()
Set r2Copy = r2
r2Copy.Name = "Li Si"
r2Copy.Skills(0) = "JavaScript"

Response.Write r1.Name & " " & r1.Skills(0)   ' Zhang San VBScript
Response.Write r2Copy.Name & " " & r2Copy.Skills(0)   ' Li Si JavaScript
```

**Axon VBScript trade-offs**:
- `ICloneable` interface guarantees all prototype classes have a `Clone` method, and you can call `Clone()` directly through interface references for automatic dispatch to concrete implementations. Remaining gap: **deep copy still requires manual implementation**. The interface solves the contract problem, not the syntax sugar problem — VBScript has no built-in deep copy or serialization mechanism. Every added field, every nested object level, requires manually adding Clone logic (if Skills were an object array instead of a string array, each object would need its own Clone call; more nesting means more code). Go also has no built-in Clone; Go's approach is for each type to implement its own `Clone()` (same as VBScript), or use `encoding/gob` serialization for generic deep copy. The real pain point here is **lacking automatic deep-copy syntax sugar**, not inheritance.

### VB.NET Version (syntactically complete baseline)

VB.NET implements the standard `System.ICloneable` interface, with manual deep copy of fields. Same scenario as Axon version: just MyResume class + Name/Age/Skills fields.

```vbnet
' ① Resume class: implements System.ICloneable interface, fields same as Axon version
Public Class MyResume
    Implements ICloneable

    Public Property Name As String
    Public Property Age As Integer
    Public Property Skills As String()   ' Same as Axon version, using array

    ' Standard ICloneable interface method: manual deep copy
    Public Function Clone() As Object Implements ICloneable.Clone
        Dim copy As New MyResume() With {
            .Name = Me.Name,
            .Age = Me.Age
        }
        ' Array deep copy: element-by-element (same logic as Axon version)
        If Me.Skills IsNot Nothing Then
            copy.Skills = New String(Me.Skills.Length - 1) {}
            Array.Copy(Me.Skills, copy.Skills, Me.Skills.Length)
        End If
        Return copy
    End Function
End Class

' Demo: modify clone, original unaffected
Dim original As New MyResume() With {
    .Name = "Zhang San",
    .Age = 25,
    .Skills = {"VBScript", "HTML"}
}

Dim clone As MyResume = DirectCast(original.Clone(), MyResume)
clone.Name = "Li Si"
clone.Skills(0) = "JavaScript"

Console.WriteLine(original.Name & " " & original.Skills(0))   ' Zhang San VBScript
Console.WriteLine(clone.Name & " " & clone.Skills(0))         ' Li Si JavaScript
```

**VB.NET version notes**:
- **Standard `System.ICloneable` interface**: .NET BCL's built-in universal contract, recognized by all framework class libraries. Axon version needs to define its own `Class ICloneable` empty shell.
- **Array deep copy with `Array.Copy`**: VB.NET has standard library `Array.Copy` for one-line array copying. Axon version needs manual `ReDim` + `For` loop element-by-element assignment.
- **No `Set` needed**: VB.NET object assignment uses `=` directly, `Dim clone As Resume = DirectCast(...)` doesn't need `Set`.
- **Deep copy still manual**: Same as Axon version — every added reference type field requires adding copy logic in Clone. This is Prototype pattern's inherent pain point, VB.NET is no exception.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Clone contract | None (method name convention) | Custom `ICloneable` interface | Standard `System.ICloneable` interface |
| Deep copy implementation | Manual `ReDim` + `For` loop | Manual `ReDim` + `For` loop | `Array.Copy` one-line array copy |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |
| Interface dispatch | No interface | `r1.Clone()` via interface dispatch | `DirectCast(original.Clone(), Resume)` |
---
