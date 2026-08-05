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
- `ICloneable` guarantees all prototype classes have a `Clone` method, and you can call `Clone()` directly through an interface reference. But manual field-by-field copying is still required — the interface solves the contract problem, not the syntactic sugar problem. VBScript still has no built-in deep copy or serialization.
- Missing syntax: **Built-in deep copy**. Go also lacks inheritance, but Go also has no built-in Clone — Go's approach is for each type to implement its own `Clone()` method (same as VBScript), or use serialization (`encoding/gob`) for deep copies. The real pain point here isn't "no inheritance" — it's the lack of automatic deep-copy syntax sugar.
---
