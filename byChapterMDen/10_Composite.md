## Chapter 10: Composite

**Core idea**: Treat individual objects and compositions of objects uniformly (tree structure).

**Example**: Leaf is a leaf node (employee), Composite is a branch node (department). Both have the same `Operation` method. Composite recursively calls `Operation` on all children — whether leaf or branch, the same method traverses the whole tree.

### Classic VBScript Version

```vbscript
' Leaf node: the end of the tree
Class Leaf
    Public Name
    ' Display self; indent controls indentation level
    Public Function Operation(indent)
        Response.Write indent & "Leaf: " & Name
    End Function
End Class

' Composite node: can contain child nodes (Leaf or Composite)
Class Composite
    Public Name
    Private m_Children()   ' Child node array
    Private m_Count        ' Current child count

    ' Constructor: initialize array
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Children(10)
    End Sub

    ' Add child (auto-resize when full)
    Public Function Add(child)
        If m_Count >= UBound(m_Children) + 1 Then
            ReDim Preserve m_Children(m_Count * 2)
        End If
        Set m_Children(m_Count) = child
        m_Count = m_Count + 1
    End Function

    ' Display self, then recursively call Operation on all children
    Public Function Operation(indent)
        Response.Write indent & "Composite: " & Name
        Dim i
        For i = 0 To m_Count - 1
            m_Children(i).Operation indent & "  "
        Next
    End Function
End Class

' Demo: build a HQ → Branch → Employee tree
Dim root, branch1, leaf1, leaf2, leaf3
Set root = New Composite
root.Name = "HQ"

Set branch1 = New Composite
branch1.Name = "Branch"

Set leaf1 = New Leaf
leaf1.Name = "Alice"
Set leaf2 = New Leaf
leaf2.Name = "Bob"
Set leaf3 = New Leaf
leaf3.Name = "Charlie"

branch1.Add leaf1
branch1.Add leaf2
root.Add branch1
root.Add leaf3

root.Operation ""   ' Traverse the whole tree uniformly
```

**Classic VBScript trade-offs**:
- **No common base class**: The classic Composite pattern requires Leaf and Composite to inherit from the same `Component` base class. VBScript has no inheritance — they're completely independent classes, relying only on the `Operation` method-name convention ("duck typing"). The compiler can't ensure type safety.
- **No type safety**: The `Add` method takes an untyped `child` parameter — any object could be passed. The error only surfaces when `Operation` is called at runtime.

### Axon VBScript Version (supports Implements)

```vbscript
' Component interface
Class IComponent
    Public Function Operation(indent As String)
    End Function
    Public Function Add(child As IComponent)
    End Function
End Class

' Leaf node
Class Leaf
    Implements IComponent
    Private m_Name As String

    Public Property Get Name As String
        Name = m_Name
    End Property
    Public Property Let Name(v As String)
        m_Name = v
    End Property

    Public Function IComponent_Operation(indent As String)
        Response.Write indent & "Leaf: " & m_Name
    End Function
    Public Function IComponent_Add(child As IComponent)
        ' Leaf has no children — no-op
    End Function
End Class

' Composite node
Class Composite
    Implements IComponent
    Private m_Name As String
    Private m_Children    ' Collection

    Public Property Get Name As String
        Name = m_Name
    End Property
    Public Property Let Name(v As String)
        m_Name = v
    End Property

    Private Sub Class_Initialize
        Set m_Children = Server.CreateObject("Collection")
    End Sub

    Public Function IComponent_Add(child As IComponent)
        m_Children.Add child
    End Function

    Public Function IComponent_Operation(indent As String)
        Response.Write indent & "Composite: " & m_Name
        Dim child As IComponent
        For Each child In m_Children
            child.Operation indent & "  "
        Next
    End Function
End Class

' Demo: traverse tree via uniform interface
Dim rootObj As Composite, branch1Obj As Composite
Dim leaf1Obj As Leaf, leaf2Obj As Leaf, leaf3Obj As Leaf
Set rootObj = New Composite
rootObj.Name = "HQ"
Set branch1Obj = New Composite
branch1Obj.Name = "Branch"
Set leaf1Obj = New Leaf: leaf1Obj.Name = "Alice"
Set leaf2Obj = New Leaf: leaf2Obj.Name = "Bob"
Set leaf3Obj = New Leaf: leaf3Obj.Name = "Charlie"

Dim root As IComponent, branch1 As IComponent
Dim leaf1 As IComponent, leaf2 As IComponent, leaf3 As IComponent
Set root = rootObj
Set branch1 = branch1Obj
Set leaf1 = leaf1Obj
Set leaf2 = leaf2Obj
Set leaf3 = leaf3Obj

root.Add branch1
root.Add leaf3
branch1.Add leaf1
branch1.Add leaf2

root.Operation ""
```

**Axon VBScript trade-offs**:
- The interface mechanism unifies the contract for leaves and composites. `Composite` calls `child.Operation` directly through `IComponent`-typed child references, recursively traversing the whole tree — matching the classic Composite's transparent composition semantics. Child storage uses the built-in `Collection` with `For Each` iteration — no manual array management needed.
- Missing syntax: **Code reuse mechanism** (inheritance or struct embedding). `Leaf` and `Composite` can't share a common `Component` base class to reuse default implementations. Go uses struct embedding to solve this; AxonASP requires each to implement separately.
- `Leaf.IComponent_Add` must be a no-op: the interface requires all implementors to provide `Add`, but a leaf shouldn't support adding children. There's no way to prohibit this at the type level — we can only silently ignore it at runtime.
---
