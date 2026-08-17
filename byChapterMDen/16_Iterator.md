## Chapter 16: Iterator

**Core idea**: Provide a way to sequentially access elements of an aggregate object without exposing its internal structure.

**Example**: MyCollection maintains an internal array. MyIterator holds the current index and traverses elements via `HasNext` and `NextItem` — the caller doesn't need to know whether the internals are an array or a linked list.

### Classic VBScript Version

```vbscript
' Collection class: stores data in an internal array
Class MyCollection
    Private m_Items()   ' Internal array
    Private m_Count     ' Current element count

    ' Constructor: initialize array
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Items(10)
    End Sub

    ' Add element (auto-resize when full)
    Public Function Add(item)
        If m_Count >= UBound(m_Items) + 1 Then
            ReDim Preserve m_Items(m_Count * 2)
        End If
        m_Items(m_Count) = item
        m_Count = m_Count + 1
    End Function

    ' Return total element count
    Public Function Count
        Count = m_Count
    End Function

    ' Return element at given index
    Public Function Item(index)
        Item = m_Items(index)
    End Function

    ' Create an iterator
    Public Function CreateIterator
        Set CreateIterator = New MyIterator
        CreateIterator.Init Me
    End Function
End Class

' Iterator: encapsulates traversal logic
Class MyIterator
    Private m_Collection   ' The collection being traversed
    Private m_Index        ' Current index

    ' Init: pass in collection and reset index
    Public Function Init(collection)
        Set m_Collection = collection
        m_Index = 0
    End Function

    ' Check if there's a next element
    Public Function HasNext
        HasNext = (m_Index < m_Collection.Count)
    End Function

    ' Return current element and advance index
    Public Function NextItem
        NextItem = m_Collection.Item(m_Index)
        m_Index = m_Index + 1
    End Function
End Class

' Demo: traverse collection with iterator, without exposing internal array
Dim coll, iter
Set coll = New MyCollection
coll.Add "Apple"
coll.Add "Banana"
coll.Add "Orange"

Set iter = coll.CreateIterator
Do While iter.HasNext
    Response.Write iter.NextItem
Loop
```

**Classic VBScript trade-offs**:
- **No unified interface**: MyCollection and MyIterator have no `ICollection`/`IIterator` interface. If different collections implement inconsistently (e.g., some use `NextItem`, others use `Next`), callers can't traverse them uniformly.
- **No IEnumerable standard**: VBScript's `For Each` only works with arrays and Dictionary, not custom collections. Custom iterators require manual `Do While iter.HasNext` loops.

### Axon VBScript Version (supports For Each custom collection iteration)

```vbscript
' Custom collection: built-in Collection + [DispId(-4)] forwarding
Class MyCollection
    Private m_Items

    ' Constructor: initialize built-in Collection
    Private Sub Class_Initialize
        Set m_Items = Server.CreateObject("Collection")
    End Sub

    ' Add element
    Public Sub Add(item)
        m_Items.Add item
    End Sub

    ' Return total element count
    Public Function Count
        Count = m_Items.Count
    End Function

    ' Return element at given index
    Public Function Item(index)
        Item = m_Items.Item(index)
    End Function

    ' For Each entry point: forward built-in Collection's enumerator
    [DispId(-4)]
    Public Property Get NewEnum
        Set NewEnum = m_Items.[_NewEnum]
    End Property
End Class

' Demo: traverse with For Each, without exposing internal structure
Dim coll, item
Set coll = New MyCollection
coll.Add "Apple"
coll.Add "Banana"
coll.Add "Orange"

For Each item In coll
    Response.Write item
Next
```

**Axon VBScript trade-offs**:
- This pattern is **fully solved in AxonASP, no compromises — the iterator class has disappeared**. AxonASP's `For Each` custom collection iteration (issue #52, now implemented) makes the Iterator pattern **disappear entirely** — no need to hand-write `IIterator`/`ICollection` interfaces or a `MyIterator` class, no `m_Index` variable maintenance, no `HasNext`/`NextItem` boilerplate methods. Callers just write `For Each item In coll`, identical to native array syntax. `[DispId(-4)]` marks the `NewEnum` property, forwarding the built-in `Collection`'s `[_NewEnum]` enumerator — this is the standard approach in VBA/VB6/twinBASIC. The appendix classifies this pattern as one of the 18 patterns where "AxonASP fully solves the core pain point", with no residual defects.

### VB.NET Version (syntactically complete baseline)

VB.NET has `IEnumerable(Of T)`/`IEnumerator(Of T)` generic interfaces + `Yield` keyword — this is the idiomatic .NET iterator pattern. `Yield` lets the compiler auto-generate a state-machine enumerator class, no hand-written `MoveNext`/`Current` needed.

```vbnet
' ① Implement IEnumerable(Of T): supports For Each traversal
Public Class MyCollection(Of T)
    Implements IEnumerable(Of T)

    Private ReadOnly m_Items As New List(Of T)()

    ' Add element
    Public Sub Add(item As T)
        m_Items.Add(item)
    End Sub

    ' Return total element count
    Public ReadOnly Property Count As Integer
        Get
            Return m_Items.Count
        End Get
    End Property

    ' Indexer
    Default Public ReadOnly Property Item(index As Integer) As T
        Get
            Return m_Items(index)
        End Get
    End Property

    ' ② Generic enumerator: Yield auto-generates IEnumerator(Of T) state machine
    Public Iterator Function GetEnumerator() As IEnumerator(Of T) _
        Implements IEnumerable(Of T).GetEnumerator
        For Each element As T In m_Items
            Yield element   ' Compiler auto-generates MoveNext/Current state machine
        Next
    End Function

    ' ③ Non-generic enumerator: IEnumerable(Of T) inherits from IEnumerable, must implement
    Private Function GetEnumeratorObj() As IEnumerator _
        Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function
End Class

' Demo: For Each traversal, consistent with Axon version call pattern
Dim coll As New MyCollection(Of String)()
coll.Add("Apple")
coll.Add("Banana")
coll.Add("Orange")

For Each item In coll
    Console.WriteLine(item)
Next
```

**VB.NET version notes**:
- **`Yield` compiler-generated state machine**: `Yield item` in an `Iterator` method makes the compiler auto-generate a nested class implementing `IEnumerator(Of T)`, including `MoveNext()`/`Current`/`Dispose()`. Compare with classic version: hand-written `MyIterator` needs 20+ lines of code (`m_Index`/`Init`/`HasNext`/`NextItem`), and is bug-prone.
- **`IEnumerable(Of T)` BCL standard interface**: After implementation, auto-supports `For Each` traversal, and can be passed as argument to any BCL method accepting `IEnumerable(Of T)`. Axon version uses `[DispId(-4)]` to forward COM enumerator, only supports `For Each`, can't plug into .NET collection ecosystem.
- **Generic type safety**: `MyCollection(Of String)` compile-time constrains to only store `String`, `Add(123)` causes compile error; `For Each item As String In coll` needs no type conversion. Axon version depends on built-in `Collection` COM object, elements are still Variant/Object.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Traversal method | Hand-written `Do While iter.HasNext` + `iter.NextItem` | Native `For Each item In coll` | Native `For Each` |
| Iterator class | Hand-written `MyIterator` class (20+ lines) | **Disappeared**, forwards built-in Collection enumerator | **Disappeared**, `Yield` compiler auto-generated |
| Interface standardization | No convention, each collection implements freely | `[DispId(-4)]` + `NewEnum` forwards COM enumerator | `IEnumerable(Of T)`/`IEnumerator(Of T)` BCL standard interface |
| Type safety | All Variant, runtime errors | Depends on COM Collection, elements still Variant/Object | Generic `(Of T)` compile-time constrained |
---
