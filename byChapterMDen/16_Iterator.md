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
- AxonASP's `For Each` custom collection iteration (issue #52, now implemented) makes the Iterator pattern **disappear entirely** — no need to hand-write `IIterator`/`ICollection` interfaces or a `MyIterator` class. The caller just writes `For Each item In coll`. `[DispId(-4)]` marks the `NewEnum` property, forwarding the built-in `Collection`'s enumerator — this is the standard approach in VBA/VB6/twinBASIC. Compared to the classic version's hand-written iterator class (`HasNext`/`NextItem`/`m_Index`), code volume drops significantly and traversal logic is built into the language.
---
