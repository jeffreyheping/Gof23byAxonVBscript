## Chapter 13: Strategy

**Core idea**: Encapsulate an algorithm as a swappable strategy object.

**Example**: Sorter holds a SortStrategy reference. Both BubbleSort and QuickSort implement `Sort`. Inject different strategies when creating a Sorter, and the `Sort` method runs different algorithms.

### Classic VBScript Version

```vbscript
' Concrete strategy: bubble sort
Class BubbleSort
    ' Return a sorted copy of the array (original is not modified)
    Public Function Sort(arr)
        Dim a, i, j, tmp
        a = arr   ' Copy — don't modify original
        For i = 0 To UBound(a)
            For j = 0 To UBound(a) - 1 - i
                If a(j) > a(j + 1) Then
                    tmp = a(j)
                    a(j) = a(j + 1)
                    a(j + 1) = tmp
                End If
            Next
        Next
        Sort = a
    End Function
End Class

' Concrete strategy: quick sort (simplified)
Class QuickSort
    ' Return a sorted copy of the array
    Public Function Sort(arr)
        Sort = QuickSortHelper(arr, 0, UBound(arr))
    End Function

    ' Quick sort recursive implementation (teaching simplified version,
    ' not suitable for large arrays; VBScript has no tail-call optimization,
    ' large data may cause stack overflow)
    Private Function QuickSortHelper(a, lo, hi)
        Dim i, j, pivot, tmp
        If lo < hi Then
            pivot = a(hi)
            i = lo
            For j = lo To hi - 1
                If a(j) <= pivot Then
                    tmp = a(i)
                    a(i) = a(j)
                    a(j) = tmp
                    i = i + 1
                End If
            Next
            tmp = a(i)
            a(i) = a(hi)
            a(hi) = tmp
            a = QuickSortHelper(a, lo, i - 1)
            a = QuickSortHelper(a, i + 1, hi)
        End If
        QuickSortHelper = a
    End Function
End Class

' Context: holds strategy reference; caller can switch strategies anytime
Class Sorter
    Private m_Strategy   ' Current sorting strategy

    ' Inject strategy object
    Public Function SetStrategy(strategy)
        Set m_Strategy = strategy
    End Function

    ' Sort array with current strategy
    Public Function Sort(arr)
        Sort = m_Strategy.Sort(arr)
    End Function
End Class

' Demo: same Sorter, different strategy = different algorithm
Dim mySorter, data
Set mySorter = New Sorter
data = Array(5, 2, 8, 1, 9)

mySorter.SetStrategy New BubbleSort
Response.Write Join(mySorter.Sort(data), ",")

mySorter.SetStrategy New QuickSort
Response.Write Join(mySorter.Sort(data), ",")
```

**Classic VBScript trade-offs**:
- **No interface constraint**: BubbleSort and QuickSort rely on the `Sort` method-name convention. No `ISortStrategy` interface guarantees consistent signatures. If a strategy class forgets to define Sort, the error only surfaces at runtime.
- **Tight coupling between algorithm and context**: The `strategy` parameter passed to Sorter has no type constraint. Passing the wrong object only reveals the problem at runtime.

### Axon VBScript Version (supports Implements)

```vbscript
' Strategy interface
Class ISortStrategy
    Public Function Sort(arr) As Variant
    End Function
End Class

' Bubble sort
Class BubbleSort
    Implements ISortStrategy
    Public Function ISortStrategy_Sort(arr) As Variant
        Dim a, i As Integer, j As Integer, tmp
        a = arr
        For i = 0 To UBound(a)
            For j = 0 To UBound(a) - 1 - i
                If a(j) > a(j + 1) Then
                    tmp = a(j)
                    a(j) = a(j + 1)
                    a(j + 1) = tmp
                End If
            Next
        Next
        ISortStrategy_Sort = a
    End Function
End Class

' Quick sort
Class QuickSort
    Implements ISortStrategy
    Public Function ISortStrategy_Sort(arr) As Variant
        ISortStrategy_Sort = QuickSortHelper(arr, 0, UBound(arr))
    End Function

    ' Quick sort recursive implementation (teaching simplified version,
    ' not suitable for large arrays; VBScript has no tail-call optimization,
    ' large data may cause stack overflow)
    Private Function QuickSortHelper(a, lo As Integer, hi As Integer) As Variant
        Dim i As Integer, j As Integer, pivot, tmp
        If lo < hi Then
            pivot = a(hi)
            i = lo
            For j = lo To hi - 1
                If a(j) <= pivot Then
                    tmp = a(i)
                    a(i) = a(j)
                    a(j) = tmp
                    i = i + 1
                End If
            Next
            tmp = a(i)
            a(i) = a(hi)
            a(hi) = tmp
            a = QuickSortHelper(a, lo, i - 1)
            a = QuickSortHelper(a, i + 1, hi)
        End If
        QuickSortHelper = a
    End Function
End Class

' Context
Class Sorter
    Private m_Strategy As ISortStrategy

    Public Function SetStrategy(strategy As ISortStrategy)
        Set m_Strategy = strategy
    End Function
    Public Function Sort(arr) As Variant
        Sort = m_Strategy.Sort(arr)
    End Function
End Class

' Demo
Dim sorter As Sorter, data As Variant
Set sorter = New Sorter
data = Array(5, 2, 8, 1, 9)

sorter.SetStrategy New BubbleSort
Response.Write Join(sorter.Sort(data), ",")

sorter.SetStrategy New QuickSort
Response.Write Join(sorter.Sort(data), ",")
```

**Axon VBScript trade-offs**:
- This pattern maps naturally to AxonASP. `ISortStrategy` interface guarantees all strategy classes have a `Sort` method. Callers can switch strategies through a type-safe interface reference. `Sorter` holds an `ISortStrategy` reference, calls `m_Strategy.Sort` directly — auto-dispatches to the concrete strategy without fully-qualified names. Remaining gap: **missing code reuse mechanism (inheritance)**. If you need to add common logic to all strategies (e.g., performance timing, logging), you need an abstract base class + subclass `Overrides`. AxonASP has no inheritance — every strategy class must duplicate the common code. Go also lacks inheritance, solved with struct embedding — embedding a `BaseStrategy` struct auto-provides common methods.

### VB.NET Version (syntactically complete baseline)

VB.NET has `MustInherit` (abstract base class) + `Overrides` (override), enabling textbook Strategy — abstract base class constrains algorithm contract, subclasses `Overrides` to implement specific algorithms.

```vbnet
' ① Abstract strategy base class: MustInherit prevents direct instantiation, MustOverride forces subclasses to implement Sort
Public MustInherit Class SortStrategy
    Public MustOverride Function Sort(arr As Integer()) As Integer()
End Class

' ② Concrete strategy: Bubble Sort
Public Class BubbleSort
    Inherits SortStrategy
    Public Overrides Function Sort(arr As Integer()) As Integer()
        Dim a As Integer() = CType(arr.Clone(), Integer())
        For i = 0 To a.Length - 1
            For j = 0 To a.Length - 2 - i
                If a(j) > a(j + 1) Then
                    Dim tmp As Integer = a(j)
                    a(j) = a(j + 1)
                    a(j + 1) = tmp
                End If
            Next
        Next
        Return a
    End Function
End Class

' ③ Concrete strategy: Quick Sort
Public Class QuickSort
    Inherits SortStrategy
    Public Overrides Function Sort(arr As Integer()) As Integer()
        Dim a As Integer() = CType(arr.Clone(), Integer())
        QuickSortHelper(a, 0, a.Length - 1)
        Return a
    End Function

    Private Sub QuickSortHelper(a As Integer(), lo As Integer, hi As Integer)
        If lo < hi Then
            Dim pivot As Integer = a(hi)
            Dim i As Integer = lo
            For j = lo To hi - 1
                If a(j) <= pivot Then
                    Dim tmp As Integer = a(i)
                    a(i) = a(j)
                    a(j) = tmp
                    i += 1
                End If
            Next
            Dim tmp2 As Integer = a(i)
            a(i) = a(hi)
            a(hi) = tmp2
            QuickSortHelper(a, lo, i - 1)
            QuickSortHelper(a, i + 1, hi)
        End If
    End Sub
End Class

' ④ Context: holds abstract base class reference, same SetStrategy switching as Axon version
Public Class Sorter
    Private m_Strategy As SortStrategy

    Public Sub SetStrategy(strategy As SortStrategy)
        m_Strategy = strategy
    End Sub

    Public Function Sort(arr As Integer()) As Integer()
        Return m_Strategy.Sort(arr)
    End Function
End Class

' Demo: same Sorter, switch strategy to change algorithm (same call pattern as Axon version)
Dim sorter As New Sorter()
Dim data As Integer() = {5, 2, 8, 1, 9}
sorter.SetStrategy(New BubbleSort())
Console.WriteLine(String.Join(",", sorter.Sort(data)))

sorter.SetStrategy(New QuickSort())
Console.WriteLine(String.Join(",", sorter.Sort(data)))
```

**VB.NET version notes**:
- **Real abstract base class + inheritance for code reuse**: `MustInherit Class SortStrategy` prevents `New SortStrategy()`, `MustOverride Sort` forces all subclasses to implement — compile-time check, missing implementation causes immediate error. If you later want to add common fields/methods to all strategies, just add once in base class, subclasses inherit automatically. Axon version can only use `ISortStrategy` interface to constrain method existence, but can't add common logic for subclass inheritance.
- **Strongly-typed arrays**: `Sort(arr As Integer()) As Integer()` constrains array element types at compile time, passing a string array causes compile error. Axon version `As Variant` can accept anything, type mismatch only crashes at runtime.
- **No `Set`/`Let` distinction**: VB.NET object assignment uses `=` directly, no need to remember `Set` for objects, `Let` for value types.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Strategy contract | Method name convention (easy to miss) | `ISortStrategy` interface constrains `Sort` | `MustInherit` + `MustOverride` compile-time enforced |
| Code reuse | None (parallel strategy classes) | None (parallel strategy classes, no inheritance) | Base class fields/methods automatically inherited by all subclasses |
| Type safety | All Variant, runtime errors | Interface reference strongly typed, Sort params/return still Variant | `Integer()` compile-time constrains element type |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |
---
