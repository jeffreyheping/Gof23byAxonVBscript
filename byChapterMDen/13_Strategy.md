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
- `ISortStrategy` guarantees all strategy classes have a `Sort` method. Callers can switch strategies through a type-safe interface reference. `Sorter` holds an `ISortStrategy` reference and calls `m_Strategy.Sort` directly — auto-dispatches to the concrete strategy. Remaining gap: **Generics**. `Sort` parameters and return value can only be untyped (Variant) arrays. The compiler can't constrain array element types. Before Go 1.18 introduced generics, it was the same — using `interface{}` for parameters and relying on runtime type assertions. AxonASP's current state is equivalent to Go 1.17 and earlier.
---
