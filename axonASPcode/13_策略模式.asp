<%
' 策略接口
Class ISortStrategy
    Public Function Sort(arr) As Variant
    End Function
End Class

' 冒泡排序
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

' 快速排序
Class QuickSort
    Implements ISortStrategy
    Public Function ISortStrategy_Sort(arr) As Variant
        ISortStrategy_Sort = QuickSortHelper(arr, 0, UBound(arr))
    End Function

    ' 快速排序递归实现（教学简化版，不适合大规模数组；VBScript 无尾调用优化，大规模数据可能栈溢出）
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

' 上下文
Class Sorter
    Private m_Strategy As ISortStrategy

    Public Function SetStrategy(strategy As ISortStrategy)
        Set m_Strategy = strategy
    End Function
    Public Function Sort(arr) As Variant
        Sort = m_Strategy.Sort(arr)
    End Function
End Class

' 演示
Dim sorter As Sorter, data As Variant
Set sorter = New Sorter
data = Array(5, 2, 8, 1, 9)

sorter.SetStrategy New BubbleSort
Response.Write Join(sorter.Sort(data), ",")

sorter.SetStrategy New QuickSort
Response.Write Join(sorter.Sort(data), ",")
%>