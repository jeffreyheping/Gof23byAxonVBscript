Option Explicit
Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
' 具体策略：冒泡排序
Class BubbleSort
    ' 返回排序后数组的副本（原数组不被修改）
    Public Function Sort(arr)
        Dim a, i, j, tmp
        a = arr   ' 复制一份，不修改原数组
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

' 具体策略：快速排序（简化版）
Class QuickSort
    ' 返回排序后数组的副本
    Public Function Sort(arr)
        Sort = QuickSortHelper(arr, 0, UBound(arr))
    End Function

    ' 快速排序递归实现（教学简化版，不适合大规模数组；VBScript 无尾调用优化，大规模数据可能栈溢出）
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

' 上下文：持有策略引用，调用方可随时切换策略
Class Sorter
    Private m_Strategy   ' 当前使用的排序策略

    ' 注入策略对象
    Public Function SetStrategy(strategy)
        Set m_Strategy = strategy
    End Function

    ' 用当前策略对数组排序
    Public Function Sort(arr)
        Sort = m_Strategy.Sort(arr)
    End Function
End Class

' 演示：同一 Sorter，换策略就换算法
Dim mySorter, data
Set mySorter = New Sorter
data = Array(5, 2, 8, 1, 9)

mySorter.SetStrategy(New BubbleSort)

Response.Write(Join(mySorter.Sort(data), ","))


mySorter.SetStrategy(New QuickSort)

Response.Write(Join(mySorter.Sort(data), ","))

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
