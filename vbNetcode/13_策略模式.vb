Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch13Module
    Public MustInherit Class SortStrategy
        Public MustOverride Function Sort(arr As Integer()) As Integer()
    End Class
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
    Public Class QuickSort
        Inherits SortStrategy
        Public Overrides Function Sort(arr As Integer()) As Integer()
            Dim a As Integer() = CType(arr.Clone(), Integer())
            QuickSortHelper(a, 0, a.Length - 1)
            Return a
        End Function

        Private Function QuickSortHelper(a As Integer(), lo As Integer, hi As Integer) As Object
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
        End Function
    End Class
    Public Class Sorter
        Private m_Strategy As SortStrategy

        Public Function SetStrategy(strategy As SortStrategy) As Object
            m_Strategy = strategy
        End Function

        Public Function Sort(arr As Integer()) As Integer()
            Return m_Strategy.Sort(arr)
        End Function
    End Class
    Sub Main()

        ' ② 具体策略：冒泡排序

        ' ③ 具体策略：快速排序

        ' ④ 上下文：持有抽象基类引用，与 Axon 版一致使用 SetStrategy 切换策略

        ' 演示：同一 Sorter，换策略就换算法（与 Axon 版调用方式一致）
        Dim sorter As New Sorter()
        Dim data As Integer() = {5, 2, 8, 1, 9}
        sorter.SetStrategy(New BubbleSort())
        Console.WriteLine(String.Join(",", sorter.Sort(data)))

        sorter.SetStrategy(New QuickSort())
        Console.WriteLine(String.Join(",", sorter.Sort(data)))
    End Sub
End Module
