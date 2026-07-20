## 第13章 策略模式（Strategy）

**核心思想**：把算法封装成可替换的策略对象。

**示例说明**：Sorter 持有 SortStrategy 引用。BubbleSort 和 QuickSort 都实现了 Sort 方法。创建 Sorter 时注入不同策略，Sort 方法就会执行不同算法。

### 传统 VBScript 版

```vbscript
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

mySorter.SetStrategy New BubbleSort
Response.Write Join(mySorter.Sort(data), ",")

mySorter.SetStrategy New QuickSort
Response.Write Join(mySorter.Sort(data), ",")
```

**传统 VBScript 版妥协说明**：
- **无接口约束**：BubbleSort 和 QuickSort 靠 `Sort` 方法名约定，没有 `ISortStrategy` 接口保证签名一致。如果某个策略类漏写 Sort 方法，运行时调用才报错。
- **算法与上下文紧耦合**：Sorter 接收 `strategy` 参数无类型约束，传错对象时只能在运行时暴露问题。

### Axon VBScript 版（支持 Implements）

```vbscript
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
```

**Axon VBScript 版妥协说明**：
- `ISortStrategy` 接口保证了所有策略类都有 `Sort` 方法，调用方可以通过接口类型安全地切换策略。`Sorter` 持有 `ISortStrategy` 引用，直接调用 `m_Strategy.Sort` 即可自动路由到具体策略实现，无需完整限定名。残留限制：缺失语法点：**泛型**。`Sort` 的参数与返回值只能用无类型（Variant）数组，编译期无法约束数组元素类型。Go 在 1.18 引入泛型前同样如此——用 `interface{}` 传参，靠运行时类型断言。AxonASP 当前状态等同于 Go 1.17 及之前。
---