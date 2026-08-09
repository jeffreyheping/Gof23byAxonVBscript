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

mySorter.SetStrategy(New BubbleSort)

Response.Write(Join(mySorter.Sort(data), ","))


mySorter.SetStrategy(New QuickSort)

Response.Write(Join(mySorter.Sort(data), ","))

```

**传统 VBScript 版妥协说明**：
- **无接口约束**：BubbleSort 和 QuickSort 靠 `Sort` 方法名约定，没有 `ISortStrategy` 接口保证签名一致。如果某个策略类漏写 Sort 方法，运行时调用才报错。
- **算法与上下文紧耦合**：Sorter 接收 `strategy` 参数无类型约束，传错对象时只能在运行时暴露问题。

### Axon VBScript 版（支持 Implements）

```vba
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

sorter.SetStrategy(New BubbleSort)

Response.Write(Join(sorter.Sort(data), ","))


sorter.SetStrategy(New QuickSort)

Response.Write(Join(sorter.Sort(data), ","))

```

**Axon VBScript 版妥协说明**：
- 此模式在 AxonASP 中实现较为自然，`ISortStrategy` 接口保证了所有策略类都有 `Sort` 方法，调用方可以通过接口类型安全地切换策略。`Sorter` 持有 `ISortStrategy` 引用，直接调用 `m_Strategy.Sort` 即可自动路由到具体策略实现，无需完整限定名。残留限制：缺失语法点：**代码复用机制（继承）**。经典策略模式如果需要给所有策略加公共逻辑（如性能计时、日志埋点），需要抽象基类 + 子类 `Overrides`，AxonASP 无继承，每个策略类需各自重复写公共代码。Go 同样无继承，用 struct embedding 解决——嵌入一个 `BaseStrategy` 结构体即自动获得公共方法。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有 `MustInherit`（抽象基类）+ `Overrides`（重写），可以写出教科书式的策略模式——抽象基类约束算法契约，子类 `Overrides` 重写具体算法。

```vbnet
' ① 抽象策略基类：MustInherit 禁止直接实例化，MustOverride 强制子类实现 Sort
Public MustInherit Class SortStrategy
    Public MustOverride Function Sort(arr As Integer()) As Integer()
End Class

' ② 具体策略：冒泡排序
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

' ③ 具体策略：快速排序
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

' ④ 上下文：持有抽象基类引用，与 Axon 版一致使用 SetStrategy 切换策略
Public Class Sorter
    Private m_Strategy As SortStrategy

    Public Function SetStrategy(strategy As SortStrategy) As Object
        m_Strategy = strategy
    End Function

    Public Function Sort(arr As Integer()) As Integer()
        Return m_Strategy.Sort(arr)
    End Function
End Class

' 演示：同一 Sorter，换策略就换算法（与 Axon 版调用方式一致）
Dim sorter As New Sorter()
Dim data As Integer() = {5, 2, 8, 1, 9}
sorter.SetStrategy(New BubbleSort())
Console.WriteLine(String.Join(",", sorter.Sort(data)))

sorter.SetStrategy(New QuickSort())
Console.WriteLine(String.Join(",", sorter.Sort(data)))
```

**VB.NET 版说明**：
- **真正的抽象基类 + 继承复用**：`MustInherit Class SortStrategy` 禁止 `New SortStrategy()`，`MustOverride Sort` 强制所有子类必须实现——编译期检查，漏写直接报错。若后续要给所有策略加公共字段或方法，只需在基类加一次，子类自动继承。Axon 版只能用 `ISortStrategy` 接口约束方法存在，但无法在基类加公共逻辑供子类继承。
- **强类型数组**：`Sort(arr As Integer()) As Integer()` 编译期约束数组元素类型，传字符串数组直接编译报错。Axon 版 `As Variant` 可以传任何东西，类型不匹配运行时才崩。
- **无需 `Set` / `Let` 区分**：VB.NET 对象赋值直接用 `=`，不再需要记忆 `Set` 给对象、`Let` 给值类型这一历史包袱。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 策略契约 | 方法名约定（易漏写） | `ISortStrategy` 接口约束 `Sort` | `MustInherit` + `MustOverride` 编译期强制 |
| 代码复用 | 无（各策略平行类） | 无（各策略平行类，无继承） | 基类字段/方法自动传给所有子类 |
| 类型安全 | 全 Variant，运行时报错 | 接口引用强类型，Sort 参数/返回值仍 Variant | `Integer()` 编译期约束元素类型 |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
---