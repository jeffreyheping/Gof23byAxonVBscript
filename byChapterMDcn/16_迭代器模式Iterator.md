## 第16章 迭代器模式（Iterator）

**核心思想**：提供一种顺序访问聚合对象元素的方法，而不暴露其内部结构。

**示例说明**：MyCollection 维护一个内部数组。MyIterator 持有当前索引，通过 HasNext 和 NextItem 方法顺序遍历集合元素，调用方无需知道内部是数组还是链表。

### 传统 VBScript 版

```vbscript
' 集合类：内部用数组存储数据
Class MyCollection
    Private m_Items()   ' 内部数组

    Private m_Count     ' 当前元素数量

    ' 构造函数：初始化数组
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Items(10)
    End Sub

    ' 添加元素（容量不足时自动扩容）
    Public Function Add(item)
        If m_Count >= UBound(m_Items) + 1 Then
            ReDim Preserve m_Items(m_Count * 2)
        End If
        m_Items(m_Count) = item
        m_Count = m_Count + 1
    End Function

    ' 返回元素总数
    Public Function Count
        Count = m_Count
    End Function

    ' 返回指定索引的元素
    Public Function Item(index)
        Item = m_Items(index)
    End Function

    ' 创建迭代器
    Public Function CreateIterator
        Set CreateIterator = New MyIterator
        CreateIterator.Init(Me)

    End Function
End Class

' 迭代器：封装遍历逻辑
Class MyIterator
    Private m_Collection   ' 被遍历的集合
    Private m_Index        ' 当前索引

    ' 初始化：传入集合并重置索引
    Public Function Init(collection)
        Set m_Collection = collection
        m_Index = 0
    End Function

    ' 检查是否还有下一个元素
    Public Function HasNext
        HasNext = (m_Index < m_Collection.Count)
    End Function

    ' 返回当前元素并将索引后移
    Public Function NextItem
        NextItem = m_Collection.Item(m_Index)
        m_Index = m_Index + 1
    End Function
End Class

' 演示：用迭代器遍历集合，不暴露内部数组
Dim coll, iter
Set coll = New MyCollection
coll.Add("苹果")

coll.Add("香蕉")

coll.Add("橙子")


Set iter = coll.CreateIterator
Do While iter.HasNext
    Response.Write(iter.NextItem)

Loop
```

**传统 VBScript 版妥协说明**：
- **无统一接口**：MyCollection 和 MyIterator 没有 `ICollection`/`IIterator` 接口。如果多个集合类实现不一致（比如有的叫 `NextItem`，有的叫 `Next`），调用方无法统一遍历。
- **无 IEnumerable 标准接口**：VBScript 的 `For Each` 只支持数组和 Dictionary，不支持自定义集合。自定义迭代器只能靠 `Do While iter.HasNext` 手动循环。

### Axon VBScript 版（支持 For Each 自定义集合迭代）

```vba
' 自定义集合：内置 Collection + [DispId(-4)] 转发
Class MyCollection
    Private m_Items

    ' 构造函数：初始化内置 Collection
    Private Sub Class_Initialize
        Set m_Items = Server.CreateObject("Collection")
    End Sub

    ' 添加元素
    Public Sub Add(item)
        m_Items.Add(item)

    End Sub

    ' 返回元素总数
    Public Function Count
        Count = m_Items.Count
    End Function

    ' 返回指定索引的元素
    Public Function Item(index)
        Item = m_Items.Item(index)
    End Function

    ' For Each 入口：转发内置 Collection 的枚举器
    [DispId(-4)]
    Public Property Get NewEnum
        Set NewEnum = m_Items.[_NewEnum]
    End Property
End Class

' 演示：用 For Each 遍历集合，不暴露内部结构
Dim coll, item
Set coll = New MyCollection
coll.Add("苹果")

coll.Add("香蕉")

coll.Add("橙子")


For Each item In coll
    Response.Write(item)

Next
```

**Axon VBScript 版妥协说明**：
- 此模式在 AxonASP 中**彻底解决核心痛点，无妥协，迭代器类已消失**。AxonASP 的 `For Each` 自定义集合迭代（issue #52 已实现）让迭代器模式**完全消失**——无需手写 `IIterator`/`ICollection` 接口和 `MyIterator` 类，无需维护 `m_Index` 索引变量，无需写 `HasNext`/`NextItem` 样板方法。调用方直接 `For Each item In coll` 即可，写法与原生数组完全一致。`[DispId(-4)]` 标记 `NewEnum` 属性，转发内置 `Collection` 的 `[_NewEnum]` 枚举器，这是 VBA/VB6/twinBASIC 的标准做法。附录将此模式归类为"AxonASP 彻底解决核心痛点"的 18 个模式之一，无残留缺陷。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有 `IEnumerable(Of T)`/`IEnumerator(Of T)` 泛型接口 + `Yield` 关键字——这是 .NET 平台迭代器的地道写法，`Yield` 让编译器自动生成状态机枚举器类，无需手写 `MoveNext`/`Current`。

```vbnet
' ① 实现 IEnumerable(Of T)：支持 For Each 遍历
Public Class MyCollection(Of T)
    Implements IEnumerable(Of T)

    Private ReadOnly m_Items As New List(Of T)()

    ' 添加元素
    Public Function Add(item As T)
        m_Items.Add(item)
    End Function

    ' 返回元素总数
    Public ReadOnly Property Count As Integer
        Get
            Return m_Items.Count
        End Get
    End Property

    ' 索引器
        Default Public ReadOnly Property Item(index As Integer) As T

        Get
            Return m_Items(index)
        End Get
    End Property

    ' ② 泛型枚举器：Yield 自动生成 IEnumerator(Of T) 状态机
    Public Iterator Function GetEnumerator() As IEnumerator(Of T) _
        Implements IEnumerable(Of T).GetEnumerator
        For Each element As T In m_Items
            Yield element   ' 编译器自动生成 MoveNext/Current 状态机
        Next
    End Function

    ' ③ 非泛型枚举器：IEnumerable(Of T) 继承自 IEnumerable，必须实现
    Private Function GetEnumeratorObj() As IEnumerator _
        Implements IEnumerable.GetEnumerator
        Return GetEnumerator()
    End Function
End Class

' 演示：For Each 遍历，与 Axon 版调用方式一致
Dim coll As New MyCollection(Of String)()
coll.Add("苹果")
coll.Add("香蕉")
coll.Add("橙子")

For Each item In coll
    Console.WriteLine(item)
Next
```

**VB.NET 版说明**：
- **`Yield` 编译器生成状态机**：`Iterator` 方法中的 `Yield item` 让编译器自动生成实现了 `IEnumerator(Of T)` 的嵌套类，包含 `MoveNext()`/`Current`/`Dispose()`。对比传统版：手写 `MyIterator` 需要 20+ 行代码（`m_Index`/`Init`/`HasNext`/`NextItem`），还容易出 bug。
- **`IEnumerable(Of T)` BCL 标准接口**：实现后自动支持 `For Each` 遍历，并可作为参数传入任何接受 `IEnumerable(Of T)` 的 BCL 方法。Axon 版用 `[DispId(-4)]` 转发 COM 枚举器，仅支持 `For Each`，无法接入 .NET 集合生态。
- **泛型类型安全**：`MyCollection(Of String)` 编译期约束只能存 `String`，`Add(123)` 直接报错；`For Each item As String In coll` 无需类型转换。Axon 版因依赖内置 `Collection` COM 对象，元素仍为 Variant/Object。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 遍历方式 | 手写 `Do While iter.HasNext` + `iter.NextItem` | 原生 `For Each item In coll` | 原生 `For Each` |
| 迭代器类 | 需要手写 `MyIterator` 类（20+ 行） | **已消失**，转发内置 Collection 枚举器 | **已消失**，`Yield` 编译器自动生成 |
| 接口标准化 | 无约定，各集合实现随意 | `[DispId(-4)]` + `NewEnum` 转发 COM 枚举器 | `IEnumerable(Of T)`/`IEnumerator(Of T)` BCL 标准接口 |
| 类型安全 | 全 Variant，运行时报错 | 依赖 COM Collection，元素仍 Variant/Object | 泛型 `(Of T)` 编译期约束 |
---