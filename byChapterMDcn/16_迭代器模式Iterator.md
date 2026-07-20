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
        CreateIterator.Init Me
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
coll.Add "苹果"
coll.Add "香蕉"
coll.Add "橙子"

Set iter = coll.CreateIterator
Do While iter.HasNext
    Response.Write iter.NextItem
Loop
```

**传统 VBScript 版妥协说明**：
- **无统一接口**：MyCollection 和 MyIterator 没有 `ICollection`/`IIterator` 接口。如果多个集合类实现不一致（比如有的叫 `NextItem`，有的叫 `Next`），调用方无法统一遍历。
- **无 IEnumerable 标准接口**：VBScript 的 `For Each` 只支持数组和 Dictionary，不支持自定义集合。自定义迭代器只能靠 `Do While iter.HasNext` 手动循环。

### Axon VBScript 版（支持 For Each 自定义集合迭代）

```vbscript
' 自定义集合：内置 Collection + [DispId(-4)] 转发
Class MyCollection
    Private m_Items

    ' 构造函数：初始化内置 Collection
    Private Sub Class_Initialize
        Set m_Items = Server.CreateObject("Collection")
    End Sub

    ' 添加元素
    Public Sub Add(item)
        m_Items.Add item
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
coll.Add "苹果"
coll.Add "香蕉"
coll.Add "橙子"

For Each item In coll
    Response.Write item
Next
```

**Axon VBScript 版妥协说明**：
- AxonASP 的 `For Each` 自定义集合迭代（issue #52 已实现）让迭代器模式**完全消失**——无需手写 `IIterator`/`ICollection` 接口和 `MyIterator` 类，调用方直接 `For Each item In coll` 即可。`[DispId(-4)]` 标记 `NewEnum` 属性，转发内置 `Collection` 的枚举器，这是 VBA/VB6/twinBASIC 的标准做法。对比传统版的手写迭代器类（`HasNext`/`NextItem`/`m_Index`），代码量大幅减少，遍历逻辑由语言内置支持。
---