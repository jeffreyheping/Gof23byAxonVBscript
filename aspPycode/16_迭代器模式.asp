<%
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
%>