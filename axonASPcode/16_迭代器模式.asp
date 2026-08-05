<%
' 迭代器接口
Class IIterator
    Public Function HasNext As Boolean
    End Function
    Public Function NextItem As Variant
    End Function
End Class

' 集合接口
Class ICollection
    Public Function Add(item As Variant)
    End Function
    Public Function Count As Integer
    End Function
    Public Function Item(index As Integer) As Variant
    End Function
    Public Function CreateIterator As IIterator
    End Function
End Class

' 具体集合
Class MyCollection
    Implements ICollection
    Private m_Items()
    Private m_Count As Integer

    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Items(10)
    End Sub

    Public Function ICollection_Add(item As Variant)
        If m_Count >= UBound(m_Items) + 1 Then
            ReDim Preserve m_Items(m_Count * 2)
        End If
        m_Items(m_Count) = item
        m_Count = m_Count + 1
    End Function

    Public Function ICollection_Count As Integer
        ICollection_Count = m_Count
    End Function

    Public Function ICollection_Item(index As Integer) As Variant
        ICollection_Item = m_Items(index)
    End Function

    Public Function ICollection_CreateIterator As IIterator
        Dim iter As MyIterator
        Set iter = New MyIterator
        iter.Init Me
        Set ICollection_CreateIterator = iter
    End Function
End Class

' 具体迭代器：持有集合引用，通过接口方法访问元素
Class MyIterator
    Implements IIterator
    Private m_Collection As ICollection
    Private m_Index As Integer

    Public Function Init(collection As ICollection)
        Set m_Collection = collection
        m_Index = 0
    End Function

    Public Function IIterator_HasNext As Boolean
        IIterator_HasNext = (m_Index < m_Collection.Count)
    End Function

    Public Function IIterator_NextItem As Variant
        IIterator_NextItem = m_Collection.Item(m_Index)
        m_Index = m_Index + 1
    End Function
End Class

' 演示
Dim coll As ICollection
Set coll = New MyCollection
coll.Add "苹果"
coll.Add "香蕉"
coll.Add "橙子"

Dim iter As IIterator
Set iter = coll.CreateIterator
Do While iter.HasNext
    Response.Write iter.NextItem
Loop
%>