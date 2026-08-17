<%
Option Explicit
' 自定义集合：内置 Collection + [DispId(-4)] 转发
Class MyCollection
    Private m_Items

    ' 构造函数：初始化内置 Collection
    Private Sub Class_Initialize
        Set m_Items = Server.CreateObject("Collection")
    End Sub

    ' 添加元素
    Public Function Add(item)
        m_Items.Add(item)

    End Function

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
%>