<%
Option Explicit
' 享元对象：树的固有属性（名称、颜色），可被多棵树共享
Class TreeType
    Public Name As String
    Public Color As String

    ' 在指定坐标绘制树
    Public Function Draw(x As Long, y As Long)
        Response.Write("在 (" & x & "," & y & ") 绘制 " & Color & Name)

    End Function
End Class

' 享元工厂：缓存并复用 TreeType 对象
Class TreeFactory
    ' 注：Scripting.Dictionary 是 COM 对象，As 不支持标注，保留 Variant
    Private m_Types

    ' 构造函数：创建字典
    Private Sub Class_Initialize
        Set m_Types = CreateObject("Scripting.Dictionary")
    End Sub

    ' 获取或创建 TreeType：相同参数返回同一个对象
    ' name: 树名, color: 颜色
    ' 返回值：共享的 TreeType 实例
    Public Function GetTreeType(name As String, color As String) As TreeType
        Dim key As String
        key = name & "|" & color
        If Not m_Types.Exists(key) Then
            Dim t As TreeType
            Set t = New TreeType
            t.Name = name
            t.Color = color
            Set m_Types(key) = t
        End If
        Set GetTreeType = m_Types(key)
    End Function
End Class

' 演示：3 棵树共享同一个 TreeType 对象
Dim factory As TreeFactory
Dim oakType As TreeType
Dim i As Long
Set factory = New TreeFactory
Set oakType = factory.GetTreeType("橡树", "绿色")

For i = 0 To 2
    oakType.Draw i, i * 2
Next
Response.Write("3 棵树，实际只有 1 个 TreeType 对象")
%>