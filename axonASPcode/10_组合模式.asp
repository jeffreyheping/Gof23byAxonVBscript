<%
' 组件接口
Class IComponent
    Public Function Operation(indent As String)
    End Function
    Public Function Add(child As IComponent)
    End Function
End Class

' 叶子节点
Class Leaf
    Implements IComponent
    Private m_Name As String

    Public Property Get Name As String
        Name = m_Name
    End Property
    Public Property Let Name(v As String)
        m_Name = v
    End Property

    Public Function IComponent_Operation(indent As String)
        Response.Write indent & "叶子：" & m_Name
    End Function
    Public Function IComponent_Add(child As IComponent)
        ' 叶子无子节点，空实现
    End Function
End Class

' 组合节点
Class Composite
    Implements IComponent
    Private m_Name As String
    Private m_Children()
    Private m_Count As Integer

    Public Property Get Name As String
        Name = m_Name
    End Property
    Public Property Let Name(v As String)
        m_Name = v
    End Property

    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Children(10)
    End Sub

    Public Function IComponent_Add(child As IComponent)
        If m_Count >= UBound(m_Children) + 1 Then
            ReDim Preserve m_Children(m_Count * 2)
        End If
        Set m_Children(m_Count) = child
        m_Count = m_Count + 1
    End Function

    Public Function IComponent_Operation(indent As String)
        Response.Write indent & "组合：" & m_Name
        Dim i As Integer, child As IComponent
        For i = 0 To m_Count - 1
            Set child = m_Children(i)
            child.Operation indent & "  "
        Next
    End Function
End Class

' 演示：统一接口遍历树
Dim rootObj As Composite, branch1Obj As Composite
Dim leaf1Obj As Leaf, leaf2Obj As Leaf, leaf3Obj As Leaf
Set rootObj = New Composite
rootObj.Name = "总部"
Set branch1Obj = New Composite
branch1Obj.Name = "分公司"
Set leaf1Obj = New Leaf: leaf1Obj.Name = "员工A"
Set leaf2Obj = New Leaf: leaf2Obj.Name = "员工B"
Set leaf3Obj = New Leaf: leaf3Obj.Name = "员工C"

Dim root As IComponent, branch1 As IComponent
Dim leaf1 As IComponent, leaf2 As IComponent, leaf3 As IComponent
Set root = rootObj
Set branch1 = branch1Obj
Set leaf1 = leaf1Obj
Set leaf2 = leaf2Obj
Set leaf3 = leaf3Obj

root.Add branch1
root.Add leaf3
branch1.Add leaf1
branch1.Add leaf2

root.Operation ""
%>