Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
' 叶子节点：树形结构的最末端
Class Leaf
    Public Name
    ' 显示自身信息，indent 控制缩进层级
    Public Function Operation(indent)
        Response.Write indent & "叶子：" & Name
    End Function
End Class

' 组合节点：可包含子节点（Leaf 或 Composite）
Class Composite
    Public Name
    Private m_Children()   ' 子节点数组
    Private m_Count        ' 当前子节点数

    ' 构造函数：初始化数组
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Children(10)
    End Sub

    ' 添加子节点（容量不足时自动扩容）
    Public Function Add(child)
        If m_Count >= UBound(m_Children) + 1 Then
            ReDim Preserve m_Children(m_Count * 2)
        End If
        Set m_Children(m_Count) = child
        m_Count = m_Count + 1
    End Function

    ' 显示自身信息，并递归调用所有子节点的 Operation
    Public Function Operation(indent)
        Response.Write indent & "组合：" & Name
        Dim i
        For i = 0 To m_Count - 1
            m_Children(i).Operation indent & "  "
        Next
    End Function
End Class

' 演示：构建 总部→分公司→员工 的树形结构
Dim root, branch1, leaf1, leaf2, leaf3
Set root = New Composite
root.Name = "总部"

Set branch1 = New Composite
branch1.Name = "分公司"

Set leaf1 = New Leaf
leaf1.Name = "员工A"
Set leaf2 = New Leaf
leaf2.Name = "员工B"
Set leaf3 = New Leaf
leaf3.Name = "员工C"

branch1.Add leaf1
branch1.Add leaf2
root.Add branch1
root.Add leaf3

root.Operation ""   ' 统一遍历整棵树

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
