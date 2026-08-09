Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch10Module
    Public Interface IComponent
        Property Name As String
        Function Add(child As IComponent) As Object
        Function Operation(indent As String) As Object
    End Interface
    Public MustInherit Class ComponentBase
        Implements IComponent

        Public Overridable Property Name As String Implements IComponent.Name

        ' 默认空实现：叶子节点继承此默认，组合节点重写
        Public Overridable Function Add(child As IComponent) As Object Implements IComponent.Add
        End Function

        Public MustOverride Function Operation(indent As String) As Object Implements IComponent.Operation
    End Class
    Public Class Leaf
        Inherits ComponentBase

        Public Overrides Function Operation(indent As String) As Object
            Console.WriteLine(indent & "叶子：" & Name)
        End Function
    End Class
    Public Class Composite
        Inherits ComponentBase

        Private ReadOnly m_Children As New List(Of IComponent)()

        Public Overrides Function Add(child As IComponent) As Object
            m_Children.Add(child)
        End Function

        Public Overrides Function Operation(indent As String) As Object
            Console.WriteLine(indent & "组合：" & Name)
            For Each child In m_Children
                child.Operation(indent & "  ")
            Next
        End Function
    End Class
    Sub Main()

        ' ② MustInherit 基类：Name 属性写一次，子类共享；默认 Add 空实现（叶子继承）

        ' ③ 叶子节点：仅重写 Operation，其余继承基类默认

        ' ④ 组合节点：用 List(Of IComponent) 管理子节点，递归遍历

        ' 演示：构建 总部→分公司→员工 树，统一接口遍历
        Dim root As IComponent = New Composite With {.Name = "总部"}
        Dim branch1 As IComponent = New Composite With {.Name = "分公司"}
        Dim leaf1 As IComponent = New Leaf With {.Name = "员工A"}
        Dim leaf2 As IComponent = New Leaf With {.Name = "员工B"}
        Dim leaf3 As IComponent = New Leaf With {.Name = "员工C"}

        root.Add(branch1)
        root.Add(leaf3)
        branch1.Add(leaf1)
        branch1.Add(leaf2)

        root.Operation("")
        ' 组合：总部
        '   组合：分公司
        '     叶子：员工A
        '     叶子：员工B
        '   叶子：员工C
    End Sub
End Module
