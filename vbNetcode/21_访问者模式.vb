Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch21Module
    Public Interface IElement
        Function Accept(visitor As IVisitor) As Object
    End Interface
    Public Interface IVisitor
        Function Visit(dot As Dot) As Object
        Function Visit(circle As Circle) As Object
    End Interface
    Public Class Dot
        Implements IElement
        Public X As Integer, Y As Integer

        ' visitor.Visit(Me) 中 Me 编译期类型为 Dot，方法重载自动匹配 IVisitor.Visit(Dot)
        Public Function Accept(visitor As IVisitor) As Object Implements IElement.Accept
            visitor.Visit(Me)
        End Function
    End Class
    Public Class Circle
        Implements IElement
        Public X As Integer, Y As Integer, Radius As Integer

        Public Function Accept(visitor As IVisitor) As Object Implements IElement.Accept
            visitor.Visit(Me)   ' 自动匹配 IVisitor.Visit(Circle)
        End Function
    End Class
    Public Class DrawingVisitor
        Implements IVisitor

        Public Function Visit(dot As Dot) As Object Implements IVisitor.Visit
            Console.WriteLine($"绘制点：({dot.X},{dot.Y})")
        End Function

        Public Function Visit(circle As Circle) As Object Implements IVisitor.Visit
            Console.WriteLine($"绘制圆：中心({circle.X},{circle.Y}) 半径{circle.Radius}")
        End Function
    End Class
    Sub Main()

        ' 访问者接口：每个具体元素对应一个 Visit 重载

        ' 具体元素：点

        ' 具体元素：圆

        ' 具体访问者：绘图

        ' 演示
        Dim d As IElement = New Dot With {.X = 10, .Y = 20}
        Dim c As IElement = New Circle With {.X = 5, .Y = 5, .Radius = 10}
        Dim drawer As IVisitor = New DrawingVisitor()
        d.Accept(drawer)   ' 绘制点：(10,20)
        c.Accept(drawer)   ' 绘制圆：中心(5,5) 半径10
    End Sub
End Module
