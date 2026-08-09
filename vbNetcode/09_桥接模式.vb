Option Strict Off
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch09Module
    Public Interface IRenderer
        Function RenderCircle(radius As Integer)
    End Interface
    Public Class VectorRenderer
        Implements IRenderer
        Public Function RenderCircle(radius As Integer) Implements IRenderer.RenderCircle
            Console.WriteLine("矢量引擎绘制半径" & radius & "的圆")
        End Function
    End Class
    Public Class RasterRenderer
        Implements IRenderer
        Public Function RenderCircle(radius As Integer) Implements IRenderer.RenderCircle
            Console.WriteLine("光栅引擎绘制半径" & radius & "的圆")
        End Function
    End Class
    Public MustInherit Class Shape
        Protected m_Renderer As IRenderer

        Protected Sub New(renderer As IRenderer)
            m_Renderer = renderer
        End Sub

        Public MustOverride Function Draw()
    End Class
    Public Class Circle
        Inherits Shape

        Private ReadOnly m_Radius As Integer

        Public Sub New(radius As Integer, renderer As IRenderer)
            MyBase.New(renderer)
            m_Radius = radius
        End Sub

        Public Overrides Function Draw()
            m_Renderer.RenderCircle(m_Radius)
        End Function
    End Class
    Sub Main()



        ' ===== 抽象层：MustInherit 基类，子类共享 m_Renderer =====


        ' ===== 具体形状：Inherits Shape，自动获得 m_Renderer =====


        ' 演示：同一个形状搭配不同引擎
        Dim c1 As Shape = New Circle(5, New VectorRenderer())
        c1.Draw()   ' 矢量引擎绘制半径5的圆

        Dim c2 As Shape = New Circle(5, New RasterRenderer())
        c2.Draw()   ' 光栅引擎绘制半径5的圆
    End Sub
End Module
