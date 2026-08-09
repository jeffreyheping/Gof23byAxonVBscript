<%
' 实现层接口
Class IRenderer
    Public Function RenderCircle(radius As Integer)
    End Function
End Class

' 抽象层接口
Class IShape
    Public Function Draw
    End Function
End Class

' 矢量渲染
Class VectorRenderer
    Implements IRenderer
    Public Function IRenderer_RenderCircle(radius As Integer)
        Response.Write("矢量引擎绘制半径" & radius & "的圆")

    End Function
End Class

' 光栅渲染
Class RasterRenderer
    Implements IRenderer
    Public Function IRenderer_RenderCircle(radius As Integer)
        Response.Write("光栅引擎绘制半径" & radius & "的圆")

    End Function
End Class

' 具体形状：通过组合持有实现层
Class Circle
    Implements IShape
    Private m_Radius As Integer
    Private m_Renderer As IRenderer

    Public Function Init(radius As Integer, renderer As IRenderer)
        m_Radius = radius
        Set m_Renderer = renderer
    End Function

    Public Function IShape_Draw
        m_Renderer.RenderCircle(m_Radius)
    End Function
End Class

' 演示：同形状搭配不同引擎，通过 IShape 接口引用调用 Draw
Dim c1 As Circle, c2 As Circle
Dim s1 As IShape, s2 As IShape

Set c1 = New Circle
c1.Init 5, New VectorRenderer
Set s1 = c1
s1.Draw()   ' 矢量引擎绘制半径5的圆


Set c2 = New Circle
c2.Init 5, New RasterRenderer
Set s2 = c2
s2.Draw()   ' 光栅引擎绘制半径5的圆
%>