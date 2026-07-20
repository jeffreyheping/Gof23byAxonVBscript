Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
' ===== 实现层：渲染引擎（可独立扩展） =====

' 矢量渲染引擎
Class VectorRenderer
    ' 用矢量方式绘制圆
    Public Function RenderCircle(radius)
        Response.Write "矢量引擎绘制半径" & radius & "的圆"
    End Function
End Class

' 光栅渲染引擎
Class RasterRenderer
    ' 用光栅方式绘制圆
    Public Function RenderCircle(radius)
        Response.Write "光栅引擎绘制半径" & radius & "的圆"
    End Function
End Class

' ===== 抽象层：形状（持有实现层引用） =====

Class Circle
    Private m_Radius
    Private m_Renderer   ' 桥接的渲染引擎

    ' 初始化：传入半径和渲染引擎
    Public Function Init(radius, renderer)
        m_Radius = radius
        Set m_Renderer = renderer
    End Function

    ' 绘制：委托给所持引擎
    Public Function Draw
        m_Renderer.RenderCircle m_Radius
    End Function
End Class

' 演示：同一个形状，搭配不同引擎
Dim c1, c2
Set c1 = New Circle
c1.Init 5, New VectorRenderer
c1.Draw   ' 矢量引擎...

Set c2 = New Circle
c2.Init 5, New RasterRenderer
c2.Draw   ' 光栅引擎...

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
