<%
Option Explicit
' 元素：点
Class Dot
    Public x, y

    ' 接受访问者，调用访问者的 VisitDot
    Public Function Accept(visitor)
        visitor.VisitDot(Me)

    End Function
End Class

' 元素：圆
Class Circle
    Public x, y, radius

    ' 接受访问者，调用访问者的 VisitCircle
    Public Function Accept(visitor)
        visitor.VisitCircle(Me)

    End Function
End Class

' 访问者：绘图
Class DrawingVisitor
    ' 访问点
    Public Function VisitDot(dot)
        Response.Write("绘制点：(" & dot.x & "," & dot.y & ")")

    End Function
    ' 访问圆
    Public Function VisitCircle(circle)
        Response.Write("绘制圆：中心(" & circle.x & "," & circle.y & ") 半径" & circle.radius)

    End Function
End Class

' 演示：对同一组元素执行不同操作
Dim d, c, drawer
Set d = New Dot
d.x = 10
d.y = 20
Set c = New Circle
c.x = 5
c.y = 5
c.radius = 10

Set drawer = New DrawingVisitor
d.Accept(drawer)

c.Accept(drawer)
%>