<%
Option Explicit
' 访问者接口
Class IVisitor
    Public Function VisitDot(dot As Dot)
    End Function
    Public Function VisitCircle(circle As Circle)
    End Function
End Class

' 元素接口
Class IElement
    Public Function Accept(visitor As IVisitor)
    End Function
End Class

' 具体元素：点
Class Dot
    Implements IElement
    Public x As Integer, y As Integer

    Public Function IElement_Accept(visitor As IVisitor)
        visitor.VisitDot(Me)

    End Function
End Class

' 具体元素：圆
Class Circle
    Implements IElement
    Public x As Integer, y As Integer, radius As Integer

    Public Function IElement_Accept(visitor As IVisitor)
        visitor.VisitCircle(Me)

    End Function
End Class

' 具体访问者：绘图
Class DrawingVisitor
    Implements IVisitor
    Public Function IVisitor_VisitDot(dot As Dot)
        Response.Write("绘制点：(" & dot.x & "," & dot.y & ")")

    End Function
    Public Function IVisitor_VisitCircle(circle As Circle)
        Response.Write("绘制圆：中心(" & circle.x & "," & circle.y & ") 半径" & circle.radius)

    End Function
End Class

' 演示
Dim dObj As Dot
Dim cObj As Circle
Set dObj = New Dot
dObj.x = 10
dObj.y = 20
Set cObj = New Circle
cObj.x = 5
cObj.y = 5
cObj.radius = 10

Dim d As IElement, c As IElement
Set d = dObj
Set c = cObj

Dim drawer As IVisitor
Set drawer = New DrawingVisitor
d.Accept(drawer)

c.Accept(drawer)
%>