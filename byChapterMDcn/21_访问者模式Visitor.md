## 第21章 访问者模式（Visitor）

**核心思想**：把对对象结构的操作封装到访问者中，新增操作无需修改元素类。

**示例说明**：Dot 和 Circle 都接受 DrawingVisitor，分别调用 VisitDot 和 VisitCircle。新增 ExportVisitor 时，无需修改 Dot 和 Circle，只需新增一个访问者类。

### 传统 VBScript 版

```vbscript
' 元素：点
Class Dot
    Public x, y

    ' 接受访问者，调用访问者的 VisitDot
    Public Function Accept(visitor)
        visitor.VisitDot Me
    End Function
End Class

' 元素：圆
Class Circle
    Public x, y, radius

    ' 接受访问者，调用访问者的 VisitCircle
    Public Function Accept(visitor)
        visitor.VisitCircle Me
    End Function
End Class

' 访问者：绘图
Class DrawingVisitor
    ' 访问点
    Public Function VisitDot(dot)
        Response.Write "绘制点：(" & dot.x & "," & dot.y & ")"
    End Function
    ' 访问圆
    Public Function VisitCircle(circle)
        Response.Write "绘制圆：中心(" & circle.x & "," & circle.y & ") 半径" & circle.radius
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
d.Accept drawer
c.Accept drawer
```

**传统 VBScript 版妥协说明**：
- **无接口约束**：Dot 和 Circle 没有 `IElement` 接口强制 `Accept` 方法，DrawingVisitor 也没有 `IVisitor` 接口强制 `VisitDot`/`VisitCircle`。如果类名或方法名不一致，运行时调用才报错。
- **双分派缺失**：经典访问者依赖语言的双分派机制（根据对象类型和访问者类型自动选择方法）。VBScript 无多态重载，必须在 `Accept` 中显式调用 `visitor.VisitDot` 或 `visitor.VisitCircle`，新增元素类型时需要修改所有访问者。

### Axon VBScript 版（支持 Implements）

```vbscript
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
        visitor.VisitDot Me
    End Function
End Class

' 具体元素：圆
Class Circle
    Implements IElement
    Public x As Integer, y As Integer, radius As Integer

    Public Function IElement_Accept(visitor As IVisitor)
        visitor.VisitCircle Me
    End Function
End Class

' 具体访问者：绘图
Class DrawingVisitor
    Implements IVisitor
    Public Function IVisitor_VisitDot(dot As Dot)
        Response.Write "绘制点：(" & dot.x & "," & dot.y & ")"
    End Function
    Public Function IVisitor_VisitCircle(circle As Circle)
        Response.Write "绘制圆：中心(" & circle.x & "," & circle.y & ") 半径" & circle.radius
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
d.Accept drawer
c.Accept drawer
```

**Axon VBScript 版妥协说明**：
- `IElement`/`IVisitor` 接口约束了元素和访问者的契约。AxonASP 接口方法派发已修复，元素在 `IElement_Accept` 中可直接调用访问者的 `VisitDot`/`VisitCircle` 接口方法，演示中通过 `IElement` 类型变量调用 `Accept` 即自动路由，无需 `DoAccept` 辅助方法或完整限定名。剩余限制：缺失语法点：**方法重载或双分派**。Go 同样无方法重载——Go 用 **type switch**（`switch v := element.(type)`）实现访问者的双分派，无需为每种元素写单独的 `Visit` 方法。AxonASP 目前只能在 `Accept` 中显式调用 `VisitDot`/`VisitCircle` 分支，新增元素类型时需修改所有访问者。
---