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

```

**传统 VBScript 版妥协说明**：
- **无接口约束**：Dot 和 Circle 没有 `IElement` 接口强制 `Accept` 方法，DrawingVisitor 也没有 `IVisitor` 接口强制 `VisitDot`/`VisitCircle`。如果类名或方法名不一致，运行时调用才报错。
- **双分派缺失**：经典访问者依赖语言的双分派机制（根据对象类型和访问者类型自动选择方法）。VBScript 无多态重载，必须在 `Accept` 中显式调用 `visitor.VisitDot` 或 `visitor.VisitCircle`，新增元素类型时需要修改所有访问者。

### Axon VBScript 版（支持 Implements）

```vba
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

```

**Axon VBScript 版妥协说明**：
- 改善但残留：缺方法重载/双分派。`IElement`/`IVisitor` 接口约束了元素和访问者的契约。AxonASP 接口方法派发已修复，元素在 `IElement_Accept` 中可直接调用访问者的 `VisitDot`/`VisitCircle` 接口方法，演示中通过 `IElement` 类型变量调用 `Accept` 即自动路由，无需 `DoAccept` 辅助方法或完整限定名。剩余限制：缺失语法点——**方法重载/双分派**。经典访问者需要语言层面的双分派：`visitor.Visit(element)` 能根据 `element` 的运行时类型自动匹配到 `Visit(Dot)` 或 `Visit(Circle)` 的重载版本。AxonASP 无方法重载，必须在每个元素的 `Accept` 中手动分支调用 `visitor.VisitDot Me` / `visitor.VisitCircle Me`，新增元素类型时需修改所有 `IVisitor` 接口 + 所有实现接口的访问者类。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有方法重载（Overloads），访问者模式与 Axon 版同结构：保留 `IElement`/`IVisitor` 接口 + `Dot`/`Circle`/`DrawingVisitor`，仅用 `Visit(Dot)`/`Visit(Circle)` 重载统一 `Accept` 中的分派写法，不引入 dynamic/DLR 等额外机制。

```vbnet
' 元素接口
Public Interface IElement
    Sub Accept(visitor As IVisitor)
End Interface

' 访问者接口：每个具体元素对应一个 Visit 重载
Public Interface IVisitor
    Sub Visit(dot As Dot)
    Sub Visit(circle As Circle)
End Interface

' 具体元素：点
Public Class Dot
    Implements IElement
    Public X As Integer, Y As Integer

    ' visitor.Visit(Me) 中 Me 编译期类型为 Dot，方法重载自动匹配 IVisitor.Visit(Dot)
    Public Sub Accept(visitor As IVisitor) Implements IElement.Accept
        visitor.Visit(Me)
    End Sub
End Class

' 具体元素：圆
Public Class Circle
    Implements IElement
    Public X As Integer, Y As Integer, Radius As Integer

    Public Sub Accept(visitor As IVisitor) Implements IElement.Accept
        visitor.Visit(Me)   ' 自动匹配 IVisitor.Visit(Circle)
    End Sub
End Class

' 具体访问者：绘图
Public Class DrawingVisitor
    Implements IVisitor

    Public Sub Visit(dot As Dot) Implements IVisitor.Visit
        Console.WriteLine($"绘制点：({dot.X},{dot.Y})")
    End Sub

    Public Sub Visit(circle As Circle) Implements IVisitor.Visit
        Console.WriteLine($"绘制圆：中心({circle.X},{circle.Y}) 半径{circle.Radius}")
    End Sub
End Class

' 演示
Dim d As IElement = New Dot With {.X = 10, .Y = 20}
Dim c As IElement = New Circle With {.X = 5, .Y = 5, .Radius = 10}
Dim drawer As IVisitor = New DrawingVisitor()
d.Accept(drawer)   ' 绘制点：(10,20)
c.Accept(drawer)   ' 绘制圆：中心(5,5) 半径10
```

**VB.NET 版说明**：
- **方法重载 = 统一的 `Visit(Me)` 分派**：Axon 版 `Accept` 里必须写死 `visitor.VisitDot Me` / `visitor.VisitCircle Me`，不同元素方法名不同；VB.NET 版所有元素 `Accept` 都写 `visitor.Visit(Me)`，编译器按 `Me` 类型自动匹配 `Visit(Dot)`/`Visit(Circle)` 重载。
- **`Interface` 重载 + `Implements` 编译期强制**：`IVisitor` 接口内 `Visit(Dot)`/`Visit(Circle)` 重载声明，漏写 `Implements IVisitor.Visit` 直接报错；与 Axon 版同结构，仅方法名免去 `VisitDot`/`VisitCircle` 分支。
- **对象初始化器一行赋值**：`New Dot With {.X = 10, .Y = 20}` 替代 Axon 版逐字段 `dObj.x = 10` 赋值。
- **无需 `Set` 区分对象赋值**：`Dim d As IElement = New Dot()` 直接赋值给接口变量，前两版需 `Set d = dObj`。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Accept 分派 | 写死 `visitor.VisitDot`/`VisitCircle` | 写死 `visitor.VisitDot`/`VisitCircle` | 统一 `visitor.Visit(Me)`，重载匹配 |
| 访问者契约 | 方法名约定 | `IVisitor`/`IElement` 接口 | `Interface` 重载 + `Implements` 编译期强制 |
| 元素字段初始化 | 逐字段赋值 | 逐字段赋值 | 对象初始化器 `With {.X=10}` 一行 |
| 对象赋值 | `Set d = New Dot` | `Set d = dObj` | 直接 `d = New Dot()` |
---