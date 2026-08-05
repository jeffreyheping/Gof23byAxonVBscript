## 第9章 桥接模式（Bridge）

**核心思想**：把抽象和实现分离，使它们可以独立变化。

**示例说明**：Circle（抽象层）持有 Renderer（实现层）的引用。同一个 Circle 搭配 VectorRenderer 或 RasterRenderer，可以产出不同渲染效果。形状和渲染引擎可以各自独立扩展。

### 传统 VBScript 版

```vbscript
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
```

**传统 VBScript 版妥协说明**：
- **无抽象类**：经典桥接模式要求 Abstraction（形状）是抽象类，子类（圆、方形）继承并扩展。VBScript 无继承，Circle 只是一个普通类，无法形成 Shape 抽象层级。
- **无接口约束**：Renderer 没有 `IRenderer` 接口保证 `RenderCircle` 存在。

### Axon VBScript 版（支持 Implements）

```vbscript
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
        Response.Write "矢量引擎绘制半径" & radius & "的圆"
    End Function
End Class

' 光栅渲染
Class RasterRenderer
    Implements IRenderer
    Public Function IRenderer_RenderCircle(radius As Integer)
        Response.Write "光栅引擎绘制半径" & radius & "的圆"
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
s1.Draw   ' 矢量引擎绘制半径5的圆

Set c2 = New Circle
c2.Init 5, New RasterRenderer
Set s2 = c2
s2.Draw   ' 光栅引擎绘制半径5的圆
```

**Axon VBScript 版妥协说明**：
- 接口机制解决了抽象与实现的分离问题：`Circle` 通过 `IRenderer` 接口引用调用渲染器，可在运行时自由搭配不同引擎；调用方通过 `IShape` 接口引用调用 `Draw`，自动路由到 `Circle.IShape_Draw`，无需知道具体形状类型，体现了接口多态派发。
- 缺失语法点：**代码复用机制**（继承或 struct embedding）。经典桥接要求 Abstraction 是抽象基类，子类继承复用。Go 同样无继承，但 Go 用 struct embedding 让 `Circle` 嵌入一个基础 `Shape` 结构体即可复用公共逻辑。AxonASP 目前只能为每种形状各写一个类，手动复制公共字段和逻辑。
---