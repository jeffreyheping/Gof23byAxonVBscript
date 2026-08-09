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
        Response.Write("矢量引擎绘制半径" & radius & "的圆")

    End Function
End Class

' 光栅渲染引擎
Class RasterRenderer
    ' 用光栅方式绘制圆
    Public Function RenderCircle(radius)
        Response.Write("光栅引擎绘制半径" & radius & "的圆")

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
        m_Renderer.RenderCircle(m_Radius)

    End Function
End Class

' 演示：同一个形状，搭配不同引擎
Dim c1, c2
Set c1 = New Circle
c1.Init 5, New VectorRenderer
c1.Draw()   ' 矢量引擎...


Set c2 = New Circle
c2.Init 5, New RasterRenderer
c2.Draw()   ' 光栅引擎...

```

**传统 VBScript 版妥协说明**：
- **无抽象类**：经典桥接模式要求 Abstraction（形状）是抽象基类，子类（圆、方形）继承并扩展。VBScript 无继承，Circle 只是一个普通类，无法形成 Shape 抽象层级。
- **无接口约束**：Renderer 没有 `IRenderer` 接口保证 `RenderCircle` 存在，也无法用 embedding 复用公共字段。
- **缺 embedding**：每种形状都要手动写一遍 `m_Renderer` 字段、`Init` 注入逻辑、`Draw` 委托骨架，形状越多样板代码越多。

### Axon VBScript 版（支持 Implements）

```vba
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

```

**Axon VBScript 版妥协说明**：
- 接口机制解决了抽象与实现的分离问题：`Circle` 通过 `IRenderer` 接口引用调用渲染器，可在运行时自由搭配不同引擎；调用方通过 `IShape` 接口引用调用 `Draw`，自动路由到 `Circle.IShape_Draw`，无需知道具体形状类型，体现了接口多态派发。
- 缺失语法点：**代码复用机制（继承或 struct embedding）**。经典桥接要求 Abstraction 是抽象基类，所有形状共享 `m_Renderer` 字段、注入逻辑、`Resize()`/`Move()` 等公共方法只需在基类写一次。Go 同样无继承，但 Go 用 struct embedding（`type Circle struct { Shape }`）让 `Circle` 嵌入基础 `Shape` 结构体即可"免费获得"全部公共字段与方法。AxonASP 目前只能为每种形状（Circle、Square、Triangle…）各写一个独立类，手动复制 `m_Renderer` 字段、`Init` 的 renderer 注入、`Draw` 之前的参数校验等公共代码——每新增一种形状样板代码就复制一份。

### VB.NET 版（语法完备的对照基准）

VB.NET 用 `MustInherit Shape` 抽象基类持有 `Protected m_Renderer As IRenderer`，具体形状通过 `Inherits Shape` 共享 renderer 引用与注入逻辑；实现层用 `Interface IRenderer` + 多个具体引擎。场景与 Axon 版一致：同一个 `Circle` 搭配不同引擎产出不同效果。

```vbnet
' ===== 实现层：渲染引擎接口 + 具体实现 =====

Public Interface IRenderer
    Function RenderCircle(radius As Integer) As Object
End Interface

Public Class VectorRenderer
    Implements IRenderer
    Public Function RenderCircle(radius As Integer) As Object Implements IRenderer.RenderCircle
        Console.WriteLine("矢量引擎绘制半径" & radius & "的圆")
    End Function
End Class

Public Class RasterRenderer
    Implements IRenderer
    Public Function RenderCircle(radius As Integer) As Object Implements IRenderer.RenderCircle
        Console.WriteLine("光栅引擎绘制半径" & radius & "的圆")
    End Function
End Class

' ===== 抽象层：MustInherit 基类，子类共享 m_Renderer =====

Public MustInherit Class Shape
    Protected m_Renderer As IRenderer

    Protected Sub New(renderer As IRenderer)
        m_Renderer = renderer
    End Sub

    Public MustOverride Function Draw() As Object
End Class

' ===== 具体形状：Inherits Shape，自动获得 m_Renderer =====

Public Class Circle
    Inherits Shape

    Private ReadOnly m_Radius As Integer

    Public Sub New(radius As Integer, renderer As IRenderer)
        MyBase.New(renderer)
        m_Radius = radius
    End Sub

    Public Overrides Function Draw() As Object
        m_Renderer.RenderCircle(m_Radius)
    End Function
End Class

' 演示：同一个形状搭配不同引擎
Dim c1 As Shape = New Circle(5, New VectorRenderer())
c1.Draw()   ' 矢量引擎绘制半径5的圆

Dim c2 As Shape = New Circle(5, New RasterRenderer())
c2.Draw()   ' 光栅引擎绘制半径5的圆
```

**VB.NET 版说明**：
- **`Protected` 字段通过继承共享**：`m_Renderer` 在 `Shape` 基类写一次，`Circle` 通过 `Inherits Shape` 直接使用，无需每个形状类重复声明。Axon 版每个形状类都要各自写 `Private m_Renderer`。
- **`MustInherit` + `MustOverride` 编译期契约**：`Shape` 禁止直接实例化，`Draw` 强制子类实现。Axon 版 `IShape` 接口能约束方法存在，但无法禁止实例化。
- **带参构造 + `MyBase.New` 替代 Init**：`New Circle(5, renderer)` 一步完成注入，不存在"先 New 后 Init"的半初始化窗口。
- **两条继承线独立扩展**：加新形状只需 `Inherits Shape`，加新引擎只需 `Implements IRenderer`，互不干扰。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 形状抽象层 | 孤立普通类 Circle | `IShape` 接口 + 独立 Circle 类 | `MustInherit Shape` 基类 + `Inherits` 具体形状 |
| 实现层约束 | 方法名约定 | `Implements IRenderer` 接口约束 | `Interface IRenderer` 编译期强制 |
| m_Renderer 复用 | 每个形状手动复制字段 | 每个形状手动复制字段 | 基类 `Protected` 字段，子类继承获得 |
| renderer 注入 | `Init` 两步（易忘） | `Init` 两步（易忘） | 带参构造 `New(radius, renderer)` 一步到位 |
---