## Chapter 9: Bridge

**Core idea**: Separate abstraction from implementation so they can vary independently.

**Example**: Circle (abstraction) holds a reference to Renderer (implementation). The same Circle paired with VectorRenderer or RasterRenderer produces different rendering results. Shapes and rendering engines can each be extended independently.

### Classic VBScript Version

```vbscript
' ===== Implementation layer: rendering engines (extensible independently) =====

' Vector rendering engine
Class VectorRenderer
    ' Draw a circle using vector graphics
    Public Function RenderCircle(radius)
        Response.Write "Vector engine draws circle with radius " & radius
    End Function
End Class

' Raster rendering engine
Class RasterRenderer
    ' Draw a circle using raster graphics
    Public Function RenderCircle(radius)
        Response.Write "Raster engine draws circle with radius " & radius
    End Function
End Class

' ===== Abstraction layer: shapes (hold implementation layer reference) =====

Class Circle
    Private m_Radius
    Private m_Renderer   ' Bridged rendering engine

    ' Init: pass in radius and rendering engine
    Public Function Init(radius, renderer)
        m_Radius = radius
        Set m_Renderer = renderer
    End Function

    ' Draw: delegate to the held engine
    Public Function Draw
        m_Renderer.RenderCircle m_Radius
    End Function
End Class

' Demo: same shape, different engines
Dim c1, c2
Set c1 = New Circle
c1.Init 5, New VectorRenderer
c1.Draw   ' Vector engine...

Set c2 = New Circle
c2.Init 5, New RasterRenderer
c2.Draw   ' Raster engine...
```

**Classic VBScript trade-offs**:
- **No abstract class**: The classic Bridge pattern requires the Abstraction (Shape) to be an abstract class with subclasses (Circle, Square) inheriting and extending it. VBScript has no inheritance. Circle is just a plain class — no Shape abstraction layer.
- **No interface constraint**: Renderer has no `IRenderer` interface guaranteeing `RenderCircle` exists.

### Axon VBScript Version (supports Implements)

```vbscript
' Implementation layer interface
Class IRenderer
    Public Function RenderCircle(radius As Integer)
    End Function
End Class

' Abstraction layer interface
Class IShape
    Public Function Draw
    End Function
End Class

' Vector rendering
Class VectorRenderer
    Implements IRenderer
    Public Function IRenderer_RenderCircle(radius As Integer)
        Response.Write "Vector engine draws circle with radius " & radius
    End Function
End Class

' Raster rendering
Class RasterRenderer
    Implements IRenderer
    Public Function IRenderer_RenderCircle(radius As Integer)
        Response.Write "Raster engine draws circle with radius " & radius
    End Function
End Class

' Concrete shape: holds implementation via composition
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

' Demo: same shape with different engines, call Draw via IShape interface reference
Dim c1 As Circle, c2 As Circle
Dim s1 As IShape, s2 As IShape

Set c1 = New Circle
c1.Init 5, New VectorRenderer
Set s1 = c1
s1.Draw   ' Vector engine draws circle with radius 5

Set c2 = New Circle
c2.Init 5, New RasterRenderer
Set s2 = c2
s2.Draw   ' Raster engine draws circle with radius 5
```

**Axon VBScript trade-offs**:
- The interface mechanism separates abstraction from implementation: `Circle` calls the renderer via `IRenderer` interface reference, freely pairing with different engines at runtime. Callers invoke `Draw` through `IShape` interface reference, which auto-dispatches to `Circle.IShape_Draw` — no need to know the concrete shape type. This demonstrates interface polymorphic dispatch.
- Missing syntax: **code reuse mechanism (inheritance or struct embedding)**. Classic Bridge requires the Abstraction to be an abstract base class — all shapes share the `m_Renderer` field, injection logic, and common methods like `Resize()`/`Move()` written once in the base class. Go also lacks inheritance, but Go uses struct embedding (`type Circle struct { Shape }`) to let `Circle` embed a base `Shape` struct and "get for free" all common fields and methods. AxonASP currently requires a separate independent class for each shape (Circle, Square, Triangle...), manually copying `m_Renderer` field, `Init`'s renderer injection, pre-`Draw` parameter validation, etc. — every new shape means another copy of the boilerplate.

### VB.NET Version (syntactically complete baseline)

VB.NET uses `MustInherit Shape` abstract base class to hold `Protected m_Renderer As IRenderer`, concrete shapes share renderer reference and injection logic via `Inherits Shape`; implementation side uses `Interface IRenderer` + multiple concrete engines. Same scenario as Axon version: same `Circle` paired with different engines produces different results.

```vbnet
' ===== Implementation side: renderer engine interface + concrete implementations =====

Public Interface IRenderer
    Sub RenderCircle(radius As Integer)
End Interface

Public Class VectorRenderer
    Implements IRenderer
    Public Sub RenderCircle(radius As Integer) Implements IRenderer.RenderCircle
        Console.WriteLine("Vector engine draws circle radius " & radius)
    End Sub
End Class

Public Class RasterRenderer
    Implements IRenderer
    Public Sub RenderCircle(radius As Integer) Implements IRenderer.RenderCircle
        Console.WriteLine("Raster engine draws circle radius " & radius)
    End Sub
End Class

' ===== Abstraction side: MustInherit base class, subclasses share m_Renderer =====

Public MustInherit Class Shape
    Protected m_Renderer As IRenderer

    Protected Sub New(renderer As IRenderer)
        m_Renderer = renderer
    End Sub

    Public MustOverride Sub Draw()
End Class

' ===== Concrete shape: Inherits Shape, automatically gets m_Renderer =====

Public Class Circle
    Inherits Shape

    Private ReadOnly m_Radius As Integer

    Public Sub New(radius As Integer, renderer As IRenderer)
        MyBase.New(renderer)
        m_Radius = radius
    End Sub

    Public Overrides Sub Draw()
        m_Renderer.RenderCircle(m_Radius)
    End Sub
End Class

' Demo: same shape with different engines
Dim c1 As Shape = New Circle(5, New VectorRenderer())
c1.Draw()   ' Vector engine draws circle radius 5

Dim c2 As Shape = New Circle(5, New RasterRenderer())
c2.Draw()   ' Raster engine draws circle radius 5
```

**VB.NET version notes**:
- **`Protected` field shared through inheritance**: `m_Renderer` written once in `Shape` base class, `Circle` uses it directly via `Inherits Shape`, no need to repeat in every shape class. Axon version must write `Private m_Renderer` in each shape class.
- **`MustInherit` + `MustOverride` compile-time contract**: `Shape` prevents direct instantiation, `Draw` forces subclass implementation. Axon's `IShape` interface can constrain method existence, but cannot prevent instantiation.
- **Parameterized constructor + `MyBase.New` replaces Init**: `New Circle(5, renderer)` completes injection in one step, no "New then Init" half-initialized window.
- **Two inheritance lines extend independently**: Adding new shapes just needs `Inherits Shape`, adding new engines just needs `Implements IRenderer`, no interference.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Shape abstraction layer | Isolated plain Circle class | `IShape` interface + independent Circle class | `MustInherit Shape` base class + `Inherits` concrete shapes |
| Implementation layer constraint | Method name convention | `Implements IRenderer` interface constraint | `Interface IRenderer` compile-time enforced |
| m_Renderer reuse | Manual field copy per shape | Manual field copy per shape | Base class `Protected` field, inherited by subclasses |
| Renderer injection | `Init` two-step (easy to forget) | `Init` two-step (easy to forget) | Parameterized constructor `New(radius, renderer)` one-step |
---
