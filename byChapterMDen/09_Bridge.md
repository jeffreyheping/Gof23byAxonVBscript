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
- Missing syntax: **Code reuse mechanism** (inheritance or struct embedding). The classic Bridge requires the Abstraction to be an abstract base class, with subclasses inheriting and reusing code. Go also lacks inheritance, but Go uses struct embedding to let `Circle` embed a base `Shape` struct and reuse common logic. AxonASP currently requires writing a separate class for each shape, manually copying common fields and logic.
---
