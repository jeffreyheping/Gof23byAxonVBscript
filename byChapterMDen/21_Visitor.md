## Chapter 21: Visitor

**Core idea**: Encapsulate operations on an object structure in a visitor — adding new operations without modifying element classes.

**Example**: Both Dot and Circle accept a DrawingVisitor, calling VisitDot and VisitCircle respectively. When adding a new ExportVisitor, there's no need to modify Dot or Circle — just add a new visitor class.

### Classic VBScript Version

```vbscript
' Element: point
Class Dot
    Public x, y

    ' Accept visitor, call visitor's VisitDot
    Public Function Accept(visitor)
        visitor.VisitDot Me
    End Function
End Class

' Element: circle
Class Circle
    Public x, y, radius

    ' Accept visitor, call visitor's VisitCircle
    Public Function Accept(visitor)
        visitor.VisitCircle Me
    End Function
End Class

' Visitor: drawing
Class DrawingVisitor
    ' Visit point
    Public Function VisitDot(dot)
        Response.Write "Draw dot: (" & dot.x & "," & dot.y & ")"
    End Function
    ' Visit circle
    Public Function VisitCircle(circle)
        Response.Write "Draw circle: center (" & circle.x & "," & circle.y & ") radius " & circle.radius
    End Function
End Class

' Demo: apply different operations to the same set of elements
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

**Classic VBScript trade-offs**:
- **No interface constraint**: Dot and Circle have no `IElement` interface forcing `Accept`. DrawingVisitor has no `IVisitor` interface forcing `VisitDot`/`VisitCircle`. If class or method names are inconsistent, the error only surfaces at runtime.
- **No double dispatch**: The classic Visitor relies on double dispatch (automatically selecting a method based on both object type and visitor type). VBScript has no polymorphic overloading — `Accept` must explicitly call `visitor.VisitDot` or `visitor.VisitCircle`. Adding a new element type requires modifying all visitors.

### Axon VBScript Version (supports Implements)

```vbscript
' Visitor interface
Class IVisitor
    Public Function VisitDot(dot As Dot)
    End Function
    Public Function VisitCircle(circle As Circle)
    End Function
End Class

' Element interface
Class IElement
    Public Function Accept(visitor As IVisitor)
    End Function
End Class

' Concrete element: point
Class Dot
    Implements IElement
    Public x As Integer, y As Integer

    Public Function IElement_Accept(visitor As IVisitor)
        visitor.VisitDot Me
    End Function
End Class

' Concrete element: circle
Class Circle
    Implements IElement
    Public x As Integer, y As Integer, radius As Integer

    Public Function IElement_Accept(visitor As IVisitor)
        visitor.VisitCircle Me
    End Function
End Class

' Concrete visitor: drawing
Class DrawingVisitor
    Implements IVisitor
    Public Function IVisitor_VisitDot(dot As Dot)
        Response.Write "Draw dot: (" & dot.x & "," & dot.y & ")"
    End Function
    Public Function IVisitor_VisitCircle(circle As Circle)
        Response.Write "Draw circle: center (" & circle.x & "," & circle.y & ") radius " & circle.radius
    End Function
End Class

' Demo
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

**Axon VBScript trade-offs**:
- `IElement`/`IVisitor` interfaces constrain the element and visitor contracts. AxonASP's interface method dispatch is fixed — elements call the visitor's `VisitDot`/`VisitCircle` interface methods directly in `IElement_Accept`. The demo calls `Accept` through `IElement`-typed variables which auto-dispatches — no `DoAccept` helper or fully-qualified names needed. Remaining gap: **Method overloading or double dispatch**. Go also lacks method overloading — Go uses **type switch** (`switch v := element.(type)`) for the visitor's double dispatch, without writing a separate `Visit` method for each element type. AxonASP currently requires explicit `VisitDot`/`VisitCircle` branches in `Accept` — adding a new element type requires modifying all visitors.
---
