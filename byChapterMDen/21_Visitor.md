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
- Improved but with residual gaps: no method overloading/double dispatch. `IElement`/`IVisitor` interfaces constrain the element and visitor contracts. AxonASP's interface method dispatch is fixed — elements call the visitor's `VisitDot`/`VisitCircle` interface methods directly in `IElement_Accept`. Calling `Accept` through `IElement`-typed variables auto-dispatches — no `DoAccept` helper or fully-qualified names needed. Remaining gap: **Method overloading/double dispatch**. Classic Visitor requires language-level double dispatch: `visitor.Visit(element)` should auto-match `Visit(Dot)` or `Visit(Circle)` overloads based on `element`'s runtime type. AxonASP has no method overloading — each element's `Accept` must manually branch to `visitor.VisitDot Me` / `visitor.VisitCircle Me`. Adding a new element type requires modifying all `IVisitor` interfaces + all visitor classes.

### VB.NET Version (syntactically complete baseline)

VB.NET has method overloading (`Overloads`). The Visitor pattern follows the same structure as the Axon version: `IElement`/`IVisitor` interfaces + `Dot`/`Circle`/`DrawingVisitor`. Using `Visit(Dot)`/`Visit(Circle)` overloads unifies the dispatch in `Accept` — no dynamic/DLR or other extra mechanisms needed.

```vbnet
' Element interface
Public Interface IElement
    Sub Accept(visitor As IVisitor)
End Interface

' Visitor interface: one Visit overload per concrete element
Public Interface IVisitor
    Sub Visit(dot As Dot)
    Sub Visit(circle As Circle)
End Interface

' Concrete element: Dot
Public Class Dot
    Implements IElement
    Public X As Integer, Y As Integer

    ' In visitor.Visit(Me), Me's compile-time type is Dot; overload auto-matches IVisitor.Visit(Dot)
    Public Sub Accept(visitor As IVisitor) Implements IElement.Accept
        visitor.Visit(Me)
    End Sub
End Class

' Concrete element: Circle
Public Class Circle
    Implements IElement
    Public X As Integer, Y As Integer, Radius As Integer

    Public Sub Accept(visitor As IVisitor) Implements IElement.Accept
        visitor.Visit(Me)   ' Auto-matches IVisitor.Visit(Circle)
    End Sub
End Class

' Concrete visitor: Drawing
Public Class DrawingVisitor
    Implements IVisitor

    Public Sub Visit(dot As Dot) Implements IVisitor.Visit
        Console.WriteLine($"Drawing dot: ({dot.X},{dot.Y})")
    End Sub

    Public Sub Visit(circle As Circle) Implements IVisitor.Visit
        Console.WriteLine($"Drawing circle: center({circle.X},{circle.Y}) radius{circle.Radius}")
    End Sub
End Class

' Demo
Dim d As IElement = New Dot With {.X = 10, .Y = 20}
Dim c As IElement = New Circle With {.X = 5, .Y = 5, .Radius = 10}
Dim drawer As IVisitor = New DrawingVisitor()
d.Accept(drawer)   ' Drawing dot: (10,20)
c.Accept(drawer)   ' Drawing circle: center(5,5) radius10
```

**VB.NET version notes**:
- **Method overloading = unified `Visit(Me)` dispatch**: Axon's `Accept` must hardcode `visitor.VisitDot Me` / `visitor.VisitCircle Me` with different method names per element; VB.NET's `Accept` always writes `visitor.Visit(Me)` — the compiler auto-matches `Visit(Dot)`/`Visit(Circle)` overloads by `Me`'s type.
- **`Interface` overloading + `Implements` compile-time enforcement**: `IVisitor` declares `Visit(Dot)`/`Visit(Circle)` overloads; forgetting `Implements IVisitor.Visit` is a compile error. Same structure as Axon, just without `VisitDot`/`VisitCircle` branching.
- **Object initializer one-liner**: `New Dot With {.X = 10, .Y = 20}` replaces Axon's field-by-field `dObj.x = 10` assignment.
- **No `Set` for object assignment**: `Dim d As IElement = New Dot()` assigns directly to an interface variable; the other two versions need `Set d = dObj`.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|-----------|-----------------|---------------|--------|
| Accept dispatch | Hardcoded `VisitDot`/`VisitCircle` | Hardcoded `VisitDot`/`VisitCircle` | Unified `visitor.Visit(Me)`, overload matching |
| Visitor contract | Method name convention | `IVisitor`/`IElement` interface | `Interface` overloading + `Implements` compile-time |
| Element field init | Field-by-field | Field-by-field | Object initializer `With {.X=10}` one-liner |
| Object assignment | `Set d = New Dot` | `Set d = dObj` | Direct `d = New Dot()` |

---
