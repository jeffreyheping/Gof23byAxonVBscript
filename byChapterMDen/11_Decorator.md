## Chapter 11: Decorator

**Core idea**: Dynamically add new behavior to an object without modifying its class.

**Example**: SimpleCoffee is the base object. MilkDecorator and SugarDecorator each hold a `coffee` reference. Pass base into milk, then milk into sugar. The final `sugar.Cost = base.Cost + 2 + 1` — costs stack up layer by layer.

### Classic VBScript Version

```vbscript
' Base component: plain coffee
Class SimpleCoffee
    ' Return price
    Public Function Cost
        Cost = 10
    End Function
    ' Return description
    Public Function Description
        Description = "Plain coffee"
    End Function
End Class

' Decorator: milk (wraps a coffee object, adds its price)
Class MilkDecorator
    Private m_Coffee   ' The wrapped inner object

    ' Inject the decorated object
    Public Function Init(coffee)
        Set m_Coffee = coffee
    End Function
    ' Price = inner object price + milk surcharge
    Public Function Cost
        Cost = m_Coffee.Cost + 2
    End Function
    ' Description = inner object description + milk
    Public Function Description
        Description = m_Coffee.Description & " + milk"
    End Function
End Class

' Decorator: sugar (same structure as MilkDecorator)
Class SugarDecorator
    Private m_Coffee

    ' Inject the decorated object
    Public Function Init(coffee)
        Set m_Coffee = coffee
    End Function
    ' Price = inner object price + sugar surcharge
    Public Function Cost
        Cost = m_Coffee.Cost + 1
    End Function
    ' Description = inner object description + sugar
    Public Function Description
        Description = m_Coffee.Description & " + sugar"
    End Function
End Class

' Demo: wrap layer by layer, dynamically stack behavior
Dim base, milk, sugar
Set base = New SimpleCoffee
Response.Write base.Description & " = " & base.Cost & " yuan"

Set milk = New MilkDecorator
milk.Init base              ' milk wraps base

Set sugar = New SugarDecorator
sugar.Init milk             ' sugar wraps milk
Response.Write sugar.Description & " = " & sugar.Cost & " yuan"
```

**Classic VBScript trade-offs**:
- **Can't inherit base class**: The classic Decorator requires Decorator to inherit Component and hold a Component reference for "type compatibility". VBScript has no inheritance — `MilkDecorator` and `SimpleCoffee` are completely different classes, can't be substituted for each other. The caller must know whether it holds a decorator or the original object.
- **No transparency**: Ideally the decorator is transparent to the caller, but in VBScript `Set milk = New MilkDecorator` and `Set coffee = New SimpleCoffee` are different types — you can't declare a single variable type for both.

### Axon VBScript Version (supports Implements)

```vbscript
' Coffee interface
Class ICoffee
    Public Function Cost As Integer
    End Function
    Public Function Description As String
    End Function
End Class

' Base coffee
Class SimpleCoffee
    Implements ICoffee
    Public Function ICoffee_Cost As Integer
        ICoffee_Cost = 10
    End Function
    Public Function ICoffee_Description As String
        ICoffee_Description = "Plain coffee"
    End Function
End Class

' Milk decorator: holds a wrapped ICoffee reference, delegates in interface methods
Class MilkDecorator
    Implements ICoffee
    Private m_Coffee As ICoffee

    ' Inject the decorated object
    Public Function Init(coffee As ICoffee)
        Set m_Coffee = coffee
    End Function

    Public Function ICoffee_Cost As Integer
        ICoffee_Cost = m_Coffee.Cost + 2
    End Function
    Public Function ICoffee_Description As String
        ICoffee_Description = m_Coffee.Description & " + milk"
    End Function
End Class

' Sugar decorator: same structure as MilkDecorator
Class SugarDecorator
    Implements ICoffee
    Private m_Coffee As ICoffee

    ' Inject the decorated object
    Public Function Init(coffee As ICoffee)
        Set m_Coffee = coffee
    End Function

    Public Function ICoffee_Cost As Integer
        ICoffee_Cost = m_Coffee.Cost + 1
    End Function
    Public Function ICoffee_Description As String
        ICoffee_Description = m_Coffee.Description & " + sugar"
    End Function
End Class

' Demo: interface references enable transparent nesting, layer by layer
Dim coffee As ICoffee
Set coffee = New SimpleCoffee
Response.Write coffee.Description & " = " & coffee.Cost & " yuan"

Dim milk As ICoffee
Dim milkObj As MilkDecorator
Set milkObj = New MilkDecorator
milkObj.Init coffee              ' milk wraps coffee
Set milk = milkObj

Dim sugar As ICoffee
Dim sugarObj As SugarDecorator
Set sugarObj = New SugarDecorator
sugarObj.Init milk               ' sugar wraps milk
Set sugar = sugarObj
Response.Write sugar.Description & " = " & sugar.Cost & " yuan"
```

**Axon VBScript trade-offs**:
- The interface mechanism lets decorators and components implement the same `ICoffee` interface, making them type-compatible and transparently nestable. The decorator holds the wrapped object (an `ICoffee` reference) and calls its `Cost`/`Description` interface methods in `ICoffee_Cost`/`ICoffee_Description`, preserving lazy evaluation.
- Missing syntax: **code reuse mechanism (inheritance or struct embedding)**. Classic Decorator needs a `Decorator : Coffee` abstract base class that holds the internal `Coffee` reference and defaults all methods to transparent forwarding — concrete decorators only `Override` the 1-2 methods they want to change, everything else passes through "for free". Go achieves this with struct embedding (`type MilkDecorator struct { Coffee }`); VB.NET uses `MustInherit CoffeeDecorator Inherits CoffeeBase` base class to write the generic forwarding skeleton once. AxonASP must manually delegate every `ICoffee` method in every concrete decorator (`MilkDecorator`, `SugarDecorator`, `WhipDecorator`, `CaramelDecorator`...) — if the interface adds N methods and there are M decorators, you hand-write N×M forwarding lines.
- Remaining limitation: **parameterized constructors**. AxonASP doesn't support `New MilkDecorator(inner)` syntax, can't inject the wrapped object at creation. Still needs explicit `Init` method injection; and since `Init` is a concrete class method (not an interface method), callers must call `Init` on the concrete type reference before assigning to the interface variable.

### VB.NET Version (syntactically complete baseline)

VB.NET uses `MustInherit CoffeeBase` abstract base class as the unified anchor, then adds a `CoffeeDecorator Inherits CoffeeBase` decorator base class — holds `m_Inner` and defaults to forwarding `Cost`/`Description`. Concrete decorators only `Override` the methods they want to change. Same scenario as Axon version: SimpleCoffee + Milk + Sugar layered.

```vbnet
' ① Abstract base class: common anchor for all coffee (concrete + decorators)
Public MustInherit Class CoffeeBase
    Public MustOverride Function Cost() As Integer
    Public MustOverride Function Description() As String
End Class

' ② Concrete coffee
Public Class SimpleCoffee
    Inherits CoffeeBase

    Public Overrides Function Cost() As Integer
        Return 10
    End Function

    Public Overrides Function Description() As String
        Return "Plain coffee"
    End Function
End Class

' ③ Decorator base class: holds wrapped object, defaults to forwarding all methods
'    Concrete decorators only Override methods they want to change
Public MustInherit Class CoffeeDecorator
    Inherits CoffeeBase

    Protected ReadOnly m_Inner As CoffeeBase

    Protected Sub New(inner As CoffeeBase)
        m_Inner = inner
    End Sub

    Public Overrides Function Cost() As Integer
        Return m_Inner.Cost()
    End Function

    Public Overrides Function Description() As String
        Return m_Inner.Description()
    End Function
End Class

' ④ Milk decorator: only Overrides Cost/Description
Public Class MilkDecorator
    Inherits CoffeeDecorator

    Public Sub New(inner As CoffeeBase)
        MyBase.New(inner)
    End Sub

    Public Overrides Function Cost() As Integer
        Return MyBase.Cost() + 2
    End Function

    Public Overrides Function Description() As String
        Return MyBase.Description() & " + milk"
    End Function
End Class

' ⑤ Sugar decorator: same structure as MilkDecorator
Public Class SugarDecorator
    Inherits CoffeeDecorator

    Public Sub New(inner As CoffeeBase)
        MyBase.New(inner)
    End Sub

    Public Overrides Function Cost() As Integer
        Return MyBase.Cost() + 1
    End Function

    Public Overrides Function Description() As String
        Return MyBase.Description() & " + sugar"
    End Function
End Class

' Demo: layered wrapping, dynamic stacking
Dim coffee As CoffeeBase = New SimpleCoffee()
Console.WriteLine(coffee.Description() & " = " & coffee.Cost() & " yuan")
' Plain coffee = 10 yuan

coffee = New MilkDecorator(coffee)
coffee = New SugarDecorator(coffee)
Console.WriteLine(coffee.Description() & " = " & coffee.Cost() & " yuan")
' Plain coffee + milk + sugar = 13 yuan
```

**VB.NET version notes**:
- **Decorator base class reduces repetition**: `CoffeeDecorator` holds `m_Inner` and writes default forwarding once. Concrete decorators only `Override` methods they want to change. Axon version must manually write delegation for all methods in every decorator.
- **Parameterized constructor + `MyBase.New` replaces Init**: `New MilkDecorator(coffee)` injects wrapped object at creation, no "New then Init" half-initialized window.
- **Type uniformity**: Decorators and concrete coffees are all `CoffeeBase` subclasses, can be transparently nested. Axon version relies on `ICoffee` interface for type compatibility.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Type compatibility | None (decorator and original coffee are different types) | `Implements ICoffee` interface unified | `MustInherit CoffeeBase` base class unified |
| Non-overridden method forwarding | Each decorator hand-writes all delegation | Each decorator hand-writes all delegation | Decorator base class writes once, concrete decorators only Override what they change |
| Wrapped object injection | `Init` two-step (easy to forget) | `Init` two-step (easy to forget) | Parameterized constructor `New(inner)` one-step |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |
---
