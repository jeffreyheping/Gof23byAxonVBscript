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
- The interface mechanism lets decorators and components implement the same `ICoffee` interface, making them type-compatible and transparently nestable. The decorator holds the wrapped object (an `ICoffee` reference) and calls its `Cost`/`Description` interface methods in `ICoffee_Cost`/`ICoffee_Description`, preserving lazy evaluation. Remaining gap: **Parameterized constructor**. AxonASP doesn't support `New CoffeeDecorator(inner)` syntax — you can't inject the wrapped object at creation time. You still need an explicit `Init` method. And since `Init` is a concrete class method (not an interface method), the caller must call `Init` on the concrete type reference before assigning to the interface variable.
---
