## Chapter 23: Interpreter

**Core idea**: Build an interpreter for a language — turn expressions into executable logic.

**Example**: Context holds variable mappings (a=5, b=3). NumberExpression returns its value directly. AddExpression recursively interprets left and right sub-expressions then adds them. Interpreting `a + b` yields 8.

### Classic VBScript Version

```vbscript
' Context: holds variable name → value mappings
Class Context
    Private m_Vars   ' Dictionary

    ' Constructor: create dictionary
    Private Sub Class_Initialize
        Set m_Vars = CreateObject("Scripting.Dictionary")
    End Sub

    ' Set variable value
    Public Function SetVar(name, value)
        m_Vars(name) = value
    End Function

    ' Get variable value
    Public Function GetVar(name)
        GetVar = m_Vars(name)
    End Function
End Class

' Terminal expression: number literal
Class NumberExpression
    Public Value

    ' Interpret: return own value directly
    Public Function Interpret(context)
        Interpret = Value
    End Function
End Class

' Non-terminal expression: addition
Class AddExpression
    Public Left, Right   ' Left and right sub-expressions

    ' Interpret: recursively interpret left and right, then add
    Public Function Interpret(context)
        Interpret = Left.Interpret(context) + Right.Interpret(context)
    End Function
End Class

' Non-terminal expression: variable reference
Class VariableExpression
    Public Name

    ' Interpret: look up variable value from context
    Public Function Interpret(context)
        Interpret = context.GetVar(Name)
    End Function
End Class

' Demo: interpret expression "a + b"
Dim ctx
Set ctx = New Context
ctx.SetVar "a", 5
ctx.SetVar "b", 3

Dim a, b, add
Set a = New VariableExpression
a.Name = "a"
Set b = New VariableExpression
b.Name = "b"
Set add = New AddExpression
Set add.Left = a
Set add.Right = b

Response.Write "a + b = " & add.Interpret(ctx)   ' 8
```

**Classic VBScript trade-offs**:
- **No interface**: NumberExpression, AddExpression, VariableExpression have no `IExpression` interface forcing `Interpret`. If an expression class forgets Interpret, the error only surfaces at runtime.
- **No recursive type safety**: AddExpression's Left and Right have no type constraint — they can point to any object. The error only surfaces when `Interpret` is called at runtime.

### Axon VBScript Version (supports Implements)

```vbscript
' Expression interface
Class IExpression
    Public Function Interpret(context As Context) As Integer
    End Function
End Class

' Context
Class Context
    Private m_Vars As Object

    Private Sub Class_Initialize
        Set m_Vars = CreateObject("Scripting.Dictionary")
    End Sub

    Public Function SetVar(name As String, value As Integer)
        m_Vars(name) = value
    End Function

    Public Function GetVar(name As String) As Integer
        GetVar = m_Vars(name)
    End Function
End Class

' Number expression
Class NumberExpression
    Implements IExpression
    Public Value As Integer

    Public Function IExpression_Interpret(context As Context) As Integer
        IExpression_Interpret = Value
    End Function
End Class

' Variable expression
Class VariableExpression
    Implements IExpression
    Public Name As String

    Public Function IExpression_Interpret(context As Context) As Integer
        IExpression_Interpret = context.GetVar(Name)
    End Function
End Class

' Addition expression
Class AddExpression
    Implements IExpression
    Private m_Left As IExpression
    Private m_Right As IExpression

    Public Function Init(left As IExpression, right As IExpression)
        Set m_Left = left
        Set m_Right = right
    End Function

    Public Function IExpression_Interpret(context As Context) As Integer
        IExpression_Interpret = m_Left.Interpret(context) + m_Right.Interpret(context)
    End Function
End Class

' Demo
Dim ctx As Context
Set ctx = New Context
ctx.SetVar "a", 5
ctx.SetVar "b", 3

Dim aObj As VariableExpression, bObj As VariableExpression
Set aObj = New VariableExpression
aObj.Name = "a"
Set bObj = New VariableExpression
bObj.Name = "b"

Dim addObj As AddExpression
Set addObj = New AddExpression
addObj.Init aObj, bObj

Dim add As IExpression
Set add = addObj
Response.Write "a + b = " & add.Interpret(ctx)
```

**Axon VBScript trade-offs**:
- No inheritance, no abstract base class. `IExpression` constrains the expression class contract. AxonASP's interface method dispatch is fixed — `AddExpression` holds `Private m_Left As IExpression`/`Private m_Right As IExpression` and recursively calls `m_Left.Interpret(context)` in `IExpression_Interpret` to evaluate sub-expressions. No pre-computed results or `SetResults` helper needed — consistent with standard OOP. Remaining limitation: **No inheritance, no abstract base class**. Classic Interpreter uses `MustInherit ExpressionBase` to unify all expression types, with shared utility methods (e.g. `ToString()`, `Clone()`). VB.NET can also combine generics for `Expression(Of T)` generic expression trees + Parser recursive descent parsing. AxonASP currently only has interfaces — no abstract base class as a "type anchor". Adding multiply/divide/subtract operators requires writing separate parallel classes for each expression type, unable to reuse base class logic.

### VB.NET Version (syntactically complete baseline)

VB.NET has `MustInherit` (abstract base class) + `MustOverride` (abstract method) + `Inherits` (inheritance). The Interpreter pattern follows the same structure as the Axon version: `MustInherit ExpressionBase` replaces `IExpression` as the type anchor, `Number`/`Variable`/`Add` concrete expressions correspond one-to-one with the Axon version — no operator overloading, Parser, or generic expression abstractions.

```vbnet
' Context: variable table
Public Class Context
    Private ReadOnly m_Vars As New Dictionary(Of String, Integer)()

    Public Sub SetVar(name As String, value As Integer)
        m_Vars(name) = value
    End Sub

    Public Function GetVar(name As String) As Integer
        Return m_Vars(name)
    End Function
End Class

' Abstract expression base class: MustInherit prevents direct instantiation, MustOverride enforces Interpret
Public MustInherit Class ExpressionBase
    Public MustOverride Function Interpret(ctx As Context) As Integer
End Class

' Number expression (terminal)
Public Class NumberExpression
    Inherits ExpressionBase

    Private ReadOnly m_Value As Integer

    Public Sub New(value As Integer)
        m_Value = value
    End Sub

    Public Overrides Function Interpret(ctx As Context) As Integer
        Return m_Value
    End Function
End Class

' Variable expression (terminal)
Public Class VariableExpression
    Inherits ExpressionBase

    Private ReadOnly m_Name As String

    Public Sub New(name As String)
        m_Name = name
    End Sub

    Public Overrides Function Interpret(ctx As Context) As Integer
        Return ctx.GetVar(m_Name)
    End Function
End Class

' Addition expression (non-terminal, left/right children)
Public Class AddExpression
    Inherits ExpressionBase

    Private ReadOnly m_Left As ExpressionBase
    Private ReadOnly m_Right As ExpressionBase

    Public Sub New(left As ExpressionBase, right As ExpressionBase)
        m_Left = left
        m_Right = right
    End Sub

    Public Overrides Function Interpret(ctx As Context) As Integer
        Return m_Left.Interpret(ctx) + m_Right.Interpret(ctx)
    End Function
End Class

' Demo: interpret expression "a + b"
Dim ctx As New Context()
ctx.SetVar("a", 5)
ctx.SetVar("b", 3)

Dim a As ExpressionBase = New VariableExpression("a")
Dim b As ExpressionBase = New VariableExpression("b")
Dim add As ExpressionBase = New AddExpression(a, b)
Console.WriteLine("a + b = " & add.Interpret(ctx))   ' 8
```

**VB.NET version notes**:
- **`MustInherit ExpressionBase` = unified type anchor for expression trees**: Axon uses `IExpression` to constrain `Interpret`, but interfaces cannot hold shared logic; VB.NET's abstract base class can both `MustOverride` enforce contracts and provide shared utility methods for subclasses. `New ExpressionBase()` is a compile error.
- **Parameterized constructor builds subtrees in one step**: `New AddExpression(a, b)` assembles left/right children in one line; Axon needs `New` then manual `Init aObj, bObj` in two steps.
- **Recursive evaluation is consistent**: `m_Left.Interpret(ctx) + m_Right.Interpret(ctx)` recursively interprets sub-expressions — same logic as Axon, just strongly typed by `ExpressionBase`.
- **No `Set` for object assignment**: `Dim a As ExpressionBase = New VariableExpression("a")` uses uniform `=`; sub-expressions can be declared as base class type directly.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|-----------|-----------------|---------------|--------|
| Expression contract | Method name convention | `IExpression` interface constrains `Interpret` | `MustInherit` + `MustOverride` compile-time enforcement |
| Type anchor | None (isolated classes) | `IExpression` interface (constraint only) | `ExpressionBase` abstract base class (constraint + shared logic) |
| Subtree assembly | Public field `Set add.Left = a` | `Init aObj, bObj` two steps | Parameterized ctor `New AddExpression(a, b)` one step |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |

---
