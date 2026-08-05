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
- `IExpression` constrains the expression class contract. AxonASP's interface method dispatch is fixed — `AddExpression` holds `Private m_Left As IExpression`/`Private m_Right As IExpression` and recursively calls `m_Left.Interpret(context)` in `IExpression_Interpret` to evaluate sub-expressions. No pre-computed results or `SetResults` helper needed. Consistent with standard OOP. Remaining notes: the expression tree must be hand-built in code (no lexer/parser), and `Context` depends on the `Scripting.Dictionary` COM object.

---
