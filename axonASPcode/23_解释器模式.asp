<%
' 表达式接口
Class IExpression
    Public Function Interpret(context As Context) As Integer
    End Function
End Class

' 上下文
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

' 数字表达式
Class NumberExpression
    Implements IExpression
    Public Value As Integer

    Public Function IExpression_Interpret(context As Context) As Integer
        IExpression_Interpret = Value
    End Function
End Class

' 变量表达式
Class VariableExpression
    Implements IExpression
    Public Name As String

    Public Function IExpression_Interpret(context As Context) As Integer
        IExpression_Interpret = context.GetVar(Name)
    End Function
End Class

' 加法表达式
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

' 演示
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
%>