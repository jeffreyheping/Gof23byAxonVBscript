<%
Option Explicit
' 上下文：保存变量名→值的映射
Class Context
    Private m_Vars   ' Dictionary

    ' 构造函数：创建字典
    Private Sub Class_Initialize
        Set m_Vars = CreateObject("Scripting.Dictionary")
    End Sub

    ' 设置变量值
    Public Function SetVar(name, value)
        m_Vars(name) = value
    End Function

    ' 获取变量值
    Public Function GetVar(name)
        GetVar = m_Vars(name)
    End Function
End Class

' 终结符表达式：数字字面量
Class NumberExpression
    Public Value

    ' 解释：直接返回自身数值
    Public Function Interpret(context)
        Interpret = Value
    End Function
End Class

' 非终结符表达式：加法
Class AddExpression
    Public Left, Right   ' 左右子表达式

    ' 解释：递归解释左右子表达式后相加
    Public Function Interpret(context)
        Interpret = Left.Interpret(context) + Right.Interpret(context)
    End Function
End Class

' 非终结符表达式：变量引用
Class VariableExpression
    Public Name

    ' 解释：从上下文中查找变量值
    Public Function Interpret(context)
        Interpret = context.GetVar(Name)
    End Function
End Class

' 演示：解释表达式 "a + b"
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

Response.Write("a + b = " & add.Interpret(ctx))   ' 8
%>