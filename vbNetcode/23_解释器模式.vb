Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch23Module
    Public Class Context
        Private ReadOnly m_Vars As New Dictionary(Of String, Integer)()

        Public Sub SetVar(name As String, value As Integer)
            m_Vars(name) = value
        End Sub

        Public Function GetVar(name As String) As Integer
            Return m_Vars(name)
        End Function
    End Class
    Public MustInherit Class ExpressionBase
        Public MustOverride Function Interpret(ctx As Context) As Integer
    End Class
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
    Sub Main()

        ' 抽象表达式基类：MustInherit 禁止直接实例化，MustOverride 强制实现 Interpret

        ' 数字表达式（终结符）

        ' 变量表达式（终结符）

        ' 加法表达式（非终结符，左右子节点）

        ' 演示：解释表达式 "a + b"
        Dim ctx As New Context()
        ctx.SetVar("a", 5)
        ctx.SetVar("b", 3)

        Dim a As ExpressionBase = New VariableExpression("a")
        Dim b As ExpressionBase = New VariableExpression("b")
        Dim add As ExpressionBase = New AddExpression(a, b)
        Console.WriteLine("a + b = " & add.Interpret(ctx))   ' 8
    End Sub
End Module
