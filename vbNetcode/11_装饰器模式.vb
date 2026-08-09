Option Strict Off
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch11Module
    Public MustInherit Class CoffeeBase
        Public MustOverride Function Cost() As Integer
        Public MustOverride Function Description() As String
    End Class
    Public Class SimpleCoffee
        Inherits CoffeeBase

        Public Overrides Function Cost() As Integer
            Return 10
        End Function

        Public Overrides Function Description() As String
            Return "普通咖啡"
        End Function
    End Class
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
    Public Class MilkDecorator
        Inherits CoffeeDecorator

        Public Sub New(inner As CoffeeBase)
            MyBase.New(inner)
        End Sub

        Public Overrides Function Cost() As Integer
            Return MyBase.Cost() + 2
        End Function

        Public Overrides Function Description() As String
            Return MyBase.Description() & " + 牛奶"
        End Function
    End Class
    Public Class SugarDecorator
        Inherits CoffeeDecorator

        Public Sub New(inner As CoffeeBase)
            MyBase.New(inner)
        End Sub

        Public Overrides Function Cost() As Integer
            Return MyBase.Cost() + 1
        End Function

        Public Overrides Function Description() As String
            Return MyBase.Description() & " + 糖"
        End Function
    End Class
    Sub Main()

        ' ② 具体咖啡

        ' ③ 装饰器基类：持有被包裹对象，默认透传所有方法
        '    具体装饰器只需 Overrides 想改的方法，其余自动继承透传

        ' ④ 牛奶装饰器：仅 Overrides Cost/Description

        ' ⑤ 糖装饰器：结构同 MilkDecorator

        ' 演示：层层包裹，动态叠加
        Dim coffee As CoffeeBase = New SimpleCoffee()
        Console.WriteLine(coffee.Description() & " = " & coffee.Cost() & "元")
        ' 普通咖啡 = 10元

        coffee = New MilkDecorator(coffee)
        coffee = New SugarDecorator(coffee)
        Console.WriteLine(coffee.Description() & " = " & coffee.Cost() & "元")
        ' 普通咖啡 + 牛奶 + 糖 = 13元
    End Sub
End Module
