<%
Option Explicit
' 咖啡接口
Class ICoffee
    Public Function Cost As Integer
    End Function
    Public Function Description As String
    End Function
End Class

' 基础咖啡
Class SimpleCoffee
    Implements ICoffee
    Public Function ICoffee_Cost As Integer
        ICoffee_Cost = 10
    End Function
    Public Function ICoffee_Description As String
        ICoffee_Description = "普通咖啡"
    End Function
End Class

' 牛奶装饰器：持有被包裹的 ICoffee 引用，在接口方法中直接委托调用
Class MilkDecorator
    Implements ICoffee
    Private m_Coffee As ICoffee

    ' 注入被装饰的对象
    Public Function Init(coffee As ICoffee)
        Set m_Coffee = coffee
    End Function

    Public Function ICoffee_Cost As Integer
        ICoffee_Cost = m_Coffee.Cost + 2
    End Function
    Public Function ICoffee_Description As String
        ICoffee_Description = m_Coffee.Description & " + 牛奶"
    End Function
End Class

' 糖装饰器：结构同 MilkDecorator
Class SugarDecorator
    Implements ICoffee
    Private m_Coffee As ICoffee

    ' 注入被装饰的对象
    Public Function Init(coffee As ICoffee)
        Set m_Coffee = coffee
    End Function

    Public Function ICoffee_Cost As Integer
        ICoffee_Cost = m_Coffee.Cost + 1
    End Function
    Public Function ICoffee_Description As String
        ICoffee_Description = m_Coffee.Description & " + 糖"
    End Function
End Class

' 演示：接口引用实现透明嵌套，层层包裹
Dim coffee As ICoffee
Set coffee = New SimpleCoffee
Response.Write(coffee.Description & " = " & coffee.Cost & "元")


Dim milk As ICoffee
Dim milkObj As MilkDecorator
Set milkObj = New MilkDecorator
milkObj.Init(coffee)              ' milk 包裹 coffee

Set milk = milkObj

Dim sugar As ICoffee
Dim sugarObj As SugarDecorator
Set sugarObj = New SugarDecorator
sugarObj.Init(milk)               ' sugar 包裹 milk

Set sugar = sugarObj
Response.Write(sugar.Description & " = " & sugar.Cost & "元")
%>