Option Explicit
Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
' 基础组件：普通咖啡
Class SimpleCoffee
    ' 返回价格
    Public Function Cost
        Cost = 10
    End Function
    ' 返回描述
    Public Function Description
        Description = "普通咖啡"
    End Function
End Class

' 装饰器：牛奶（包裹一个 coffee 对象，在其基础上加价）
Class MilkDecorator
    Private m_Coffee   ' 被包裹的内部对象

    ' 注入被装饰的对象
    Public Function Init(coffee)
        Set m_Coffee = coffee
    End Function
    ' 价格 = 内部对象价格 + 牛奶加价
    Public Function Cost
        Cost = m_Coffee.Cost + 2
    End Function
    ' 描述 = 内部对象描述 + 牛奶
    Public Function Description
        Description = m_Coffee.Description & " + 牛奶"
    End Function
End Class

' 装饰器：糖（结构同 MilkDecorator）
Class SugarDecorator
    Private m_Coffee

    ' 注入被装饰的对象
    Public Function Init(coffee)
        Set m_Coffee = coffee
    End Function
    ' 价格 = 内部对象价格 + 糖加价
    Public Function Cost
        Cost = m_Coffee.Cost + 1
    End Function
    ' 描述 = 内部对象描述 + 糖
    Public Function Description
        Description = m_Coffee.Description & " + 糖"
    End Function
End Class

' 演示：层层包裹，动态叠加功能
Dim base, milk, sugar
Set base = New SimpleCoffee
Response.Write(base.Description & " = " & base.Cost & "元")


Set milk = New MilkDecorator
milk.Init(base)              ' milk 包裹 base


Set sugar = New SugarDecorator
sugar.Init(milk)             ' sugar 包裹 milk

Response.Write(sugar.Description & " = " & sugar.Cost & "元")

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
