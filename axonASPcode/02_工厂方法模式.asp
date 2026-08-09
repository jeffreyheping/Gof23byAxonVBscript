<%
Option Explicit
' 产品接口
Class IAnimal
    Public Function Speak
    End Function
End Class

' 工厂接口
Class IFactory
    Public Function CreateAnimal As IAnimal
    End Function
End Class

' 具体产品：狗
Class Dog
    Implements IAnimal
    Public Function IAnimal_Speak
        Response.Write("汪汪")

    End Function
End Class

' 具体产品：猫
Class Cat
    Implements IAnimal
    Public Function IAnimal_Speak
        Response.Write("喵喵")

    End Function
End Class

' 具体工厂：创建狗
Class DogFactory
    Implements IFactory
    Public Function IFactory_CreateAnimal As IAnimal
        Set IFactory_CreateAnimal = New Dog
    End Function
End Class

' 具体工厂：创建猫
Class CatFactory
    Implements IFactory
    Public Function IFactory_CreateAnimal As IAnimal
        Set IFactory_CreateAnimal = New Cat
    End Function
End Class

' 演示：通过工厂接口创建动物，调用方不依赖具体类
Dim factory As IFactory
Dim myPet As IAnimal
Set factory = New DogFactory
Set myPet = factory.CreateAnimal
myPet.Speak()   ' 汪汪
%>