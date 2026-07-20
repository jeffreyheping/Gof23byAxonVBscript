Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
' 产品类：狗
Class Dog
    ' 让狗叫一声
    Public Function Speak
        Response.Write "汪汪"
    End Function
End Class

' 产品类：猫
Class Cat
    ' 让猫叫一声
    Public Function Speak
        Response.Write "喵喵"
    End Function
End Class

' 工厂类：根据类型字符串创建对应的动物对象
' animalType: "dog" 或 "cat"
' 返回值: Dog 或 Cat 实例
Class AnimalFactory
    Public Function CreateAnimal(animalType)
        Select Case LCase(animalType)
            Case "dog"
                Set CreateAnimal = New Dog
            Case "cat"
                Set CreateAnimal = New Cat
            Case Else
                Set CreateAnimal = Nothing
        End Select
    End Function
End Class

' 演示：通过工厂创建对象，不直接 New
Dim factory
Set factory = New AnimalFactory
Set myPet = factory.CreateAnimal("dog")
myPet.Speak   ' 汪汪

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
