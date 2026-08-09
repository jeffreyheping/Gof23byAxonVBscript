Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch02Module
    Public MustInherit Class Animal
        Public MustOverride Function Speak() As Object
    End Class
    Public MustInherit Class AnimalFactory
        Public MustOverride Function CreateAnimal() As Animal
    End Class
    Public Class Dog
        Inherits Animal
        Public Overrides Function Speak() As Object
            Console.WriteLine("汪汪")
        End Function
    End Class
    Public Class Cat
        Inherits Animal
        Public Overrides Function Speak() As Object
            Console.WriteLine("喵喵")
        End Function
    End Class
    Public Class DogFactory
        Inherits AnimalFactory
        Public Overrides Function CreateAnimal() As Animal
            Return New Dog()   ' 不需要 Set，VB.NET 对象赋值直接 =
        End Function
    End Class
    Public Class CatFactory
        Inherits AnimalFactory
        Public Overrides Function CreateAnimal() As Animal
            Return New Cat()
        End Function
    End Class
    Sub Main()

        ' ② 抽象工厂基类：定义 CreateAnimal 骨架，不写具体创建逻辑，留给子类

        ' ③ 具体产品


        ' ④ 具体工厂：每个 Factory 子类只负责一种产品，天然符合开闭原则


        ' 演示：调用方只依赖抽象基类，完全不关心具体 Factory / Product
        Dim factory As AnimalFactory = New DogFactory()
        Dim myPet As Animal = factory.CreateAnimal()
        myPet.Speak()   ' 汪汪
    End Sub
End Module
