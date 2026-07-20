## 第2章 工厂方法模式（Factory Method）

**核心思想**：把"创建哪种对象"的决定封装到工厂中。

**示例说明**：定义 Dog、Cat 两个产品类，AnimalFactory 根据传入的类型字符串决定创建哪种动物。调用方不需要知道具体类名。

### 传统 VBScript 版

```vbscript
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
```

**传统 VBScript 版妥协说明**：
- **无继承、无抽象方法**：经典工厂方法依赖"抽象 Creator + 子类重写 FactoryMethod"。VBScript 无继承，只能把 `Select Case` 写在工厂类内部，每加一种产品就要改工厂代码——违背了开闭原则。
- **无接口约束**：Dog 和 Cat 没有 `IAnimal` 接口保证它们都有 `Speak` 方法，完全靠开发者自觉。传错类型运行时才报错。

### Axon VBScript 版（支持 Implements）

```vbscript
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
        Response.Write "汪汪"
    End Function
End Class

' 具体产品：猫
Class Cat
    Implements IAnimal
    Public Function IAnimal_Speak
        Response.Write "喵喵"
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
myPet.Speak   ' 汪汪
```

**Axon VBScript 版妥协说明**：
- 此模式在 AxonASP 中实现较为自然，接口机制解决了核心多态问题，无显著妥协。
---