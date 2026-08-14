## 第2章 工厂方法模式（Factory Method）

**核心思想**：把"创建哪种对象"的决定封装到工厂中。

**示例说明**：定义 Dog、Cat 两个产品类，AnimalFactory 根据传入的类型字符串决定创建哪种动物。调用方不需要知道具体类名。

### 传统 VBScript 版

```vbscript
' 产品类：狗
Class Dog
    ' 让狗叫一声
    Public Function Speak
        Response.Write("汪汪")

    End Function
End Class

' 产品类：猫
Class Cat
    ' 让猫叫一声
    Public Function Speak
        Response.Write("喵喵")

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
Dim factory, myPet
Set factory = New AnimalFactory
Set myPet = factory.CreateAnimal("dog")
myPet.Speak()   ' 汪汪

```

**传统 VBScript 版妥协说明**：
- **无继承、无抽象方法**：经典工厂方法依赖"抽象 Creator + 子类重写 FactoryMethod"。VBScript 无继承，只能把 `Select Case` 写在工厂类内部，每加一种产品就要改工厂代码——违背了开闭原则。
- **无接口约束**：Dog 和 Cat 没有 `IAnimal` 接口保证它们都有 `Speak` 方法，完全靠开发者自觉。传错类型运行时才报错。

### Axon VBScript 版（支持 Implements）

```vba
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

```

**Axon VBScript 版妥协说明**：
- 此模式在 AxonASP 中实现较为自然，接口机制解决了核心多态问题。残留限制：**缺失代码复用机制（继承）**。经典工厂方法要求"抽象 Creator 基类 + 子类重写 FactoryMethod"，AxonASP 目前只能用两个独立接口（`IFactory`/`IAnimal`）替代，缺少抽象基类来约束和复用 Creator 的公共逻辑。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有 `MustInherit`（抽象类）+ `MustOverride`（抽象方法）+ `Inherits`（继承），可以写出教科书式的工厂方法——抽象 Creator 定义骨架，具体 Factory 子类重写工厂方法。

```vbnet
' ① 抽象产品基类：MustInherit 禁止直接实例化，MustOverride 强制子类实现 Speak
Public MustInherit Class Animal
    Public MustOverride Sub Speak()
End Class

' ② 抽象工厂基类：定义 CreateAnimal 骨架，不写具体创建逻辑，留给子类
Public MustInherit Class AnimalFactory
    Public MustOverride Function CreateAnimal() As Animal
End Class

' ③ 具体产品
Public Class Dog
    Inherits Animal
    Public Overrides Sub Speak()
        Console.WriteLine("汪汪")
    End Sub
End Class

Public Class Cat
    Inherits Animal
    Public Overrides Sub Speak()
        Console.WriteLine("喵喵")
    End Sub
End Class

' ④ 具体工厂：每个 Factory 子类只负责一种产品，天然符合开闭原则
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

' 演示：调用方只依赖抽象基类，完全不关心具体 Factory / Product
Dim factory As AnimalFactory = New DogFactory()
Dim myPet As Animal = factory.CreateAnimal()
myPet.Speak()   ' 汪汪
```

**VB.NET 版说明**：
- **真正的抽象基类**：`MustInherit Class Animal` 禁止 `New Animal()`，`MustOverride Speak` 强制所有子类必须实现——编译期检查，漏写直接报错。对比 Axon 版：`IAnimal` 接口能约束 `Speak` 存在，但无法禁止外部 `New IAnimal`（VBScript 类不能标记为抽象/不可实例化）。
- **继承带来的代码复用**：如果后续要给所有产品加 `Name` 属性、`Born()` 方法，只需在 `Animal` 基类加一次，所有子类自动继承。Axon 版只能手动给 Dog、Cat 各写一份，产品越多重复越多。
- **开闭原则真正落地**：新增产品只需新增一个 `*Product` 类 + 一个 `*Factory` 类，**现有代码零修改**。Axon 版虽然也满足开闭，但因为没有抽象基类作"契约锚点"，所有工厂和产品都是孤立的平行类，缺少架构层面的约束感。
- **无需 `Set` / `Let` 区分**：VB.NET 对象赋值直接用 `=`，不再需要记忆 `Set` 给对象、`Let` 给值类型这一历史包袱。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 产品约束 | 方法名约定（易漏写） | `IAnimal` 接口约束 `Speak` | `MustInherit` + `MustOverride` 编译期强制 |
| 工厂实现 | 单工厂内部 `Select Case`（改一处全改） | 每种产品一个 Factory 类 + `IFactory` 接口 | `MustInherit` 抽象基类 + 子类 `Overrides` 重写 |
| 开闭原则 | 否（新增产品要改 Factory 的 Case 链） | 是（新增产品 = 新增两个类） | 是（继承关系更清晰，基类作契约锚点） |
| 代码复用 | 无 | 无（各产品/工厂平行类，无继承） | 基类字段/方法自动传给所有子类 |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
---