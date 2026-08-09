## 第11章 装饰器模式（Decorator）

**核心思想**：动态地给对象添加新功能，不用修改原类。

**示例说明**：SimpleCoffee 是基础对象，MilkDecorator 和 SugarDecorator 各持有一个 coffee 引用。将 base 传入 milk，再将 milk 传入 sugar，最终 sugar.Cost = base.Cost + 2 + 1，层层叠加。

### 传统 VBScript 版

```vbscript
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

```

**传统 VBScript 版妥协说明**：
- **无法继承基类**：经典装饰器要求 Decorator 继承 Component 并持有 Component 引用，实现"类型兼容"。VBScript 无继承，`MilkDecorator` 和 `SimpleCoffee` 是完全不同的类，无法互相替换，调用方必须知道当前持有的是装饰器还是原对象。
- **缺继承 embedding**：如果咖啡接口再加 `Temperature()`（温度）、`IsHot()`（是否热饮）、`Calories()`（卡路里）等 10 个方法，每个装饰器都要手动把 10 个方法全部委托给 `m_Coffee`；装饰器越多，重复的委托样板代码越多。
- **无透明性**：理想状态下装饰器对调用方透明，但 VBScript 中 `Set milk = New MilkDecorator` 与 `Set coffee = New SimpleCoffee` 类型不同，无法声明统一变量类型。

### Axon VBScript 版（支持 Implements）

```vba
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

```

**Axon VBScript 版妥协说明**：
- 接口机制使装饰器与组件实现同一接口 `ICoffee`，类型兼容、可透明嵌套。装饰器直接持有被包裹对象（`ICoffee` 引用），并在 `ICoffee_Cost`/`ICoffee_Description` 中调用其 `Cost`/`Description` 接口方法，保留了惰性求值特性。
- 缺失语法点：**代码复用机制（继承或 struct embedding）**。经典装饰器需要一个 `Decorator : Coffee` 的抽象基类，持有内部 `Coffee` 引用并把所有方法默认透传给它——具体装饰器只需 `Overrides` 自己想修改的那 1~2 个方法，其余方法"免费"透传。Go 用 struct embedding（`type MilkDecorator struct { Coffee }`）实现；VB.NET 用 `MustInherit CoffeeDecorator Inherits CoffeeBase` 基类写一次通用透传骨架。AxonASP 则必须在每个具体装饰器（`MilkDecorator`、`SugarDecorator`、`WhipDecorator`、`CaramelDecorator`…）里把 ICoffee 的全部方法逐个手动委托一次，接口加 N 个方法、M 个装饰器，就要手写 N×M 个样板转发。
- 残留限制：**带参构造函数**。AxonASP 不支持 `New MilkDecorator(inner)` 语法，无法在创建时注入被包裹对象，仍需通过 `Init` 方法显式注入；且 `Init` 是具体类方法而非接口方法，调用方需在具体类型引用上调用 `Init` 后再赋值给接口变量。

### VB.NET 版（语法完备的对照基准）

VB.NET 用 `MustInherit CoffeeBase` 抽象基类做统一锚点，再写一层 `CoffeeDecorator Inherits CoffeeBase` 装饰器基类——持有 `m_Inner` 并默认透传 `Cost`/`Description`，具体装饰器只需 `Overrides` 自己想修改的方法。场景与 Axon 版一致：SimpleCoffee + Milk + Sugar 层层叠加。

```vbnet
' ① 抽象基类：所有咖啡（具体咖啡 + 装饰器）的共同锚点
Public MustInherit Class CoffeeBase
    Public MustOverride Function Cost() As Integer
    Public MustOverride Function Description() As String
End Class

' ② 具体咖啡
Public Class SimpleCoffee
    Inherits CoffeeBase

    Public Overrides Function Cost() As Integer
        Return 10
    End Function

    Public Overrides Function Description() As String
        Return "普通咖啡"
    End Function
End Class

' ③ 装饰器基类：持有被包裹对象，默认透传所有方法
'    具体装饰器只需 Overrides 想改的方法，其余自动继承透传
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

' ④ 牛奶装饰器：仅 Overrides Cost/Description
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

' ⑤ 糖装饰器：结构同 MilkDecorator
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

' 演示：层层包裹，动态叠加
Dim coffee As CoffeeBase = New SimpleCoffee()
Console.WriteLine(coffee.Description() & " = " & coffee.Cost() & "元")
' 普通咖啡 = 10元

coffee = New MilkDecorator(coffee)
coffee = New SugarDecorator(coffee)
Console.WriteLine(coffee.Description() & " = " & coffee.Cost() & "元")
' 普通咖啡 + 牛奶 + 糖 = 13元
```

**VB.NET 版说明**：
- **装饰器基类减少重复**：`CoffeeDecorator` 持有 `m_Inner` 并写一次默认透传，具体装饰器只需 `Overrides` 想改的方法。Axon 版每个装饰器都要手动写全部方法的委托。
- **带参构造 + `MyBase.New` 替代 Init**：`New MilkDecorator(coffee)` 创建时即注入被包裹对象，不存在"先 New 后 Init"的半初始化窗口。
- **类型统一**：装饰器和具体咖啡都是 `CoffeeBase` 子类，可透明嵌套。Axon 版靠 `ICoffee` 接口实现类型兼容。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 类型兼容 | 无（装饰器与原咖啡类型不同） | `Implements ICoffee` 接口统一 | `MustInherit CoffeeBase` 基类统一 |
| 非重写方法透传 | 每个装饰器手写全部委托 | 每个装饰器手写全部委托 | 装饰器基类写一次，具体装饰器仅 Overrides 需改的 |
| 被包裹对象注入 | `Init` 两步（易忘） | `Init` 两步（易忘） | 带参构造 `New(inner)` 一步到位 |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
---