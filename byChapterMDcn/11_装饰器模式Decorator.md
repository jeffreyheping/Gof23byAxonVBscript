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
Response.Write base.Description & " = " & base.Cost & "元"

Set milk = New MilkDecorator
milk.Init base              ' milk 包裹 base

Set sugar = New SugarDecorator
sugar.Init milk             ' sugar 包裹 milk
Response.Write sugar.Description & " = " & sugar.Cost & "元"
```

**传统 VBScript 版妥协说明**：
- **无法继承基类**：经典装饰器要求 Decorator 继承 Component 并持有 Component 引用，实现"类型兼容"。VBScript 无继承，`MilkDecorator` 和 `SimpleCoffee` 是完全不同的类，无法互相替换，调用方必须知道当前持有的是装饰器还是原对象。
- **无透明性**：理想状态下装饰器对调用方透明，但 VBScript 中 `Set milk = New MilkDecorator` 与 `Set coffee = New SimpleCoffee` 类型不同，无法声明统一变量类型。

### Axon VBScript 版（支持 Implements）

```vbscript
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
Response.Write coffee.Description & " = " & coffee.Cost & "元"

Dim milk As ICoffee
Dim milkObj As MilkDecorator
Set milkObj = New MilkDecorator
milkObj.Init coffee              ' milk 包裹 coffee
Set milk = milkObj

Dim sugar As ICoffee
Dim sugarObj As SugarDecorator
Set sugarObj = New SugarDecorator
sugarObj.Init milk               ' sugar 包裹 milk
Set sugar = sugarObj
Response.Write sugar.Description & " = " & sugar.Cost & "元"
```

**Axon VBScript 版妥协说明**：
- 接口机制使装饰器与组件实现同一接口 `ICoffee`，类型兼容、可透明嵌套。装饰器直接持有被包裹对象（`ICoffee` 引用），并在 `ICoffee_Cost`/`ICoffee_Description` 中调用其 `Cost`/`Description` 接口方法，保留了惰性求值特性。残留限制：缺失语法点：**带参构造函数**。AxonASP 不支持 `New CoffeeDecorator(inner)` 语法，无法在创建时注入被包裹对象，仍需通过 `Init` 方法显式注入；且 `Init` 是具体类方法而非接口方法，调用方需在具体类型引用上调用 `Init` 后再赋值给接口变量。
---