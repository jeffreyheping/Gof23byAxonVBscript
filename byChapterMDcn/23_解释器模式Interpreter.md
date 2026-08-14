## 第23章 解释器模式（Interpreter）

**核心思想**：为语言创建解释器，把表达式转成执行逻辑。

**示例说明**：Context 保存变量映射（a=5, b=3）。NumberExpression 直接返回值，AddExpression 递归解释左右子表达式后相加。最终解释 `a + b` 得到 8。

### 传统 VBScript 版

```vbscript
' 上下文：保存变量名→值的映射
Class Context
    Private m_Vars   ' Dictionary

    ' 构造函数：创建字典
    Private Sub Class_Initialize
        Set m_Vars = CreateObject("Scripting.Dictionary")
    End Sub

    ' 设置变量值
    Public Function SetVar(name, value)
        m_Vars(name) = value
    End Function

    ' 获取变量值
    Public Function GetVar(name)
        GetVar = m_Vars(name)
    End Function
End Class

' 终结符表达式：数字字面量
Class NumberExpression
    Public Value

    ' 解释：直接返回自身数值
    Public Function Interpret(context)
        Interpret = Value
    End Function
End Class

' 非终结符表达式：加法
Class AddExpression
    Public Left, Right   ' 左右子表达式

    ' 解释：递归解释左右子表达式后相加
    Public Function Interpret(context)
        Interpret = Left.Interpret(context) + Right.Interpret(context)
    End Function
End Class

' 非终结符表达式：变量引用
Class VariableExpression
    Public Name

    ' 解释：从上下文中查找变量值
    Public Function Interpret(context)
        Interpret = context.GetVar(Name)
    End Function
End Class

' 演示：解释表达式 "a + b"
Dim ctx
Set ctx = New Context
ctx.SetVar "a", 5
ctx.SetVar "b", 3

Dim a, b, add
Set a = New VariableExpression
a.Name = "a"
Set b = New VariableExpression
b.Name = "b"
Set add = New AddExpression
Set add.Left = a
Set add.Right = b

Response.Write("a + b = " & add.Interpret(ctx))   ' 8

```

**传统 VBScript 版妥协说明**：
- **无接口**：NumberExpression、AddExpression、VariableExpression 没有 `IExpression` 接口强制 `Interpret` 方法。如果某个表达式类漏写 Interpret，运行时调用才报错。
- **无递归类型安全**：AddExpression 的 Left 和 Right 没有类型约束，可以指向任何对象，运行时调用 Interpret 才报错。

### Axon VBScript 版（支持 Implements）

```vba
' 表达式接口
Class IExpression
    Public Function Interpret(context As Context) As Integer
    End Function
End Class

' 上下文
Class Context
    Private m_Vars As Object

    Private Sub Class_Initialize
        Set m_Vars = CreateObject("Scripting.Dictionary")
    End Sub

    Public Function SetVar(name As String, value As Integer)
        m_Vars(name) = value
    End Function

    Public Function GetVar(name As String) As Integer
        GetVar = m_Vars(name)
    End Function
End Class

' 数字表达式
Class NumberExpression
    Implements IExpression
    Public Value As Integer

    Public Function IExpression_Interpret(context As Context) As Integer
        IExpression_Interpret = Value
    End Function
End Class

' 变量表达式
Class VariableExpression
    Implements IExpression
    Public Name As String

    Public Function IExpression_Interpret(context As Context) As Integer
        IExpression_Interpret = context.GetVar(Name)
    End Function
End Class

' 加法表达式
Class AddExpression
    Implements IExpression
    Private m_Left As IExpression
    Private m_Right As IExpression

    Public Function Init(left As IExpression, right As IExpression)
        Set m_Left = left
        Set m_Right = right
    End Function

    Public Function IExpression_Interpret(context As Context) As Integer
        IExpression_Interpret = m_Left.Interpret(context) + m_Right.Interpret(context)
    End Function
End Class

' 演示
Dim ctx As Context
Set ctx = New Context
ctx.SetVar "a", 5
ctx.SetVar "b", 3

Dim aObj As VariableExpression, bObj As VariableExpression
Set aObj = New VariableExpression
aObj.Name = "a"
Set bObj = New VariableExpression
bObj.Name = "b"

Dim addObj As AddExpression
Set addObj = New AddExpression
addObj.Init aObj, bObj

Dim add As IExpression
Set add = addObj
Response.Write("a + b = " & add.Interpret(ctx))

```

**Axon VBScript 版妥协说明**：
- 缺继承无抽象基类。`IExpression` 接口约束了表达式类的契约。AxonASP 接口方法派发已修复，`AddExpression` 持有 `Private m_Left As IExpression`/`Private m_Right As IExpression`，可在 `IExpression_Interpret` 中递归调用 `m_Left.Interpret(context)` 求值子表达式，无需预计算结果或 `SetResults` 辅助方法，写法与标准 OOP 一致。剩余限制：**缺继承，无抽象基类**。经典解释器模式用 `MustInherit ExpressionBase` 抽象基类统一所有表达式类型，并可在基类内置公共工具方法（如 `ToString()`、`Clone()`），VB.NET 还能配合泛型写 `Expression(Of T)` 通用表达式树 + Parser 递归下降解析。AxonASP 目前只有接口，没有抽象基类作"类型锚点"，新增乘/除/减等运算符需要为每类表达式独立写平行类，无法复用基类公共逻辑。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有 `MustInherit`（抽象基类）+ `MustOverride`（抽象方法）+ `Inherits`（继承），解释器模式与 Axon 版同结构：`MustInherit ExpressionBase` 替代 `IExpression` 接口作类型锚点，`Number`/`Variable`/`Add` 三类具体表达式与 Axon 版一一对应，不引入运算符重载、Parser 解析器或泛型表达式等额外抽象。

```vbnet
' 上下文：变量表
Public Class Context
    Private ReadOnly m_Vars As New Dictionary(Of String, Integer)()

    Public Sub SetVar(name As String, value As Integer)
        m_Vars(name) = value
    End Sub

    Public Function GetVar(name As String) As Integer
        Return m_Vars(name)
    End Function
End Class

' 抽象表达式基类：MustInherit 禁止直接实例化，MustOverride 强制实现 Interpret
Public MustInherit Class ExpressionBase
    Public MustOverride Function Interpret(ctx As Context) As Integer
End Class

' 数字表达式（终结符）
Public Class NumberExpression
    Inherits ExpressionBase

    Private ReadOnly m_Value As Integer

    Public Sub New(value As Integer)
        m_Value = value
    End Sub

    Public Overrides Function Interpret(ctx As Context) As Integer
        Return m_Value
    End Function
End Class

' 变量表达式（终结符）
Public Class VariableExpression
    Inherits ExpressionBase

    Private ReadOnly m_Name As String

    Public Sub New(name As String)
        m_Name = name
    End Sub

    Public Overrides Function Interpret(ctx As Context) As Integer
        Return ctx.GetVar(m_Name)
    End Function
End Class

' 加法表达式（非终结符，左右子节点）
Public Class AddExpression
    Inherits ExpressionBase

    Private ReadOnly m_Left As ExpressionBase
    Private ReadOnly m_Right As ExpressionBase

    Public Sub New(left As ExpressionBase, right As ExpressionBase)
        m_Left = left
        m_Right = right
    End Sub

    Public Overrides Function Interpret(ctx As Context) As Integer
        Return m_Left.Interpret(ctx) + m_Right.Interpret(ctx)
    End Function
End Class

' 演示：解释表达式 "a + b"
Dim ctx As New Context()
ctx.SetVar("a", 5)
ctx.SetVar("b", 3)

Dim a As ExpressionBase = New VariableExpression("a")
Dim b As ExpressionBase = New VariableExpression("b")
Dim add As ExpressionBase = New AddExpression(a, b)
Console.WriteLine("a + b = " & add.Interpret(ctx))   ' 8
```

**VB.NET 版说明**：
- **`MustInherit ExpressionBase` = 表达式树统一类型锚点**：Axon 版用 `IExpression` 接口约束 `Interpret`，但接口无法挂公共逻辑；VB.NET 的抽象基类既能 `MustOverride` 强制契约，又能在基类内置公共方法供子类复用，`New ExpressionBase()` 直接编译报错。
- **带参构造函数一步构建子树**：`New AddExpression(a, b)` 一行装配左右子节点；Axon 版需先 `New` 再手动 `Init aObj, bObj` 两步走。
- **递归求值写法一致**：`m_Left.Interpret(ctx) + m_Right.Interpret(ctx)` 递归解释子表达式，与 Axon 版逻辑相同，仅类型由 `ExpressionBase` 强约束。
- **无需 `Set` 区分对象赋值**：`Dim a As ExpressionBase = New VariableExpression("a")` 统一用 `=`，子表达式可直接声明为基类类型。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 表达式契约 | 方法名约定 | `IExpression` 接口约束 `Interpret` | `MustInherit` + `MustOverride` 编译期强制 |
| 类型锚点 | 无（孤立类） | `IExpression` 接口（仅约束） | `ExpressionBase` 抽象基类（约束 + 可挂公共逻辑） |
| 子树装配 | 公共字段 `Set add.Left = a` | `Init aObj, bObj` 两步 | 带参构造 `New AddExpression(a, b)` 一步 |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
---