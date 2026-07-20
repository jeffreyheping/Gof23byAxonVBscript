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

Response.Write "a + b = " & add.Interpret(ctx)   ' 8
```

**传统 VBScript 版妥协说明**：
- **无接口**：NumberExpression、AddExpression、VariableExpression 没有 `IExpression` 接口强制 `Interpret` 方法。如果某个表达式类漏写 Interpret，运行时调用才报错。
- **无递归类型安全**：AddExpression 的 Left 和 Right 没有类型约束，可以指向任何对象，运行时调用 Interpret 才报错。

### Axon VBScript 版（支持 Implements）

```vbscript
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
Response.Write "a + b = " & add.Interpret(ctx)
```

**Axon VBScript 版妥协说明**：
- `IExpression` 接口约束了表达式类的契约。AxonASP 接口方法派发已修复，`AddExpression` 持有 `Private m_Left As IExpression`/`Private m_Right As IExpression`，可在 `IExpression_Interpret` 中递归调用 `m_Left.Interpret(context)` 求值子表达式，无需预计算结果或 `SetResults` 辅助方法，写法与标准 OOP 一致。剩余限制：表达式树需在代码中手工构建（无词法/语法分析器），`Context` 依赖 `Scripting.Dictionary` COM 对象。

---