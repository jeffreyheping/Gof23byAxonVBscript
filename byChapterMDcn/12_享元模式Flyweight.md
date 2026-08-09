## 第12章 享元模式（Flyweight）

**核心思想**：共享细粒度对象，减少内存占用。

**示例说明**：森林中有大量树，但树的"型"（名称+颜色）只有少数几种。TreeFactory 用 Dictionary 缓存 TreeType，相同配置只创建一次，多个 Tree 实例共享同一个 TreeType。

### 传统 VBScript 版

```vbscript
' 享元对象：树的固有属性（名称、颜色），可被多棵树共享
Class TreeType
    Public Name, Color

    ' 在指定坐标绘制树
    Public Function Draw(x, y)
        Response.Write("在 (" & x & "," & y & ") 绘制 " & Color & Name)

    End Function
End Class

' 享元工厂：缓存并复用 TreeType 对象
Class TreeFactory
    Private m_Types   ' Dictionary：key→TreeType

    ' 构造函数：创建字典
    Private Sub Class_Initialize
        Set m_Types = CreateObject("Scripting.Dictionary")
    End Sub

    ' 获取或创建 TreeType：相同参数返回同一个对象
    ' name: 树名, color: 颜色
    ' 返回值：共享的 TreeType 实例
    Public Function GetTreeType(name, color)
        Dim key
        key = name & "|" & color
        If Not m_Types.Exists(key) Then
            Dim t
            Set t = New TreeType
            t.Name = name
            t.Color = color
            Set m_Types(key) = t
        End If
        Set GetTreeType = m_Types(key)
    End Function
End Class

' 演示：3 棵树共享同一个 TreeType 对象
Dim factory, oakType, i
Set factory = New TreeFactory
Set oakType = factory.GetTreeType("橡树", "绿色")

For i = 0 To 2
    oakType.Draw i, i * 2
Next
Response.Write("3 棵树，实际只有 1 个 TreeType 对象")

```

**传统 VBScript 版妥协说明**：
- 此模式在 VBScript 中实现较为自然（传统已地道）。`Scripting.Dictionary` 恰好提供了享元工厂所需的"按 key 缓存对象"能力，与模式需求契合。唯一限制是 Dictionary 存取对象必须显式使用 `Set`，语法上稍显繁琐。

### Axon VBScript 版（支持强类型）

```vba
' 享元对象：树的固有属性（名称、颜色），可被多棵树共享
Class TreeType
    Public Name As String
    Public Color As String

    ' 在指定坐标绘制树
    Public Function Draw(x As Long, y As Long)
        Response.Write("在 (" & x & "," & y & ") 绘制 " & Color & Name)

    End Function
End Class

' 享元工厂：缓存并复用 TreeType 对象
Class TreeFactory
    ' 注：Scripting.Dictionary 是 COM 对象，As 不支持标注，保留 Variant
    Private m_Types

    ' 构造函数：创建字典
    Private Sub Class_Initialize
        Set m_Types = CreateObject("Scripting.Dictionary")
    End Sub

    ' 获取或创建 TreeType：相同参数返回同一个对象
    ' name: 树名, color: 颜色
    ' 返回值：共享的 TreeType 实例
    Public Function GetTreeType(name As String, color As String) As TreeType
        Dim key As String
        key = name & "|" & color
        If Not m_Types.Exists(key) Then
            Dim t As TreeType
            Set t = New TreeType
            t.Name = name
            t.Color = color
            Set m_Types(key) = t
        End If
        Set GetTreeType = m_Types(key)
    End Function
End Class

' 演示：3 棵树共享同一个 TreeType 对象
Dim factory As TreeFactory
Dim oakType As TreeType
Dim i As Long
Set factory = New TreeFactory
Set oakType = factory.GetTreeType("橡树", "绿色")

For i = 0 To 2
    oakType.Draw i, i * 2
Next
Response.Write("3 棵树，实际只有 1 个 TreeType 对象")

```

**Axon VBScript 版妥协说明**：
- 传统版的 `Scripting.Dictionary` 做对象缓存仍是最优解（传统已地道）。强类型版本将 `TreeType` 的字段、`GetTreeType` 的参数与返回值、循环变量全部标注类型，使 IDE 能在 `oakType.` 后智能补全成员。唯一无法标注的是 `m_Types`——`Scripting.Dictionary` 是 COM 对象，`As` 不支持，这是 COM 互操作层面的限制，与享元模式本身无关。

### VB.NET 版（语法完备的对照基准）

传统版已地道，VB.NET 用泛型 `Dictionary(Of String, TreeType)` 替代 COM `Scripting.Dictionary`，编译期类型安全，存取无需 `Set`。场景与 Axon 版一致：3 棵树共享同一个 TreeType。

```vbnet
' ① 享元对象：树的"型"（名称、颜色），可被多棵树共享
Public Class TreeType
    Public ReadOnly Property Name As String
    Public ReadOnly Property Color As String

    Public Sub New(name As String, color As String)
        Me.Name = name
        Me.Color = color
    End Sub

    ' 在指定坐标绘制树（x, y 是外部状态，由调用方传入）
    Public Function Draw(x As Long, y As Long)
        Console.WriteLine($"在 ({x},{y}) 绘制 {Color}{Name}")
    End Function
End Class

' ② 享元工厂：缓存并复用 TreeType
Public Class TreeFactory
    ' 泛型字典：Key 一定是 String，Value 一定是 TreeType
    Private ReadOnly m_Types As New Dictionary(Of String, TreeType)()

    ' 获取或创建 TreeType：相同参数返回同一个对象
    Public Function GetTreeType(name As String, color As String) As TreeType
        Dim key = name & "|" & color
        If Not m_Types.ContainsKey(key) Then
            m_Types(key) = New TreeType(name, color)
        End If
        Return m_Types(key)
    End Function
End Class

' 演示：3 棵树共享同一个 TreeType 对象
Dim factory As New TreeFactory()
Dim oakType = factory.GetTreeType("橡树", "绿色")

For i = 0 To 2
    oakType.Draw(i, i * 2)
Next
Console.WriteLine("3 棵树，实际只有 1 个 TreeType 对象")
' 在 (0,0) 绘制 绿色橡树
' 在 (1,2) 绘制 绿色橡树
' 在 (2,4) 绘制 绿色橡树
```

**VB.NET 版说明**：
- **泛型 `Dictionary(Of String, TreeType)` 编译期类型安全**：Key 和 Value 类型在尖括号里锁定，`m_Types(key)` 取出来就是强类型 `TreeType`，无需 `CType` 转换。COM `Scripting.Dictionary` 的 Key/Value 都是 `Object`，转型错误到运行时才暴露。
- **`ReadOnly` 属性 + 构造函数赋值**：`TreeType` 创建后 Name/Color 不可变，避免共享对象被意外修改。Axon/传统版 `Public Name` 是可变字段。
- **无需 `Set`**：对象赋值直接用 `=`，不再需要 `Set m_Types(key) = t` 的繁琐语法。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 缓存字典类型 | COM `Scripting.Dictionary`（全 Object） | COM `Scripting.Dictionary`（无法标注类型） | 泛型 `Dictionary(Of String, TreeType)`（编译期锁定类型） |
| 类型安全 | 无，存取全 Variant 转型 | 外围强类型，字典内部仍 Object | 全链路强类型（存取无需转型） |
| 享元对象可变性 | `Public Name` 可变字段 | 同左（加 `As String`） | `ReadOnly` 属性 + 构造函数赋值，不可变 |
| 对象赋值 | `Set m_Types(key) = t` | 同左（COM 限制） | 直接 `m_Types(key) = New TreeType(...)` |
---