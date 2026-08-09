## 第10章 组合模式（Composite）

**核心思想**：统一处理单个对象和对象组合（树形结构）。

**示例说明**：Leaf 是叶子节点（员工），Composite 是组合节点（部门），两者都有相同的 Operation 方法。Composite 内部递归调用所有子节点的 Operation，实现"不管叶子还是分支，用同一方法遍历"。

### 传统 VBScript 版

```vbscript
' 叶子节点：树形结构的最末端
Class Leaf
    Public Name
    ' 显示自身信息，indent 控制缩进层级
    Public Function Operation(indent)
        Response.Write(indent & "叶子：" & Name)

    End Function
End Class

' 组合节点：可包含子节点（Leaf 或 Composite）
Class Composite
    Public Name
    Private m_Children()   ' 子节点数组

    Private m_Count        ' 当前子节点数

    ' 构造函数：初始化数组
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Children(10)
    End Sub

    ' 添加子节点（容量不足时自动扩容）
    Public Function Add(child)
        If m_Count >= UBound(m_Children) + 1 Then
            ReDim Preserve m_Children(m_Count * 2)
        End If
        Set m_Children(m_Count) = child
        m_Count = m_Count + 1
    End Function

    ' 显示自身信息，并递归调用所有子节点的 Operation
    Public Function Operation(indent)
        Response.Write(indent & "组合：" & Name)

        Dim i
        For i = 0 To m_Count - 1
            m_Children(i).Operation indent & "  "
        Next
    End Function
End Class

' 演示：构建 总部→分公司→员工 的树形结构
Dim root, branch1, leaf1, leaf2, leaf3
Set root = New Composite
root.Name = "总部"

Set branch1 = New Composite
branch1.Name = "分公司"

Set leaf1 = New Leaf
leaf1.Name = "员工A"
Set leaf2 = New Leaf
leaf2.Name = "员工B"
Set leaf3 = New Leaf
leaf3.Name = "员工C"

branch1.Add(leaf1)

branch1.Add(leaf2)

root.Add(branch1)

root.Add(leaf3)


root.Operation("")   ' 统一遍历整棵树

```

**传统 VBScript 版妥协说明**：
- **无共同基类**：经典组合模式要求 Leaf 和 Composite 继承同一个 `Component` 基类。VBScript 无继承，两者是完全独立的类，仅靠 `Operation` 方法名约定实现"鸭子类型"。编译器无法保证类型安全。
- **缺继承**：如果要给所有节点加 `Parent` 字段、`GetPath()` 方法、`Name` 属性，必须在 Leaf 和 Composite 里各写一份；节点种类越多（再加 ProjectGroup、Team…），公共代码复制的份数越多。
- **无类型安全**：`Add` 方法接收 `child` 参数无类型约束，理论上可以传入任何对象，运行时调用 `Operation` 才报错。
- **手动数组管理**：没有内置集合类，要手动 `ReDim Preserve` 扩容，边界处理易错。

### Axon VBScript 版（支持 Implements）

```vba
' 组件接口
Class IComponent
    Public Function Operation(indent As String)
    End Function
    Public Function Add(child As IComponent)
    End Function
End Class

' 叶子节点
Class Leaf
    Implements IComponent
    Private m_Name As String

    Public Property Get Name As String
        Name = m_Name
    End Property
    Public Property Let Name(v As String)
        m_Name = v
    End Property

    Public Function IComponent_Operation(indent As String)
        Response.Write(indent & "叶子：" & m_Name)

    End Function
    Public Function IComponent_Add(child As IComponent)
        ' 叶子无子节点，空实现
    End Function
End Class

' 组合节点
Class Composite
    Implements IComponent
    Private m_Name As String
    Private m_Children    ' Collection

    Public Property Get Name As String
        Name = m_Name
    End Property
    Public Property Let Name(v As String)
        m_Name = v
    End Property

    Private Sub Class_Initialize
        Set m_Children = Server.CreateObject("Collection")
    End Sub

    Public Function IComponent_Add(child As IComponent)
        m_Children.Add(child)

    End Function

    Public Function IComponent_Operation(indent As String)
        Response.Write(indent & "组合：" & m_Name)

        Dim child As IComponent
        For Each child In m_Children
            child.Operation(indent & "  ")

        Next
    End Function
End Class

' 演示：统一接口遍历树
Dim rootObj As Composite, branch1Obj As Composite
Dim leaf1Obj As Leaf, leaf2Obj As Leaf, leaf3Obj As Leaf
Set rootObj = New Composite
rootObj.Name = "总部"
Set branch1Obj = New Composite
branch1Obj.Name = "分公司"
Set leaf1Obj = New Leaf: leaf1Obj.Name = "员工A"
Set leaf2Obj = New Leaf: leaf2Obj.Name = "员工B"
Set leaf3Obj = New Leaf: leaf3Obj.Name = "员工C"

Dim root As IComponent, branch1 As IComponent
Dim leaf1 As IComponent, leaf2 As IComponent, leaf3 As IComponent
Set root = rootObj
Set branch1 = branch1Obj
Set leaf1 = leaf1Obj
Set leaf2 = leaf2Obj
Set leaf3 = leaf3Obj

root.Add(branch1)

root.Add(leaf3)

branch1.Add(leaf1)

branch1.Add(leaf2)


root.Operation("")

```

**Axon VBScript 版妥协说明**：
- 接口机制统一了叶子和组合节点的契约，`Composite` 通过 `IComponent` 类型的子节点引用直接调用 `child.Operation`，递归遍历整棵树，符合经典组合模式的透明组合语义。子节点存储用内置 `Collection` + `For Each` 迭代，无需手动管理数组。
- 缺失语法点：**代码复用机制（继承）**。`Leaf` 与 `Composite` 无法共享公共 `Component` 基类来复用默认实现——`Name` 属性在两个类里各写了一遍，若再加 `Parent` 字段、`GetPath()` 相对路径方法、`Depth` 深度属性，需要每个实现类手动复制。Go 用 struct embedding 解决此问题（`type Leaf struct { Component }`，嵌入基类自动提升字段方法），AxonASP 只能各自实现。
- `Leaf.IComponent_Add` 只能空实现：透明组合要求接口包含 `Add`，但叶子节点本不应支持添加子节点，无法在类型层面禁止，只能靠运行时空方法体"静默忽略"。

### VB.NET 版（语法完备的对照基准）

VB.NET 用 `MustInherit ComponentBase` 抽象基类统一 Leaf 与 Composite 的公共代码（`Name` 属性写一次，子类共享），Composite 用强类型 `List(Of IComponent)` 持有子节点并递归 `Operation`。场景与 Axon 版一致：总部→分公司→员工 的树形遍历。

```vbnet
' ① 组件接口：定义所有节点必须支持的契约
Public Interface IComponent
    Property Name As String
    Function Add(child As IComponent) As Object
    Function Operation(indent As String) As Object
End Interface

' ② MustInherit 基类：Name 属性写一次，子类共享；默认 Add 空实现（叶子继承）
Public MustInherit Class ComponentBase
    Implements IComponent

    Public Overridable Property Name As String Implements IComponent.Name

    ' 默认空实现：叶子节点继承此默认，组合节点重写
    Public Overridable Function Add(child As IComponent) As Object Implements IComponent.Add
    End Function

    Public MustOverride Function Operation(indent As String) As Object Implements IComponent.Operation
End Class

' ③ 叶子节点：仅重写 Operation，其余继承基类默认
Public Class Leaf
    Inherits ComponentBase

    Public Overrides Function Operation(indent As String) As Object
        Console.WriteLine(indent & "叶子：" & Name)
    End Function
End Class

' ④ 组合节点：用 List(Of IComponent) 管理子节点，递归遍历
Public Class Composite
    Inherits ComponentBase

    Private ReadOnly m_Children As New List(Of IComponent)()

    Public Overrides Function Add(child As IComponent) As Object
        m_Children.Add(child)
    End Function

    Public Overrides Function Operation(indent As String) As Object
        Console.WriteLine(indent & "组合：" & Name)
        For Each child In m_Children
            child.Operation(indent & "  ")
        Next
    End Function
End Class

' 演示：构建 总部→分公司→员工 树，统一接口遍历
Dim root As IComponent = New Composite With {.Name = "总部"}
Dim branch1 As IComponent = New Composite With {.Name = "分公司"}
Dim leaf1 As IComponent = New Leaf With {.Name = "员工A"}
Dim leaf2 As IComponent = New Leaf With {.Name = "员工B"}
Dim leaf3 As IComponent = New Leaf With {.Name = "员工C"}

root.Add(branch1)
root.Add(leaf3)
branch1.Add(leaf1)
branch1.Add(leaf2)

root.Operation("")
' 组合：总部
'   组合：分公司
'     叶子：员工A
'     叶子：员工B
'   叶子：员工C
```

**VB.NET 版说明**：
- **`MustInherit` 基类消除公共代码重复**：`Name` 属性在 `ComponentBase` 写一次，`Leaf` 和 `Composite` 通过 `Inherits` 自动获得。Axon 版 `Name` 在两个类里各写一遍。
- **`List(Of IComponent)` 强类型容器**：编译期保证子节点只能是 `IComponent`，且原生支持 `For Each`。传统版手动 `ReDim` 数组易越界，Axon 版 `Collection` 无类型安全。
- **叶子 Add 空实现**：与 Axon 版一致，叶子节点的 `Add` 继承基类的空实现，静默忽略。`MustOverride Operation` 强制每个节点类型实现自己的遍历逻辑。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 节点契约 | 方法名约定 | `Implements IComponent` 接口约束 | `MustInherit ComponentBase` + `Interface IComponent` |
| Name 属性复用 | Leaf/Composite 各写一份 | Leaf/Composite 各写一份 | 基类写一次，子类继承获得 |
| 子节点容器 | 手动 `ReDim` 数组（易越界） | 内置 `Collection`（无类型安全） | 泛型 `List(Of IComponent)`（编译期类型检查） |
| 叶子 Add 行为 | 方法不存在 | 空实现（继承基类默认） | 空实现（继承基类默认） |
---