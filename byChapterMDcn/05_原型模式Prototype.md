## 第5章 原型模式（Prototype）

**核心思想**：通过复制现有对象来创建新对象。

**示例说明**：创建一份简历模板，通过 Clone 方法复制出多份。修改副本的技能数组不影响原件，证明是深拷贝。

### 传统 VBScript 版

```vbscript
' 简历类：包含姓名、年龄、技能数组
Class MyResume
    Public Name, Age, Skills

    ' 浅拷贝克隆：创建新 MyResume 并逐字段复制
    ' 数组元素逐值复制，当前示例中 Skills 为字符串数组，修改副本不影响原件
    ' 返回值：新的 MyResume 实例
    Public Function Clone
        Dim copy
        Set copy = New MyResume
        copy.Name = Me.Name
        copy.Age = Me.Age
        Dim ub, i, arr
        ub = UBound(Me.Skills)
        ReDim arr(ub)
        For i = 0 To ub
            arr(i) = Me.Skills(i)
        Next
        copy.Skills = arr
        Set Clone = copy
    End Function
End Class

' 演示：克隆后修改副本，原件不受影响
Dim r1, r2
Set r1 = New MyResume
r1.Name = "张三"
r1.Age = 25
r1.Skills = Array("VBScript", "HTML")

Set r2 = r1.Clone
r2.Name = "李四"
r2.Skills(0) = "JavaScript"

Response.Write(r1.Name & " " & r1.Skills(0)   ' 张三 VBScript)

Response.Write(r2.Name & " " & r2.Skills(0)   ' 李四 JavaScript)

```

**传统 VBScript 版妥协说明**：
- **无内置 Clone**：VBScript 没有 `Clone()` 方法或序列化机制，必须手动逐字段拷贝。字段越多，Clone 方法越冗长，且新增字段容易忘记在 Clone 中添加。
- **无 ICloneable 接口**：无法约定所有类都必须实现 Clone，靠开发者自觉。

### Axon VBScript 版（支持 Implements）

```vba
' 克隆接口
Class ICloneable
    Public Function Clone As ICloneable
    End Function
End Class

' 简历类：实现 ICloneable 接口
Class MyResume
    Implements ICloneable
    Public Name As String
    Public Age As Integer
    Public Skills

    ' 深拷贝克隆：逐字段复制，数组逐元素拷贝
    Public Function ICloneable_Clone As ICloneable
        Dim copy As MyResume
        Set copy = New MyResume
        copy.Name = Me.Name
        copy.Age = Me.Age
        Dim ub As Integer, i As Integer, arr
        ub = UBound(Me.Skills)
        ReDim arr(ub)
        For i = 0 To ub
            arr(i) = Me.Skills(i)
        Next
        copy.Skills = arr
        Set ICloneable_Clone = copy
    End Function
End Class

' 演示：通过接口引用调用 Clone
Dim r1 As MyResume
Dim r2 As ICloneable
Dim r2Copy As MyResume
Set r1 = New MyResume
r1.Name = "张三"
r1.Age = 25
r1.Skills = Array("VBScript", "HTML")
Set r2 = r1.Clone()
Set r2Copy = r2
r2Copy.Name = "李四"
r2Copy.Skills(0) = "JavaScript"

Response.Write(r1.Name & " " & r1.Skills(0)   ' 张三 VBScript)

Response.Write(r2Copy.Name & " " & r2Copy.Skills(0)   ' 李四 JavaScript)

```

**Axon VBScript 版妥协说明**：
- `ICloneable` 接口保证了所有原型类都有 `Clone` 方法，且通过接口引用可直接调用 `Clone()` 自动派发到具体实现。残留限制：**深拷贝仍需手动实现**。接口解决的是契约问题，不是语法糖问题——VBScript 没有内置的深拷贝或序列化机制，每加一个字段、每嵌套一层对象，都要手动补 Clone 逻辑（如果 Skills 是对象数组而非字符串数组，就需要每个对象再各自 Clone 一遍，嵌套越多代码越长）。Go 同样无内置 Clone，Go 的做法是让每个类型自行实现 `Clone()`（与 VBScript 相同），或借助 `encoding/gob` 序列化做通用深拷贝。此处真正的痛点是**缺少自动深拷贝语法糖**，不是继承。

### VB.NET 版（语法完备的对照基准）

VB.NET 实现 `System.ICloneable` 标准接口，手动深拷贝字段。与 Axon 版场景一致，只保留 MyResume 类 + Name/Age/Skills 字段。

```vbnet
' ① 简历类：实现 System.ICloneable 接口，字段与 Axon 版一致
Public Class MyResume
    Implements ICloneable

    Public Property Name As String
    Public Property Age As Integer
    Public Property Skills As String()   ' 与 Axon 版一致，用数组

    ' 标准 ICloneable 接口方法：手动深拷贝
    Public Function Clone() As Object Implements ICloneable.Clone
        Dim copy As New MyResume() With {
            .Name = Me.Name,
            .Age = Me.Age
        }
        ' 数组深拷贝：逐元素复制（与 Axon 版逻辑一致）
        If Me.Skills IsNot Nothing Then
            copy.Skills = New String(Me.Skills.Length - 1) {}
            Array.Copy(Me.Skills, copy.Skills, Me.Skills.Length)
        End If
        Return copy
    End Function
End Class

' 演示：克隆后修改副本，原件不受影响
Dim original As New MyResume() With {
    .Name = "张三",
    .Age = 25,
    .Skills = {"VBScript", "HTML"}
}

Dim clone As MyResume = DirectCast(original.Clone(), MyResume)
clone.Name = "李四"
clone.Skills(0) = "JavaScript"

Console.WriteLine(original.Name & " " & original.Skills(0))   ' 张三 VBScript
Console.WriteLine(clone.Name & " " & clone.Skills(0))         ' 李四 JavaScript
```

**VB.NET 版说明**：
- **标准 `System.ICloneable` 接口**：.NET BCL 自带的通用契约，所有框架类库都认识它。Axon 版需要自己定义 `Class ICloneable` 空壳类。
- **数组深拷贝用 `Array.Copy`**：VB.NET 有标准库的 `Array.Copy` 一行搞定数组复制，Axon 版需要手动 `ReDim` + `For` 循环逐元素赋值。
- **无需 `Set`**：VB.NET 对象赋值直接用 `=`，`Dim clone As Resume = DirectCast(...)` 不需要 `Set`。
- **深拷贝仍需手动**：与 Axon 版一样，每加一个引用类型字段都要在 Clone 里补拷贝逻辑，这是原型模式的固有痛点，VB.NET 也不例外。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 克隆契约 | 无（方法名约定） | 自定义 `ICloneable` 接口 | 标准 `System.ICloneable` 接口 |
| 深拷贝实现 | 手动 `ReDim` + `For` 循环 | 手动 `ReDim` + `For` 循环 | `Array.Copy` 一行复制数组 |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
| 接口调用 | 无接口 | `r1.Clone()` 经接口派发 | `DirectCast(original.Clone(), Resume)` |
---