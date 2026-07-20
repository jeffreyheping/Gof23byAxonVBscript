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

Response.Write r1.Name & " " & r1.Skills(0)   ' 张三 VBScript
Response.Write r2.Name & " " & r2.Skills(0)   ' 李四 JavaScript
```

**传统 VBScript 版妥协说明**：
- **无内置 Clone**：VBScript 没有 `Clone()` 方法或序列化机制，必须手动逐字段拷贝。字段越多，Clone 方法越冗长，且新增字段容易忘记在 Clone 中添加。
- **无 ICloneable 接口**：无法约定所有类都必须实现 Clone，靠开发者自觉。

### Axon VBScript 版（支持 Implements）

```vbscript
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

Response.Write r1.Name & " " & r1.Skills(0)   ' 张三 VBScript
Response.Write r2Copy.Name & " " & r2Copy.Skills(0)   ' 李四 JavaScript
```

**Axon VBScript 版妥协说明**：
- `ICloneable` 接口保证了所有原型类都有 `Clone` 方法，且通过接口引用可直接调用 `Clone()` 自动派发到具体实现。但仍需手动逐字段拷贝——接口解决的是契约问题，不是语法糖问题，VBScript 仍无内置的深拷贝或序列化机制。
- 缺失语法点：**内置深拷贝机制**。Go 同样无继承，但 Go 也无内置 Clone——Go 的做法是让每个类型自行实现 `Clone()` 方法（与 VBScript 相同），或借助序列化（`encoding/gob`）做深拷贝。此处"无继承"不是真正的痛点，真正的痛点是缺少自动深拷贝语法糖。
---