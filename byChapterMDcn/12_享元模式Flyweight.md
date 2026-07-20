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
        Response.Write "在 (" & x & "," & y & ") 绘制 " & Color & Name
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
Response.Write "3 棵树，实际只有 1 个 TreeType 对象"
```

**传统 VBScript 版妥协说明**：
此模式在 VBScript 中实现较为自然。`Scripting.Dictionary` 恰好提供了享元工厂所需的"按 key 缓存对象"能力，与模式需求契合。唯一限制是 Dictionary 存取对象必须显式使用 `Set`，语法上稍显繁琐。

### Axon VBScript 版

> 此模式在 AxonASP 中的实现与传统 VBScript 完全一致。`Scripting.Dictionary` 做享元对象缓存已是最优解，AxonASP 的现代化扩展对此模式没有额外的改善价值，直接沿用传统版本即可。

---