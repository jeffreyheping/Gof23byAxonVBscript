## 第22章 备忘录模式（Memento）

**核心思想**：保存对象的内部状态，以便之后恢复。

**示例说明**：TextEditor 编辑文档后，EditorMemento 保存当前内容。编辑器可以恢复到之前的任意状态，实现撤销功能。

### 传统 VBScript 版

```vbscript
' 备忘录：保存编辑器状态（通过 Property 封装，模拟访问控制）
Class EditorMemento
    Private m_Content
    Private m_CursorPos

    Public Property Get Content
        Content = m_Content
    End Property
    Public Property Let Content(v)
        m_Content = v
    End Property

    Public Property Get CursorPos
        CursorPos = m_CursorPos
    End Property
    Public Property Let CursorPos(v)
        m_CursorPos = v
    End Property
End Class

' 编辑器：可以保存和恢复状态
Class TextEditor
    Private m_Content
    Private m_CursorPos

    Private Sub Class_Initialize
        m_Content = ""
        m_CursorPos = 0
    End Sub

    Public Function Write(text)
        m_Content = m_Content & text
        m_CursorPos = Len(m_Content)
    End Function

    ' 保存当前状态到备忘录
    Public Function SaveState
        Dim memento
        Set memento = New EditorMemento
        memento.Content = m_Content
        memento.CursorPos = m_CursorPos
        Set SaveState = memento
    End Function

    ' 从备忘录恢复状态
    Public Function RestoreState(memento)
        m_Content = memento.Content
        m_CursorPos = memento.CursorPos
    End Function

    Public Property Get Content
        Content = m_Content
    End Property

    Public Property Get CursorPos
        CursorPos = m_CursorPos
    End Property
End Class

' 演示：编辑、保存、再编辑、再恢复
Dim editor
Set editor = New TextEditor
editor.Write("Hello")


Dim saved
Set saved = editor.SaveState

editor.Write(" World")

Response.Write("编辑后: " & editor.Content & " (光标: " & editor.CursorPos & ")" & vbCrLf)


editor.RestoreState(saved)

Response.Write("恢复后: " & editor.Content & " (光标: " & editor.CursorPos & ")" & vbCrLf)

```

**传统 VBScript 版妥协说明**：
- **Property 封装有限**：虽然用 `Property Get/Let` 模拟了封装，但 VBScript 的 Property 本身是 Public 的，外部代码仍然可以读写 `EditorMemento` 的属性。备忘录理想状态下内部数据应仅对 Originator（TextEditor）可见，但 VBScript 无类间访问控制机制。
- **无类型约束**：`RestoreState` 的参数可以是任何对象，没有编译期类型检查。

### Axon VBScript 版（支持 UDT）

```vba
' 备忘录：保存编辑器状态（UDT 值类型，自动拷贝语义）
Type EditorMemento

    Content As String

    CursorPos As Integer

End Type

' 编辑器：可以保存和恢复状态
Class TextEditor
    Private m_Content As String
    Private m_CursorPos As Integer

    Private Sub Class_Initialize
        m_Content = ""
        m_CursorPos = 0
    End Sub

    Public Function Write(text As String)
        m_Content = m_Content & text
        m_CursorPos = Len(m_Content)
    End Function

    ' 保存当前状态到备忘录（UDT 作为返回值，值拷贝）
    Public Function SaveState As EditorMemento
        SaveState.Content = m_Content
        SaveState.CursorPos = m_CursorPos
    End Function

    ' 从备忘录恢复状态（UDT 作为参数，值拷贝）
    Public Function RestoreState(memento As EditorMemento)
        m_Content = memento.Content
        m_CursorPos = memento.CursorPos
    End Function

    Public Property Get Content As String
        Content = m_Content
    End Property

    Public Property Get CursorPos As Integer
        CursorPos = m_CursorPos
    End Property
End Class

' 演示：编辑、保存、再编辑、再恢复
Dim editor As TextEditor
Set editor = New TextEditor
editor.Write("Hello")


Dim saved As EditorMemento
saved = editor.SaveState

editor.Write(" World")

Response.Write("编辑后: " & editor.Content & " (光标: " & editor.CursorPos & ")" & vbCrLf)


editor.RestoreState(saved)

Response.Write("恢复后: " & editor.Content & " (光标: " & editor.CursorPos & ")" & vbCrLf)

```

**Axon VBScript 版妥协说明**：
- 传统已地道但 VB.NET 用 UDT/ReadOnly Struct 值语义。传统版的 `Class EditorMemento` + `Property Get/Let` 已是 VBScript 中地道的实现方式。UDT + 强类型（`As String`/`As Integer`）提供了更优雅的写法：值拷贝语义天然防止外部通过引用意外修改快照，且无需 Property 样板代码。这是三个"传统已地道"模式中最早引入强类型的，但同样是表达/质量层面的提升，而非对不地道代码的修复——故仍归为"传统已地道"。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有 `Structure`（值类型）+ `Dictionary(Of TKey, TValue)` 泛型集合 + 完整 `Property` 语法，备忘录模式与 Axon 版同结构：`Structure EditorMemento` 对应 Axon 的 UDT 值类型，`TextEditor` 类同样的保存/恢复方法，不引入 Caretaker 历史栈或 `Friend` 访问控制等额外抽象。

```vbnet
' 备忘录：Structure 值类型，赋值即拷贝（对应 Axon 版 UDT）
Public Structure EditorMemento
    Public Content As String
    Public CursorPos As Integer
End Structure

' 编辑器：可以保存和恢复状态
Public Class TextEditor
    Private m_Content As String = ""
    Private m_CursorPos As Integer = 0

    Public Function Write(text As String) As Object
        m_Content &= text

        m_CursorPos = m_Content.Length
    End Function

    ' 保存当前状态到备忘录（Structure 值拷贝返回）
    Public Function SaveState() As EditorMemento
        Return New EditorMemento With {
            .Content = m_Content,
            .CursorPos = m_CursorPos
        }
    End Function

    ' 从备忘录恢复状态（Structure 值拷贝传入，不修改原快照）
    Public Function RestoreState(memento As EditorMemento) As Object
        m_Content = memento.Content
        m_CursorPos = memento.CursorPos
    End Function

    Public ReadOnly Property Content As String
        Get
            Return m_Content
        End Get
    End Property

    Public ReadOnly Property CursorPos As Integer
        Get
            Return m_CursorPos
        End Get
    End Property
End Class

' 演示：编辑、保存、再编辑、再恢复
Dim editor As New TextEditor()
editor.Write("Hello")

Dim saved As EditorMemento = editor.SaveState()

editor.Write(" World")
Console.WriteLine($"编辑后: {editor.Content} (光标: {editor.CursorPos})")

editor.RestoreState(saved)
Console.WriteLine($"恢复后: {editor.Content} (光标: {editor.CursorPos})")
```

**VB.NET 版说明**：
- **`Structure` 值类型对应 Axon 版 UDT**：赋值即值拷贝，快照天然独立，外部无法通过引用意外修改原快照——与 Axon 版 UDT 语义一致。
- **对象初始化器一行构造备忘录**：`New EditorMemento With {.Content = ..., .CursorPos = ...}` 替代 Axon 版逐字段 `SaveState.Content = ...` 赋值。
- **`ReadOnly Property` 只读暴露**：编辑器对外只暴露 `Content`/`CursorPos` 的 `Get`，外部无法直接改写编辑器内部状态。
- **`&=` 运算符 + 无需 `Set`**：`m_Content &= text` 原地拼接，对象赋值统一 `=`，无需 `Set`/`Let` 区分。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 备忘录载体 | Class（引用类型） | UDT（值类型） | `Structure`（值类型，对应 UDT） |
| 快照独立性 | 否（引用共享） | 是（UDT 值拷贝） | 是（Structure 值拷贝） |
| 创建备忘录 | 逐字段 Property Let | 逐字段 UDT 赋值 | 对象初始化器 `With {...}` 一行 |
| 对象赋值 | `Set a = New X` | `saved = editor.SaveState` | 直接 `= New X()` |
---