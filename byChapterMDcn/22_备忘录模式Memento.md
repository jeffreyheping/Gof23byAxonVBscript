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
editor.Write "Hello"

Dim saved
Set saved = editor.SaveState

editor.Write " World"
Response.Write "编辑后: " & editor.Content & " (光标: " & editor.CursorPos & ")" & vbCrLf

editor.RestoreState saved
Response.Write "恢复后: " & editor.Content & " (光标: " & editor.CursorPos & ")" & vbCrLf
```

**传统 VBScript 版妥协说明**：
- **Property 封装有限**：虽然用 `Property Get/Let` 模拟了封装，但 VBScript 的 Property 本身是 Public 的，外部代码仍然可以读写 `EditorMemento` 的属性。备忘录理想状态下内部数据应仅对 Originator（TextEditor）可见，但 VBScript 无类间访问控制机制。
- **无类型约束**：`RestoreState` 的参数可以是任何对象，没有编译期类型检查。

### Axon VBScript 版（支持 UDT）

```vbscript
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
editor.Write "Hello"

Dim saved As EditorMemento
saved = editor.SaveState

editor.Write " World"
Response.Write "编辑后: " & editor.Content & " (光标: " & editor.CursorPos & ")" & vbCrLf

editor.RestoreState saved
Response.Write "恢复后: " & editor.Content & " (光标: " & editor.CursorPos & ")" & vbCrLf
```

**Axon VBScript 版妥协说明**：
- 传统版的 `Class EditorMemento` + `Property Get/Let` 已是 VBScript 中地道的实现方式。UDT 提供了一种更优雅的替代写法：值拷贝语义天然防止外部通过引用意外修改快照，且无需 Property 样板代码。这是表达层面的提升，而非对不地道代码的修复。

---