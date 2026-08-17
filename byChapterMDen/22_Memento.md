## Chapter 22: Memento

**Core idea**: Save an object's internal state so it can be restored later.

**Example**: After editing a document, TextEditor saves the current state to EditorMemento. The editor can then revert to any previous state — implementing undo.

### Classic VBScript Version

```vbscript
' Memento: saves editor state (Property simulates access control)
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

' Editor: can save and restore state
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

    ' Save current state to memento
    Public Function SaveState
        Dim memento
        Set memento = New EditorMemento
        memento.Content = m_Content
        memento.CursorPos = m_CursorPos
        Set SaveState = memento
    End Function

    ' Restore state from memento
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

' Demo: edit, save, edit more, then restore
Dim editor
Set editor = New TextEditor
editor.Write "Hello"

Dim saved
Set saved = editor.SaveState

editor.Write " World"
Response.Write "After edit: " & editor.Content & " (cursor: " & editor.CursorPos & ")" & vbCrLf

editor.RestoreState saved
Response.Write "After restore: " & editor.Content & " (cursor: " & editor.CursorPos & ")" & vbCrLf
```

**Classic VBScript trade-offs**:
- **Limited Property encapsulation**: Although `Property Get/Let` simulates encapsulation, VBScript Properties are inherently Public — external code can still read and write `EditorMemento`'s properties. Ideally, memento internals should only be visible to the Originator (TextEditor), but VBScript has no cross-class access control.
- **No type constraint**: `RestoreState`'s parameter can be any object — no compile-time type checking.

### Axon VBScript Version (supports UDT)

```vbscript
' Memento: saves editor state (UDT value type, automatic copy semantics)
Type EditorMemento
    Content As String
    CursorPos As Integer
End Type

' Editor: can save and restore state
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

    ' Save current state to memento (UDT as return value, value copy)
    Public Function SaveState As EditorMemento
        SaveState.Content = m_Content
        SaveState.CursorPos = m_CursorPos
    End Function

    ' Restore state from memento (UDT as parameter, value copy)
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

' Demo: edit, save, edit more, then restore
Dim editor As TextEditor
Set editor = New TextEditor
editor.Write "Hello"

Dim saved As EditorMemento
saved = editor.SaveState

editor.Write " World"
Response.Write "After edit: " & editor.Content & " (cursor: " & editor.CursorPos & ")" & vbCrLf

editor.RestoreState saved
Response.Write "After restore: " & editor.Content & " (cursor: " & editor.CursorPos & ")" & vbCrLf
```

**Axon VBScript trade-offs**:
- Classic already idiomatic, but VB.NET uses UDT/ReadOnly Struct value semantics. The classic version's `Class EditorMemento` + `Property Get/Let` is already the idiomatic VBScript approach. UDT + strong typing (`As String`/`As Integer`) provides a cleaner alternative: value-copy semantics naturally prevent external modification of snapshots through references, and no Property boilerplate is needed. This is the first of the three "classic already idiomatic" patterns to introduce strong typing, but it's still an expression/quality improvement rather than a fix for non-idiomatic code — so it remains classified as "classic already idiomatic".

### VB.NET Version (syntactically complete baseline)

VB.NET has `Structure` (value type) + `Dictionary(Of TKey, TValue)` generic collections + complete `Property` syntax. The Memento pattern follows the same structure as the Axon version: `Structure EditorMemento` corresponds to Axon's UDT value type, `TextEditor` class has the same save/restore methods — no Caretaker history stack or `Friend` access control or other extra abstractions.

```vbnet
' Memento: Structure value type, assignment = copy (corresponds to Axon UDT)
Public Structure EditorMemento
    Public Content As String
    Public CursorPos As Integer
End Structure

' Editor: can save and restore state
Public Class TextEditor
    Private m_Content As String = ""
    Private m_CursorPos As Integer = 0

    Public Sub Write(text As String)
        m_Content &= text
        m_CursorPos = m_Content.Length
    End Sub

    ' Save current state to memento (Structure value-copy return)
    Public Function SaveState() As EditorMemento
        Return New EditorMemento With {
            .Content = m_Content,
            .CursorPos = m_CursorPos
        }
    End Function

    ' Restore state from memento (Structure value-copy in, doesn't modify original snapshot)
    Public Sub RestoreState(memento As EditorMemento)
        m_Content = memento.Content
        m_CursorPos = memento.CursorPos
    End Sub

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

' Demo: edit, save, edit more, then restore
Dim editor As New TextEditor()
editor.Write("Hello")

Dim saved As EditorMemento = editor.SaveState()

editor.Write(" World")
Console.WriteLine($"After edit: {editor.Content} (cursor: {editor.CursorPos})")

editor.RestoreState(saved)
Console.WriteLine($"After restore: {editor.Content} (cursor: {editor.CursorPos})")
```

**VB.NET version notes**:
- **`Structure` value type corresponds to Axon UDT**: Assignment = value copy, snapshots are naturally independent, external code cannot accidentally modify the original snapshot through references — same semantics as Axon UDT.
- **Object initializer one-liner memento construction**: `New EditorMemento With {.Content = ..., .CursorPos = ...}` replaces Axon's field-by-field `SaveState.Content = ...` assignment.
- **`ReadOnly Property` read-only exposure**: Editor only exposes `Get` for `Content`/`CursorPos` — external code cannot directly modify internal editor state.
- **`&=` operator + no `Set` needed**: `m_Content &= text` concatenates in place; object assignment uses uniform `=`, no `Set`/`Let` distinction.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|-----------|-----------------|---------------|--------|
| Memento carrier | Class (reference type) | UDT (value type) | `Structure` (value type, corresponds to UDT) |
| Snapshot independence | No (reference shared) | Yes (UDT value copy) | Yes (Structure value copy) |
| Memento creation | Field-by-field Property Let | Field-by-field UDT assignment | Object initializer `With {...}` one-liner |
| Object assignment | `Set a = New X` | `saved = editor.SaveState` | Direct `= New X()` |

---
