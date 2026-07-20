Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
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

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
