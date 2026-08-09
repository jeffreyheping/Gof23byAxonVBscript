<%
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
%>