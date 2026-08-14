Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch22Module
    Public Structure EditorMemento
        Public Content As String
        Public CursorPos As Integer
    End Structure
    Public Class TextEditor
        Private m_Content As String = ""
        Private m_CursorPos As Integer = 0

        Public Sub Write(text As String)
            m_Content &= text

            m_CursorPos = m_Content.Length
        End Sub

        ' 保存当前状态到备忘录（Structure 值拷贝返回）
        Public Function SaveState() As EditorMemento
            Return New EditorMemento With {
                .Content = m_Content,
                .CursorPos = m_CursorPos
            }
        End Function

        ' 从备忘录恢复状态（Structure 值拷贝传入，不修改原快照）
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
    Sub Main()

        ' 编辑器：可以保存和恢复状态

        ' 演示：编辑、保存、再编辑、再恢复
        Dim editor As New TextEditor()
        editor.Write("Hello")

        Dim saved As EditorMemento = editor.SaveState()

        editor.Write(" World")
        Console.WriteLine($"编辑后: {editor.Content} (光标: {editor.CursorPos})")

        editor.RestoreState(saved)
        Console.WriteLine($"恢复后: {editor.Content} (光标: {editor.CursorPos})")
    End Sub
End Module
