<%
Class Singleton
    Private m_Data As String

    Private Sub Class_Initialize
        m_Data = "我是唯一实例"
    End Sub

    Public Property Get Data As String
        Data = m_Data
    End Property

    Public Property Let Data(value As String)
        m_Data = value
    End Property
End Class

' 全局访问点：Static 变量在函数调用间保持值，支持对象引用
Function GetInstance() As Singleton
    Static instance As Singleton

    If instance Is Nothing Then
        Set instance = New Singleton
    End If
    Set GetInstance = instance
End Function

' 演示：保证同一实例
Dim s1 As Singleton, s2 As Singleton
Set s1 = GetInstance()
Set s2 = GetInstance()
s1.Data = "已修改"
Response.Write(s2.Data)   ' 已修改
%>