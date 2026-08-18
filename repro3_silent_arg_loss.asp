<%
Option Explicit
' 最小复现：数组元素作调用方 + 无括号调用，实参静默丢失
Class Obs
    Public Name
    Public Function Update(news)
        Response.Write(Name & " 收到：[" & news & "] TypeName=" & TypeName(news))
    End Function
End Class

Class Pub
    Private m_Obs()
    Private m_Count
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Obs(10)
    End Sub
    Public Function Add(o)
        Set m_Obs(m_Count) = o
        m_Count = m_Count + 1
    End Function
    Public Function Notify(news)
        Dim i
        For i = 0 To m_Count - 1
            m_Obs(i).Update news          ' ← 无括号：AxonASP 实参丢失
        Next
    End Function
End Class

Dim p, a
Set p = New Pub
Set a = New Obs
a.Name = "A"
p.Add(a)

' 对照组：简单变量作调用方 + 无括号
Dim simple
Set simple = New Obs
simple.Name = "S"

Response.Write("== 简单变量调用方（对照）==")
simple.Update "重大新闻"
Response.Write("== 数组元素调用方（问题）==")
p.Notify("重大新闻")
%>
