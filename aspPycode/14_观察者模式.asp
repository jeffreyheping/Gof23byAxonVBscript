<%
Option Explicit
' 观察者：报纸
Class Newspaper
    Public Name

    ' 收到新闻时的响应
    Public Function Update(news)
        Response.Write(Name & " 收到新闻：" & news)

    End Function
End Class

' 被观察者：新闻社
Class NewsAgency
    Private m_Observers()   ' 观察者数组

    Private m_Count         ' 当前观察者数量

    ' 构造函数：初始化数组
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Observers(10)
    End Sub

    ' 注册观察者（容量不足时自动扩容）
    Public Function Subscribe(observer)
        If m_Count >= UBound(m_Observers) + 1 Then
            ReDim Preserve m_Observers(m_Count * 2)
        End If
        Set m_Observers(m_Count) = observer
        m_Count = m_Count + 1
    End Function

    ' 通知所有观察者
    Public Function Notify(news)
        Dim i
        For i = 0 To m_Count - 1
            m_Observers(i).Update news
        Next
    End Function

    ' 发布新闻：先更新自身状态，再通知所有观察者
    Public Function Publish(news)
        Notify(news)

    End Function
End Class

' 演示：多个观察者同时收到通知
Dim agency, paper1, paper2
Set agency = New NewsAgency
Set paper1 = New Newspaper
paper1.Name = "晨报"
Set paper2 = New Newspaper
paper2.Name = "晚报"

agency.Subscribe(paper1)

agency.Subscribe(paper2)

agency.Publish("重大新闻！")
%>