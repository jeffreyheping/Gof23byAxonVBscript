<%
Option Explicit
' 被观察者：声明事件
Class NewsAgency
    Event OnNews(news As String)

    ' 发布新闻：触发事件，所有订阅者自动收到通知
    Public Function Publish(news As String)
        RaiseEvent OnNews(news)
    End Function
End Class

' 具体观察者
Class Newspaper
    Public Name As String

    ' 收到新闻时的响应
    Public Function Update(news As String)
        Response.Write(Name & " 收到新闻：" & news)

    End Function
End Class

' 用 WithEvents 声明事件接收变量
Dim WithEvents agency As NewsAgency

' 事件处理程序：命名规则为 变量名_事件名
Sub agency_OnNews(news As String)
    ' 通过全局引用分发到具体观察者
    paper1.Update(news)

    paper2.Update(news)

End Sub

Dim paper1 As Newspaper, paper2 As Newspaper
Set paper1 = New Newspaper
paper1.Name = "晨报"
Set paper2 = New Newspaper
paper2.Name = "晚报"

Set agency = New NewsAgency
agency.Publish("重大新闻！")
%>