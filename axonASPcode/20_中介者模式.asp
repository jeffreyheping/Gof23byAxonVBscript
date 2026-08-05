<%
' 中介者接口（契约声明）
Class IMediator
    Public Function SendMessage(msg As String, fromUser As User)
    End Function
End Class

' 同事接口（契约声明）
Class IColleague
    Public Function Receive(msg As String, fromUser As User)
    End Function
End Class

' 聊天室：实现中介者接口，持有所有同事对象
Class ChatRoom
    Implements IMediator
    Private m_Users    ' Collection

    Private Sub Class_Initialize
        Set m_Users = Server.CreateObject("Collection")
    End Sub

    ' 注册同事（非接口公共方法）
    Public Function Register(user As User)
        m_Users.Add user
    End Function

    ' 转发消息：遍历同事，调用接口方法 Receive，发送者除外
    Public Function IMediator_SendMessage(msg As String, fromUser As User)
        Dim u As IColleague
        For Each u In m_Users
            If Not u Is fromUser Then
                u.Receive msg, fromUser
            End If
        Next
    End Function
End Class

' 用户：实现同事接口
Class User
    Implements IColleague
    Public Name As String
    Private m_Mediator As IMediator

    ' 加入中介者：保存引用并注册自身
    Public Function Join(mediator As IMediator)
        Set m_Mediator = mediator
        mediator.Register Me
    End Function

    ' 发送消息：交给中介者转发，Me 关键字传递自身对象
    Public Function Send(msg As String)
        m_Mediator.SendMessage msg, Me
    End Function

    ' 接收消息：显示收到内容
    Public Function IColleague_Receive(msg As String, fromUser As User)
        Response.Write Name & " 收到 " & fromUser.Name & " 的消息：" & msg & vbCrLf
    End Function
End Class

' 演示
Dim room As ChatRoom
Set room = New ChatRoom

Dim alice As User
Set alice = New User
alice.Name = "Alice"

Dim bob As User
Set bob = New User
bob.Name = "Bob"

alice.Join room
bob.Join room
alice.Send "大家好！"
%>