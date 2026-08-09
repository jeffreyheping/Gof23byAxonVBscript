<%
' 用户类：通过中介发送和接收消息
Class User
    Public Name
    Private m_Mediator   ' 所持中介者引用

    ' 注册到中介
    Public Function Join(mediator)
        Set m_Mediator = mediator
        mediator.Register(Me)

    End Function

    ' 发送消息：交给中介转发
    Public Function Send(msg)
        m_Mediator.SendMessage msg, Me
    End Function

    ' 接收消息：显示收到内容
    Public Function Receive(msg, fromUser)
        Response.Write(Name & " 收到 " & fromUser.Name & " 的消息：" & msg)

    End Function
End Class

' 中介者：聊天室
Class ChatRoom
    Private m_Users()   ' 在线用户数组

    Private m_Count     ' 当前用户数

    ' 构造函数：初始化数组
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Users(10)
    End Sub

    ' 注册用户（容量不足时自动扩容）
    Public Function Register(user)
        If m_Count >= UBound(m_Users) + 1 Then
            ReDim Preserve m_Users(m_Count * 2)
        End If
        Set m_Users(m_Count) = user
        m_Count = m_Count + 1
    End Function

    ' 转发消息：发给除发送者外的所有用户
    Public Function SendMessage(msg, fromUser)
        Dim i
        For i = 0 To m_Count - 1
            If Not m_Users(i) Is fromUser Then
                m_Users(i).Receive msg, fromUser
            End If
        Next
    End Function
End Class

' 演示：用户之间不直接交互，全部通过聊天室
Dim room, alice, bob
Set room = New ChatRoom
Set alice = New User
alice.Name = "Alice"
Set bob = New User
bob.Name = "Bob"

alice.Join(room)

bob.Join(room)

alice.Send("大家好！")
%>