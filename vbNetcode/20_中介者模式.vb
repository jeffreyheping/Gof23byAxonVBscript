Option Strict Off
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch20Module
    Public Interface IMediator
        Function Register(user As User)
        Function SendMessage(msg As String, fromUser As User)
    End Interface
    Public Interface IColleague
        Function Receive(msg As String, fromUser As User)
    End Interface
    Public Class ChatRoom
        Implements IMediator
        Private ReadOnly m_Users As New List(Of User)()

        ' 注册同事
        Public Function Register(user As User) Implements IMediator.Register
            m_Users.Add(user)
        End Function

        ' 转发消息：遍历同事，调用 Receive，发送者除外
        Public Function SendMessage(msg As String, fromUser As User) Implements IMediator.SendMessage
            For Each u As User In m_Users
                If Not u Is fromUser Then
                    u.Receive(msg, fromUser)
                End If
            Next
        End Function
    End Class
    Public Class User
        Implements IColleague
        Public Name As String
        Private m_Mediator As IMediator

        ' 加入中介者：保存引用并注册自身
        Public Function Join(mediator As IMediator)
            m_Mediator = mediator
            mediator.Register(Me)
        End Function

        ' 发送消息：交给中介者转发
        Public Function Send(msg As String)
            m_Mediator.SendMessage(msg, Me)
        End Function

        ' 接收消息：显示收到内容
        Public Function Receive(msg As String, fromUser As User) Implements IColleague.Receive
            Console.WriteLine(Name & " 收到 " & fromUser.Name & " 的消息：" & msg)
        End Function
    End Class
    Sub Main()

        ' 同事接口

        ' 聊天室：实现中介者接口，持有所有同事对象

        ' 用户：实现同事接口

        ' 演示：用户之间不直接交互，全部通过聊天室
        Dim room As New ChatRoom()
        Dim alice As New User() With {.Name = "Alice"}
        Dim bob As New User() With {.Name = "Bob"}

        alice.Join(room)
        bob.Join(room)
        alice.Send("大家好！")
    End Sub
End Module
