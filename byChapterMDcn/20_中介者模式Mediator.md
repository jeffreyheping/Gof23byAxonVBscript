## 第20章 中介者模式（Mediator）

**核心思想**：用一个中介对象封装多个对象之间的交互，避免对象直接引用。

**示例说明**：ChatRoom 作为中介，User 发送消息时不直接发给其他 User，而是交给 ChatRoom，由 ChatRoom 转发给所有在线用户。用户之间完全解耦。

### 传统 VBScript 版

```vbscript
' 用户类：通过中介发送和接收消息
Class User
    Public Name
    Private m_Mediator   ' 所持中介者引用

    ' 注册到中介
    Public Function Join(mediator)
        Set m_Mediator = mediator
        mediator.Register Me
    End Function

    ' 发送消息：交给中介转发
    Public Function Send(msg)
        m_Mediator.SendMessage msg, Me
    End Function

    ' 接收消息：显示收到内容
    Public Function Receive(msg, fromUser)
        Response.Write Name & " 收到 " & fromUser.Name & " 的消息：" & msg
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

alice.Join room
bob.Join room
alice.Send "大家好！"
```

**传统 VBScript 版妥协说明**：
- **无接口约束**：ChatRoom 和 User 没有 `IMediator`/`IColleague` 接口约束。如果某个类漏写 `SendMessage` 或 `Receive`，运行时调用才报错。
- **Mediator 职责过重**：所有交互逻辑都集中在 ChatRoom 中，如果用户类型增多，ChatRoom 会越来越臃肿，且无法拆分到子类（无继承）。

### Axon VBScript 版（支持 Implements）

```vbscript
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
```

**Axon VBScript 版妥协说明**：
- `IMediator`/`IColleague` 接口约束了中介者和同事的契约。AxonASP 接口方法派发已修复且 `Me` 关键字工作正常，`ChatRoom` 在 `IMediator_SendMessage` 中用 `For Each` 遍历同事集合并直接调用 `IColleague` 接口方法 `Receive`，`User.Send` 通过 `Me` 将自身作为发送者传递给中介者，模式得以自然实现，无需存储用户名或引入辅助类。剩余限制：`Register` 为非接口公共方法，未纳入 `IMediator` 契约。Mediator 职责集中是模式本身的特点，非语言缺陷——Go 实现中介者模式同样如此，可通过拆分多个中介者接口缓解。
---