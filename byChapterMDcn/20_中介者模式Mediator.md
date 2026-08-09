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

```

**传统 VBScript 版妥协说明**：
- **无接口约束**：ChatRoom 和 User 没有 `IMediator`/`IColleague` 接口约束。如果某个类漏写 `SendMessage` 或 `Receive`，运行时调用才报错。
- **Mediator 职责过重**：所有交互逻辑都集中在 ChatRoom 中，如果用户类型增多，ChatRoom 会越来越臃肿，且无法拆分到子类（无继承）。

### Axon VBScript 版（支持 Implements）

```vba
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
        m_Users.Add(user)

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
        mediator.Register(Me)

    End Function

    ' 发送消息：交给中介者转发，Me 关键字传递自身对象
    Public Function Send(msg As String)
        m_Mediator.SendMessage msg, Me
    End Function

    ' 接收消息：显示收到内容
    Public Function IColleague_Receive(msg As String, fromUser As User)
        Response.Write(Name & " 收到 " & fromUser.Name & " 的消息：" & msg & vbCrLf)

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

alice.Join(room)

bob.Join(room)

alice.Send("大家好！")

```

**Axon VBScript 版妥协说明**：
- 缺继承，每个 Observer 手动 Register。`IMediator`/`IColleague` 接口约束了中介者和同事的契约。AxonASP 接口方法派发已修复且 `Me` 关键字工作正常，`ChatRoom` 在 `IMediator_SendMessage` 中用 `For Each` 遍历同事集合并直接调用 `IColleague` 接口方法 `Receive`，`User.Send` 通过 `Me` 将自身作为发送者传递给中介者，模式得以自然实现，无需存储用户名或引入辅助类。剩余限制：缺失继承机制——经典中介者模式可以用抽象基类 `ColleagueBase` 封装 `Join`/`Register` 逻辑，子类自动获得注册能力；Axon 无继承，每个同事类（Observer）都必须手动写一遍 `Join` 保存引用 + 调用 `mediator.Register Me` 的注册代码，同事类越多重复越多。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有 `Interface` 接口、`Implements` 显式实现、`List(Of T)` 泛型集合，中介者模式结构与 Axon 版一致：同样保留 `IMediator`/`IColleague` 接口 + `User`/`ChatRoom` 类 + 手动 `Join` 注册，不引入共享字典或抽象基类骨架。

```vbnet
' 中介者接口
Public Interface IMediator
    Function Register(user As User) As Object
    Function SendMessage(msg As String, fromUser As User) As Object
End Interface

' 同事接口
Public Interface IColleague
    Function Receive(msg As String, fromUser As User) As Object
End Interface

' 聊天室：实现中介者接口，持有所有同事对象
Public Class ChatRoom
    Implements IMediator
    Private ReadOnly m_Users As New List(Of User)()

    ' 注册同事
    Public Function Register(user As User) As Object Implements IMediator.Register
        m_Users.Add(user)
    End Function

    ' 转发消息：遍历同事，调用 Receive，发送者除外
    Public Function SendMessage(msg As String, fromUser As User) As Object Implements IMediator.SendMessage
        For Each u As User In m_Users
            If Not u Is fromUser Then
                u.Receive(msg, fromUser)
            End If
        Next
    End Function
End Class

' 用户：实现同事接口
Public Class User
    Implements IColleague
    Public Name As String
    Private m_Mediator As IMediator

    ' 加入中介者：保存引用并注册自身
    Public Function Join(mediator As IMediator) As Object
        m_Mediator = mediator
        mediator.Register(Me)
    End Function

    ' 发送消息：交给中介者转发
    Public Function Send(msg As String) As Object
        m_Mediator.SendMessage(msg, Me)
    End Function

    ' 接收消息：显示收到内容
    Public Function Receive(msg As String, fromUser As User) As Object Implements IColleague.Receive
        Console.WriteLine(Name & " 收到 " & fromUser.Name & " 的消息：" & msg)
    End Function
End Class

' 演示：用户之间不直接交互，全部通过聊天室
Dim room As New ChatRoom()
Dim alice As New User() With {.Name = "Alice"}
Dim bob As New User() With {.Name = "Bob"}

alice.Join(room)
bob.Join(room)
alice.Send("大家好！")
```

**VB.NET 版说明**：
- **`Interface` + `Implements` 编译期强制契约**：`IMediator`/`IColleague` 接口约束中介者与同事方法，`Implements` 让漏写直接编译报错；与 Axon 版同结构，仅方法名免去 `IMediator_`/`IColleague_` 前缀限定。
- **`List(Of User)` 泛型集合类型安全**：注册与遍历编译期保证元素类型为 `User`，无需 Axon 版的 COM `Collection` 或传统版的 `ReDim Preserve` 扩容。
- **`Me` 关键字传递自身**：`Send` 中 `m_Mediator.SendMessage(msg, Me)` 把自身作为发送者传给中介者，与 Axon 版写法一致。
- **无需 `Set` 区分对象赋值**：`m_Mediator = mediator` 统一用 `=`，`New User() With {.Name = "Alice"}` 一行完成创建与字段初始化。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 中介者-同事契约 | 方法名约定（易漏写） | `IMediator`/`IColleague` 接口约束 | `Interface` + `Implements` 编译期强制 |
| 注册表实现 | 动态数组 + ReDim Preserve | COM `Collection` | `List(Of User)` 泛型集合 |
| 方法命名 | 自由命名 | `IMediator_`/`IColleague_` 前缀限定 | 直接 `SendMessage`/`Receive`，`Implements` 绑定 |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
---