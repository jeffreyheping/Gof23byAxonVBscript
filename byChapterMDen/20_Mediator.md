## Chapter 20: Mediator

**Core idea**: Use a mediator object to encapsulate interactions between multiple objects, avoiding direct references between them.

**Example**: ChatRoom acts as the mediator. When a User sends a message, they don't send it directly to other Users — they hand it to ChatRoom, which forwards it to all online users. Users are fully decoupled from each other.

### Classic VBScript Version

```vbscript
' User class: sends and receives messages through the mediator
Class User
    Public Name
    Private m_Mediator   ' Held mediator reference

    ' Register with mediator
    Public Function Join(mediator)
        Set m_Mediator = mediator
        mediator.Register Me
    End Function

    ' Send message: hand off to mediator for forwarding
    Public Function Send(msg)
        m_Mediator.SendMessage msg, Me
    End Function

    ' Receive message: display received content
    Public Function Receive(msg, fromUser)
        Response.Write Name & " received message from " & fromUser.Name & ": " & msg
    End Function
End Class

' Mediator: chat room
Class ChatRoom
    Private m_Users()   ' Online user array
    Private m_Count     ' Current user count

    ' Constructor: initialize array
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Users(10)
    End Sub

    ' Register user (auto-resize when full)
    Public Function Register(user)
        If m_Count >= UBound(m_Users) + 1 Then
            ReDim Preserve m_Users(m_Count * 2)
        End If
        Set m_Users(m_Count) = user
        m_Count = m_Count + 1
    End Function

    ' Forward message: send to all users except the sender
    Public Function SendMessage(msg, fromUser)
        Dim i
        For i = 0 To m_Count - 1
            If Not m_Users(i) Is fromUser Then
                m_Users(i).Receive msg, fromUser
            End If
        Next
    End Function
End Class

' Demo: users don't interact directly — everything goes through the chat room
Dim room, alice, bob
Set room = New ChatRoom
Set alice = New User
alice.Name = "Alice"
Set bob = New User
bob.Name = "Bob"

alice.Join room
bob.Join room
alice.Send "Hello everyone!"
```

**Classic VBScript trade-offs**:
- **No interface constraint**: ChatRoom and User have no `IMediator`/`IColleague` interface. If a class forgets `SendMessage` or `Receive`, the error only surfaces at runtime.
- **Mediator is too heavy**: All interaction logic is concentrated in ChatRoom. As user types grow, ChatRoom gets bloated — and can't be split into subclasses (no inheritance).

### Axon VBScript Version (supports Implements)

```vbscript
' Mediator interface (contract declaration)
Class IMediator
    Public Function SendMessage(msg As String, fromUser As User)
    End Function
End Class

' Colleague interface (contract declaration)
Class IColleague
    Public Function Receive(msg As String, fromUser As User)
    End Function
End Class

' Chat room: implements mediator interface, holds all colleague objects
Class ChatRoom
    Implements IMediator
    Private m_Users    ' Collection

    Private Sub Class_Initialize
        Set m_Users = Server.CreateObject("Collection")
    End Sub

    ' Register colleague (non-interface public method)
    Public Function Register(user As User)
        m_Users.Add user
    End Function

    ' Forward message: iterate colleagues, call interface method Receive, skip sender
    Public Function IMediator_SendMessage(msg As String, fromUser As User)
        Dim u As IColleague
        For Each u In m_Users
            If Not u Is fromUser Then
                u.Receive msg, fromUser
            End If
        Next
    End Function
End Class

' User: implements colleague interface
Class User
    Implements IColleague
    Public Name As String
    Private m_Mediator As IMediator

    ' Join mediator: save reference and register self
    Public Function Join(mediator As IMediator)
        Set m_Mediator = mediator
        mediator.Register Me
    End Function

    ' Send message: hand off to mediator; Me keyword passes self
    Public Function Send(msg As String)
        m_Mediator.SendMessage msg, Me
    End Function

    ' Receive message: display received content
    Public Function IColleague_Receive(msg As String, fromUser As User)
        Response.Write Name & " received message from " & fromUser.Name & ": " & msg & vbCrLf
    End Function
End Class

' Demo
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
alice.Send "Hello everyone!"
```

**Axon VBScript trade-offs**:
- Missing inheritance, each Observer manually Registers. `IMediator`/`IColleague` interfaces constrain the mediator and colleague contracts. AxonASP's interface method dispatch is fixed and `Me` keyword works correctly. `ChatRoom` uses `For Each` to iterate the colleague collection in `IMediator_SendMessage` and directly calls the `IColleague` interface method `Receive`. `User.Send` passes itself as sender to the mediator via `Me`. The pattern is implemented naturally — no need to store usernames or introduce helper classes. Remaining limitation: missing inheritance mechanism — classic Mediator can use an abstract base class `ColleagueBase` to encapsulate `Join`/`Register` logic, subclasses auto-get registration capability. Axon has no inheritance — every colleague class (Observer) must manually write `Join` to save reference + call `mediator.Register Me` registration code. More colleague classes means more repetition.

### VB.NET Version (syntactically complete baseline)

VB.NET has `Interface`, `Implements` explicit implementation, `List(Of T)` generic collection. Mediator pattern structure is consistent with Axon version: same `IMediator`/`IColleague` interfaces + `User`/`ChatRoom` classes + manual `Join` registration, without introducing shared dictionaries or abstract base class skeletons.

```vbnet
' Mediator interface
Public Interface IMediator
    Sub Register(user As User)
    Sub SendMessage(msg As String, fromUser As User)
End Interface

' Colleague interface
Public Interface IColleague
    Sub Receive(msg As String, fromUser As User)
End Interface

' Chat room: implements mediator interface, holds all colleague objects
Public Class ChatRoom
    Implements IMediator
    Private ReadOnly m_Users As New List(Of User)()

    ' Register colleague
    Public Sub Register(user As User) Implements IMediator.Register
        m_Users.Add(user)
    End Sub

    ' Forward message: iterate colleagues, call Receive, excluding sender
    Public Sub SendMessage(msg As String, fromUser As User) Implements IMediator.SendMessage
        For Each u As User In m_Users
            If Not u Is fromUser Then
                u.Receive(msg, fromUser)
            End If
        Next
    End Sub
End Class

' User: implements colleague interface
Public Class User
    Implements IColleague
    Public Name As String
    Private m_Mediator As IMediator

    ' Join mediator: save reference and register self
    Public Sub Join(mediator As IMediator)
        m_Mediator = mediator
        mediator.Register(Me)
    End Sub

    ' Send message: hand off to mediator for forwarding
    Public Sub Send(msg As String)
        m_Mediator.SendMessage(msg, Me)
    End Sub

    ' Receive message: display received content
    Public Sub Receive(msg As String, fromUser As User) Implements IColleague.Receive
        Console.WriteLine(Name & " received message from " & fromUser.Name & ": " & msg)
    End Sub
End Class

' Demo: users don't interact directly, all through chat room
Dim room As New ChatRoom()
Dim alice As New User() With {.Name = "Alice"}
Dim bob As New User() With {.Name = "Bob"}

alice.Join(room)
bob.Join(room)
alice.Send("Hello everyone!")
```

**VB.NET version notes**:
- **`Interface` + `Implements` compile-time enforced contracts**: `IMediator`/`IColleague` interfaces constrain mediator and colleague methods. `Implements` makes missing methods cause compile errors; same structure as Axon version, only method names without `IMediator_`/`IColleague_` prefix qualification.
- **`List(Of User)` generic collection type safety**: Registration and traversal compile-time guarantee element type is `User`, no need for Axon's COM `Collection` or classic version's `ReDim Preserve` resizing.
- **`Me` keyword passes self**: `Send` uses `m_Mediator.SendMessage(msg, Me)` to pass itself as sender to mediator, same as Axon version.
- **No `Set` for object assignment**: `m_Mediator = mediator` uniformly uses `=`, `New User() With {.Name = "Alice"}` completes creation and field initialization in one line.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Mediator-colleague contract | Method name convention (easy to miss) | `IMediator`/`IColleague` interface constraints | `Interface` + `Implements` compile-time enforced |
| Registry implementation | Dynamic array + ReDim Preserve | COM `Collection` | `List(Of User)` generic collection |
| Method naming | Free naming | `IMediator_`/`IColleague_` prefix qualification | Direct `SendMessage`/`Receive`, `Implements` binding |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |
---
