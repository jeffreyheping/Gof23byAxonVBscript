## Chapter 18: Command

**Core idea**: Encapsulate a request as an object, enabling undo, queuing, logging, etc.

**Example**: LightOnCommand and LightOffCommand encapsulate turning the light on and off. RemoteControl holds a Command object — pressing the button executes that command. You can also store commands in an array for macro functionality.

### Classic VBScript Version

```vbscript
' Receiver: light
Class Light
    Public Function TurnOn
        Response.Write "Light is on"
    End Function
    Public Function TurnOff
        Response.Write "Light is off"
    End Function
End Class

' Command: turn light on
Class LightOnCommand
    Private m_Light

    Public Function Init(light)
        Set m_Light = light
    End Function

    Public Function Execute
        m_Light.TurnOn
    End Function
End Class

' Command: turn light off
Class LightOffCommand
    Private m_Light

    Public Function Init(light)
        Set m_Light = light
    End Function

    Public Function Execute
        m_Light.TurnOff
    End Function
End Class

' Invoker: remote control
Class RemoteControl
    Private m_Command

    Public Function SetCommand(cmd)
        Set m_Command = cmd
    End Function

    Public Function PressButton
        m_Command.Execute
    End Function
End Class

' Demo: encapsulate operations as objects for storage, passing, and delayed execution
Dim myLight, onCmd, offCmd, remote
Set myLight = New Light

Set onCmd = New LightOnCommand
onCmd.Init myLight

Set offCmd = New LightOffCommand
offCmd.Init myLight

Set remote = New RemoteControl
remote.SetCommand onCmd
remote.PressButton   ' Light is on

remote.SetCommand offCmd
remote.PressButton   ' Light is off
```

**Classic VBScript trade-offs**:
- **No interface constraint**: LightOnCommand and LightOffCommand have no `ICommand` interface forcing `Execute`. If a command class has an inconsistent method name, the error only surfaces at runtime.
- **No typed queue**: When storing multiple commands in an array for batch execution, array elements share no common base class or interface — you rely on the "they all have Execute" convention.

### Axon VBScript Version (supports Implements)

```vbscript
' Command interface
Class ICommand
    Public Function Execute
    End Function
End Class

' Receiver
Class Light
    Public Function TurnOn
        Response.Write "Light is on"
    End Function
    Public Function TurnOff
        Response.Write "Light is off"
    End Function
End Class

' Command: turn light on
Class LightOnCommand
    Implements ICommand
    Private m_Light As Light

    Public Function Init(light As Light)
        Set m_Light = light
    End Function

    Public Function ICommand_Execute
        m_Light.TurnOn
    End Function
End Class

' Command: turn light off
Class LightOffCommand
    Implements ICommand
    Private m_Light As Light

    Public Function Init(light As Light)
        Set m_Light = light
    End Function

    Public Function ICommand_Execute
        m_Light.TurnOff
    End Function
End Class

' Invoker
Class RemoteControl
    Private m_Command As ICommand

    Public Function SetCommand(cmd As ICommand)
        Set m_Command = cmd
    End Function

    Public Function PressButton
        m_Command.Execute
    End Function
End Class

' Demo
Dim light As Light
Set light = New Light

Dim onCmd As LightOnCommand
Set onCmd = New LightOnCommand
onCmd.Init light

Dim offCmd As LightOffCommand
Set offCmd = New LightOffCommand
offCmd.Init light

Dim remote As RemoteControl
Set remote = New RemoteControl
remote.SetCommand onCmd
remote.PressButton

remote.SetCommand offCmd
remote.PressButton
```

**Axon VBScript trade-offs**:
- Fully solved, no compromises. `ICommand` interface guarantees all command classes have an `Execute` method. Command objects can be uniformly stored in arrays or queues. AxonASP's interface method dispatch is fixed — `RemoteControl` holds `Private m_Command As ICommand`, `PressButton` calls `m_Command.Execute` which auto-dispatches to the concrete command's `ICommand_Execute`. Consistent with standard OOP — no fully-qualified names needed.

### VB.NET Version (syntactically complete baseline)

VB.NET has `Interface`, `Implements` explicit implementation, parameterized constructors. Command pattern structure is consistent with Axon version, only replacing `Init` two-step assembly with constructors, without introducing Undo/stack features not present in Axon version.

```vbnet
' Command interface
Public Interface ICommand
    Sub Execute()
End Interface

' Receiver: Light
Public Class Light
    Public Sub TurnOn()
        Console.WriteLine("Light is on")
    End Sub

    Public Sub TurnOff()
        Console.WriteLine("Light is off")
    End Sub
End Class

' LightOn command
Public Class LightOnCommand
    Implements ICommand

    Private ReadOnly m_Light As Light

    Public Sub New(light As Light)
        m_Light = light
    End Sub

    Public Sub Execute() Implements ICommand.Execute
        m_Light.TurnOn()
    End Sub
End Class

' LightOff command
Public Class LightOffCommand
    Implements ICommand

    Private ReadOnly m_Light As Light

    Public Sub New(light As Light)
        m_Light = light
    End Sub

    Public Sub Execute() Implements ICommand.Execute
        m_Light.TurnOff()
    End Sub
End Class

' Invoker: Remote control
Public Class RemoteControl
    Private m_Command As ICommand

    Public Sub SetCommand(cmd As ICommand)
        m_Command = cmd
    End Sub

    Public Sub PressButton()
        m_Command.Execute()
    End Sub
End Class

' Demo
Dim light As New Light()
Dim onCmd As New LightOnCommand(light)
Dim offCmd As New LightOffCommand(light)
Dim remote As New RemoteControl()

remote.SetCommand(onCmd)
remote.PressButton()   ' Light is on

remote.SetCommand(offCmd)
remote.PressButton()   ' Light is off
```

**VB.NET version notes**:
- **`Interface ICommand` + `Implements` compile-time enforces `Execute`**: Same as Axon version using interface to constrain command contract. `Implements ICommand.Execute` makes missing methods cause compile errors — no more runtime failures from misspelled method names.
- **Parameterized constructor one-step assembly**: `New LightOnCommand(light)` injects receiver in one line; first two versions need `New` then manual `Init(light)` two-step, easy to forget the second step.
- **`ReadOnly` field locks receiver reference**: `Private ReadOnly m_Light` immutable after construction, prevents command object from having its receiver swapped midway.
- **No `Set`/`Let` distinction**: Object assignment uniformly uses `=`, `m_Command = cmd` no longer needs to remember `Set` for objects.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Command contract | Method name convention (easy to miss) | `ICommand` interface constrains `Execute` | `Interface` + `Implements` compile-time enforced |
| Assembly method | `New` + manual `Init` two-step | `New` + manual `Init` two-step | Parameterized constructor `New(light)` one-step |
| Receiver reference protection | `Private` field (can be rewritten) | `Private m_Light As Light` | `Private ReadOnly` immutable after construction |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |
---
