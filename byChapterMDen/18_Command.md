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
- `ICommand` guarantees all command classes have an `Execute` method. Command objects can be uniformly stored in arrays or queues. AxonASP's interface method dispatch is fixed — `RemoteControl` holds `Private m_Command As ICommand`, and `PressButton` calls `m_Command.Execute` which auto-dispatches to the concrete command's `ICommand_Execute`. Consistent with standard OOP — no fully-qualified names needed.
---
