## 第18章 命令模式（Command）

**核心思想**：把请求封装成对象，支持撤销、排队、日志等功能。

**示例说明**：LightOnCommand 和 LightOffCommand 分别封装了开灯和关灯操作。RemoteControl 持有 Command 对象，按下按钮时执行该命令。还可以把命令存入数组实现宏功能。

### 传统 VBScript 版

```vbscript
' 接收者：灯
Class Light
    Public Function TurnOn
        Response.Write "灯已打开"
    End Function
    Public Function TurnOff
        Response.Write "灯已关闭"
    End Function
End Class

' 命令：开灯
Class LightOnCommand
    Private m_Light

    Public Function Init(light)
        Set m_Light = light
    End Function

    Public Function Execute
        m_Light.TurnOn
    End Function
End Class

' 命令：关灯
Class LightOffCommand
    Private m_Light

    Public Function Init(light)
        Set m_Light = light
    End Function

    Public Function Execute
        m_Light.TurnOff
    End Function
End Class

' 调用者：遥控器
Class RemoteControl
    Private m_Command

    Public Function SetCommand(cmd)
        Set m_Command = cmd
    End Function

    Public Function PressButton
        m_Command.Execute
    End Function
End Class

' 演示：把操作封装成对象，可以存储、传递、延迟执行
Dim myLight, onCmd, offCmd, remote
Set myLight = New Light

Set onCmd = New LightOnCommand
onCmd.Init myLight

Set offCmd = New LightOffCommand
offCmd.Init myLight

Set remote = New RemoteControl
remote.SetCommand onCmd
remote.PressButton   ' 灯已打开

remote.SetCommand offCmd
remote.PressButton   ' 灯已关闭
```

**传统 VBScript 版妥协说明**：
- **无接口约束**：LightOnCommand 和 LightOffCommand 没有 `ICommand` 接口强制要求 `Execute` 方法。如果某个命令类方法名不一致，运行时调用才报错。
- **无法类型化队列**：想把多个命令放入数组统一执行时，数组元素没有共同基类或接口，只能依靠"都有 Execute 方法"的约定。

### Axon VBScript 版（支持 Implements）

```vbscript
' 命令接口
Class ICommand
    Public Function Execute
    End Function
End Class

' 接收者
Class Light
    Public Function TurnOn
        Response.Write "灯已打开"
    End Function
    Public Function TurnOff
        Response.Write "灯已关闭"
    End Function
End Class

' 开灯命令
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

' 关灯命令
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

' 调用者
Class RemoteControl
    Private m_Command As ICommand

    Public Function SetCommand(cmd As ICommand)
        Set m_Command = cmd
    End Function

    Public Function PressButton
        m_Command.Execute
    End Function
End Class

' 演示
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

**Axon VBScript 版妥协说明**：
- `ICommand` 接口保证了所有命令类都有 `Execute` 方法，命令对象可以统一存入数组或队列。AxonASP 接口方法派发已修复，`RemoteControl` 持有 `Private m_Command As ICommand`，`PressButton` 中调用 `m_Command.Execute` 即自动路由到具体命令的 `ICommand_Execute`，写法与标准 OOP 一致，无需完整限定名。
---