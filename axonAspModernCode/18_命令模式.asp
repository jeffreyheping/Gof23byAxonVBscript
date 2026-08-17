<%
Option Explicit
' 命令接口
Class ICommand
    Public Function Execute
    End Function
End Class

' 接收者
Class Light
    Public Function TurnOn
        Response.Write("灯已打开")

    End Function
    Public Function TurnOff
        Response.Write("灯已关闭")

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
onCmd.Init(light)


Dim offCmd As LightOffCommand
Set offCmd = New LightOffCommand
offCmd.Init(light)


Dim remote As RemoteControl
Set remote = New RemoteControl
remote.SetCommand(onCmd)

remote.PressButton

remote.SetCommand(offCmd)

remote.PressButton
%>