<%
' 接收者：灯
Class Light
    Public Function TurnOn
        Response.Write("灯已打开")

    End Function
    Public Function TurnOff
        Response.Write("灯已关闭")

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
onCmd.Init(myLight)


Set offCmd = New LightOffCommand
offCmd.Init(myLight)


Set remote = New RemoteControl
remote.SetCommand(onCmd)

remote.PressButton()   ' 灯已打开


remote.SetCommand(offCmd)

remote.PressButton()   ' 灯已关闭
%>