Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch18Module
    Public Interface ICommand
        Sub Execute()
    End Interface
    Public Class Light
        Public Sub TurnOn()
            Console.WriteLine("灯已打开")
        End Sub

        Public Sub TurnOff()
            Console.WriteLine("灯已关闭")
        End Sub
    End Class
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
    Public Class RemoteControl
        Private m_Command As ICommand

        Public Sub SetCommand(cmd As ICommand)
            m_Command = cmd
        End Sub

        Public Sub PressButton()
            m_Command.Execute()
        End Sub
    End Class
    Sub Main()

        ' 接收者：灯

        ' 开灯命令

        ' 关灯命令

        ' 调用者：遥控器

        ' 演示
        Dim light As New Light()
        Dim onCmd As New LightOnCommand(light)
        Dim offCmd As New LightOffCommand(light)
        Dim remote As New RemoteControl()

        remote.SetCommand(onCmd)
        remote.PressButton()   ' 灯已打开

        remote.SetCommand(offCmd)
        remote.PressButton()   ' 灯已关闭
    End Sub
End Module
