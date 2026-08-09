Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch18Module
    Public Interface ICommand
        Function Execute() As Object
    End Interface
    Public Class Light
        Public Function TurnOn() As Object
            Console.WriteLine("灯已打开")
        End Function

        Public Function TurnOff() As Object
            Console.WriteLine("灯已关闭")
        End Function
    End Class
    Public Class LightOnCommand
        Implements ICommand

        Private ReadOnly m_Light As Light

        Public Sub New(light As Light)
            m_Light = light
        End Sub

        Public Function Execute() As Object Implements ICommand.Execute
            m_Light.TurnOn()
        End Function
    End Class
    Public Class LightOffCommand
        Implements ICommand

        Private ReadOnly m_Light As Light

        Public Sub New(light As Light)
            m_Light = light
        End Sub

        Public Function Execute() As Object Implements ICommand.Execute
            m_Light.TurnOff()
        End Function
    End Class
    Public Class RemoteControl
        Private m_Command As ICommand

        Public Function SetCommand(cmd As ICommand) As Object
            m_Command = cmd
        End Function

        Public Function PressButton() As Object
            m_Command.Execute()
        End Function
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
