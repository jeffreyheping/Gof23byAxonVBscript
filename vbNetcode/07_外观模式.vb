Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch07Module
    Public Class CPU
        Public Function Freeze() As Object
            Console.WriteLine("CPU 冻结")
        End Function
        Public Function Jump(position As Long) As Object
            Console.WriteLine("CPU 跳转到 " & position)
        End Function
        Public Function Execute() As Object
            Console.WriteLine("CPU 执行")
        End Function
    End Class
    Public Class Memory
        Public Function Load(position As Long, data As String) As Object
            Console.WriteLine("内存加载 " & data & " 到 " & position)
        End Function
    End Class
    Public Class HardDrive
        Public Function Read(lba As Long) As String
            Return "数据块(" & lba & ")"
        End Function
    End Class
    Public Class ComputerFacade
        Private m_CPU As CPU
        Private m_Mem As Memory
        Private m_HD As HardDrive

        ' 构造函数：直接 New 创建子系统（与 Axon 版 Class_Initialize 一致）
        Public Sub New()
            m_CPU = New CPU()
            m_Mem = New Memory()
            m_HD = New HardDrive()
        End Sub

        ' 一键开机：内部按顺序调用各子系统
        Public Function Start() As Object
            m_CPU.Freeze()
            Dim bootData As String = m_HD.Read(0)
            m_Mem.Load(0, bootData)
            m_CPU.Jump(0)
            m_CPU.Execute()
        End Function
    End Class
    Sub Main()



        ' ② 外观类：组合持有子系统，对外只暴露 Start

        ' 演示：外部只需调用 Start，无需了解内部细节
        Dim pc As New ComputerFacade()
        pc.Start()   ' 对外只暴露一个 Start
    End Sub
End Module
