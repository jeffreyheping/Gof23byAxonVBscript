Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch04Module
    Public Class Computer
        Public Property CPU As String
        Public Property RAM As String
        Public Property Disk As String

        Public Sub ShowConfig()
            Console.WriteLine($"配置：{CPU} / {RAM} / {Disk}")
        End Sub
    End Class
    Public Interface IComputerBuilder
        Function BuildCPU(cpu As String) As IComputerBuilder
        Function BuildRAM(ram As String) As IComputerBuilder
        Function BuildDisk(disk As String) As IComputerBuilder
        Function GetResult() As Computer
    End Interface
    Public Class ComputerBuilder
        Implements IComputerBuilder

        Private m_CPU As String
        Private m_RAM As String
        Private m_Disk As String

        Public Function BuildCPU(cpu As String) As IComputerBuilder Implements IComputerBuilder.BuildCPU
            m_CPU = cpu
            Return Me
        End Function

        Public Function BuildRAM(ram As String) As IComputerBuilder Implements IComputerBuilder.BuildRAM
            m_RAM = ram
            Return Me
        End Function

        Public Function BuildDisk(disk As String) As IComputerBuilder Implements IComputerBuilder.BuildDisk
            m_Disk = disk
            Return Me
        End Function

        Public Function GetResult() As Computer Implements IComputerBuilder.GetResult
            Return New Computer() With {.CPU = m_CPU, .RAM = m_RAM, .Disk = m_Disk}
        End Function
    End Class
    Public Class Director
        Private ReadOnly m_Builder As IComputerBuilder

        Public Sub New(builder As IComputerBuilder)
            m_Builder = builder
        End Sub

        Public Function ConstructGamingPC() As Computer
            Return m_Builder.BuildCPU("i9").BuildRAM("32GB").BuildDisk("2TB SSD").GetResult()
        End Function

        Public Function ConstructOfficePC() As Computer
            Return m_Builder.BuildCPU("i5").BuildRAM("16GB").BuildDisk("512GB SSD").GetResult()
        End Function
    End Class
    Sub Main()

        ' ② 建造者接口：每个 Build* 方法返回自身类型，支持链式

        ' ③ 具体建造者：每个方法 Return Me，实现链式调用

        ' ④ Director：封装预配置方案，内部用链式调用更简洁

        ' 演示：通过 Director 走预设方案
        Dim director As New Director(New ComputerBuilder())
        Dim gamingPC As Computer = director.ConstructGamingPC()
        gamingPC.ShowConfig()   ' 配置：i9 / 32GB / 2TB SSD
    End Sub
End Module
