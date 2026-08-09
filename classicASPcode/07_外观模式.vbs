Option Explicit
Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
' 子系统：CPU
Class CPU
    ' 冻结当前状态
    Public Function Freeze
        Response.Write("CPU 冻结")

    End Function
    ' 跳转到指定地址
    Public Function Jump(pos)
        Response.Write("CPU 跳转到 " & pos)

    End Function
    ' 开始执行
    Public Function Execute
        Response.Write("CPU 执行")

    End Function
End Class

' 子系统：内存
Class Memory
    ' 将数据加载到指定地址
    Public Function Load(pos, data)
        Response.Write("内存加载 " & data & " 到 " & pos)

    End Function
End Class

' 子系统：硬盘
Class HardDrive
    ' 从指定扇区读取数据
    ' 返回值：模拟的数据块字符串
    Public Function Read(lba)
        Read = "数据块(" & lba & ")"
    End Function
End Class

' 外观类：封装子系统的复杂调用，对外只暴露 Start
Class ComputerFacade
    Private m_CPU, m_Mem, m_HD

    ' 构造函数：初始化所有子系统
    Private Sub Class_Initialize
        Set m_CPU = New CPU
        Set m_Mem = New Memory
        Set m_HD = New HardDrive
    End Sub

    ' 一键开机：内部按顺序调用各子系统
    Public Function Start
        m_CPU.Freeze
        Dim bootData
        bootData = m_HD.Read(0)
        m_Mem.Load 0, bootData
        m_CPU.Jump(0)

        m_CPU.Execute
    End Function
End Class

' 演示：外部只需调用 Start，无需了解内部细节
Dim pc
Set pc = New ComputerFacade
pc.Start()   ' 对外只暴露一个 Start

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
