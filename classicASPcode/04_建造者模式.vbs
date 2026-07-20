Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
' 产品类：电脑
Class Computer
    Public CPU, RAM, Disk
    ' 打印当前配置
    Public Function ShowConfig
        Response.Write "配置：" & CPU & " / " & RAM & " / " & Disk
    End Function
End Class

' 建造者：逐步组装 Computer 的各个部件
Class ComputerBuilder
    Private m_Computer

    ' 构造函数：创建空白产品实例
    Private Sub Class_Initialize
        Set m_Computer = New Computer
    End Sub

    ' 安装 CPU
    Public Function BuildCPU(v)
        m_Computer.CPU = v
    End Function
    ' 安装内存
    Public Function BuildRAM(v)
        m_Computer.RAM = v
    End Function
    ' 安装硬盘
    Public Function BuildDisk(v)
        m_Computer.Disk = v
    End Function
    ' 返回组装完成的产品
    Public Function GetResult
        Set GetResult = m_Computer
    End Function
End Class

' 指挥者：按固定步骤调用 Builder，封装不同配置方案
Class Director
    ' 方案一：组装游戏 PC
    Public Function ConstructGamingPC(b)
        b.BuildCPU "i9"
        b.BuildRAM "32GB"
        b.BuildDisk "2TB SSD"
    End Function
    ' 方案二：组装办公 PC
    Public Function ConstructOfficePC(b)
        b.BuildCPU "i5"
        b.BuildRAM "16GB"
        b.BuildDisk "512GB SSD"
    End Function
End Class

' 演示：Director 指挥 Builder 组装
Dim myBuilder, myDirector, pc
Set myBuilder = New ComputerBuilder
Set myDirector = New Director
myDirector.ConstructGamingPC myBuilder
Set pc = myBuilder.GetResult
pc.ShowConfig

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
