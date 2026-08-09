<%
' 产品
Class Computer
    Public CPU As String, RAM As String, Disk As String
    Public Function ShowConfig
        Response.Write("配置：" & CPU & " / " & RAM & " / " & Disk)

    End Function
End Class

' 建造者接口
Class IBuilder
    Public Function BuildCPU(v As String)
    End Function
    Public Function BuildRAM(v As String)
    End Function
    Public Function BuildDisk(v As String)
    End Function
    Public Function GetResult As Computer
    End Function
End Class

' 具体建造者
Class ComputerBuilder
    Implements IBuilder
    Private m_Computer As Computer

    Private Sub Class_Initialize
        Set m_Computer = New Computer
    End Sub

    Public Function IBuilder_BuildCPU(v As String)
        m_Computer.CPU = v
    End Function
    Public Function IBuilder_BuildRAM(v As String)
        m_Computer.RAM = v
    End Function
    Public Function IBuilder_BuildDisk(v As String)
        m_Computer.Disk = v
    End Function
    Public Function IBuilder_GetResult As Computer
        Set IBuilder_GetResult = m_Computer
    End Function
End Class

' 指挥者
Class Director
    Public Function ConstructGamingPC(builder As IBuilder)
        builder.BuildCPU("i9")

        builder.BuildRAM("32GB")

        builder.BuildDisk("2TB SSD")

    End Function

    Public Function ConstructOfficePC(builder As IBuilder)
        builder.BuildCPU("i5")

        builder.BuildRAM("16GB")

        builder.BuildDisk("512GB SSD")

    End Function
End Class

' 演示
Dim myBuilder As IBuilder
Dim myDirector As Director
Dim pc As Computer
Set myBuilder = New ComputerBuilder
Set myDirector = New Director
myDirector.ConstructGamingPC(myBuilder)

Set pc = myBuilder.GetResult
pc.ShowConfig
%>