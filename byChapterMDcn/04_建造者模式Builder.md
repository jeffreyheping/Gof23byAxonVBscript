## 第4章 建造者模式（Builder）

**核心思想**：分步骤构建复杂对象，同样的构建过程可以创建不同配置。

**示例说明**：Director 指挥 Builder 分步组装 Computer（CPU→RAM→Disk）。同一个 Builder 可以被指挥出游戏 PC 或办公 PC 两种配置。

### 传统 VBScript 版

```vbscript
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
```

**传统 VBScript 版妥协说明**：
- **无接口**：Director 接收的 `builder` 参数没有 `IBuilder` 接口约束。如果传入的对象没有 `BuildCPU` 等方法，运行时才报错。
- **无链式调用**：VBScript 虽然用 Function 可以返回 `Me`，但语法上无法实现 `builder.BuildCPU("i9").BuildRAM("32GB")` 这样的链式调用（`Set FunctionName = Me` 写法可行但调用方仍需逐行接收返回值），只能逐行调用。

### Axon VBScript 版（支持 Implements）

```vbscript
' 产品
Class Computer
    Public CPU As String, RAM As String, Disk As String
    Public Function ShowConfig
        Response.Write "配置：" & CPU & " / " & RAM & " / " & Disk
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
        builder.BuildCPU "i9"
        builder.BuildRAM "32GB"
        builder.BuildDisk "2TB SSD"
    End Function

    Public Function ConstructOfficePC(builder As IBuilder)
        builder.BuildCPU "i5"
        builder.BuildRAM "16GB"
        builder.BuildDisk "512GB SSD"
    End Function
End Class

' 演示
Dim myBuilder As IBuilder
Dim myDirector As Director
Dim pc As Computer
Set myBuilder = New ComputerBuilder
Set myDirector = New Director
myDirector.ConstructGamingPC myBuilder
Set pc = myBuilder.GetResult
pc.ShowConfig
```

**Axon VBScript 版妥协说明**：
- 接口机制约束了建造者的方法契约，Director 可通过 `IBuilder` 类型安全地接收建造者。但仍无链式调用语法，无法 `builder.BuildCPU("i9").BuildRAM("32GB")`。
---