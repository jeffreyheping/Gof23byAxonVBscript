## 第7章 外观模式（Facade）

**核心思想**：为复杂子系统提供一个简单的统一入口。

**示例说明**：开机过程涉及 CPU 冻结→硬盘读取→内存加载→CPU 跳转→CPU 执行，ComputerFacade 把这些步骤封装成一个 Start 方法，外部只需调用一次。

### 传统 VBScript 版

```vbscript
' 子系统：CPU
Class CPU
    ' 冻结当前状态
    Public Function Freeze
        Response.Write "CPU 冻结"
    End Function
    ' 跳转到指定地址
    Public Function Jump(pos)
        Response.Write "CPU 跳转到 " & pos
    End Function
    ' 开始执行
    Public Function Execute
        Response.Write "CPU 执行"
    End Function
End Class

' 子系统：内存
Class Memory
    ' 将数据加载到指定地址
    Public Function Load(pos, data)
        Response.Write "内存加载 " & data & " 到 " & pos
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
        m_CPU.Jump 0
        m_CPU.Execute
    End Function
End Class

' 演示：外部只需调用 Start，无需了解内部细节
Dim pc
Set pc = New ComputerFacade
pc.Start   ' 对外只暴露一个 Start
```

**传统 VBScript 版妥协说明**：
此模式在 VBScript 中实现较为自然，外观类只是组合调用子系统，不依赖继承或接口，无显著妥协。

### Axon VBScript 版（支持强类型）

```vbscript
' 子系统：CPU
Class CPU
    ' 冻结当前状态
    Public Function Freeze
        Response.Write "CPU 冻结"
    End Function
    ' 跳转到指定地址
    Public Function Jump(pos As Long)
        Response.Write "CPU 跳转到 " & pos
    End Function
    ' 开始执行
    Public Function Execute
        Response.Write "CPU 执行"
    End Function
End Class

' 子系统：内存
Class Memory
    ' 将数据加载到指定地址
    Public Function Load(pos As Long, data As String)
        Response.Write "内存加载 " & data & " 到 " & pos
    End Function
End Class

' 子系统：硬盘
Class HardDrive
    ' 从指定扇区读取数据
    ' 返回值：模拟的数据块字符串
    Public Function Read(lba As Long) As String
        Read = "数据块(" & lba & ")"
    End Function
End Class

' 外观类：封装子系统的复杂调用，对外只暴露 Start
Class ComputerFacade
    Private m_CPU As CPU
    Private m_Mem As Memory
    Private m_HD As HardDrive

    ' 构造函数：初始化所有子系统
    Private Sub Class_Initialize
        Set m_CPU = New CPU
        Set m_Mem = New Memory
        Set m_HD = New HardDrive
    End Sub

    ' 一键开机：内部按顺序调用各子系统
    Public Function Start
        m_CPU.Freeze
        Dim bootData As String
        bootData = m_HD.Read(0)
        m_Mem.Load 0, bootData
        m_CPU.Jump 0
        m_CPU.Execute
    End Function
End Class

' 演示：外部只需调用 Start，无需了解内部细节
Dim pc As ComputerFacade
Set pc = New ComputerFacade
pc.Start   ' 对外只暴露一个 Start
```

**Axon VBScript 版妥协说明**：
- 传统版已无结构性妥协：外观类只是组合调用子系统，不依赖继承或接口。强类型版本将三个子系统引用标注为 `As CPU`/`As Memory`/`As HardDrive`，将地址/扇区参数标注为 `As Long`、数据参数标注为 `As String`，使 IDE 能在编译期发现"误传字符串给地址参数"等类型错误。这是实现质量的提升，不改变外观模式的组合结构——故仍归为"传统已地道"。

---