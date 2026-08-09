## 第7章 外观模式（Facade）

**核心思想**：为复杂子系统提供一个简单的统一入口。

**示例说明**：开机过程涉及 CPU 冻结→硬盘读取→内存加载→CPU 跳转→CPU 执行，ComputerFacade 把这些步骤封装成一个 Start 方法，外部只需调用一次。

### 传统 VBScript 版

```vbscript
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

```

**传统 VBScript 版妥协说明**：
此模式在 VBScript 中实现较为自然，外观类只是组合调用子系统，不依赖继承或接口，无显著妥协。

### Axon VBScript 版（支持强类型）

```vba
' 子系统：CPU
Class CPU
    ' 冻结当前状态
    Public Function Freeze
        Response.Write("CPU 冻结")

    End Function
    ' 跳转到指定地址
    Public Function Jump(pos As Long)
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
    Public Function Load(pos As Long, data As String)
        Response.Write("内存加载 " & data & " 到 " & pos)

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
        m_CPU.Jump(0)

        m_CPU.Execute
    End Function
End Class

' 演示：外部只需调用 Start，无需了解内部细节
Dim pc As ComputerFacade
Set pc = New ComputerFacade
pc.Start()   ' 对外只暴露一个 Start

```

**Axon VBScript 版妥协说明**：
- 传统版已无结构性妥协：外观类只是组合调用子系统，不依赖继承或接口，本身就是地道实现。强类型版本将三个子系统引用标注为 `As CPU`/`As Memory`/`As HardDrive`，将地址/扇区参数标注为 `As Long`、数据参数标注为 `As String`，使 IDE 能在编译期发现"误传字符串给地址参数"等类型错误。附录分类为"传统已地道"——这是 23 个模式中仅有的三个之一（外观 #7、享元 #12、备忘录 #22），原因是这三个模式的核心结构完全不依赖接口与继承，只要有 `Class` + 组合调用就能地道实现。AxonASP 的强类型标注是实现质量的提升，不改变外观模式的组合结构。

### VB.NET 版（语法完备的对照基准）

VB.NET 用 `Sub`/`Function` 替代 VBScript 的统一 `Function`，构造函数用 `Sub New()` 替代 `Class_Initialize`。场景与 Axon 版一致——子系统是普通 Class，外观类组合调用即可。

```vbnet
' ① 子系统：普通 Class，与 Axon 版一一对应
Public Class CPU
    Public Function Freeze()
        Console.WriteLine("CPU 冻结")
    End Function
    Public Function Jump(position As Long)
        Console.WriteLine("CPU 跳转到 " & position)
    End Function
    Public Function Execute()
        Console.WriteLine("CPU 执行")
    End Function
End Class

Public Class Memory
    Public Function Load(position As Long, data As String)
        Console.WriteLine("内存加载 " & data & " 到 " & position)
    End Function
End Class

Public Class HardDrive
    Public Function Read(lba As Long) As String
        Return "数据块(" & lba & ")"
    End Function
End Class

' ② 外观类：组合持有子系统，对外只暴露 Start
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
    Public Function Start()
        m_CPU.Freeze()
        Dim bootData As String = m_HD.Read(0)
        m_Mem.Load(0, bootData)
        m_CPU.Jump(0)
        m_CPU.Execute()
    End Function
End Class

' 演示：外部只需调用 Start，无需了解内部细节
Dim pc As New ComputerFacade()
pc.Start()   ' 对外只暴露一个 Start
```

**VB.NET 版说明**：
- **外观结构不需要继承或接口**：三个版本的核心骨架完全一致——外观类组合持有多个子系统，一个方法串起多个子系统调用。附录把外观 #7 列为"传统已地道"的根本原因。
- **`Sub` 替代统一 `Function`**：VB.NET 区分 `Sub`（无返回值）和 `Function`（有返回值），不返回值的方法用 `Sub` 更规范。Axon 版所有方法都是 `Function`，即使不返回值。
- **`Sub New()` 替代 `Class_Initialize`**：VB.NET 用显式构造函数 `Public Sub New()` 初始化对象，Axon 版用 `Private Sub Class_Initialize`。
- **无需 `Set`**：VB.NET 对象赋值直接用 `=`，`m_CPU = New CPU()` 不需要 `Set`。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 模式地道性 | ✅ 地道（组合调用即可） | ✅ 地道（结构未变，仅加类型标注） | ✅ 地道（同骨架） |
| 方法语法 | 统一 `Function` | 统一 `Function` | `Sub`（无返回）/ `Function`（有返回） |
| 构造函数 | `Class_Initialize` | `Class_Initialize` | `Sub New()` |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
---