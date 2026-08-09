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
        Response.Write("配置：" & CPU & " / " & RAM & " / " & Disk)

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
        b.BuildCPU("i9")

        b.BuildRAM("32GB")

        b.BuildDisk("2TB SSD")

    End Function
    ' 方案二：组装办公 PC
    Public Function ConstructOfficePC(b)
        b.BuildCPU("i5")

        b.BuildRAM("16GB")

        b.BuildDisk("512GB SSD")

    End Function
End Class

' 演示：Director 指挥 Builder 组装
Dim myBuilder, myDirector, pc
Set myBuilder = New ComputerBuilder
Set myDirector = New Director
myDirector.ConstructGamingPC(myBuilder)

Set pc = myBuilder.GetResult
pc.ShowConfig
```

**传统 VBScript 版妥协说明**：
- **无接口**：Director 接收的 `builder` 参数没有 `IBuilder` 接口约束。如果传入的对象没有 `BuildCPU` 等方法，运行时才报错。
- **无链式调用**：VBScript 虽然用 Function 可以返回 `Me`，但语法上无法实现 `builder.BuildCPU("i9").BuildRAM("32GB")` 这样的链式调用（`Set FunctionName = Me` 写法可行但调用方仍需逐行接收返回值），只能逐行调用。

### Axon VBScript 版（支持 Implements）

```vba
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
```

**Axon VBScript 版妥协说明**：
- 接口机制约束了建造者的方法契约，Director 可通过 `IBuilder` 类型安全地接收建造者。残留限制：**缺失链式调用语法糖**。经典建造者模式鼓励 `builder.BuildCPU("i9").BuildRAM("32GB").BuildDisk("2TB SSD")` 链式写法——每个 Build* 方法返回 `Me`，调用方一行写完所有步骤。AxonASP 即使让每个方法 `Set FunctionName = Me`，调用方仍需逐行写 `builder.BuildCPU "i9"`、`builder.BuildRAM "32GB"`，因为 VBScript 不支持对象方法的链式调用表达式（无法把前一个方法返回值当作下一个方法的调用者继续写在同一行）。Go 同样有链式调用（`builder.CPU("i9").RAM("32GB")`），这是表达力的缺失而非功能缺失。

### VB.NET 版（语法完备的对照基准）

VB.NET 建造者的地道写法：每个 Build* 方法 `Return Me` 实现流畅链式调用（Fluent Builder），弥补 Axon 版无法链式的语法缺失。

```vbnet
' ① 产品：字段 Public，与 Axon 版结构一致
Public Class Computer
    Public Property CPU As String
    Public Property RAM As String
    Public Property Disk As String

    Public Function ShowConfig()
        Console.WriteLine($"配置：{CPU} / {RAM} / {Disk}")
    End Function
End Class

' ② 建造者接口：每个 Build* 方法返回自身类型，支持链式
Public Interface IComputerBuilder
    Function BuildCPU(cpu As String) As IComputerBuilder
    Function BuildRAM(ram As String) As IComputerBuilder
    Function BuildDisk(disk As String) As IComputerBuilder
    Function GetResult() As Computer
End Interface

' ③ 具体建造者：每个方法 Return Me，实现链式调用
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

' ④ Director：封装预配置方案，内部用链式调用更简洁
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

' 演示：通过 Director 走预设方案
Dim director As New Director(New ComputerBuilder())
Dim gamingPC As Computer = director.ConstructGamingPC()
gamingPC.ShowConfig()   ' 配置：i9 / 32GB / 2TB SSD
```

**VB.NET 版说明**：
- **真正的链式调用（Fluent Interface）**：每个 `Build*` 方法 `Return Me`，调用方一行用 `.` 连到底：`builder.BuildCPU("i9").BuildRAM("32GB").BuildDisk("2TB SSD")`。Axon 版即使 `Set FunctionName = Me`，VBScript 语法也不允许把多个调用串在同一行，只能逐行拆开。
- **接口返回自身类型**：`IComputerBuilder` 的每个 Build* 方法返回 `IComputerBuilder` 而非 `void`，这是链式调用的类型基础。Axon 版的 `IBuilder` 接口方法没有返回值，无法链式。
- **Director 内部也用链式**：`ConstructGamingPC` 一行写完全部步骤，比 Axon 版的逐行写法更紧凑、可读性更高。
- **无需 `Set`**：VB.NET 对象赋值直接用 `=`，`Return New Computer()` 不需要 `Set`。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 建造者契约 | 方法名约定 | `IBuilder` 接口约束方法签名 | `IComputerBuilder` 接口 + 每个方法返回自身类型 |
| 链式调用 | 不能（只能逐行） | 不能（语法不支持串接表达式） | 地道 Fluent Interface：`a().b().c().GetResult()` |
| Director 写法 | 逐行赋值 3~6 句 | 逐行赋值 3~6 句 | 链式一行连到底，可读性高 |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
---