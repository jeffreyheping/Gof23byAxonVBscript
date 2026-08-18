## 第17章 责任链模式（Chain of Responsibility）

**核心思想**：把请求沿链传递，直到某个对象处理它。

**示例说明**：Logger 形成 DEBUG→INFO→ERRORLEVEL 的链条。调用 Log 时，若当前级别匹配则处理并继续传递，否则直接传递。最终 ERRORLEVEL 级别会被三个 Logger 都输出。

### 传统 VBScript 版

```vbscript
' 日志处理器：形成责任链
Class Logger
    Public Name
    Public Level        ' 当前处理器能处理的最低级别（DEBUG/INFO/ERRORLEVEL）
    Private m_Next      ' 链上的下一个处理器

    Private Sub Class_Initialize
        Set m_Next = Nothing
    End Sub

    ' 设置下一个处理器
    Public Function SetNext(nextHandler)
        Set m_Next = nextHandler
    End Function

    ' 处理日志请求：若级别匹配则输出，然后传递给下一个
    ' 注：参数名用 msgLevel，避免与公共字段 Level 同名遮蔽（VBScript 大小写不敏感）
    Public Function Log(msg, msgLevel)
        If ShouldHandle(msgLevel) Then
            Response.Write("【" & Name & "】" & msg)

        End If
        If Not m_Next Is Nothing Then
            m_Next.Log msg, msgLevel
        End If
    End Function

    ' 判断当前级别是否应该处理
    Private Function ShouldHandle(msgLevel)
        Dim levels(2)
        levels(0) = "DEBUG"
        levels(1) = "INFO"
        levels(2) = "ERRORLEVEL"

        Dim currentIdx, msgIdx, i
        currentIdx = -1
        msgIdx = -1
        For i = 0 To 2
            If levels(i) = Level Then currentIdx = i
            If levels(i) = msgLevel Then msgIdx = i
        Next
        ShouldHandle = (msgIdx >= currentIdx)
    End Function
End Class

' 演示：构建 DEBUG→INFO→ERRORLEVEL 的责任链
Dim debugLog, infoLog, errorLog
Set debugLog = New Logger
debugLog.Name = "控制台"
debugLog.Level = "DEBUG"

Set infoLog = New Logger
infoLog.Name = "文件"
infoLog.Level = "INFO"

Set errorLog = New Logger
errorLog.Name = "邮件"
errorLog.Level = "ERRORLEVEL"

debugLog.SetNext(infoLog)

infoLog.SetNext(errorLog)


debugLog.Log "系统启动", "INFO"     ' 文件和邮件都输出
debugLog.Log "严重错误", "ERRORLEVEL"    ' 三个都输出
```

**传统 VBScript 版妥协说明**：
- **无接口约束**：Logger 靠 `SetNext` 和 `Log` 方法名约定形成链条，没有 `IHandler` 接口强制要求这两个方法。如果某个类漏写 `SetNext`，链条就断了，运行时才发现。

### Axon VBScript 版（支持 Implements + Enum）

```vba
' 日志级别枚举
Enum LogLevel
    Debug = 0
    Info = 1
    ErrorLevel = 2
End Enum

' 处理器接口
Class IHandler
    Public Function SetNext(handler As IHandler)
    End Function
    Public Function Log(msg As String, level As LogLevel)
    End Function
End Class

' 具体处理器
Class Logger
    Implements IHandler
    Private m_Name As String
    Private m_Level As LogLevel
    Private m_Next As IHandler

    Public Property Get Name As String
        Name = m_Name
    End Property
    Public Property Let Name(v As String)
        m_Name = v
    End Property

    Public Property Get Level As LogLevel
        Level = m_Level
    End Property
    Public Property Let Level(v As LogLevel)
        m_Level = v
    End Property

    Private Sub Class_Initialize
        Set m_Next = Nothing
    End Sub

    Public Function IHandler_SetNext(handler As IHandler)
        Set m_Next = handler
    End Function

    Public Function IHandler_Log(msg As String, level As LogLevel)
        If ShouldHandle(level) Then
            Response.Write("【" & m_Name & "】" & msg)

        End If
        If m_Next IsNot Nothing Then
            m_Next.Log msg, level
        End If
    End Function

    Private Function ShouldHandle(level As LogLevel) As Boolean
        ShouldHandle = (level >= m_Level)
    End Function
End Class

' 演示
Dim debugLog As IHandler, infoLog As IHandler, errorLog As IHandler
Dim dbgObj As Logger, infoObj As Logger, errObj As Logger
Set dbgObj = New Logger
dbgObj.Name = "控制台"
dbgObj.Level = LogLevel.Debug
Set debugLog = dbgObj

Set infoObj = New Logger
infoObj.Name = "文件"
infoObj.Level = LogLevel.Info
Set infoLog = infoObj

Set errObj = New Logger
errObj.Name = "邮件"
errObj.Level = LogLevel.ErrorLevel
Set errorLog = errObj

debugLog.SetNext(infoLog)

infoLog.SetNext(errorLog)


debugLog.Log "系统启动", LogLevel.Info
debugLog.Log "严重错误", LogLevel.ErrorLevel
```

**Axon VBScript 版妥协说明**：
- `IHandler` 接口约束了链节点的契约。`Logger` 持有 `IHandler` 引用（`m_Next`），在 `IHandler_Log` 中用 `m_Next IsNot Nothing` 检查链尾后直接调用 `m_Next.Log` 转发请求，无需辅助类。`Enum LogLevel` 替代了传统版的字符串比较，`ShouldHandle` 从 14 行数组查找简化为一行整数比较 `level >= m_Level`，既安全又高效。残留限制：缺失语法点：**代码复用机制（继承）**。每个具体处理器都需自行维护 `m_Next` 字段、`SetNext` 方法与链尾转发判断，无法提取到公共基类。Go 用 struct embedding 解决——嵌入一个 `BaseHandler` 结构体即自动获得 `SetNext` 和转发逻辑，只需覆盖 `ShouldHandle`。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有 `MustInherit` 抽象基类（继承复用 `SetNext` + 转发逻辑）+ `MustOverride`（强制子类实现级别判断），可以写出教科书式的责任链——`MustInherit HandlerBase` 封装 `m_Next` 字段、`SetNext` 方法、`Log` 骨架转发逻辑，子类只需 `Overrides ShouldHandle` 即可。

```vbnet
' ① 日志级别枚举（与 Axon 版一致，普通 Enum）
Public Enum LogLevel
    Debug = 0
    Info = 1
    ErrorLevel = 2
End Enum

' ② MustInherit 抽象基类：封装 m_Next + SetNext + Log 转发骨架
Public MustInherit Class HandlerBase
    Protected m_Next As HandlerBase

    ' SetNext 是 Sub，与 Axon 版一致（不返回 Me，不支持链式构建）
    Public Sub SetNext(nextHandler As HandlerBase)
        m_Next = nextHandler
    End Sub

    ' 模板方法骨架：处理当前节点 → 若有下一个则继续转发
    Public Sub Log(msg As String, level As LogLevel)
        If ShouldHandle(level) Then
            Console.WriteLine($"【{Name}】{msg}")
        End If
        If m_Next IsNot Nothing Then
            m_Next.Log(msg, level)
        End If
    End Sub

    ' MustOverride：强制子类实现级别判断
    Protected MustOverride Function ShouldHandle(level As LogLevel) As Boolean

    ' 子类提供处理器名称（用于输出）
    Public MustOverride ReadOnly Property Name As String
End Class

' ③ 具体处理器：与 Axon 版一致的 Logger 类
Public Class Logger
    Inherits HandlerBase

    Private ReadOnly m_Name As String
    Private ReadOnly m_Level As LogLevel

    Public Sub New(name As String, level As LogLevel)
        m_Name = name
        m_Level = level
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return m_Name
        End Get
    End Property

    Protected Overrides Function ShouldHandle(level As LogLevel) As Boolean
        Return level >= m_Level
    End Function
End Class

' 演示：构建 DEBUG→INFO→ERRORLEVEL 的责任链（与 Axon 版一致）
Dim debugLog As New Logger("控制台", LogLevel.Debug)
Dim infoLog As New Logger("文件", LogLevel.Info)
Dim errorLog As New Logger("邮件", LogLevel.ErrorLevel)

debugLog.SetNext(infoLog)
infoLog.SetNext(errorLog)

debugLog.Log("系统启动", LogLevel.Info)
debugLog.Log("严重错误", LogLevel.ErrorLevel)
```

**VB.NET 版说明**：
- **`MustInherit HandlerBase` 继承复用公共逻辑**：`m_Next` 字段、`SetNext` 方法、`Log` 转发骨架在基类写一次，子类自动获得。Axon 版每个 Logger 类都要各自重复写 `Private m_Next`/`IHandler_SetNext`/转发判断三段样板代码。
- **`MustOverride ShouldHandle` 编译期强制**：子类漏写直接编译报错。Axon 版 `ShouldHandle` 是 Private 辅助方法，无法通过接口约束，漏写只会运行时行为异常。
- **`Enum LogLevel` 与 Axon 版一致**：`Enum` 底层是整型，`level >= m_Level` 直接比较，无需像传统版那样用字符串数组查索引。两版在此点无差异。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 级别比较 | 字符串数组查找（14 行 ShouldHandle） | `Enum LogLevel` + 一行整数比较 | `Enum LogLevel` + 一行整数比较 |
| 链节点契约 | 方法名约定（漏写 SetNext 运行时断链） | `IHandler` 接口约束 `SetNext` + `Log` | `MustInherit HandlerBase` + `MustOverride ShouldHandle` 编译期强制 |
| 公共逻辑复用 | 无（每个 Logger 各写一份） | 无（每个 Logger 各写一份） | `MustInherit HandlerBase` 基类写一次，子类继承获得 |
| 类型安全 | 字符串比较，易拼错 | `Enum` 强类型 | `Enum` 强类型 |
---