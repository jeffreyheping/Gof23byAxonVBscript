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
    Public Function Log(msg, level)
        If ShouldHandle(level) Then
            Response.Write "【" & Name & "】" & msg
        End If
        If Not m_Next Is Nothing Then
            m_Next.Log msg, level
        End If
    End Function

    ' 判断当前级别是否应该处理
    Private Function ShouldHandle(level)
        Dim levels(2)
        levels(0) = "DEBUG"
        levels(1) = "INFO"
        levels(2) = "ERRORLEVEL"

        Dim currentIdx, msgIdx, i
        currentIdx = -1
        msgIdx = -1
        For i = 0 To 2
            If levels(i) = Level Then currentIdx = i
            If levels(i) = level Then msgIdx = i
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

debugLog.SetNext infoLog
infoLog.SetNext errorLog

debugLog.Log "系统启动", "INFO"     ' 文件和邮件都输出
debugLog.Log "严重错误", "ERRORLEVEL"    ' 三个都输出
```

**传统 VBScript 版妥协说明**：
- **无接口约束**：Logger 靠 `SetNext` 和 `Log` 方法名约定形成链条，没有 `IHandler` 接口强制要求这两个方法。如果某个类漏写 `SetNext`，链条就断了，运行时才发现。

### Axon VBScript 版（支持 Implements + Enum）

```vbscript
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
            Response.Write "【" & m_Name & "】" & msg
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

debugLog.SetNext infoLog
infoLog.SetNext errorLog

debugLog.Log "系统启动", LogLevel.Info
debugLog.Log "严重错误", LogLevel.ErrorLevel
```


**Axon VBScript 版妥协说明**：
- `IHandler` 接口约束了链节点的契约。`Logger` 持有 `IHandler` 引用（`m_Next`），在 `IHandler_Log` 中用 `m_Next IsNot Nothing` 检查链尾后直接调用 `m_Next.Log` 转发请求，无需辅助类。`Enum LogLevel` 替代了传统版的字符串比较，`ShouldHandle` 从 14 行数组查找简化为一行整数比较 `level >= m_Level`，既安全又高效。残留限制：缺失语法点：**代码复用机制**（继承或 struct embedding）。每个具体处理器都需自行维护 `m_Next` 字段与转发判断，无法提取到公共基类。Go 用 struct embedding 解决——嵌入一个 `BaseHandler` 结构体即自动获得 `SetNext` 和转发逻辑，只需覆盖 `ShouldHandle`。此外，链条的构建（`SetNext` 调用顺序）需调用方保证，编译期无法校验链是否闭合。
---

---