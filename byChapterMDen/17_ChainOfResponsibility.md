## Chapter 17: Chain of Responsibility

**Core idea**: Pass a request along a chain until some handler processes it.

**Example**: Loggers form a DEBUG → INFO → ERRORLEVEL chain. When `Log` is called, if the current level matches it handles and continues passing; otherwise it just passes through. An ERRORLEVEL message ends up being output by all three loggers.

### Classic VBScript Version

```vbscript
' Log handler: forms a chain of responsibility
Class Logger
    Public Name
    Public Level        ' Lowest level this handler can process (DEBUG/INFO/ERRORLEVEL)
    Private m_Next      ' Next handler in the chain

    Private Sub Class_Initialize
        Set m_Next = Nothing
    End Sub

    ' Set next handler
    Public Function SetNext(nextHandler)
        Set m_Next = nextHandler
    End Function

    ' Handle log request: output if level matches, then pass to next
    Public Function Log(msg, level)
        If ShouldHandle(level) Then
            Response.Write "[" & Name & "] " & msg
        End If
        If Not m_Next Is Nothing Then
            m_Next.Log msg, level
        End If
    End Function

    ' Check if current level should handle this
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

' Demo: build a DEBUG → INFO → ERRORLEVEL chain
Dim debugLog, infoLog, errorLog
Set debugLog = New Logger
debugLog.Name = "Console"
debugLog.Level = "DEBUG"

Set infoLog = New Logger
infoLog.Name = "File"
infoLog.Level = "INFO"

Set errorLog = New Logger
errorLog.Name = "Email"
errorLog.Level = "ERRORLEVEL"

debugLog.SetNext infoLog
infoLog.SetNext errorLog

debugLog.Log "System started", "INFO"          ' File and Email both output
debugLog.Log "Critical error", "ERRORLEVEL"    ' All three output
```

**Classic VBScript trade-offs**:
- **No interface constraint**: Logger relies on `SetNext` and `Log` method-name conventions to form the chain. No `IHandler` interface enforces these two methods. If a class forgets `SetNext`, the chain breaks — only discovered at runtime.

### Axon VBScript Version (supports Implements + Enum)

```vbscript
' Log level enum
Enum LogLevel
    Debug = 0
    Info = 1
    ErrorLevel = 2
End Enum

' Handler interface
Class IHandler
    Public Function SetNext(handler As IHandler)
    End Function
    Public Function Log(msg As String, level As LogLevel)
    End Function
End Class

' Concrete handler
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
            Response.Write "[" & m_Name & "] " & msg
        End If
        If m_Next IsNot Nothing Then
            m_Next.Log msg, level
        End If
    End Function

    Private Function ShouldHandle(level As LogLevel) As Boolean
        ShouldHandle = (level >= m_Level)
    End Function
End Class

' Demo
Dim debugLog As IHandler, infoLog As IHandler, errorLog As IHandler
Dim dbgObj As Logger, infoObj As Logger, errObj As Logger
Set dbgObj = New Logger
dbgObj.Name = "Console"
dbgObj.Level = LogLevel.Debug
Set debugLog = dbgObj

Set infoObj = New Logger
infoObj.Name = "File"
infoObj.Level = LogLevel.Info
Set infoLog = infoObj

Set errObj = New Logger
errObj.Name = "Email"
errObj.Level = LogLevel.ErrorLevel
Set errorLog = errObj

debugLog.SetNext infoLog
infoLog.SetNext errorLog

debugLog.Log "System started", LogLevel.Info
debugLog.Log "Critical error", LogLevel.ErrorLevel
```


**Axon VBScript trade-offs**:
- `IHandler` enforces the chain node contract. `Logger` holds an `IHandler` reference (`m_Next`) and uses `m_Next IsNot Nothing` to check for chain end before forwarding via `m_Next.Log` — no helper classes needed. `Enum LogLevel` replaces the classic version's string comparison. `ShouldHandle` simplifies from a 14-line array search to a single integer comparison `level >= m_Level` — both safer and faster. Remaining gap: **Code reuse mechanism** (inheritance or struct embedding). Each concrete handler must maintain its own `m_Next` field and forwarding logic — can't extract to a common base class. Go uses struct embedding — embed a `BaseHandler` struct to automatically get `SetNext` and forwarding logic, only override `ShouldHandle`. Also, chain construction (`SetNext` call order) is the caller's responsibility — the compiler can't verify the chain is complete.
---

---
