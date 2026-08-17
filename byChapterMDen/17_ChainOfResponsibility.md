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
- `IHandler` interface constrains the chain node contract. `Logger` holds an `IHandler` reference (`m_Next`), uses `m_Next IsNot Nothing` to check chain end in `IHandler_Log` before forwarding via `m_Next.Log` — no helper classes needed. `Enum LogLevel` replaces the classic version's string comparison. `ShouldHandle` simplifies from a 14-line array search to a single integer comparison `level >= m_Level` — both safer and more efficient. Remaining gap: **missing code reuse mechanism (inheritance)**. Each concrete handler must maintain its own `m_Next` field, `SetNext` method, and chain-end forwarding check — can't extract to a common base class. Go uses struct embedding — embed a `BaseHandler` struct to automatically get `SetNext` and forwarding logic, only override `ShouldHandle`.

### VB.NET Version (syntactically complete baseline)

VB.NET has `MustInherit` abstract base class (inheritance reuses `SetNext` + forwarding logic) + `MustOverride` (forces subclass level judgment), enabling textbook Chain of Responsibility — `MustInherit HandlerBase` encapsulates `m_Next` field, `SetNext` method, `Log` skeleton forwarding logic. Subclasses only need to `Overrides ShouldHandle`.

```vbnet
' ① Log level enum (same as Axon version, regular Enum)
Public Enum LogLevel
    Debug = 0
    Info = 1
    ErrorLevel = 2
End Enum

' ② MustInherit abstract base class: encapsulates m_Next + SetNext + Log forwarding skeleton
Public MustInherit Class HandlerBase
    Protected m_Next As HandlerBase

    ' SetNext is a Sub, same as Axon version (doesn't return Me, no chained construction)
    Public Sub SetNext(nextHandler As HandlerBase)
        m_Next = nextHandler
    End Sub

    ' Template method skeleton: handle current node → if next exists, continue forwarding
    Public Sub Log(msg As String, level As LogLevel)
        If ShouldHandle(level) Then
            Console.WriteLine($"[{Name}] {msg}")
        End If
        If m_Next IsNot Nothing Then
            m_Next.Log(msg, level)
        End If
    End Sub

    ' MustOverride: forces subclasses to implement level judgment
    Protected MustOverride Function ShouldHandle(level As LogLevel) As Boolean

    ' Subclasses provide processor name (for output)
    Public MustOverride ReadOnly Property Name As String
End Class

' ③ Concrete handler: same Logger class as Axon version
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

' Demo: build DEBUG→INFO→ERRORLEVEL chain (same as Axon version)
Dim debugLog As New Logger("Console", LogLevel.Debug)
Dim infoLog As New Logger("File", LogLevel.Info)
Dim errorLog As New Logger("Email", LogLevel.ErrorLevel)

debugLog.SetNext(infoLog)
infoLog.SetNext(errorLog)

debugLog.Log("System started", LogLevel.Info)
debugLog.Log("Critical error", LogLevel.ErrorLevel)
```

**VB.NET version notes**:
- **`MustInherit HandlerBase` inheritance reuses common logic**: `m_Next` field, `SetNext` method, `Log` forwarding skeleton written once in base class, subclasses automatically get them. Axon version every Logger class must duplicate `Private m_Next`/`IHandler_SetNext`/forwarding check — three blocks of boilerplate.
- **`MustOverride ShouldHandle` compile-time enforced**: Subclass missing it causes compile error. Axon version `ShouldHandle` is a Private helper method, can't be constrained by interface, missing it only causes runtime behavioral issues.
- **`Enum LogLevel` same as Axon version**: `Enum` underlying type is integer, `level >= m_Level` direct comparison, no need for string array index lookup like classic version. No difference between these two versions on this point.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Level comparison | String array lookup (14-line ShouldHandle) | `Enum LogLevel` + single integer comparison | `Enum LogLevel` + single integer comparison |
| Chain node contract | Method name convention (missing SetNext = broken chain at runtime) | `IHandler` interface constrains `SetNext` + `Log` | `MustInherit HandlerBase` + `MustOverride ShouldHandle` compile-time enforced |
| Common logic reuse | None (each Logger writes their own) | None (each Logger writes their own) | `MustInherit HandlerBase` base class writes once, subclasses inherit |
| Type safety | String comparison, easy to misspell | `Enum` strongly typed | `Enum` strongly typed |
---

---
