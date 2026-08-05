<%
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
%>