<%
Option Explicit
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
%>