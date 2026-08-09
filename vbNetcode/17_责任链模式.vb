Option Strict Off
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch17Module
    Public Enum LogLevel
        Debug = 0
        Info = 1
        ErrorLevel = 2
    End Enum
    Public MustInherit Class HandlerBase
        Protected m_Next As HandlerBase

        ' SetNext 是 Sub，与 Axon 版一致（不返回 Me，不支持链式构建）
        Public Function SetNext(nextHandler As HandlerBase)
            m_Next = nextHandler
        End Function

        ' 模板方法骨架：处理当前节点 → 若有下一个则继续转发
        Public Function Log(msg As String, level As LogLevel)
            If ShouldHandle(level) Then
                Console.WriteLine($"【{Name}】{msg}")
            End If
            If m_Next IsNot Nothing Then
                m_Next.Log(msg, level)
            End If
        End Function

        ' MustOverride：强制子类实现级别判断
        Protected MustOverride Function ShouldHandle(level As LogLevel) As Boolean

        ' 子类提供处理器名称（用于输出）
        Public MustOverride ReadOnly Property Name As String
    End Class
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
    Sub Main()

        ' ② MustInherit 抽象基类：封装 m_Next + SetNext + Log 转发骨架

        ' ③ 具体处理器：与 Axon 版一致的 Logger 类

        ' 演示：构建 DEBUG→INFO→ERRORLEVEL 的责任链（与 Axon 版一致）
        Dim debugLog As New Logger("控制台", LogLevel.Debug)
        Dim infoLog As New Logger("文件", LogLevel.Info)
        Dim errorLog As New Logger("邮件", LogLevel.ErrorLevel)

        debugLog.SetNext(infoLog)
        infoLog.SetNext(errorLog)

        debugLog.Log("系统启动", LogLevel.Info)
        debugLog.Log("严重错误", LogLevel.ErrorLevel)
    End Sub
End Module
