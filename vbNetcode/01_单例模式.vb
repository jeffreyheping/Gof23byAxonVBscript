Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch01Module
    Public Class Singleton
        ' ① Private 构造函数：外部彻底无法 New，只有类自己能创建
        Private Sub New()
            m_Data = "我是唯一实例"
        End Sub

        ' ② Shared ReadOnly 字段 + 初始化器：
        '    CLR 保证首次访问类型时初始化、且只初始化一次（线程安全）
        Private Shared ReadOnly m_Instance As New Singleton()

        Private m_Data As String

        ' ③ Shared 访问点：类级别成员，无需函数包装
        Public Shared ReadOnly Property Instance As Singleton
            Get
                Return m_Instance
            End Get
        End Property

        Public Property Data As String
            Get
                Return m_Data
            End Get
            Set(value As String)
                m_Data = value
            End Set
        End Property
    End Class
    Sub Main()

        ' 演示
        Dim s1 As Singleton = Singleton.Instance
        Dim s2 As Singleton = Singleton.Instance
        s1.Data = "已修改"
        Console.WriteLine(s2.Data)   ' 已修改（同一个对象）
    End Sub
End Module
