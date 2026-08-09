Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch14Module
    Public Class NewsEventArgs
        Inherits EventArgs
        Public ReadOnly Property News As String

        Public Sub New(news As String)
            News = news
        End Sub
    End Class
    Public Class NewsAgency
        ' 声明事件：EventHandler(Of NewsEventArgs) 是 .NET 标准泛型委托
        Public Event NewsPublished As EventHandler(Of NewsEventArgs)

        ' 发布新闻：触发事件，所有订阅者自动收到通知
        Public Function Publish(news As String) As Object
            RaiseEvent NewsPublished(Me, New NewsEventArgs(news))
        End Function
    End Class
    Public Class Newspaper
        Private ReadOnly m_Name As String
        Private WithEvents m_Agency As NewsAgency

        Public Sub New(name As String, agency As NewsAgency)
            m_Name = name
            m_Agency = agency
        End Sub

        ' Handles 关键字：声明式绑定 m_Agency 的 NewsPublished 事件
        Private Sub OnNewsPublished(sender As Object, e As NewsEventArgs) _
            Handles m_Agency.NewsPublished
            Console.WriteLine($"{m_Name} 收到新闻：{e.News}")
        End Sub
    End Class
    Sub Main()

        ' ② 被观察者：用普通 Event + EventHandler(Of T) 声明强类型事件

        ' ③ 观察者：通过 WithEvents + Handles 声明式订阅

        ' 演示：与 Axon 版新闻场景一致
        Dim agency As New NewsAgency()
        Dim paper1 As New Newspaper("晨报", agency)
        Dim paper2 As New Newspaper("晚报", agency)
        agency.Publish("重大新闻！")
    End Sub
End Module
