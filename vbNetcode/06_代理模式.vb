Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch06Module
    Public MustInherit Class Image
        Public MustOverride Function Init(filename As String) As Object
        Public MustOverride Function Display() As Object
    End Class
    Public Class RealImage
        Inherits Image

        Private m_Filename As String

        Public Overrides Function Init(filename As String) As Object
            m_Filename = filename
            Console.WriteLine("【加载大图】" & filename)
        End Function

        Public Overrides Function Display() As Object
            Console.WriteLine("显示图片：" & m_Filename)
        End Function
    End Class
    Public Class ProxyImage
        Inherits Image

        Private m_Filename As String
        Private m_RealImage As Image   ' 基类引用，初始为 Nothing

        Public Overrides Function Init(filename As String) As Object
            m_Filename = filename
        End Function

        Public Overrides Function Display() As Object
            If m_RealImage Is Nothing Then
                m_RealImage = New RealImage()
                m_RealImage.Init(m_Filename)
            End If
            m_RealImage.Display()
        End Function
    End Class
    Sub Main()

        ' ② RealSubject：真实对象

        ' ③ Proxy：延迟加载，通过基类引用持有真实对象

        ' 演示：通过抽象基类引用透明使用代理
        Dim img As Image = New ProxyImage()
        img.Init("photo.jpg")
        Console.WriteLine("代理已创建，真实大图尚未加载")
        img.Display()   ' 此时才触发真实加载
        img.Display()   ' 第二次不再加载
    End Sub
End Module
