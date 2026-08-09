Option Strict Off
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch12Module
    Public Class TreeType
        Public ReadOnly Property Name As String
        Public ReadOnly Property Color As String

        Public Sub New(name As String, color As String)
            Me.Name = name
            Me.Color = color
        End Sub

        ' 在指定坐标绘制树（x, y 是外部状态，由调用方传入）
        Public Function Draw(x As Long, y As Long)
            Console.WriteLine($"在 ({x},{y}) 绘制 {Color}{Name}")
        End Function
    End Class
    Public Class TreeFactory
        ' 泛型字典：Key 一定是 String，Value 一定是 TreeType
        Private ReadOnly m_Types As New Dictionary(Of String, TreeType)()

        ' 获取或创建 TreeType：相同参数返回同一个对象
        Public Function GetTreeType(name As String, color As String) As TreeType
            Dim key = name & "|" & color
            If Not m_Types.ContainsKey(key) Then
                m_Types(key) = New TreeType(name, color)
            End If
            Return m_Types(key)
        End Function
    End Class
    Sub Main()

        ' ② 享元工厂：缓存并复用 TreeType

        ' 演示：3 棵树共享同一个 TreeType 对象
        Dim factory As New TreeFactory()
        Dim oakType = factory.GetTreeType("橡树", "绿色")

        For i = 0 To 2
            oakType.Draw(i, i * 2)
        Next
        Console.WriteLine("3 棵树，实际只有 1 个 TreeType 对象")
        ' 在 (0,0) 绘制 绿色橡树
        ' 在 (1,2) 绘制 绿色橡树
        ' 在 (2,4) 绘制 绿色橡树
    End Sub
End Module
