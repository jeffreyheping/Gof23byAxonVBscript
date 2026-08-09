Option Strict Off
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch05Module
    Public Class MyResume
        Implements ICloneable

        Public Property Name As String
        Public Property Age As Integer
        Public Property Skills As String()   ' 与 Axon 版一致，用数组

        ' 标准 ICloneable 接口方法：手动深拷贝
        Public Function Clone() As Object Implements ICloneable.Clone
            Dim copy As New MyResume() With {
                .Name = Me.Name,
                .Age = Me.Age
            }
            ' 数组深拷贝：逐元素复制（与 Axon 版逻辑一致）
            If Me.Skills IsNot Nothing Then
                copy.Skills = New String(Me.Skills.Length - 1) {}
                Array.Copy(Me.Skills, copy.Skills, Me.Skills.Length)
            End If
            Return copy
        End Function
    End Class
    Sub Main()

        ' 演示：克隆后修改副本，原件不受影响
        Dim original As New MyResume() With {
            .Name = "张三",
            .Age = 25,
            .Skills = {"VBScript", "HTML"}
        }

        Dim clone As MyResume = DirectCast(original.Clone(), MyResume)
        clone.Name = "李四"
        clone.Skills(0) = "JavaScript"

        Console.WriteLine(original.Name & " " & original.Skills(0))   ' 张三 VBScript
        Console.WriteLine(clone.Name & " " & clone.Skills(0))         ' 李四 JavaScript
    End Sub
End Module
