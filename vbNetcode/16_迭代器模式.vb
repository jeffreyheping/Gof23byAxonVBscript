Option Strict Off
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch16Module
    Public Class MyCollection(Of T)
        Implements IEnumerable(Of T)

        Private ReadOnly m_Items As New List(Of T)()

        ' 添加元素
        Public Function Add(item As T)
            m_Items.Add(item)
        End Function

        ' 返回元素总数
        Public ReadOnly Property Count As Integer
            Get
                Return m_Items.Count
            End Get
        End Property

        ' 索引器
            Default Public ReadOnly Property Item(index As Integer) As T

            Get
                Return m_Items(index)
            End Get
        End Property

        ' ② 泛型枚举器：Yield 自动生成 IEnumerator(Of T) 状态机
        Public Iterator Function GetEnumerator() As IEnumerator(Of T) _
            Implements IEnumerable(Of T).GetEnumerator
            For Each element As T In m_Items
                Yield element   ' 编译器自动生成 MoveNext/Current 状态机
            Next
        End Function

        ' ③ 非泛型枚举器：IEnumerable(Of T) 继承自 IEnumerable，必须实现
        Private Function GetEnumeratorObj() As IEnumerator _
            Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function
    End Class
    Sub Main()

        ' 演示：For Each 遍历，与 Axon 版调用方式一致
        Dim coll As New MyCollection(Of String)()
        coll.Add("苹果")
        coll.Add("香蕉")
        coll.Add("橙子")

        For Each item In coll
            Console.WriteLine(item)
        Next
    End Sub
End Module
