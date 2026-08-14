Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch19Module
    Public MustInherit Class StateBase
        Public MustOverride Sub Handle()
        Public MustOverride Function NextState() As StateBase
    End Class
    Public Class RedState
        Inherits StateBase

        Public Overrides Sub Handle()
            Console.WriteLine("红灯：停止")
        End Sub

        Public Overrides Function NextState() As StateBase
            Return New GreenState()
        End Function
    End Class
    Public Class GreenState
        Inherits StateBase

        Public Overrides Sub Handle()
            Console.WriteLine("绿灯：通行")
        End Sub

        Public Overrides Function NextState() As StateBase
            Return New YellowState()
        End Function
    End Class
    Public Class YellowState
        Inherits StateBase

        Public Overrides Sub Handle()
            Console.WriteLine("黄灯：注意")
        End Sub

        Public Overrides Function NextState() As StateBase
            Return New RedState()
        End Function
    End Class
    Public Class TrafficLight
        Private m_State As StateBase

        Public Sub New()
            m_State = New RedState()
        End Sub

        Public Sub Change()
            m_State = m_State.NextState()
        End Sub

        Public Sub Operate()
            m_State.Handle()
        End Sub
    End Class
    Sub Main()

        ' 红灯

        ' 绿灯

        ' 黄灯

        ' 上下文：持有当前状态，委托给状态类处理行为和切换

        ' 演示
        Dim light As New TrafficLight()
        light.Operate()   ' 红灯：停止
        light.Change()
        light.Operate()   ' 绿灯：通行
        light.Change()
        light.Operate()   ' 黄灯：注意
    End Sub
End Module
