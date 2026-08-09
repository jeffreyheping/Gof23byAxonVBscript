<%
' 状态接口：行为 + 切换
Class IState
    Public Function Handle
    End Function
    Public Function NextState As IState
    End Function
End Class

' 红灯
Class RedState
    Implements IState
    Public Function IState_Handle
        Response.Write("红灯：停止")

    End Function
    Public Function IState_NextState As IState
        Set IState_NextState = New GreenState
    End Function
End Class

' 绿灯
Class GreenState
    Implements IState
    Public Function IState_Handle
        Response.Write("绿灯：通行")

    End Function
    Public Function IState_NextState As IState
        Set IState_NextState = New YellowState
    End Function
End Class

' 黄灯
Class YellowState
    Implements IState
    Public Function IState_Handle
        Response.Write("黄灯：注意")

    End Function
    Public Function IState_NextState As IState
        Set IState_NextState = New RedState
    End Function
End Class

' 上下文：通过接口引用持有当前状态
Class TrafficLight
    Private m_State As IState

    Private Sub Class_Initialize
        Set m_State = New RedState
    End Sub

    ' 切换状态：一行委托，编译期类型安全
    Public Function Change
        Set m_State = m_State.NextState
    End Function

    ' 执行当前状态的行为
    Public Function Operate
        m_State.Handle
    End Function
End Class

' 演示
Dim light As TrafficLight
Set light = New TrafficLight
light.Operate()   ' 红灯：停止

light.Change
light.Operate()   ' 绿灯：通行

light.Change
light.Operate()   ' 黄灯：注意
%>