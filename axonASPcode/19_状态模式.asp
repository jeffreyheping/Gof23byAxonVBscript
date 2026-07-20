<%
' 状态接口
Class IState
    Public Function Handle
    End Function
End Class

' 红灯
Class RedState
    Implements IState
    Public Function IState_Handle
        Response.Write "红灯：停止"
    End Function
End Class

' 绿灯
Class GreenState
    Implements IState
    Public Function IState_Handle
        Response.Write "绿灯：通行"
    End Function
End Class

' 黄灯
Class YellowState
    Implements IState
    Public Function IState_Handle
        Response.Write "黄灯：注意"
    End Function
End Class

' 上下文
Class TrafficLight
    Private m_State As IState

    Private Sub Class_Initialize
        Set m_State = New RedState
    End Sub

    Public Function Change
        If TypeName(m_State) = "RedState" Then
            Set m_State = New GreenState
        ElseIf TypeName(m_State) = "GreenState" Then
            Set m_State = New YellowState
        ElseIf TypeName(m_State) = "YellowState" Then
            Set m_State = New RedState
        End If
    End Function

    Public Function Operate
        m_State.Handle
    End Function
End Class

' 演示
Dim light As TrafficLight
Set light = New TrafficLight
light.Operate
light.Change
light.Operate
light.Change
light.Operate
%>