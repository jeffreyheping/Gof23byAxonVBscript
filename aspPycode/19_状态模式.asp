<%
' 状态：红灯
Class RedState
    ' 当前状态的行为
    Public Function Handle
        Response.Write "红灯：停止"
    End Function
End Class

' 状态：绿灯
Class GreenState
    Public Function Handle
        Response.Write "绿灯：通行"
    End Function
End Class

' 状态：黄灯
Class YellowState
    Public Function Handle
        Response.Write "黄灯：注意"
    End Function
End Class

' 上下文：持有当前状态，集中管理状态切换
Class TrafficLight
    Private m_State   ' 当前状态对象

    ' 构造函数：初始为红灯
    Private Sub Class_Initialize
        Set m_State = New RedState
    End Sub

    ' 集中管理状态切换逻辑（替代分散在各状态类中的 NextState）
    Public Function Change
        If TypeName(m_State) = "RedState" Then
            Set m_State = New GreenState
        ElseIf TypeName(m_State) = "GreenState" Then
            Set m_State = New YellowState
        ElseIf TypeName(m_State) = "YellowState" Then
            Set m_State = New RedState
        End If
    End Function

    ' 执行当前状态的行为
    Public Function Operate
        m_State.Handle
    End Function
End Class

' 演示：状态切换自动改变行为
Dim light
Set light = New TrafficLight
light.Operate   ' 红灯：停止
light.Change
light.Operate   ' 绿灯：通行
light.Change
light.Operate   ' 黄灯：注意
%>