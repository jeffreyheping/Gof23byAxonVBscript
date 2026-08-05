Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
' 状态：红灯
Class RedState
    ' 当前状态的行为
    Public Function Handle
        Response.Write "红灯：停止"
    End Function

    ' 切换到下一个状态：红灯自己知道下一个是绿灯
    Public Function NextState
        Set NextState = New GreenState
    End Function
End Class

' 状态：绿灯
Class GreenState
    Public Function Handle
        Response.Write "绿灯：通行"
    End Function

    ' 切换到下一个状态：绿灯自己知道下一个是黄灯
    Public Function NextState
        Set NextState = New YellowState
    End Function
End Class

' 状态：黄灯
Class YellowState
    Public Function Handle
        Response.Write "黄灯：注意"
    End Function

    ' 切换到下一个状态：黄灯自己知道下一个是红灯（循环）
    Public Function NextState
        Set NextState = New RedState
    End Function
End Class

' 上下文：持有当前状态，委托给状态类处理行为和切换
Class TrafficLight
    Private m_State   ' 当前状态对象

    ' 构造函数：初始为红灯
    Private Sub Class_Initialize
        Set m_State = New RedState
    End Sub

    ' 切换状态：委托给当前状态类，无需 TypeName 判断
    Public Function Change
        Set m_State = m_State.NextState
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

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
