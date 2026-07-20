## 第19章 状态模式（State）

**核心思想**：对象行为随内部状态改变而改变。

**示例说明**：TrafficLight 持有 State 引用。红灯时 Stop，绿灯时 Go，黄灯时 Caution。调用 Change 会切换到下一个状态，状态切换逻辑封装在各自的状态类中。

### 传统 VBScript 版

```vbscript
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
```

**传统 VBScript 版妥协说明**：
- **无接口约束**：RedState、GreenState、YellowState 没有 `IState` 接口强制要求 `Handle`。如果某个状态类漏写方法，运行时调用才报错。
- **TypeName 判断不够优雅**：上下文用 `TypeName(m_State)` 判断当前状态类型来切换，是字符串匹配，编译期无法检查。如果类名拼写错误，切换会失败。

### Axon VBScript 版（支持 Implements）

```vbscript
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
```

**Axon VBScript 版妥协说明**：
- `IState` 接口约束了状态类的契约，状态切换逻辑封装在各状态类的 `IState_NextState` 实现中。AxonASP 接口方法派发已修复，`TrafficLight` 持有 `Private m_State As IState`，`Change`/`Operate` 中通过 `m_State.NextState()`/`m_State.Handle` 即自动路由到具体状态的 `IState_NextState`/`IState_Handle`，写法与标准 OOP 一致，无需完整限定名。
---