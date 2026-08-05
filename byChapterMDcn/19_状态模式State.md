## 第19章 状态模式（State）

**核心思想**：对象行为随内部状态改变而改变。

**示例说明**：TrafficLight 持有 State 引用。红灯时输出"停止"，绿灯时输出"通行"，黄灯时输出"注意"。调用 Change 会切换到下一个状态，状态切换逻辑封装在各状态类的 `NextState` 方法中——每个状态自己知道下一个状态是谁，上下文只需一行委托。

### 传统 VBScript 版

```vbscript
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
```

**传统 VBScript 版妥协说明**：
- **无接口约束**：RedState、GreenState、YellowState 没有 `IState` 接口强制要求 `Handle` 和 `NextState`。如果某个状态类漏写方法，运行时调用才报错。
- **无类型安全**：`m_State` 无类型约束，`m_State.NextState` 的返回值可以是任何对象，编译期无法校验。

### Axon VBScript 版（支持 Implements）

```vbscript
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
        Response.Write "红灯：停止"
    End Function
    Public Function IState_NextState As IState
        Set IState_NextState = New GreenState
    End Function
End Class

' 绿灯
Class GreenState
    Implements IState
    Public Function IState_Handle
        Response.Write "绿灯：通行"
    End Function
    Public Function IState_NextState As IState
        Set IState_NextState = New YellowState
    End Function
End Class

' 黄灯
Class YellowState
    Implements IState
    Public Function IState_Handle
        Response.Write "黄灯：注意"
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
light.Operate   ' 红灯：停止
light.Change
light.Operate   ' 绿灯：通行
light.Change
light.Operate   ' 黄灯：注意
```

**Axon VBScript 版妥协说明**：
- `IState` 接口同时约束了状态类的行为契约（`Handle`）和切换契约（`NextState`），状态切换逻辑下发到各状态类，`TrafficLight.Change` 缩为一行 `Set m_State = m_State.NextState`，符合状态模式"状态自己决定下一个状态"的精髓。新增状态只需新增一个类并修改相邻状态类的 `NextState`，符合开闭原则。
- `TrafficLight` 持有 `Private m_State As IState`，`Change`/`Operate` 中调用 `m_State.NextState`/`m_State.Handle` 即自动路由到具体状态实现，写法与标准 OOP 一致，无需完整限定名。
---