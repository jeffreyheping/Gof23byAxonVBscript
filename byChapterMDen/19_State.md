## Chapter 19: State

**Core idea**: An object's behavior changes along with its internal state.

**Example**: TrafficLight holds a State reference. Red light outputs "stop", green light outputs "go", yellow light outputs "caution". Calling `Change` switches to the next state. The transition logic is encapsulated in each state class's `NextState` method — each state knows its own successor, and the context only needs one line of delegation.

### Classic VBScript Version

```vbscript
' State: red light
Class RedState
    ' Current state's behavior
    Public Function Handle
        Response.Write "Red light: stop"
    End Function

    ' Switch to next state: red knows next is green
    Public Function NextState
        Set NextState = New GreenState
    End Function
End Class

' State: green light
Class GreenState
    Public Function Handle
        Response.Write "Green light: go"
    End Function

    ' Switch to next state: green knows next is yellow
    Public Function NextState
        Set NextState = New YellowState
    End Function
End Class

' State: yellow light
Class YellowState
    Public Function Handle
        Response.Write "Yellow light: caution"
    End Function

    ' Switch to next state: yellow knows next is red (cycle)
    Public Function NextState
        Set NextState = New RedState
    End Function
End Class

' Context: holds current state, delegates behavior and transition to state classes
Class TrafficLight
    Private m_State   ' Current state object

    ' Constructor: starts with red
    Private Sub Class_Initialize
        Set m_State = New RedState
    End Sub

    ' Switch state: delegate to current state, no TypeName check needed
    Public Function Change
        Set m_State = m_State.NextState
    End Function

    ' Execute current state's behavior
    Public Function Operate
        m_State.Handle
    End Function
End Class

' Demo: behavior changes automatically with state transitions
Dim light
Set light = New TrafficLight
light.Operate   ' Red light: stop
light.Change
light.Operate   ' Green light: go
light.Change
light.Operate   ' Yellow light: caution
```

**Classic VBScript trade-offs**:
- **No interface constraint**: RedState, GreenState, YellowState have no `IState` interface forcing `Handle` and `NextState`. If a state class forgets a method, the error only surfaces at runtime.
- **No type safety**: `m_State` has no type constraint. `m_State.NextState` can return any object — the compiler can't verify it.

### Axon VBScript Version (supports Implements)

```vbscript
' State interface: behavior + transition
Class IState
    Public Function Handle
    End Function
    Public Function NextState As IState
    End Function
End Class

' Red light
Class RedState
    Implements IState
    Public Function IState_Handle
        Response.Write "Red light: stop"
    End Function
    Public Function IState_NextState As IState
        Set IState_NextState = New GreenState
    End Function
End Class

' Green light
Class GreenState
    Implements IState
    Public Function IState_Handle
        Response.Write "Green light: go"
    End Function
    Public Function IState_NextState As IState
        Set IState_NextState = New YellowState
    End Function
End Class

' Yellow light
Class YellowState
    Implements IState
    Public Function IState_Handle
        Response.Write "Yellow light: caution"
    End Function
    Public Function IState_NextState As IState
        Set IState_NextState = New RedState
    End Function
End Class

' Context: holds current state via interface reference
Class TrafficLight
    Private m_State As IState

    Private Sub Class_Initialize
        Set m_State = New RedState
    End Sub

    ' Switch state: one-line delegation, compile-time type-safe
    Public Function Change
        Set m_State = m_State.NextState
    End Function

    ' Execute current state's behavior
    Public Function Operate
        m_State.Handle
    End Function
End Class

' Demo
Dim light As TrafficLight
Set light = New TrafficLight
light.Operate   ' Red light: stop
light.Change
light.Operate   ' Green light: go
light.Change
light.Operate   ' Yellow light: caution
```

**Axon VBScript trade-offs**:
- `IState` constrains both the behavior contract (`Handle`) and the transition contract (`NextState`) for state classes. Transition logic is pushed down to each state class. `TrafficLight.Change` shrinks to one line: `Set m_State = m_State.NextState` — capturing the State pattern's essence: "each state decides its own successor." Adding a new state only requires a new class and updating the adjacent state's `NextState` — consistent with the Open-Closed Principle.
- `TrafficLight` holds `Private m_State As IState`. `Change`/`Operate` call `m_State.NextState`/`m_State.Handle` which auto-dispatches to the concrete state. Consistent with standard OOP — no fully-qualified names needed.
---
