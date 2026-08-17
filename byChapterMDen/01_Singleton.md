## Chapter 1: Singleton

**Core idea**: Only one instance exists globally.

**Example**: A script-level variable holds the sole instance, and `GetInstance` controls creation. Both calls return the same object — changing one changes the other.

### Classic VBScript Version

```vbscript
' Script-level variable: the single global instance reference
Dim gInstance
Set gInstance = Nothing

Class Singleton
    Private m_Data

    ' Constructor: set default data
    Private Sub Class_Initialize
        m_Data = "I am the only instance"
    End Sub

    ' Read internal data
    Public Property Get Data
        Data = m_Data
    End Property

    ' Write internal data
    Public Property Let Data(value)
        m_Data = value
    End Property
End Class

' Global access point: create if absent, otherwise return existing
' Returns: the one and only Singleton instance
Function GetInstance()
    If gInstance Is Nothing Then
        Set gInstance = New Singleton
    End If
    Set GetInstance = gInstance
End Function

' Demo: both calls get the same object
Dim s1, s2
Set s1 = GetInstance()
Set s2 = GetInstance()
s1.Data = "Modified"
Response.Write s2.Data   ' Modified (same object)
```

**Classic VBScript trade-offs**:
- **No static variables**: VBScript classes don't support `Static`, so we must use a script-level (module-level) variable `gInstance` to hold the instance, breaking encapsulation.
- **Cannot prevent external `New`**: VBScript has no private constructor. External code can always `New Singleton` and bypass `GetInstance` — true singleton enforcement is impossible.

### Axon VBScript Version (supports Static)

```vbscript
Class Singleton
    Private m_Data As String

    Private Sub Class_Initialize
        m_Data = "I am the only instance"
    End Sub

    Public Property Get Data As String
        Data = m_Data
    End Property

    Public Property Let Data(value As String)
        m_Data = value
    End Property
End Class

' Global access point: Static variable persists across calls, supports object references
Function GetInstance() As Singleton
    Static instance As Singleton
    If instance Is Nothing Then
        Set instance = New Singleton
    End If
    Set GetInstance = instance
End Function

' Demo: guarantees the same instance
Dim s1 As Singleton, s2 As Singleton
Set s1 = GetInstance()
Set s2 = GetInstance()
s1.Data = "Modified"
Response.Write s2.Data   ' Modified
```

**Axon VBScript trade-offs**:
- AxonASP's `Static` variable supports object references, so the sole instance can live inside the function — no module-level global needed. Encapsulation is better than the classic version. Remaining gap: **Private constructor** is still missing. External code can still `New Singleton` and bypass singleton control. Go solves this with **package-private lowercase names** — `type singleton` is unexported, so external code can only get an instance via `GetInstance()`. VBScript classes have no access control, so external `New` cannot be prevented.

### VB.NET Version (syntactically complete baseline)

VB.NET is a syntactically complete OOP language with `Private` constructors, `Shared` (static) members, `ReadOnly` fields, and full `Property` syntax. The following implementation needs no compromises — **this is the textbook Singleton pattern**.

```vbnet
Public Class Singleton
    ' ① Private constructor: external code absolutely cannot New — only the class itself can create
    Private Sub New()
        m_Data = "I am the sole instance"
    End Sub

    ' ② Shared ReadOnly field + initializer:
    '    CLR guarantees initialization on first type access, and only once (thread-safe)
    Private Shared ReadOnly m_Instance As New Singleton()

    Private m_Data As String

    ' ③ Shared access point: class-level member, no function wrapper needed
    Public Shared ReadOnly Property Instance As Singleton
        Get
            Return m_Instance
        End Get
    End Property

    Public Property Data As String
        Get
            Return m_Data
        End Get
        Set(value As String)
            m_Data = value
        End Set
    End Property
End Class

' Demo: guarantees the same instance
Dim s1 As Singleton = Singleton.Instance
Dim s2 As Singleton = Singleton.Instance
s1.Data = "Modified"
Console.WriteLine(s2.Data)   ' Modified (same object)
```

**VB.NET version notes**:
- **True prevention of external New**: `Private Sub New()` makes the constructor invisible externally. Writing `New Singleton()` in external code causes a compile error — the "no external creation" constraint is enforced at compile time, not by developer discipline.
- **No global variables, no function wrappers**: `Shared` members belong to the class itself, not instances. Access directly via `Singleton.Instance`. Compare with the first two versions: classic needs `gInstance` module-level global + `GetInstance()` function; Axon needs `Static instance` + `GetInstance()` function.
- **Thread safety guaranteed by CLR**: `Shared ReadOnly m_Instance As New Singleton()` relies on CLR's static field initializer — CLR automatically locks during initialization on first type access, guaranteeing single initialization. Neither classic nor Axon versions are thread-safe (classic has no locking; Axon's `If instance Is Nothing` check has a race condition).
- **Only remaining note**: This is eager initialization (initializes on first type access). For true lazy initialization (on first `Instance` call), VB.NET also has `Lazy(Of Singleton)`, but that's an optimization concern, not a pattern idiomaticity issue.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Sole instance storage | Module-level global variable | Function-level `Static` | Class-level `Shared ReadOnly` |
| Access method | `GetInstance()` function | `GetInstance()` function | `Singleton.Instance` property |
| Prevent external `New` | Cannot | Cannot | `Private Sub New()` compile-time enforced |
| Encapsulation | Poor (global variable leak) | Medium (encapsulated in function) | Good (encapsulated in class) |
| Thread safety | No | No | Yes (CLR guaranteed) |
---
