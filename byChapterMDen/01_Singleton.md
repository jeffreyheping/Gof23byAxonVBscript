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
---
