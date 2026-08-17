## Chapter 7: Facade

**Core idea**: Provide a single entry point for a complex subsystem.

**Example**: Booting up involves CPU freeze → hard drive read → memory load → CPU jump → CPU execute. ComputerFacade wraps all these steps into one `Start` method. The caller only needs one call.

### Classic VBScript Version

```vbscript
' Subsystem: CPU
Class CPU
    ' Freeze current state
    Public Function Freeze
        Response.Write "CPU freeze"
    End Function
    ' Jump to address
    Public Function Jump(pos)
        Response.Write "CPU jump to " & pos
    End Function
    ' Start executing
    Public Function Execute
        Response.Write "CPU execute"
    End Function
End Class

' Subsystem: Memory
Class Memory
    ' Load data to address
    Public Function Load(pos, data)
        Response.Write "Memory load " & data & " to " & pos
    End Function
End Class

' Subsystem: Hard Drive
Class HardDrive
    ' Read data from sector
    ' Returns: simulated data block string
    Public Function Read(lba)
        Read = "DATA(" & lba & ")"
    End Function
End Class

' Facade: wraps complex subsystem calls, exposes only Start
Class ComputerFacade
    Private m_CPU, m_Mem, m_HD

    ' Constructor: initialize all subsystems
    Private Sub Class_Initialize
        Set m_CPU = New CPU
        Set m_Mem = New Memory
        Set m_HD = New HardDrive
    End Sub

    ' One-call boot: internally calls each subsystem in order
    Public Function Start
        m_CPU.Freeze
        Dim bootData
        bootData = m_HD.Read(0)
        m_Mem.Load 0, bootData
        m_CPU.Jump 0
        m_CPU.Execute
    End Function
End Class

' Demo: caller only needs Start, no need to know the internals
Dim pc
Set pc = New ComputerFacade
pc.Start   ' Only Start is exposed
```

**Classic VBScript trade-offs**:
This pattern maps naturally to VBScript. The facade just composes subsystem calls — no inheritance or interfaces needed. No significant trade-offs.

### Axon VBScript Version (supports strong typing)

```vbscript
' Subsystem: CPU
Class CPU
    ' Freeze current state
    Public Function Freeze
        Response.Write "CPU freeze"
    End Function
    ' Jump to address
    Public Function Jump(pos As Long)
        Response.Write "CPU jump to " & pos
    End Function
    ' Start executing
    Public Function Execute
        Response.Write "CPU execute"
    End Function
End Class

' Subsystem: Memory
Class Memory
    ' Load data to address
    Public Function Load(pos As Long, data As String)
        Response.Write "Memory load " & data & " to " & pos
    End Function
End Class

' Subsystem: Hard Drive
Class HardDrive
    ' Read data from sector
    ' Returns: simulated data block string
    Public Function Read(lba As Long) As String
        Read = "DATA(" & lba & ")"
    End Function
End Class

' Facade: wraps complex subsystem calls, exposes only Start
Class ComputerFacade
    Private m_CPU As CPU
    Private m_Mem As Memory
    Private m_HD As HardDrive

    ' Constructor: initialize all subsystems
    Private Sub Class_Initialize
        Set m_CPU = New CPU
        Set m_Mem = New Memory
        Set m_HD = New HardDrive
    End Sub

    ' One-call boot: internally calls each subsystem in order
    Public Function Start
        m_CPU.Freeze
        Dim bootData As String
        bootData = m_HD.Read(0)
        m_Mem.Load 0, bootData
        m_CPU.Jump 0
        m_CPU.Execute
    End Function
End Class

' Demo: caller only needs Start, no need to know the internals
Dim pc As ComputerFacade
Set pc = New ComputerFacade
pc.Start   ' Only Start is exposed
```

**Axon VBScript trade-offs**:
- The classic version already has no structural trade-offs: the facade class just composes subsystem calls, no dependency on inheritance or interfaces — it's already idiomatic. The strongly-typed version annotates the three subsystem references as `As CPU`/`As Memory`/`As HardDrive`, types address/sector parameters as `As Long` and data parameters as `As String`, letting the IDE catch errors like "passing a string to an address parameter" at compile time. Classified in the appendix as "already idiomatic in classic" — one of only three patterns out of 23 (Facade #7, Flyweight #12, Memento #22). The reason: these three patterns' core structures don't depend on interfaces or inheritance at all — as long as you have `Class` + composition calls, you can implement them idiomatically. AxonASP's strong typing is an implementation quality improvement, not a change to the Facade's compositional structure.

### VB.NET Version (syntactically complete baseline)

VB.NET uses `Sub`/`Function` instead of VBScript's unified `Function`, and `Sub New()` constructor instead of `Class_Initialize`. Same scenario as Axon version — subsystems are regular Classes, facade composes and calls them.

```vbnet
' ① Subsystems: regular Classes, one-to-one with Axon version
Public Class CPU
    Public Sub Freeze()
        Console.WriteLine("CPU freeze")
    End Sub
    Public Sub Jump(position As Long)
        Console.WriteLine("CPU jump to " & position)
    End Sub
    Public Sub Execute()
        Console.WriteLine("CPU execute")
    End Sub
End Class

Public Class Memory
    Public Sub Load(position As Long, data As String)
        Console.WriteLine("Memory load " & data & " to " & position)
    End Sub
End Class

Public Class HardDrive
    Public Function Read(lba As Long) As String
        Return "DataBlock(" & lba & ")"
    End Function
End Class

' ② Facade class: composes subsystems, exposes only Start
Public Class ComputerFacade
    Private m_CPU As CPU
    Private m_Mem As Memory
    Private m_HD As HardDrive

    ' Constructor: directly New to create subsystems (same as Axon's Class_Initialize)
    Public Sub New()
        m_CPU = New CPU()
        m_Mem = New Memory()
        m_HD = New HardDrive()
    End Sub

    ' One-call boot: internally calls each subsystem in order
    Public Sub Start()
        m_CPU.Freeze()
        Dim bootData As String = m_HD.Read(0)
        m_Mem.Load(0, bootData)
        m_CPU.Jump(0)
        m_CPU.Execute()
    End Sub
End Class

' Demo: external callers only need Start, no need to know internals
Dim pc As New ComputerFacade()
pc.Start()   ' Only Start is exposed externally
```

**VB.NET version notes**:
- **Facade structure needs no inheritance or interfaces**: All three versions' core skeletons are identical — facade class composes multiple subsystems, one method chains multiple subsystem calls. This is the fundamental reason the appendix lists Facade #7 as "already idiomatic in classic".
- **`Sub` replaces unified `Function`**: VB.NET distinguishes `Sub` (no return value) from `Function` (returns value). Methods that don't return values should use `Sub` for clarity. Axon version uses `Function` for all methods, even those without return values.
- **`Sub New()` replaces `Class_Initialize`**: VB.NET uses explicit constructor `Public Sub New()` to initialize objects; Axon version uses `Private Sub Class_Initialize`.
- **No `Set` needed**: VB.NET object assignment uses `=` directly, `m_CPU = New CPU()` doesn't need `Set`.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Pattern idiomaticity | ✅ Idiomatic (composition calls suffice) | ✅ Idiomatic (structure unchanged, only type annotations added) | ✅ Idiomatic (same skeleton) |
| Method syntax | Unified `Function` | Unified `Function` | `Sub` (no return) / `Function` (returns value) |
| Constructor | `Class_Initialize` | `Class_Initialize` | `Sub New()` |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |
---
