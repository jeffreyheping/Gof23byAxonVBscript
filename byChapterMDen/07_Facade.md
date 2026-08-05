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
- The classic version already has no structural trade-offs: the facade just composes subsystem calls with no dependency on inheritance or interfaces. The strongly-typed version annotates the three subsystem references as `As CPU`/`As Memory`/`As HardDrive`, and types address/sector parameters as `As Long` and data parameters as `As String`, letting the IDE catch errors like "passing a string to an address parameter" at compile time. This is an implementation quality improvement, not a change to the Facade's compositional structure — so it remains classified as "classic already idiomatic".

---
