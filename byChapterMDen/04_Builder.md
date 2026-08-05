## Chapter 4: Builder

**Core idea**: Build a complex object step by step. The same construction process can produce different configurations.

**Example**: A Director drives a Builder to assemble a Computer step by step (CPU → RAM → Disk). The same Builder can be directed to produce either a gaming PC or an office PC.

### Classic VBScript Version

```vbscript
' Product: Computer
Class Computer
    Public CPU, RAM, Disk
    ' Print current config
    Public Function ShowConfig
        Response.Write "Config: " & CPU & " / " & RAM & " / " & Disk
    End Function
End Class

' Builder: assembles Computer parts step by step
Class ComputerBuilder
    Private m_Computer

    ' Constructor: create a blank product
    Private Sub Class_Initialize
        Set m_Computer = New Computer
    End Sub

    ' Install CPU
    Public Function BuildCPU(v)
        m_Computer.CPU = v
    End Function
    ' Install RAM
    Public Function BuildRAM(v)
        m_Computer.RAM = v
    End Function
    ' Install disk
    Public Function BuildDisk(v)
        m_Computer.Disk = v
    End Function
    ' Return the finished product
    Public Function GetResult
        Set GetResult = m_Computer
    End Function
End Class

' Director: calls Builder in a fixed sequence for different configurations
Class Director
    ' Config 1: build a gaming PC
    Public Function ConstructGamingPC(b)
        b.BuildCPU "i9"
        b.BuildRAM "32GB"
        b.BuildDisk "2TB SSD"
    End Function
    ' Config 2: build an office PC
    Public Function ConstructOfficePC(b)
        b.BuildCPU "i5"
        b.BuildRAM "16GB"
        b.BuildDisk "512GB SSD"
    End Function
End Class

' Demo: Director drives Builder to assemble
Dim myBuilder, myDirector, pc
Set myBuilder = New ComputerBuilder
Set myDirector = New Director
myDirector.ConstructGamingPC myBuilder
Set pc = myBuilder.GetResult
pc.ShowConfig
```

**Classic VBScript trade-offs**:
- **No interface**: The `builder` parameter passed to Director has no `IBuilder` interface constraint. If the passed object lacks `BuildCPU` etc., the error only surfaces at runtime.
- **No chained calls**: VBScript can return `Me` from a Function, but the syntax doesn't support `builder.BuildCPU("i9").BuildRAM("32GB")` chaining (`Set FunctionName = Me` works but the caller still needs to capture each return value), so calls must be made line by line.

### Axon VBScript Version (supports Implements)

```vbscript
' Product
Class Computer
    Public CPU As String, RAM As String, Disk As String
    Public Function ShowConfig
        Response.Write "Config: " & CPU & " / " & RAM & " / " & Disk
    End Function
End Class

' Builder interface
Class IBuilder
    Public Function BuildCPU(v As String)
    End Function
    Public Function BuildRAM(v As String)
    End Function
    Public Function BuildDisk(v As String)
    End Function
    Public Function GetResult As Computer
    End Function
End Class

' Concrete builder
Class ComputerBuilder
    Implements IBuilder
    Private m_Computer As Computer

    Private Sub Class_Initialize
        Set m_Computer = New Computer
    End Sub

    Public Function IBuilder_BuildCPU(v As String)
        m_Computer.CPU = v
    End Function
    Public Function IBuilder_BuildRAM(v As String)
        m_Computer.RAM = v
    End Function
    Public Function IBuilder_BuildDisk(v As String)
        m_Computer.Disk = v
    End Function
    Public Function IBuilder_GetResult As Computer
        Set IBuilder_GetResult = m_Computer
    End Function
End Class

' Director
Class Director
    Public Function ConstructGamingPC(builder As IBuilder)
        builder.BuildCPU "i9"
        builder.BuildRAM "32GB"
        builder.BuildDisk "2TB SSD"
    End Function

    Public Function ConstructOfficePC(builder As IBuilder)
        builder.BuildCPU "i5"
        builder.BuildRAM "16GB"
        builder.BuildDisk "512GB SSD"
    End Function
End Class

' Demo
Dim myBuilder As IBuilder
Dim myDirector As Director
Dim pc As Computer
Set myBuilder = New ComputerBuilder
Set myDirector = New Director
myDirector.ConstructGamingPC myBuilder
Set pc = myBuilder.GetResult
pc.ShowConfig
```

**Axon VBScript trade-offs**:
- The interface mechanism enforces the builder's method contract. Director can safely accept a builder via `IBuilder`. However, chained call syntax is still unavailable — `builder.BuildCPU("i9").BuildRAM("32GB")` remains impossible.
---
