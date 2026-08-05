## Chapter 14: Observer

**Core idea**: Notify all interested observers when an object's state changes.

**Example**: NewsAgency (the subject) maintains a list of observers. After registering multiple Newspapers, publishing news automatically notifies all of them.

### Classic VBScript Version

```vbscript
' Observer: newspaper
Class Newspaper
    Public Name

    ' Response when news arrives
    Public Function Update(news)
        Response.Write Name & " received news: " & news
    End Function
End Class

' Subject: news agency
Class NewsAgency
    Private m_Observers()   ' Observer array
    Private m_Count         ' Current observer count

    ' Constructor: initialize array
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Observers(10)
    End Sub

    ' Register observer (auto-resize when full)
    Public Function Subscribe(observer)
        If m_Count >= UBound(m_Observers) + 1 Then
            ReDim Preserve m_Observers(m_Count * 2)
        End If
        Set m_Observers(m_Count) = observer
        m_Count = m_Count + 1
    End Function

    ' Notify all observers
    Public Function Notify(news)
        Dim i
        For i = 0 To m_Count - 1
            m_Observers(i).Update news
        Next
    End Function

    ' Publish news: update own state, then notify all observers
    Public Function Publish(news)
        Notify news
    End Function
End Class

' Demo: multiple observers receive notification simultaneously
Dim agency, paper1, paper2
Set agency = New NewsAgency
Set paper1 = New Newspaper
paper1.Name = "Morning Post"
Set paper2 = New Newspaper
paper2.Name = "Evening Post"

agency.Subscribe paper1
agency.Subscribe paper2
agency.Publish "Breaking News!"
```

**Classic VBScript trade-offs**:
- **No event/delegate mechanism**: VBScript has no built-in event system. The subject must manually maintain the observer array (Dim + ReDim + loop) — much more verbose than .NET's `event +=`.
- **No interface constraint**: Newspaper has no `IObserver` interface forcing `Update`. If a class has an inconsistent method name, the error only surfaces at runtime.

### Axon VBScript Version (supports Event)

```vbscript
' Subject: declares event
Class NewsAgency
    Event OnNews(news As String)

    ' Publish news: raise event, all subscribers are notified automatically
    Public Function Publish(news As String)
        RaiseEvent OnNews(news)
    End Function
End Class

' Concrete observer
Class Newspaper
    Public Name As String

    ' Response when news arrives
    Public Function Update(news As String)
        Response.Write Name & " received news: " & news
    End Function
End Class

' Declare event receiver variable with WithEvents
Dim WithEvents agency As NewsAgency

' Event handler: naming convention is variableName_eventName
Sub agency_OnNews(news As String)
    ' Dispatch to concrete observers via global references
    paper1.Update news
    paper2.Update news
End Sub

Dim paper1 As Newspaper, paper2 As Newspaper
Set paper1 = New Newspaper
paper1.Name = "Morning Post"
Set paper2 = New Newspaper
paper2.Name = "Evening Post"

Set agency = New NewsAgency
agency.Publish "Breaking News!"
```

**Axon VBScript trade-offs**:
- This pattern maps naturally to AxonASP. `Event`/`RaiseEvent`/`WithEvents` provides a built-in observer mechanism — no manual array maintenance needed. However, `WithEvents` variables can't be local to a procedure; they must be class members or global variables.
---
