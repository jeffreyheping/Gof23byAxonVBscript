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
- This pattern **fully solves the core pain point in AxonASP, no compromises**. `Event`/`RaiseEvent`/`WithEvents` provides a built-in observer mechanism: `Event OnNews` declares the event contract, `RaiseEvent OnNews(news)` triggers notification, `WithEvents agency` + `Sub agency_OnNews(...)` auto-subscribes — no manual observer array maintenance, no `Subscribe`/`Notify` boilerplate. The appendix classifies this pattern as one of the 18 patterns where "AxonASP fully solves the core pain point", with no residual defects.

### VB.NET Version (syntactically complete baseline)

VB.NET has a native event system with `Event`/`EventHandler(Of T)`/`Handles` keywords — this is the idiomatic .NET Observer pattern. Event delegates auto-manage subscription lists, `Handles` keyword provides declarative binding.

```vbnet
' ① Event args class: inherits EventArgs, carries news content
Public Class NewsEventArgs
    Inherits EventArgs
    Public ReadOnly Property News As String

    Public Sub New(news As String)
        News = news
    End Sub
End Class

' ② Subject: uses standard Event + EventHandler(Of T) for strongly-typed events
Public Class NewsAgency
    ' Declare event: EventHandler(Of NewsEventArgs) is .NET's standard generic delegate
    Public Event NewsPublished As EventHandler(Of NewsEventArgs)

    ' Publish news: trigger event, all subscribers auto-notified
    Public Sub Publish(news As String)
        RaiseEvent NewsPublished(Me, New NewsEventArgs(news))
    End Sub
End Class

' ③ Observer: declarative subscription via WithEvents + Handles
Public Class Newspaper
    Private ReadOnly m_Name As String
    Private WithEvents m_Agency As NewsAgency

    Public Sub New(name As String, agency As NewsAgency)
        m_Name = name
        m_Agency = agency
    End Sub

    ' Handles keyword: declarative binding to m_Agency's NewsPublished event
    Private Sub OnNewsPublished(sender As Object, e As NewsEventArgs) _
        Handles m_Agency.NewsPublished
        Console.WriteLine($"{m_Name} received news: {e.News}")
    End Sub
End Class

' Demo: same news scenario as Axon version
Dim agency As New NewsAgency()
Dim paper1 As New Newspaper("Morning Post", agency)
Dim paper2 As New Newspaper("Evening Post", agency)
agency.Publish("Breaking News!")
```

**VB.NET version notes**:
- **`EventHandler(Of T)` standard .NET event signature**: Event args `NewsEventArgs` inherits `EventArgs`, uses `(sender, e)` standard signature, compile-time constrains parameter types. Axon's `Event OnNews(news As String)` supports strongly-typed single parameter, but lacks .NET standard `sender/e` convention — less interoperability across libraries.
- **`WithEvents` + `Handles` declarative subscription**: Corresponds to Axon's `WithEvents` + `Sub Var_Event()` naming convention, but VB.NET uses `Handles` keyword for explicit binding, compile-time validates event name exists. Axon relies on naming convention — misspelled event name only errors at runtime.
- **Multicast delegate auto-manages subscription list**: .NET's `Event` underlying mechanism is `MulticastDelegate`, auto-maintains subscriber linked list. Compare with classic version: manual `ReDim m_Observers(m_Count * 2)` resizing, `For i = 0 To m_Count - 1` traversal notification — all maintained by the developer.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Observer management | Manual array + `ReDim` resize + traversal | `Event`/`RaiseEvent`/`WithEvents` built-in | `MulticastDelegate` auto-managed |
| Subscription method | Manual `Subscribe(observer)` registration | `WithEvents` + `Sub Var_Event()` naming convention | `WithEvents` + `Handles` compile-time validated binding |
| Event arguments | No constraint, pass anything | `Event OnNews(news As String)` strongly-typed single param | `EventHandler(Of T)` standard `(sender, e)` signature |
| Event contract validation | None (runtime errors) | Naming convention (misspelled name = runtime error) | `Handles` compile-time validates event name exists |
---
