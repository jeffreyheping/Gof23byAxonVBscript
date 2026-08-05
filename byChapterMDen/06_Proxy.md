## Chapter 6: Proxy

**Core idea**: Provide a surrogate for the real object to control access or defer loading.

**Example**: ProxyImage doesn't load the large image on creation. Only when `Display` is actually called does it create the RealImage and load it. The second `Display` call reuses the already-loaded object.

### Classic VBScript Version

```vbscript
' Real object: loads and displays a large image
Class RealImage
    Private m_Filename

    ' Init: simulate the cost of loading a large file
    Public Function Init(filename)
        m_Filename = filename
        Response.Write "[Loading large image] " & filename
    End Function

    ' Display the image
    Public Function Display
        Response.Write "Displaying image: " & m_Filename
    End Function
End Class

' Proxy: defers loading, controls access to RealImage
Class ProxyImage
    Private m_Filename
    Private m_RealImage   ' The real object, initially Nothing

    Private Sub Class_Initialize
        Set m_RealImage = Nothing
    End Sub

    ' Init: just store the filename, don't load yet
    Public Function Init(filename)
        m_Filename = filename
    End Function

    ' Display: create the real object on first call, reuse afterwards
    Public Function Display
        If m_RealImage Is Nothing Then
            Set m_RealImage = New RealImage
            m_RealImage.Init m_Filename
        End If
        m_RealImage.Display
    End Function
End Class

' Demo: proxy doesn't load on creation; Display triggers the real load
Dim img
Set img = New ProxyImage
img.Init "photo.jpg"
Response.Write "Proxy created; real image not yet loaded"
img.Display   ' Real load happens here
img.Display   ' Second call — no reload
```

**Classic VBScript trade-offs**:
- **No shared interface**: ProxyImage and RealImage have no `IImage` interface. External code can't swap them transparently. The classic Proxy pattern requires both to implement the same interface — VBScript can't enforce this.

### Axon VBScript Version (supports Implements)

```vbscript
' Image interface
Class IImage
    Public Function Init(filename As String)
    End Function
    Public Function Display
    End Function
End Class

' Real object
Class RealImage
    Implements IImage
    Private m_Filename As String

    Public Function IImage_Init(filename As String)
        m_Filename = filename
        Response.Write "[Loading large image] " & filename
    End Function
    Public Function IImage_Display
        Response.Write "Displaying image: " & m_Filename
    End Function
End Class

' Proxy: defers loading, holds the real object via interface
Class ProxyImage
    Implements IImage
    Private m_Filename As String
    Private m_RealImage As IImage

    Public Function IImage_Init(filename As String)
        m_Filename = filename
    End Function

    Public Function IImage_Display
        If m_RealImage Is Nothing Then
            Set m_RealImage = New RealImage
            m_RealImage.Init m_Filename
        End If
        m_RealImage.Display
    End Function
End Class

' Demo: use proxy or real object transparently via interface
Dim img As IImage
Set img = New ProxyImage
img.Init "photo.jpg"
Response.Write "Proxy created; real image not yet loaded"
img.Display
img.Display
```

**Axon VBScript trade-offs**:
- The interface mechanism lets both proxy and real object implement the same `IImage` interface, enabling transparent substitution. `ProxyImage` holds the real object as an `IImage` field, creates it on demand in `Display`, and delegates the call — matching the classic proxy's lazy-loading semantics.
- Missing syntax: **Delegation forwarding** (or inheritance, or struct embedding). Every interface method in the proxy requires a manual forwarding line. The more methods, the more boilerplate. Go also lacks inheritance, but Go has **struct embedding** — embed the real type and automatically get its methods; the proxy only overrides what it needs to intercept. AxonASP currently requires manual forwarding for every method — no automatic delegation syntax.
---
