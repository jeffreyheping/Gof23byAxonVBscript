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
- The interface mechanism lets both proxy and real object implement the same `IImage` interface, enabling transparent substitution. `ProxyImage` holds the real object via an `IImage`-typed field, creates it on demand in `Display`, and delegates the call — matching classic proxy lazy-loading semantics. Remaining gap: **missing code reuse mechanism (inheritance / struct embedding) causes verbose delegation forwarding boilerplate**. The classic Proxy structure is "abstract Subject base class + RealSubject + Proxy", all sharing base class code (e.g., common `Filename` property, logging logic), with Proxy only overriding methods it needs to intercept. AxonASP has no inheritance — even if the interface has many methods, each must be manually forwarded in Proxy with one line like `m_RealImage.SomeMethod(...)`. Ten methods means ten forwarding lines; if the interface adds methods later, Proxy must follow. Go also lacks inheritance but uses **struct embedding** to solve this: embedding `*RealImage` automatically gets all its methods, Proxy only writes the `Display` it needs to intercept, all other methods pass through at zero cost. AxonASP currently requires manual forwarding for every method.

### VB.NET Version (syntactically complete baseline)

VB.NET uses `MustInherit`/`Inherits`/`Overrides` to upgrade Axon's interface into an abstract base class. Same scenario as Axon version — lazy loading.

```vbnet
' ① Subject abstract base class (replacing Axon's IImage interface)
Public MustInherit Class Image
    Public MustOverride Sub Init(filename As String)
    Public MustOverride Sub Display()
End Class

' ② RealSubject: real object
Public Class RealImage
    Inherits Image

    Private m_Filename As String

    Public Overrides Sub Init(filename As String)
        m_Filename = filename
        Console.WriteLine("[Loading large image] " & filename)
    End Sub

    Public Overrides Sub Display()
        Console.WriteLine("Displaying image: " & m_Filename)
    End Sub
End Class

' ③ Proxy: lazy loading, holds real object via base class reference
Public Class ProxyImage
    Inherits Image

    Private m_Filename As String
    Private m_RealImage As Image   ' Base class reference, initially Nothing

    Public Overrides Sub Init(filename As String)
        m_Filename = filename
    End Sub

    Public Overrides Sub Display()
        If m_RealImage Is Nothing Then
            m_RealImage = New RealImage()
            m_RealImage.Init(m_Filename)
        End If
        m_RealImage.Display()
    End Sub
End Class

' Demo: transparently use proxy via abstract base class reference
Dim img As Image = New ProxyImage()
img.Init("photo.jpg")
Console.WriteLine("Proxy created; real image not yet loaded")
img.Display()   ' Only now triggers real loading
img.Display()   ' Second time: no loading
```

**VB.NET version notes**:
- **Abstract base class instead of interface**: `MustInherit Class Image` prevents `New Image()`, `MustOverride` forces subclass implementation — compile-time check. Axon's `IImage` is just an empty shell; external code can still `New IImage`.
- **Code reuse through inheritance**: If you want to add shared fields/methods to RealImage and ProxyImage, just add once in `Image` base class, subclasses inherit automatically. Axon's RealImage and ProxyImage are parallel classes, `m_Filename` must be written in both.
- **`Overrides` explicitly marks overrides**: Subclass overrides must write `Public Overrides Sub Display()`, missing or wrong signature causes compile error. Axon's `IImage_Display` is just a method name prefix convention.
- **No `Set` needed**: VB.NET object assignment uses `=` directly, `m_RealImage = New RealImage()` doesn't need `Set`.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Subject contract | None (method name convention) | `IImage` interface constraints | `MustInherit` base class + `MustOverride` compile-time enforced |
| Code reuse | None (fields written separately) | None (Proxy and Real are parallel classes, no inheritance) | Base class fields/methods automatically inherited by all subclasses |
| Abstract non-instantiable | Cannot | Cannot (interface classes are just regular Classes) | `MustInherit` compile-time prevents `New` |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |
---
