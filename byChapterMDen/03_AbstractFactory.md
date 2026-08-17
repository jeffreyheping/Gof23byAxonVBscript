## Chapter 3: Abstract Factory

**Core idea**: Create a family of related objects. Swap the factory, swap the whole style.

**Example**: WinFactory creates Windows-style buttons and checkboxes; MacFactory creates Mac-style ones. Switch the factory, switch the entire UI look — no need to replace items one by one.

### Classic VBScript Version

```vbscript
' ===== Windows-style products =====
Class WinButton
    ' Draw a Windows-style button
    Public Function Paint
        Response.Write "Drawing Windows-style button"
    End Function
End Class

Class WinCheckbox
    ' Draw a Windows-style checkbox
    Public Function Paint
        Response.Write "Drawing Windows-style checkbox"
    End Function
End Class

' ===== Mac-style products =====
Class MacButton
    ' Draw a Mac-style button
    Public Function Paint
        Response.Write "Drawing Mac-style button"
    End Function
End Class

Class MacCheckbox
    ' Draw a Mac-style checkbox
    Public Function Paint
        Response.Write "Drawing Mac-style checkbox"
    End Function
End Class

' ===== Windows factory: creates a full set of Windows-style controls =====
Class WinFactory
    ' Create a Windows button
    Public Function CreateButton
        Set CreateButton = New WinButton
    End Function
    ' Create a Windows checkbox
    Public Function CreateCheckbox
        Set CreateCheckbox = New WinCheckbox
    End Function
End Class

' ===== Mac factory: creates a full set of Mac-style controls =====
Class MacFactory
    ' Create a Mac button
    Public Function CreateButton
        Set CreateButton = New MacButton
    End Function
    ' Create a Mac checkbox
    Public Function CreateCheckbox
        Set CreateCheckbox = New MacCheckbox
    End Function
End Class

' Demo: swap the factory, swap the whole UI style
Dim uiFactory
Set uiFactory = New MacFactory   ' Switch to WinFactory for Windows style
Dim btn, chk
Set btn = uiFactory.CreateButton
Set chk = uiFactory.CreateCheckbox
btn.Paint
chk.Paint
```

**Classic VBScript trade-offs**:
- **No interfaces**: `WinFactory` and `MacFactory` share no `IGUIFactory` interface. The compiler can't guarantee both have `CreateButton`/`CreateCheckbox`. If a factory is missing a method, you'll only find out at runtime.
- **No product constraints**: All Button and Checkbox classes rely on method-name conventions alone. No `IButton`/`ICheckbox` interface ensures consistency.

### Axon VBScript Version (supports Implements)

```vbscript
' Product interfaces
Class IButton
    Public Function Paint
    End Function
End Class

Class ICheckbox
    Public Function Paint
    End Function
End Class

' Factory interface
Class IGUIFactory
    Public Function CreateButton As IButton
    End Function
    Public Function CreateCheckbox As ICheckbox
    End Function
End Class

' Windows products
Class WinButton
    Implements IButton
    Public Function IButton_Paint
        Response.Write "Drawing Windows-style button"
    End Function
End Class

Class WinCheckbox
    Implements ICheckbox
    Public Function ICheckbox_Paint
        Response.Write "Drawing Windows-style checkbox"
    End Function
End Class

' Mac products
Class MacButton
    Implements IButton
    Public Function IButton_Paint
        Response.Write "Drawing Mac-style button"
    End Function
End Class

Class MacCheckbox
    Implements ICheckbox
    Public Function ICheckbox_Paint
        Response.Write "Drawing Mac-style checkbox"
    End Function
End Class

' Windows factory
Class WinFactory
    Implements IGUIFactory
    Public Function IGUIFactory_CreateButton As IButton
        Set IGUIFactory_CreateButton = New WinButton
    End Function
    Public Function IGUIFactory_CreateCheckbox As ICheckbox
        Set IGUIFactory_CreateCheckbox = New WinCheckbox
    End Function
End Class

' Mac factory
Class MacFactory
    Implements IGUIFactory
    Public Function IGUIFactory_CreateButton As IButton
        Set IGUIFactory_CreateButton = New MacButton
    End Function
    Public Function IGUIFactory_CreateCheckbox As ICheckbox
        Set IGUIFactory_CreateCheckbox = New MacCheckbox
    End Function
End Class

' Demo: switch the entire style via interface reference
Dim uiFactory As IGUIFactory
Dim btn As IButton
Dim chk As ICheckbox
Set uiFactory = New MacFactory
Set btn = uiFactory.CreateButton
Set chk = uiFactory.CreateCheckbox
btn.Paint
chk.Paint
```

**Axon VBScript trade-offs**:
- The interface mechanism solves the contract constraints for product families and factory families. Remaining gap: **missing code reuse mechanism (inheritance)**. Classic Abstract Factory requires "abstract Factory base class + product family base class" to share common logic (e.g., all Button base classes share `Click` event, all Factory base classes share helper methods). AxonASP currently can only use interfaces — all common logic must be duplicated in every concrete class. Go also lacks inheritance but solves this with struct embedding; VBScript has no equivalent mechanism.

### VB.NET Version (syntactically complete baseline)

VB.NET uses `MustInherit`/`MustOverride`/`Inherits`/`Overrides` to upgrade Axon's interfaces into abstract base classes — both product families and factory families can share base class code.

```vbnet
' ① Product abstract base classes (replacing Axon's IButton/ICheckbox interfaces)
Public MustInherit Class Button
    Public MustOverride Sub Paint()
End Class

Public MustInherit Class Checkbox
    Public MustOverride Sub Paint()
End Class

' ② Factory abstract base class (replacing Axon's IGUIFactory interface)
Public MustInherit Class GUIFactory
    Public MustOverride Function CreateButton() As Button
    Public MustOverride Function CreateCheckbox() As Checkbox
End Class

' ③ Windows product family
Public Class WinButton
    Inherits Button
    Public Overrides Sub Paint()
        Console.WriteLine("Rendering Windows-style button")
    End Sub
End Class

Public Class WinCheckbox
    Inherits Checkbox
    Public Overrides Sub Paint()
        Console.WriteLine("Rendering Windows-style checkbox")
    End Sub
End Class

' ④ Mac product family
Public Class MacButton
    Inherits Button
    Public Overrides Sub Paint()
        Console.WriteLine("Rendering Mac-style button")
    End Sub
End Class

Public Class MacCheckbox
    Inherits Checkbox
    Public Overrides Sub Paint()
        Console.WriteLine("Rendering Mac-style checkbox")
    End Sub
End Class

' ⑤ Windows factory
Public Class WinFactory
    Inherits GUIFactory
    Public Overrides Function CreateButton() As Button
        Return New WinButton()
    End Function
    Public Overrides Function CreateCheckbox() As Checkbox
        Return New WinCheckbox()
    End Function
End Class

' ⑥ Mac factory
Public Class MacFactory
    Inherits GUIFactory
    Public Overrides Function CreateButton() As Button
        Return New MacButton()
    End Function
    Public Overrides Function CreateCheckbox() As Checkbox
        Return New MacCheckbox()
    End Function
End Class

' Demo: change just one line New WinFactory() to switch the entire style
Dim factory As GUIFactory = New MacFactory()
Dim btn As Button = factory.CreateButton()
Dim chk As Checkbox = factory.CreateCheckbox()
btn.Paint()
chk.Paint()
```

**VB.NET version notes**:
- **Abstract base classes instead of interfaces**: `MustInherit Class Button` prevents `New Button()`, `MustOverride Paint` forces subclass implementation — compile-time check. Axon's `IButton` is just an empty shell; external code can still `New IButton` and call empty `Paint`.
- **Code reuse through inheritance**: If you want to add `Font` property to all Buttons, just add once in `Button` base class, subclasses inherit automatically. Axon version must manually write it for `WinButton` and `MacButton` separately.
- **`Overrides` explicitly marks overrides**: Subclass overrides must write `Public Overrides Sub Paint()`, missing or wrong signature causes compile error. Axon's `IButton_Paint` is just a method name prefix convention.
- **No `Set` needed**: VB.NET object assignment uses `=` directly, no need for `Set` for objects, `Let` for value types.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Product contract | Method name convention | `IButton`/`ICheckbox` interface constraints | `MustInherit` base class + `MustOverride` compile-time enforced |
| Factory contract | Method name convention | `IGUIFactory` interface constraints | `MustInherit` base class + `Overrides` |
| Code reuse | None | None (parallel interface classes, no inheritance) | Base class fields/methods automatically inherited by all subclasses |
| Abstract non-instantiable | Cannot | Cannot (interface classes are just regular Classes) | `MustInherit` compile-time prevents `New` |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |
---
