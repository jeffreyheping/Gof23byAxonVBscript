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
- This pattern maps naturally to AxonASP. The interface mechanism solves the core polymorphism problem with no significant trade-offs.
---
