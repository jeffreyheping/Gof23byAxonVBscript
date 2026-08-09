<%
Option Explicit
' 产品接口
Class IButton
    Public Function Paint
    End Function
End Class

Class ICheckbox
    Public Function Paint
    End Function
End Class

' 工厂接口
Class IGUIFactory
    Public Function CreateButton As IButton
    End Function
    Public Function CreateCheckbox As ICheckbox
    End Function
End Class

' Windows 产品
Class WinButton
    Implements IButton
    Public Function IButton_Paint
        Response.Write("绘制 Windows 风格按钮")

    End Function
End Class

Class WinCheckbox
    Implements ICheckbox
    Public Function ICheckbox_Paint
        Response.Write("绘制 Windows 风格复选框")

    End Function
End Class

' Mac 产品
Class MacButton
    Implements IButton
    Public Function IButton_Paint
        Response.Write("绘制 Mac 风格按钮")

    End Function
End Class

Class MacCheckbox
    Implements ICheckbox
    Public Function ICheckbox_Paint
        Response.Write("绘制 Mac 风格复选框")

    End Function
End Class

' Windows 工厂
Class WinFactory
    Implements IGUIFactory
    Public Function IGUIFactory_CreateButton As IButton
        Set IGUIFactory_CreateButton = New WinButton
    End Function
    Public Function IGUIFactory_CreateCheckbox As ICheckbox
        Set IGUIFactory_CreateCheckbox = New WinCheckbox
    End Function
End Class

' Mac 工厂
Class MacFactory
    Implements IGUIFactory
    Public Function IGUIFactory_CreateButton As IButton
        Set IGUIFactory_CreateButton = New MacButton
    End Function
    Public Function IGUIFactory_CreateCheckbox As ICheckbox
        Set IGUIFactory_CreateCheckbox = New MacCheckbox
    End Function
End Class

' 演示：通过接口引用切换整套风格
Dim uiFactory As IGUIFactory
Dim btn As IButton
Dim chk As ICheckbox
Set uiFactory = New MacFactory
Set btn = uiFactory.CreateButton
Set chk = uiFactory.CreateCheckbox
btn.Paint
chk.Paint
%>