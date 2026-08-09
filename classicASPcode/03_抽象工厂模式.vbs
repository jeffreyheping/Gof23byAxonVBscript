Option Explicit
Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
' ===== Windows 风格产品 =====
Class WinButton
    ' 绘制 Windows 风格按钮
    Public Function Paint
        Response.Write("绘制 Windows 风格按钮")

    End Function
End Class

Class WinCheckbox
    ' 绘制 Windows 风格复选框
    Public Function Paint
        Response.Write("绘制 Windows 风格复选框")

    End Function
End Class

' ===== Mac 风格产品 =====
Class MacButton
    ' 绘制 Mac 风格按钮
    Public Function Paint
        Response.Write("绘制 Mac 风格按钮")

    End Function
End Class

Class MacCheckbox
    ' 绘制 Mac 风格复选框
    Public Function Paint
        Response.Write("绘制 Mac 风格复选框")

    End Function
End Class

' ===== Windows 工厂：创建一整套 Windows 风格控件 =====
Class WinFactory
    ' 创建 Windows 按钮
    Public Function CreateButton
        Set CreateButton = New WinButton
    End Function
    ' 创建 Windows 复选框
    Public Function CreateCheckbox
        Set CreateCheckbox = New WinCheckbox
    End Function
End Class

' ===== Mac 工厂：创建一整套 Mac 风格控件 =====
Class MacFactory
    ' 创建 Mac 按钮
    Public Function CreateButton
        Set CreateButton = New MacButton
    End Function
    ' 创建 Mac 复选框
    Public Function CreateCheckbox
        Set CreateCheckbox = New MacCheckbox
    End Function
End Class

' 演示：换一个工厂，整套 UI 风格全换
Dim uiFactory
Set uiFactory = New MacFactory   ' 改成 WinFactory 就是 Windows 风格
Dim btn, chk
Set btn = uiFactory.CreateButton
Set chk = uiFactory.CreateCheckbox
btn.Paint
chk.Paint

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
