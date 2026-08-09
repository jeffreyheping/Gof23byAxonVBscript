Option Strict Off
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Collections
Imports System.Linq
Module Ch03Module
    Public MustInherit Class Button
        Public MustOverride Function Paint()
    End Class
    Public MustInherit Class Checkbox
        Public MustOverride Function Paint()
    End Class
    Public MustInherit Class GUIFactory
        Public MustOverride Function CreateButton() As Button
        Public MustOverride Function CreateCheckbox() As Checkbox
    End Class
    Public Class WinButton
        Inherits Button
        Public Overrides Function Paint()
            Console.WriteLine("绘制 Windows 风格按钮")
        End Function
    End Class
    Public Class WinCheckbox
        Inherits Checkbox
        Public Overrides Function Paint()
            Console.WriteLine("绘制 Windows 风格复选框")
        End Function
    End Class
    Public Class MacButton
        Inherits Button
        Public Overrides Function Paint()
            Console.WriteLine("绘制 Mac 风格按钮")
        End Function
    End Class
    Public Class MacCheckbox
        Inherits Checkbox
        Public Overrides Function Paint()
            Console.WriteLine("绘制 Mac 风格复选框")
        End Function
    End Class
    Public Class WinFactory
        Inherits GUIFactory
        Public Overrides Function CreateButton() As Button
            Return New WinButton()
        End Function
        Public Overrides Function CreateCheckbox() As Checkbox
            Return New WinCheckbox()
        End Function
    End Class
    Public Class MacFactory
        Inherits GUIFactory
        Public Overrides Function CreateButton() As Button
            Return New MacButton()
        End Function
        Public Overrides Function CreateCheckbox() As Checkbox
            Return New MacCheckbox()
        End Function
    End Class
    Sub Main()


        ' ② 工厂抽象基类（替代 Axon 版的 IGUIFactory 接口）

        ' ③ Windows 产品族


        ' ④ Mac 产品族


        ' ⑤ Windows 工厂

        ' ⑥ Mac 工厂

        ' 演示：只需改一行 New WinFactory()，整套风格全换
        Dim factory As GUIFactory = New MacFactory()
        Dim btn As Button = factory.CreateButton()
        Dim chk As Checkbox = factory.CreateCheckbox()
        btn.Paint()
        chk.Paint()
    End Sub
End Module
