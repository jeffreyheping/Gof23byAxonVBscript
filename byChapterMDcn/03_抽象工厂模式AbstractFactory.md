## 第3章 抽象工厂模式（Abstract Factory）

**核心思想**：创建一系列相关对象（产品族），换一套工厂就换整套风格。

**示例说明**：WinFactory 创建 Windows 风格的按钮+复选框，MacFactory 创建 Mac 风格的按钮+复选框。切换工厂即可切换整套 UI 风格，无需逐个替换。

### 传统 VBScript 版

```vbscript
' ===== Windows 风格产品 =====
Class WinButton
    ' 绘制 Windows 风格按钮
    Public Function Paint
        Response.Write "绘制 Windows 风格按钮"
    End Function
End Class

Class WinCheckbox
    ' 绘制 Windows 风格复选框
    Public Function Paint
        Response.Write "绘制 Windows 风格复选框"
    End Function
End Class

' ===== Mac 风格产品 =====
Class MacButton
    ' 绘制 Mac 风格按钮
    Public Function Paint
        Response.Write "绘制 Mac 风格按钮"
    End Function
End Class

Class MacCheckbox
    ' 绘制 Mac 风格复选框
    Public Function Paint
        Response.Write "绘制 Mac 风格复选框"
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
```

**传统 VBScript 版妥协说明**：
- **无接口**：`WinFactory` 和 `MacFactory` 没有共同接口 `IGUIFactory`，编译器无法保证两者都有 `CreateButton`/`CreateCheckbox`。如果某个工厂漏写了方法，运行时调用才报错。
- **产品类无约束**：所有 Button 和 Checkbox 类仅靠方法名约定，没有 `IButton`/`ICheckbox` 接口保证一致性。

### Axon VBScript 版（支持 Implements）

```vbscript
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
        Response.Write "绘制 Windows 风格按钮"
    End Function
End Class

Class WinCheckbox
    Implements ICheckbox
    Public Function ICheckbox_Paint
        Response.Write "绘制 Windows 风格复选框"
    End Function
End Class

' Mac 产品
Class MacButton
    Implements IButton
    Public Function IButton_Paint
        Response.Write "绘制 Mac 风格按钮"
    End Function
End Class

Class MacCheckbox
    Implements ICheckbox
    Public Function ICheckbox_Paint
        Response.Write "绘制 Mac 风格复选框"
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
```

**Axon VBScript 版妥协说明**：
- 此模式在 AxonASP 中实现较为自然，接口机制解决了核心多态问题，无显著妥协。
---