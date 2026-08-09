## 第3章 抽象工厂模式（Abstract Factory）

**核心思想**：创建一系列相关对象（产品族），换一套工厂就换整套风格。

**示例说明**：WinFactory 创建 Windows 风格的按钮+复选框，MacFactory 创建 Mac 风格的按钮+复选框。切换工厂即可切换整套 UI 风格，无需逐个替换。

### 传统 VBScript 版

```vbscript
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
```

**传统 VBScript 版妥协说明**：
- **无接口**：`WinFactory` 和 `MacFactory` 没有共同接口 `IGUIFactory`，编译器无法保证两者都有 `CreateButton`/`CreateCheckbox`。如果某个工厂漏写了方法，运行时调用才报错。
- **产品类无约束**：所有 Button 和 Checkbox 类仅靠方法名约定，没有 `IButton`/`ICheckbox` 接口保证一致性。

### Axon VBScript 版（支持 Implements）

```vba
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
```

**Axon VBScript 版妥协说明**：
- 接口机制解决了产品族和工厂族的契约约束问题。残留限制：**缺失代码复用机制（继承）**。经典抽象工厂要求"抽象 Factory 基类 + 产品族基类"来共享公共逻辑（例如所有 Button 基类统一 `Click` 事件、所有 Factory 基类统一辅助方法），AxonASP 目前只能用接口，所有公共逻辑必须在每个具体类中重复编写。Go 同样无继承但用 struct embedding 解决此问题，VBScript 无对应机制。

### VB.NET 版（语法完备的对照基准）

VB.NET 用 `MustInherit`/`MustOverride`/`Inherits`/`Overrides` 把 Axon 版的接口升级为抽象基类，产品族和工厂族都能共享基类代码。

```vbnet
' ① 产品抽象基类（替代 Axon 版的 IButton/ICheckbox 接口）
Public MustInherit Class Button
    Public MustOverride Function Paint() As Object
End Class

Public MustInherit Class Checkbox
    Public MustOverride Function Paint() As Object
End Class

' ② 工厂抽象基类（替代 Axon 版的 IGUIFactory 接口）
Public MustInherit Class GUIFactory
    Public MustOverride Function CreateButton() As Button
    Public MustOverride Function CreateCheckbox() As Checkbox
End Class

' ③ Windows 产品族
Public Class WinButton
    Inherits Button
    Public Overrides Function Paint() As Object
        Console.WriteLine("绘制 Windows 风格按钮")
    End Function
End Class

Public Class WinCheckbox
    Inherits Checkbox
    Public Overrides Function Paint() As Object
        Console.WriteLine("绘制 Windows 风格复选框")
    End Function
End Class

' ④ Mac 产品族
Public Class MacButton
    Inherits Button
    Public Overrides Function Paint() As Object
        Console.WriteLine("绘制 Mac 风格按钮")
    End Function
End Class

Public Class MacCheckbox
    Inherits Checkbox
    Public Overrides Function Paint() As Object
        Console.WriteLine("绘制 Mac 风格复选框")
    End Function
End Class

' ⑤ Windows 工厂
Public Class WinFactory
    Inherits GUIFactory
    Public Overrides Function CreateButton() As Button
        Return New WinButton()
    End Function
    Public Overrides Function CreateCheckbox() As Checkbox
        Return New WinCheckbox()
    End Function
End Class

' ⑥ Mac 工厂
Public Class MacFactory
    Inherits GUIFactory
    Public Overrides Function CreateButton() As Button
        Return New MacButton()
    End Function
    Public Overrides Function CreateCheckbox() As Checkbox
        Return New MacCheckbox()
    End Function
End Class

' 演示：只需改一行 New WinFactory()，整套风格全换
Dim factory As GUIFactory = New MacFactory()
Dim btn As Button = factory.CreateButton()
Dim chk As Checkbox = factory.CreateCheckbox()
btn.Paint()
chk.Paint()
```

**VB.NET 版说明**：
- **抽象基类替代接口**：`MustInherit Class Button` 禁止 `New Button()`，`MustOverride Paint` 强制子类实现——编译期检查。Axon 版的 `IButton` 只是空壳类，外部照样可以 `New IButton` 并调用空的 `Paint`。
- **继承带来代码复用**：若要给所有 Button 加 `Font` 属性，只需在 `Button` 基类加一次，子类自动继承。Axon 版只能手动给 `WinButton`、`MacButton` 各写一份。
- **`Overrides` 显式标记重写**：子类重写必须写 `Public Overrides Sub Paint()`，漏写或签名不对编译期直接报错。Axon 版的 `IButton_Paint` 只是方法名前缀约定。
- **无需 `Set`**：VB.NET 对象赋值直接用 `=`，不再需要 `Set` 给对象、`Let` 给值类型这一历史包袱。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 产品契约 | 方法名约定 | `IButton`/`ICheckbox` 接口约束 | `MustInherit` 基类 + `MustOverride` 编译期强制 |
| 工厂契约 | 方法名约定 | `IGUIFactory` 接口约束 | `MustInherit` 基类 + `Overrides` 重写 |
| 代码复用 | 无 | 无（接口类平行，无继承） | 基类字段/方法自动传给所有子类 |
| 抽象不可实例化 | 不能 | 不能（接口类只是普通 Class） | `MustInherit` 编译期禁止 `New` |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
---