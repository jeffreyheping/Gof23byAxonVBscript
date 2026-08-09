## 第6章 代理模式（Proxy）

**核心思想**：为真实对象提供替身，控制访问或延迟加载。

**示例说明**：ProxyImage 在创建时不加载大图，只有真正调用 Display 时才创建 RealImage 并加载。第二次 Display 直接复用已加载的真实对象。

### 传统 VBScript 版

```vbscript
' 真实对象：加载并显示大图
Class RealImage
    Private m_Filename

    ' 初始化：模拟加载大文件的耗时操作
    Public Function Init(filename)
        m_Filename = filename
        Response.Write("【加载大图】" & filename)

    End Function

    ' 显示图片
    Public Function Display
        Response.Write("显示图片：" & m_Filename)

    End Function
End Class

' 代理对象：延迟加载，控制对 RealImage 的访问
Class ProxyImage
    Private m_Filename
    Private m_RealImage   ' 被代理的真实对象，初始为 Nothing

    Private Sub Class_Initialize
        Set m_RealImage = Nothing
    End Sub

    ' 初始化：只记录文件名，不加载
    Public Function Init(filename)
        m_Filename = filename
    End Function

    ' 显示图片：首次调用时创建真实对象，后续直接复用
    Public Function Display
        If m_RealImage Is Nothing Then
            Set m_RealImage = New RealImage
            m_RealImage.Init(m_Filename)

        End If
        m_RealImage.Display
    End Function
End Class

' 演示：代理创建时不加载，调用 Display 才加载
Dim img
Set img = New ProxyImage
img.Init("photo.jpg")

Response.Write("代理已创建，真实大图尚未加载")

img.Display()   ' 此时才触发真实加载

img.Display()   ' 第二次不再加载

```

**传统 VBScript 版妥协说明**：
- **无共享接口**：ProxyImage 和 RealImage 没有 `IImage` 接口，外部代码无法透明替换。经典代理模式要求代理和真实对象实现同一接口，VBScript 做不到强制保证。

### Axon VBScript 版（支持 Implements）

```vba
' 图像接口
Class IImage
    Public Function Init(filename As String)
    End Function
    Public Function Display
    End Function
End Class

' 真实对象
Class RealImage
    Implements IImage
    Private m_Filename As String

    Public Function IImage_Init(filename As String)
        m_Filename = filename
        Response.Write("【加载大图】" & filename)

    End Function
    Public Function IImage_Display
        Response.Write("显示图片：" & m_Filename)

    End Function
End Class

' 代理对象：延迟加载，通过接口持有真实对象
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
            m_RealImage.Init(m_Filename)

        End If
        m_RealImage.Display
    End Function
End Class

' 演示：通过接口透明使用代理或真实对象
Dim img As IImage
Set img = New ProxyImage
img.Init("photo.jpg")

Response.Write("代理已创建，真实大图尚未加载")

img.Display
img.Display
```

**Axon VBScript 版妥协说明**：
- 接口机制使代理与真实对象实现同一 `IImage` 接口，外部可透明替换。`ProxyImage` 通过 `IImage` 类型的字段持有真实对象，在 `Display` 时按需创建并委托调用，符合经典代理模式的延迟加载语义。残留限制：**缺失代码复用机制（继承 / struct embedding）导致委托转发样板冗长**。代理模式的经典结构是"抽象 Subject 基类 + RealSubject + Proxy"，三者共享基类代码（比如公共的 `Filename` 属性、日志逻辑），Proxy 只需覆盖需要拦截的方法。AxonASP 没有继承，接口方法再多也必须在 Proxy 里每个都写一行转发——10 个方法就要写 10 行 `m_RealImage.SomeMethod(...)`，后续接口加方法 Proxy 还得跟着补。Go 同样无继承但用 **struct embedding** 解决：嵌入 `*RealImage` 即可自动获得其全部方法，Proxy 只需写需要拦截的 `Display`，其余方法零成本透传。AxonASP 目前只能手动逐个转发。

### VB.NET 版（语法完备的对照基准）

VB.NET 用 `MustInherit`/`Inherits`/`Overrides` 把 Axon 版的接口升级为抽象基类。场景与 Axon 版一致——延迟加载即可。

```vbnet
' ① Subject 抽象基类（替代 Axon 版的 IImage 接口）
Public MustInherit Class Image
    Public MustOverride Function Init(filename As String)
    Public MustOverride Function Display()
End Class

' ② RealSubject：真实对象
Public Class RealImage
    Inherits Image

    Private m_Filename As String

    Public Overrides Function Init(filename As String)
        m_Filename = filename
        Console.WriteLine("【加载大图】" & filename)
    End Function

    Public Overrides Function Display()
        Console.WriteLine("显示图片：" & m_Filename)
    End Function
End Class

' ③ Proxy：延迟加载，通过基类引用持有真实对象
Public Class ProxyImage
    Inherits Image

    Private m_Filename As String
    Private m_RealImage As Image   ' 基类引用，初始为 Nothing

    Public Overrides Function Init(filename As String)
        m_Filename = filename
    End Function

    Public Overrides Function Display()
        If m_RealImage Is Nothing Then
            m_RealImage = New RealImage()
            m_RealImage.Init(m_Filename)
        End If
        m_RealImage.Display()
    End Function
End Class

' 演示：通过抽象基类引用透明使用代理
Dim img As Image = New ProxyImage()
img.Init("photo.jpg")
Console.WriteLine("代理已创建，真实大图尚未加载")
img.Display()   ' 此时才触发真实加载
img.Display()   ' 第二次不再加载
```

**VB.NET 版说明**：
- **抽象基类替代接口**：`MustInherit Class Image` 禁止 `New Image()`，`MustOverride` 强制子类实现——编译期检查。Axon 版的 `IImage` 只是空壳类，外部照样可以 `New IImage`。
- **继承带来代码复用**：若要给 RealImage 和 ProxyImage 加共享字段/方法，只需在 `Image` 基类加一次，子类自动继承。Axon 版的 RealImage 和 ProxyImage 是平行类，`m_Filename` 要各写一份。
- **`Overrides` 显式标记重写**：子类重写必须写 `Public Overrides Sub Display()`，漏写或签名不对编译期直接报错。Axon 版的 `IImage_Display` 只是方法名前缀约定。
- **无需 `Set`**：VB.NET 对象赋值直接用 `=`，`m_RealImage = New RealImage()` 不需要 `Set`。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Subject 契约 | 无（方法名约定） | `IImage` 接口约束 | `MustInherit` 基类 + `MustOverride` 编译期强制 |
| 代码复用 | 无（字段各自写） | 无（Proxy 和 Real 平行类，无继承） | 基类字段/方法自动传给所有子类 |
| 抽象不可实例化 | 不能 | 不能（接口类只是普通 Class） | `MustInherit` 编译期禁止 `New` |
| 对象赋值 | `Set a = New X` | `Set a = New X` | 直接 `a = New X()` |
---