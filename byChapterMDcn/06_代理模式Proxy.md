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
        Response.Write "【加载大图】" & filename
    End Function

    ' 显示图片
    Public Function Display
        Response.Write "显示图片：" & m_Filename
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
            m_RealImage.Init m_Filename
        End If
        m_RealImage.Display
    End Function
End Class

' 演示：代理创建时不加载，调用 Display 才加载
Dim img
Set img = New ProxyImage
img.Init "photo.jpg"
Response.Write "代理已创建，真实大图尚未加载"
img.Display   ' 此时才触发真实加载
img.Display   ' 第二次不再加载
```

**传统 VBScript 版妥协说明**：
- **无共享接口**：ProxyImage 和 RealImage 没有 `IImage` 接口，外部代码无法透明替换。经典代理模式要求代理和真实对象实现同一接口，VBScript 做不到强制保证。

### Axon VBScript 版（支持 Implements）

```vbscript
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
        Response.Write "【加载大图】" & filename
    End Function
    Public Function IImage_Display
        Response.Write "显示图片：" & m_Filename
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
            m_RealImage.Init m_Filename
        End If
        m_RealImage.Display
    End Function
End Class

' 演示：通过接口透明使用代理或真实对象
Dim img As IImage
Set img = New ProxyImage
img.Init "photo.jpg"
Response.Write "代理已创建，真实大图尚未加载"
img.Display
img.Display
```

**Axon VBScript 版妥协说明**：
- 接口机制使代理与真实对象实现同一 `IImage` 接口，外部可透明替换。`ProxyImage` 通过 `IImage` 类型的字段持有真实对象，在 `Display` 时按需创建并委托调用，符合经典代理模式的延迟加载语义。
- 缺失语法点：**委托转发机制**（或继承、或 struct embedding）。代理类每个接口方法都要手写一行转发逻辑，接口方法越多样板越冗长。Go 同样无继承，但 Go 有 **struct embedding**——嵌入门类型即可自动获得其方法，代理只需覆盖需要拦截的方法。AxonASP 目前只能手动逐个转发，无自动委托语法。
---