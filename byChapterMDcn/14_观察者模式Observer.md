## 第14章 观察者模式（Observer）

**核心思想**：对象状态变化时通知所有关注它的观察者。

**示例说明**：NewsAgency（被观察者）维护一个观察者列表。注册多个 Newspaper 后，一旦有新闻发布，所有报纸自动收到通知。

### 传统 VBScript 版

```vbscript
' 观察者：报纸
Class Newspaper
    Public Name

    ' 收到新闻时的响应
    Public Function Update(news)
        Response.Write(Name & " 收到新闻：" & news)

    End Function
End Class

' 被观察者：新闻社
Class NewsAgency
    Private m_Observers()   ' 观察者数组

    Private m_Count         ' 当前观察者数量

    ' 构造函数：初始化数组
    Private Sub Class_Initialize
        m_Count = 0
        ReDim m_Observers(10)
    End Sub

    ' 注册观察者（容量不足时自动扩容）
    Public Function Subscribe(observer)
        If m_Count >= UBound(m_Observers) + 1 Then
            ReDim Preserve m_Observers(m_Count * 2)
        End If
        Set m_Observers(m_Count) = observer
        m_Count = m_Count + 1
    End Function

    ' 通知所有观察者
    Public Function Notify(news)
        Dim i
        For i = 0 To m_Count - 1
            m_Observers(i).Update news
        Next
    End Function

    ' 发布新闻：先更新自身状态，再通知所有观察者
    Public Function Publish(news)
        Notify(news)

    End Function
End Class

' 演示：多个观察者同时收到通知
Dim agency, paper1, paper2
Set agency = New NewsAgency
Set paper1 = New Newspaper
paper1.Name = "晨报"
Set paper2 = New Newspaper
paper2.Name = "晚报"

agency.Subscribe(paper1)

agency.Subscribe(paper2)

agency.Publish("重大新闻！")

```

**传统 VBScript 版妥协说明**：
- **无事件/委托**：VBScript 没有内置事件机制，Subject 必须手动维护观察者数组（Dim + ReDim + 遍历），代码比 .NET 的 event += 笨重得多。
- **无接口约束**：Newspaper 没有 `IObserver` 接口强制实现 `Update`，如果某个类方法名不一致，运行时调用才报错。

### Axon VBScript 版（支持 Event）

```vba
' 被观察者：声明事件
Class NewsAgency
    Event OnNews(news As String)

    ' 发布新闻：触发事件，所有订阅者自动收到通知
    Public Function Publish(news As String)
        RaiseEvent OnNews(news)
    End Function
End Class

' 具体观察者
Class Newspaper
    Public Name As String

    ' 收到新闻时的响应
    Public Function Update(news As String)
        Response.Write(Name & " 收到新闻：" & news)

    End Function
End Class

' 用 WithEvents 声明事件接收变量
Dim WithEvents agency As NewsAgency

' 事件处理程序：命名规则为 变量名_事件名
Sub agency_OnNews(news As String)
    ' 通过全局引用分发到具体观察者
    paper1.Update(news)

    paper2.Update(news)

End Sub

Dim paper1 As Newspaper, paper2 As Newspaper
Set paper1 = New Newspaper
paper1.Name = "晨报"
Set paper2 = New Newspaper
paper2.Name = "晚报"

Set agency = New NewsAgency
agency.Publish("重大新闻！")

```

**Axon VBScript 版妥协说明**：
- 此模式在 AxonASP 中**彻底解决核心痛点，无妥协**。`Event`/`RaiseEvent`/`WithEvents` 提供了内置的观察者机制，`Event OnNews` 声明事件契约，`RaiseEvent OnNews(news)` 触发通知，`WithEvents agency` + `Sub agency_OnNews(...)` 自动订阅，无需手动维护观察者数组、无需写 `Subscribe`/`Notify` 样板代码。附录将此模式归类为"AxonASP 彻底解决核心痛点"的 18 个模式之一，无残留缺陷。

### VB.NET 版（语法完备的对照基准）

VB.NET 拥有 `Event`/`EventHandler(Of T)`/`Handles` 关键字的原生事件系统——这是 .NET 平台观察者模式的地道写法，事件委托自动管理订阅列表，`Handles` 关键字声明式绑定。

```vbnet
' ① 事件参数类：继承 EventArgs，承载新闻内容
Public Class NewsEventArgs
    Inherits EventArgs
    Public ReadOnly Property News As String

    Public Sub New(news As String)
        News = news
    End Sub
End Class

' ② 被观察者：用普通 Event + EventHandler(Of T) 声明强类型事件
Public Class NewsAgency
    ' 声明事件：EventHandler(Of NewsEventArgs) 是 .NET 标准泛型委托
    Public Event NewsPublished As EventHandler(Of NewsEventArgs)

    ' 发布新闻：触发事件，所有订阅者自动收到通知
    Public Function Publish(news As String)
        RaiseEvent NewsPublished(Me, New NewsEventArgs(news))
    End Function
End Class

' ③ 观察者：通过 WithEvents + Handles 声明式订阅
Public Class Newspaper
    Private ReadOnly m_Name As String
    Private WithEvents m_Agency As NewsAgency

    Public Sub New(name As String, agency As NewsAgency)
        m_Name = name
        m_Agency = agency
    End Sub

    ' Handles 关键字：声明式绑定 m_Agency 的 NewsPublished 事件
    Private Sub OnNewsPublished(sender As Object, e As NewsEventArgs) _
        Handles m_Agency.NewsPublished
        Console.WriteLine($"{m_Name} 收到新闻：{e.News}")
    End Sub
End Class

' 演示：与 Axon 版新闻场景一致
Dim agency As New NewsAgency()
Dim paper1 As New Newspaper("晨报", agency)
Dim paper2 As New Newspaper("晚报", agency)
agency.Publish("重大新闻！")
```

**VB.NET 版说明**：
- **`EventHandler(Of T)` 标准 .NET 事件签名**：事件参数 `NewsEventArgs` 继承 `EventArgs`，采用 `(sender, e)` 标准签名，编译期约束参数类型。Axon 版 `Event OnNews(news As String)` 虽支持强类型单参，但缺少 .NET 标准 `sender/e` 约定，跨类库协作时缺少互操作性。
- **`WithEvents` + `Handles` 声明式订阅**：与 Axon 版的 `WithEvents` + `Sub Var_Event()` 命名约定对应，但 VB.NET 用 `Handles` 关键字显式绑定事件，编译期校验事件名是否存在；Axon 版靠命名约定，事件名拼错运行时才报错。
- **多播委托自动管理订阅列表**：.NET 的 `Event` 底层是 `MulticastDelegate`，自动维护订阅者链表。对比传统版：手动 `ReDim m_Observers(m_Count * 2)` 扩容、`For i = 0 To m_Count - 1` 遍历通知，都是开发者自己维护。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 观察者管理 | 手动数组 + `ReDim` 扩容 + 遍历 | `Event`/`RaiseEvent`/`WithEvents` 内置 | `MulticastDelegate` 多播委托自动管理 |
| 订阅方式 | 手动 `Subscribe(observer)` 注册 | `WithEvents` + `Sub Var_Event()` 命名约定 | `WithEvents` + `Handles` 编译期校验绑定 |
| 事件参数 | 无约束，任意传 | `Event OnNews(news As String)` 强类型单参 | `EventHandler(Of T)` 标准 `(sender, e)` 签名 |
| 事件契约校验 | 无（运行时报错） | 命名约定（事件名拼错运行时报错） | `Handles` 编译期校验事件名存在 |
---