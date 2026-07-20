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
        Response.Write Name & " 收到新闻：" & news
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
        Notify news
    End Function
End Class

' 演示：多个观察者同时收到通知
Dim agency, paper1, paper2
Set agency = New NewsAgency
Set paper1 = New Newspaper
paper1.Name = "晨报"
Set paper2 = New Newspaper
paper2.Name = "晚报"

agency.Subscribe paper1
agency.Subscribe paper2
agency.Publish "重大新闻！"
```

**传统 VBScript 版妥协说明**：
- **无事件/委托**：VBScript 没有内置事件机制，Subject 必须手动维护观察者数组（Dim + ReDim + 遍历），代码比 .NET 的 event += 笨重得多。
- **无接口约束**：Newspaper 没有 `IObserver` 接口强制实现 `Update`，如果某个类方法名不一致，运行时调用才报错。

### Axon VBScript 版（支持 Event）

```vbscript
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
        Response.Write Name & " 收到新闻：" & news
    End Function
End Class

' 用 WithEvents 声明事件接收变量
Dim WithEvents agency As NewsAgency

' 事件处理程序：命名规则为 变量名_事件名
Sub agency_OnNews(news As String)
    ' 通过全局引用分发到具体观察者
    paper1.Update news
    paper2.Update news
End Sub

Dim paper1 As Newspaper, paper2 As Newspaper
Set paper1 = New Newspaper
paper1.Name = "晨报"
Set paper2 = New Newspaper
paper2.Name = "晚报"

Set agency = New NewsAgency
agency.Publish "重大新闻！"
```

**Axon VBScript 版妥协说明**：
- 此模式在 AxonASP 中实现较为自然，`Event`/`RaiseEvent`/`WithEvents` 提供了内置的观察者机制，无需手动维护数组。但 `WithEvents` 变量不能是过程内局部变量，必须为类成员或全局变量。
---