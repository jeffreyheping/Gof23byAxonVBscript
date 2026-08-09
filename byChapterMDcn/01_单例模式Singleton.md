## 第1章 单例模式（Singleton）

**核心思想**：全局只存在一个实例。

**示例说明**：用脚本级变量保存唯一实例，通过 `GetInstance` 函数控制创建。两次获取返回同一个对象，修改一个另一个也变。

### 传统 VBScript 版

```vbscript
' 脚本级变量：全局唯一的实例引用
Dim gInstance
Set gInstance = Nothing

Class Singleton
    Private m_Data

    ' 构造函数：初始化默认数据
    Private Sub Class_Initialize
        m_Data = "我是唯一实例"
    End Sub

    ' 读取内部数据
    Public Property Get Data
        Data = m_Data
    End Property

    ' 写入内部数据
    Public Property Let Data(value)
        m_Data = value
    End Property
End Class

' 全局访问点：若实例不存在则创建，已存在则直接返回
' 返回值：Singleton 类的唯一实例
Function GetInstance()
    If gInstance Is Nothing Then
        Set gInstance = New Singleton
    End If
    Set GetInstance = gInstance
End Function

' 演示：两次获取的是同一个对象
Dim s1, s2
Set s1 = GetInstance()
Set s2 = GetInstance()
s1.Data = "已修改"
Response.Write(s2.Data)   ' 已修改（同一个对象）

```

**传统 VBScript 版妥协说明**：
- **无静态变量**：VBScript 的 Class 内部不支持 `Static` 变量，只能用脚本级（模块级）全局变量 `gInstance` 来保存唯一实例，破坏了类的封装性。
- **无法禁止外部 New**：VBScript 没有 `Private` 构造函数，外部代码随时可以 `New Singleton` 绕过 `GetInstance`，无法真正强制单例。

### Axon VBScript 版（支持 Static）

```vba
Class Singleton
    Private m_Data As String

    Private Sub Class_Initialize
        m_Data = "我是唯一实例"
    End Sub

    Public Property Get Data As String
        Data = m_Data
    End Property

    Public Property Let Data(value As String)
        m_Data = value
    End Property
End Class

' 全局访问点：Static 变量在函数调用间保持值，支持对象引用
Function GetInstance() As Singleton
    Static instance As Singleton

    If instance Is Nothing Then
        Set instance = New Singleton
    End If
    Set GetInstance = instance
End Function

' 演示：保证同一实例
Dim s1 As Singleton, s2 As Singleton
Set s1 = GetInstance()
Set s2 = GetInstance()
s1.Data = "已修改"
Response.Write(s2.Data)   ' 已修改

```

**Axon VBScript 版妥协说明**：
- AxonASP 的 `Static` 变量已支持对象引用，可以在函数内部保持唯一实例，无需模块级全局变量，封装性优于传统版本。残留限制：缺失语法点：**Private 构造函数**。外部仍可 `New Singleton` 绕过单例控制。Go 用**包级私有首字母小写**解决此问题——`type singleton` 不导出，外部只能通过 `GetInstance()` 获取。VBScript 的类没有访问控制，无法禁止外部 `New`。

### VB.NET 版（语法完备的对照基准）

VB.NET 是语法完备的 OOP 语言，拥有 `Private` 构造函数、`Shared`（静态）成员、`ReadOnly` 字段、完整 `Property` 语法。下面的写法不需要任何妥协——**这就是教科书式的单例模式**。

```vbnet
Public Class Singleton
    ' ① Private 构造函数：外部彻底无法 New，只有类自己能创建
    Private Sub New()
        m_Data = "我是唯一实例"
    End Sub

    ' ② Shared ReadOnly 字段 + 初始化器：
    '    CLR 保证首次访问类型时初始化、且只初始化一次（线程安全）
    Private Shared ReadOnly m_Instance As New Singleton()

    Private m_Data As String

    ' ③ Shared 访问点：类级别成员，无需函数包装
    Public Shared ReadOnly Property Instance As Singleton
        Get
            Return m_Instance
        End Get
    End Property

    Public Property Data As String
        Get
            Return m_Data
        End Get
        Set(value As String)
            m_Data = value
        End Set
    End Property
End Class

' 演示
Dim s1 As Singleton = Singleton.Instance
Dim s2 As Singleton = Singleton.Instance
s1.Data = "已修改"
Console.WriteLine(s2.Data)   ' 已修改（同一个对象）
```

**VB.NET 版说明**：
- **真正的禁止外部 New**：`Private Sub New()` 让构造函数对外不可见，外部代码写 `New Singleton()` 直接编译报错——单例的"禁止创建"约束在编译期就强制生效，而非靠开发者自觉。
- **无需全局变量、无需函数包装**：`Shared` 成员属于类本身而非实例，直接通过 `Singleton.Instance` 访问。对比前两版：传统版需要 `gInstance` 模块级全局变量 + `GetInstance()` 函数；Axon 版需要 `Static instance` + `GetInstance()` 函数。
- **线程安全由 CLR 兜底**：`Shared ReadOnly m_Instance As New Singleton()` 依赖 CLR 的静态字段初始化器——CLR 在首次访问该类型时自动加锁初始化，且保证只初始化一次。传统版和 Axon 版在多线程下都不是安全的（前者无锁，后者 `If instance Is Nothing` 检查存在竞态）。
- **唯一残留**：这是饿汉式（首次访问类型即初始化）。若要真正的懒加载（首次调用 `Instance` 才初始化），VB.NET 还有 `Lazy(Of Singleton)` 写法，但那属于优化范畴，不影响模式地道性。

**三版对照**：

| 维度 | 传统 VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| 唯一实例存储 | 模块级全局变量 | 函数内 `Static` | 类内 `Shared ReadOnly` |
| 访问方式 | `GetInstance()` 函数 | `GetInstance()` 函数 | `Singleton.Instance` 属性 |
| 禁止外部 `New` | 不能 | 不能 | `Private Sub New()` 编译期强制 |
| 封装性 | 差（全局变量泄漏） | 中（封装进函数） | 好（封装进类） |
| 线程安全 | 否 | 否 | 是（CLR 保证） |
---