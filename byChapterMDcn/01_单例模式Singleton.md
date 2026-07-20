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
Response.Write s2.Data   ' 已修改（同一个对象）
```

**传统 VBScript 版妥协说明**：
- **无静态变量**：VBScript 的 Class 内部不支持 `Static` 变量，只能用脚本级（模块级）全局变量 `gInstance` 来保存唯一实例，破坏了类的封装性。
- **无法禁止外部 New**：VBScript 没有 `Private` 构造函数，外部代码随时可以 `New Singleton` 绕过 `GetInstance`，无法真正强制单例。

### Axon VBScript 版（支持 Static）

```vbscript
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
Response.Write s2.Data   ' 已修改
```

**Axon VBScript 版妥协说明**：
- AxonASP 的 `Static` 变量已支持对象引用，可以在函数内部保持唯一实例，无需模块级全局变量，封装性优于传统版本。残留限制：缺失语法点：**Private 构造函数**。外部仍可 `New Singleton` 绕过单例控制。Go 用**包级私有首字母小写**解决此问题——`type singleton` 不导出，外部只能通过 `GetInstance()` 获取。VBScript 的类没有访问控制，无法禁止外部 `New`。
---