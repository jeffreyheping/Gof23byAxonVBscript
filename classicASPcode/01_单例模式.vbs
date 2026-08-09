Option Explicit
Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
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

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
