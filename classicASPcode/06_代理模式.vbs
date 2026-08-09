Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
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

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
