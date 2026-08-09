<%
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
%>