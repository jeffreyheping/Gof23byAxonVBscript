Option Explicit
Dim Response: Set Response = New ResponseStub
' -- inject: ResponseStub class below user code --
' 简历类：包含姓名、年龄、技能数组
Class MyResume
    Public Name, Age, Skills

    ' 浅拷贝克隆：创建新 MyResume 并逐字段复制
    ' 数组元素逐值复制，当前示例中 Skills 为字符串数组，修改副本不影响原件
    ' 返回值：新的 MyResume 实例
    Public Function Clone
        Dim copy
        Set copy = New MyResume
        copy.Name = Me.Name
        copy.Age = Me.Age
        Dim ub, i, arr
        ub = UBound(Me.Skills)
        ReDim arr(ub)
        For i = 0 To ub
            arr(i) = Me.Skills(i)
        Next
        copy.Skills = arr
        Set Clone = copy
    End Function
End Class

' 演示：克隆后修改副本，原件不受影响
Dim r1, r2
Set r1 = New MyResume
r1.Name = "张三"
r1.Age = 25
r1.Skills = Array("VBScript", "HTML")

Set r2 = r1.Clone
r2.Name = "李四"
r2.Skills(0) = "JavaScript"

Response.Write(r1.Name & " " & r1.Skills(0))   ' 张三 VBScript

Response.Write(r2.Name & " " & r2.Skills(0))   ' 李四 JavaScript

Class ResponseStub
    Public Sub Write(s)
        WScript.Echo s
    End Sub
End Class
