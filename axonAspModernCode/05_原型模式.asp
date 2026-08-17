<%
Option Explicit
' 克隆接口
Class ICloneable
    Public Function Clone As ICloneable
    End Function
End Class

' 简历类：实现 ICloneable 接口
Class MyResume
    Implements ICloneable
    Public Name As String
    Public Age As Integer
    Public Skills

    ' 深拷贝克隆：逐字段复制，数组逐元素拷贝
    Public Function ICloneable_Clone As ICloneable
        Dim copy As MyResume
        Set copy = New MyResume
        copy.Name = Me.Name
        copy.Age = Me.Age
        Dim ub As Integer, i As Integer, arr
        ub = UBound(Me.Skills)
        ReDim arr(ub)
        For i = 0 To ub
            arr(i) = Me.Skills(i)
        Next
        copy.Skills = arr
        Set ICloneable_Clone = copy
    End Function
End Class

' 演示：通过接口引用调用 Clone
Dim r1 As MyResume
Dim r2 As ICloneable
Dim r2Copy As MyResume
Set r1 = New MyResume
r1.Name = "张三"
r1.Age = 25
r1.Skills = Array("VBScript", "HTML")
Set r2 = r1.Clone()
Set r2Copy = r2
r2Copy.Name = "李四"
r2Copy.Skills(0) = "JavaScript"

Response.Write(r1.Name & " " & r1.Skills(0))   ' 张三 VBScript

Response.Write(r2Copy.Name & " " & r2Copy.Skills(0))   ' 李四 JavaScript
%>