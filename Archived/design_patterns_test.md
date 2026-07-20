# 23个设计模式_VBScript版

## 1. 单例模式 (Singleton)

### ClassicASP 实现

```vbscript
' 01_singleton.vbs
Class Singleton
    Private m_value

    Private Sub Class_Initialize()
        m_value = 0
    End Sub

    Public Property Let Value(v)
        m_value = v
    End Property

    Public Property Get Value()
        Value = m_value
    End Property
End Class

' 测试代码
Dim obj1, obj2
Set obj1 = New Singleton
obj1.Value = 10
Set obj2 = New Singleton
obj2.Value = 20

WScript.Echo "ClassicASP Singleton Test:"
WScript.Echo "obj1.Value = " & obj1.Value
WScript.Echo "obj2.Value = " & obj2.Value
If obj1.Value = 10 And obj2.Value = 20 Then
    WScript.Echo "SUCCESS: Singleton properties work correctly"
Else
    WScript.Echo "FAIL: Singleton properties incorrect"
End If
```

### AxonASP 实现

```vbscript
' 01_singleton.asp
Class Singleton
    Private Shared m_instance
    Private m_value

    Private Sub Class_Initialize()
        m_value = 0
    End Sub

    Public Function GetInstance()
        If m_instance Is Nothing Then
            Set m_instance = New Singleton
        End If
        Set GetInstance = m_instance
    End Function

    Public Property Let Value(v)
        m_value = v
    End Property

    Public Property Get Value()
        Value = m_value
    End Property
End Class

' 测试代码
Dim obj1, obj2
Set obj1 = New Singleton
obj1.Value = 10
Set obj2 = New Singleton
obj2.Value = 20

Response.Write "AxonASP Singleton Test:" & vbCrLf
Response.Write "obj1.Value = " & obj1.Value & vbCrLf
Response.Write "obj2.Value = " & obj2.Value & vbCrLf
If obj1.Value = 10 And obj2.Value = 20 Then
    Response.Write "SUCCESS: Singleton properties work correctly"
Else
    Response.Write "FAIL: Singleton properties incorrect"
End If
```

## 2. 工厂模式 (Factory)

### ClassicASP 实现

```vbscript
' 02_factory.vbs
Class Product
    Private m_name

    Public Property Let Name(v)
        m_name = v
    End Property

    Public Property Get Name()
        Name = m_name
    End Property
End Class

Class Factory
    Public Function CreateProduct(name)
        Dim p
        Set p = New Product
        p.Name = name
        Set CreateProduct = p
    End Function
End Class

' 测试代码
Dim f, p1, p2
Set f = New Factory
Set p1 = f.CreateProduct("Product A")
Set p2 = f.CreateProduct("Product B")

WScript.Echo "ClassicASP Factory Test:"
WScript.Echo "p1.Name = " & p1.Name
WScript.Echo "p2.Name = " & p2.Name
If p1.Name = "Product A" And p2.Name = "Product B" Then
    WScript.Echo "SUCCESS: Factory created correct products"
Else
    WScript.Echo "FAIL: Products not created correctly"
End If
```

### AxonASP 实现

```vbscript
' 02_factory.asp
Class Product
    Private m_name

    Public Property Let Name(v)
        m_name = v
    End Property

    Public Property Get Name()
        Name = m_name
    End Property
End Class

Class Factory
    Public Function CreateProduct(name)
        Dim p
        Set p = New Product
        p.Name = name
        Set CreateProduct = p
    End Function
End Class

' 测试代码
Dim f, p1, p2
Set f = New Factory
Set p1 = f.CreateProduct("Product A")
Set p2 = f.CreateProduct("Product B")

Response.Write "AxonASP Factory Test:" & vbCrLf
Response.Write "p1.Name = " & p1.Name & vbCrLf
Response.Write "p2.Name = " & p2.Name & vbCrLf
If p1.Name = "Product A" And p2.Name = "Product B" Then
    Response.Write "SUCCESS: Factory created correct products"
Else
    Response.Write "FAIL: Products not created correctly"
End If
```

## 3. 策略模式 (Strategy)

### ClassicASP 实现

```vbscript
' 03_strategy.vbs
Class Context
    Private m_strategy

    Public Sub SetStrategy(strategy)
        Set m_strategy = strategy
    End Sub

    Public Function Execute(a, b)
        Execute = m_strategy.Calculate(a, b)
    End Function
End Class

Class AddStrategy
    Public Function Calculate(a, b)
        Calculate = a + b
    End Function
End Class

Class MultiplyStrategy
    Public Function Calculate(a, b)
        Calculate = a * b
    End Function
End Class

' 测试代码
Dim ctx, addS, mulS
Set ctx = New Context
Set addS = New AddStrategy
Set mulS = New MultiplyStrategy

ctx.SetStrategy addS
WScript.Echo "ClassicASP Strategy Test:"
WScript.Echo "Add 3+5 = " & ctx.Execute(3, 5)

ctx.SetStrategy mulS
WScript.Echo "Multiply 3*5 = " & ctx.Execute(3, 5)

If ctx.Execute(3,5) = 15 Then
    ctx.SetStrategy addS
    If ctx.Execute(3,5) = 8 Then
        WScript.Echo "SUCCESS: Strategy pattern works correctly"
    Else
        WScript.Echo "FAIL: Add strategy incorrect"
    End If
Else
    WScript.Echo "FAIL: Multiply strategy incorrect"
End If
```

### AxonASP 实现

```vbscript
' 03_strategy.asp
Class Context
    Private m_strategy

    Public Sub SetStrategy(strategy)
        Set m_strategy = strategy
    End Sub

    Public Function Execute(a, b)
        Execute = m_strategy.Calculate(a, b)
    End Function
End Class

Class AddStrategy
    Public Function Calculate(a, b)
        Calculate = a + b
    End Function
End Class

Class MultiplyStrategy
    Public Function Calculate(a, b)
        Calculate = a * b
    End Function
End Class

' 测试代码
Dim ctx, addS, mulS
Set ctx = New Context
Set addS = New AddStrategy
Set mulS = New MultiplyStrategy

ctx.SetStrategy addS
Response.Write "AxonASP Strategy Test:" & vbCrLf
Response.Write "Add 3+5 = " & ctx.Execute(3, 5) & vbCrLf

ctx.SetStrategy mulS
Response.Write "Multiply 3*5 = " & ctx.Execute(3, 5) & vbCrLf

If ctx.Execute(3,5) = 15 Then
    ctx.SetStrategy addS
    If ctx.Execute(3,5) = 8 Then
        Response.Write "SUCCESS: Strategy pattern works correctly"
    Else
        Response.Write "FAIL: Add strategy incorrect"
    End If
Else
    Response.Write "FAIL: Multiply strategy incorrect"
End If
```