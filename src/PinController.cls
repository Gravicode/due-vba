VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PinController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim time As Date

    Private Function get_button_a() As Integer
        get_button_a 97
    End Function

    Private Function get_button_b() As Integer
        get_button_b = 98
    End Function

    Private Function get_led() As Integer
        get_led = 108
    End Function

    Private Function get_piezo() As Integer
        get_piezo = 112
    End Function

    Private Function get_pullnone() As Integer
        get_pullnone = 0
    End Function

    Private Function get_pullup() As Integer
        get_pullup = 1
    End Function

    Private Function get_pulldown() As Integer
        get_pulldown = 2
    End Function

    Private Function set_empty(ByVal value As Integer) As Integer
        set_empty = 0
    End Function
    
Public Property Get ButtonA() As String
    ButtonA = get_button_a()
End Property

Public Property Get ButtonB() As String
    ButtonB = get_button_b()
End Property

Public Property Get Led() As String
    Led = get_led()
End Property

Public Property Get Piezo() As String
    Piezo = get_piezo()
End Property

Public Property Get PullNone() As String
    PullNone = get_pullnone()
End Property

Public Property Get PullUp() As String
    PullUp = get_pullup()
End Property

Public Property Get PullDown() As String
    PullDown = get_pulldown()
End Property

