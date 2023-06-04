VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeviceConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public p_IsPulse As Boolean
Public p_IsFlea As Boolean
Public p_IsPico As Boolean
Public p_IsEdge As Boolean
Public p_MaxPinIO As Integer
Public p_MaxPinAnalog As Integer
               

Public Property Get IsPulse() As Boolean
    IsPulse = p_IsPulse
End Property

Public Property Let IsPulse(ByVal val1 As Boolean)
    p_IsPulse = val1
End Property

Public Property Get IsFlea() As Boolean
    IsFlea = p_IsFlea
End Property

Public Property Let IsFlea(ByVal val1 As Boolean)
    p_IsFlea = val1
End Property

Public Property Get IsPico() As Boolean
    IsPico = p_IsPico
End Property

Public Property Let IsPico(ByVal val1 As Boolean)
    p_IsPico = val1
End Property

Public Property Get IsEdge() As Boolean
    IsEdge = p_IsEdge
End Property

Public Property Let IsEdge(ByVal val1 As Boolean)
    p_IsEdge = val1
End Property

Public Property Get MaxPinIO() As Integer
    MaxPinIO = p_MaxPinIO
End Property

Public Property Let MaxPinIO(ByVal val1 As Integer)
    p_MaxPinIO = val1
End Property

Public Property Get MaxPinAnalog() As Integer
    MaxPinAnalog = p_MaxPinAnalog
End Property

Public Property Let MaxPinAnalog(ByVal val1 As Integer)
    p_MaxPinAnalog = val1
End Property

 Private Sub Class_Initialize()
    Me.IsPulse = False
    Me.IsFlea = False
    Me.IsPico = False
    Me.IsEdge = False
    Me.MaxPinIO = 0
    Me.MaxPinAnalog = 0
End Sub
