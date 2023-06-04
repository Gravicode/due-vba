VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Analog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim serialPort As New SerialInterface
Private Fixed_Frequency As Integer

Public Property Get Frequency() As String
    Frequency = Fixed_Frequency
End Property

Public Property Let Frequency(ByVal Freq As Integer)
    Fixed_Frequency = Freq
End Property

 Private Sub Class_Initialize()
    Fixed_Frequency = 50
End Sub

    Public Function Read(ByVal pin As Integer) As Integer
        If pin < 0 Or pin >= serialPort.DeviceConfig.MaxPinAnalog Then
            Debug.Print "Invalid pin"
        End If

        Dim cmd As String
        cmd = "print(aread(" & pin & "))"
        serialPort.WriteCommand (cmd)

        Dim res As String
        res = serialPort.ReadRespone()

        Read = -1
        If res.success Then
            On Error GoTo out
                Read = CInt(res.respone)
out:
        End If

    End Function

    Public Function WriteData(ByVal pin As Integer, ByVal duty_cycle As Integer) As Boolean
        If pin = "l" Or pin = "L" Then
            pin = 108
        End If

        If pin < 0 Or (pin >= serialPort.DeviceConfig.MaxPinIO And pin <> 108) Then
            Err.Raise "Invalid pin"
        End If

        If duty_cycle < 0 Or duty_cycle > 1000 Then
            Err.Raise "Duty cycle must be in the range 0..1000"
        End If

        Dim cmd As String
        cmd = "awrite(" & pin.ToString() & ", " & duty_cycle.ToString() & ")"
        serialPort.WriteCommand (cmd)

        Dim res As respone
        res = serialPort.ReadRespone()
        If res.success Then
            WriteData = True
        End If

        WriteData = False
    End Function

