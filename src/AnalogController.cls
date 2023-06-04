VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalogController"
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

    Public Function Read(ByVal Pin As Integer) As Integer
        If Pin < 0 Or Pin >= serialPort.DeviceConfig.MaxPinAnalog Then
            Debug.Print "Invalid pin"
        End If

        Dim cmd As String
        cmd = "print(aread(" & Pin & "))"
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

    Public Function WriteData(ByVal Pin As Integer, ByVal duty_cycle As Integer) As Boolean
        If Pin = "l" Or Pin = "L" Then
            Pin = 108
        End If

        If Pin < 0 Or (Pin >= serialPort.DeviceConfig.MaxPinIO And Pin <> 108) Then
            Err.Raise "Invalid pin"
        End If

        If duty_cycle < 0 Or duty_cycle > 1000 Then
            Err.Raise "Duty cycle must be in the range 0..1000"
        End If

        Dim cmd As String
        cmd = "awrite(" & Pin.ToString() & ", " & duty_cycle.ToString() & ")"
        serialPort.WriteCommand (cmd)

        Dim res As respone
        res = serialPort.ReadRespone()
        If res.success Then
            WriteData = True
        End If

        WriteData = False
    End Function

