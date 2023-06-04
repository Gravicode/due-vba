VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServoController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim serialPort As New SerialInterface


    Public Function SetServo(ByVal pin As Integer, ByVal position As Integer) As Boolean
        If pin < 0 Or pin >= Me.serialPort.DeviceConfig.MaxPinIO Then
            Err.Raise "Invalid pin"
        End If
        If position < 0 Or position > 180 Then
            Err.Raise "Position must be in the range 0..180"
        End If

        Dim cmd As String
        cmd = "servoset(" + pin + ", " + position + ")"
        Call serialPort.WriteCommand(cmd)

        Dim response As CmdRespone
        response = serialPort.ReadRespone()

        SetServo = response.success
    End Function
