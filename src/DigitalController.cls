VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DigitalController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim serialPort As New SerialInterface
    Public Function Read(ByVal pin As Integer, ByVal inputType As Integer) As Boolean
        If pin = Asc("a") Or pin = Asc("A") Then
            pin = 97
        End If

        If pin = Asc("b") Or pin = Asc("B") Then
            pin = 98
        End If

        If pin < 0 Or (pin >= serialPort.DeviceConfig.MaxPinIO And pin <> 97 And pin <> 98 And pin <> 108) Then 'A, B, Led
            Err.Raise "Invalid pin"
        End If

        Dim pull As String
        pull = "0"
        If inputType = 1 Then
            pull = "1"
        ElseIf inputType = 2 Then
            pull = "2"
        End If

        Dim cmd As String
        cmd = "print(dread(" & pin & "," & pull & "))"
        Call serialPort.WriteCommand(cmd)

        Dim response As CmdRespone
        response = serialPort.ReadRespone()
        Read = False
        If response.success Then
        On Error GoTo Down
                If (CInt(response.respone) = 1) Then Read = True
Down:
        End If
    End Function

    Public Function WriteData(ByVal pin As Integer, ByVal value As Boolean) As Boolean
        If pin = Asc("l") Or pin = Asc("L") Then
            pin = 108
        End If

        If pin < 0 Or (pin >= serialPort.DeviceConfig.MaxPinIO And pin <> 108) Then ' Led
            Err.Raise "Invalid pin"
        End If

        Dim cmd As String
        If value Then
             valstr = 1
         Else
             valstr = 0
         End If
        cmd = "dwrite(" & pin & "," & valstr & ")"
        Call serialPort.WriteCommand(cmd)

        Dim response As CmdRespone
        response = serialPort.ReadRespone()

        WriteData = response.success
    End Function

