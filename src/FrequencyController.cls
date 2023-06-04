VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FrequencyController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Dim MaxFrequency As Integer
    Dim MinFrequency As Integer
    Private serialPort As New SerialInterface
    Private Sub Class_Initialize()
      
        MaxFrequency = 1000000
        MinFrequency = 16
    End Sub

    Public Function WriteData(ByVal pin As String, ByVal frequency As Integer, Optional ByVal duration_ms As Integer = 0, Optional ByVal dutycyle As Integer = 500) As Boolean
        If frequency < MinFrequency Or frequency > MaxFrequency Then
            Err.Raise "Frequency must be in range 16Hz..1000000Hz"
        End If

        If duration_ms > 99999999 Then
            Err.Raise "duration_ms must be in range 0..99999999"
        End If

        If dutycyle < 0 Or dutycyle > 1000 Then
            Err.Raise "dutycyle must be in range 0..1000"
        End If

        If pin = "p" Or pin = "P" Then
            pin = 112
        End If

        Dim cmd As String
        cmd = "freq(" + pin + ", " + frequency + ", " + duration_ms + ", " + dutycyle + ")"
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        WriteData = res.success
    End Function
