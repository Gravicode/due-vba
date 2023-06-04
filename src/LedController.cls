VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LedController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Dim serialPort As New SerialInterface

    Public Function SetLed(ByVal highPeriod As Integer, ByVal lowPeriod As Integer, ByVal count As Integer) As Boolean
        Dim cmd As String
        cmd = "led(" & highPeriod & "," & lowPeriod & "," & count & ")"
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadResponse()
        SetLed = res.success
    End Function
