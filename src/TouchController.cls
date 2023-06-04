VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TouchController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

    Private serialPort As New SerialInterface
    
    Public Function Read(ByVal pin As Integer) As Boolean
        Dim cmd As String
        cmd = "print(touchread(" + pin + "))"
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        Dim val As Boolean
        val = False
        If res.success Then
            On Error GoTo next1
                val = CInt(res.respone) = 1
                Read = val
next1:
        End If
        Read = val
    End Function

