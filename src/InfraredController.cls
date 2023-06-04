VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfraredController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Dim serialPort As New SerialInterface


    Public Function Read() As Integer
        Dim cmd As String
        cmd = "print(irread())"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        If res.success Then
            On Error GoTo next1
                Read = CInt(res.respone)
next1:
        End If
        Read = -1
    End Function

    Public Function enable(ByVal enable As Boolean) As Boolean
        Dim en As Integer
        en = 0
        If enable = True Then
            en = 1
        End If
        Dim cmd As String
        cmd = "irenable(" + str(en) + ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        If res.success Then
            enable = True
        Else
            enable = False
        End If
    End Function
