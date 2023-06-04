VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UartController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private serialport As New SerialInterface
    
    Public Function Enable(ByVal baudrate As Integer) As Boolean
        Dim cmd As String
        cmd = "uartinit(" & baudrate & ")"
        serialport.WriteCommand (cmd)
        Dim res As CmdRespone
        res = serialport.ReadRespone()
        Enable = res.success
    End Function

    Public Function WriteData(ByVal data As String) As Boolean
        Dim cmd As String
        cmd = "uartwrite(" & data & ")"
        Call serialport.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialport.ReadRespone()
        WriteData = res.success
    End Function

    Public Function BytesToRead() As Integer
        Dim cmd As String
        cmd = "x=uartcount():print(x)"
        serialport.WriteCommand (cmd)
        Dim res As CmdRespone
        res = serialport.ReadRespone()
        If res.success Then
            On Error GoTo next1
            Dim ready As Integer
            ready = CInt(res.respone)
                BytesToRead = ready
next1:
        End If
        Err.Raise "BytesToRead error!"
    End Function

    Public Function Read() As Integer
        Dim cmd As String
        cmd = "x=uartread():print(x)"
        Call serialport.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialport.ReadRespone()
        If res.success Then
           On Error GoTo next2
                Dim data As Integer
                data = CInt(res.respone)
                Read = data
next2:
        End If
        Err.Raise "Uart receving error!"
    End Function

