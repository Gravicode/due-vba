VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SerialInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

    Public CommandCompleteText As String
    Public DefaultBaudRate As Integer
    Public DeviceConfig As DeviceConfiguration
    Private leftOver As String
    Private ReadTimeout As Integer
    'Private portName As String
    Private echo As Boolean

    Public TransferBlockSizeMax As Integer
    
    Public TransferBlockDelay As Double
    'Private port As serialPort

    Private Sub Class_Initialize()
        CommandCompleteText = ">"
        DefaultBaudRate = 115200
        ReadTimeout = 3
        leftOver = ""
         TransferBlockDelay = 0.005
        TransferBlockSizeMax = 512
        'Me.portName = portName
        echo = True
    End Sub

    Public Sub Connect()
        'port = New SerialPort(portName, DefaultBaudRate, Parity.None, 8, StopBits.One)
        'port.ReadTimeout = ReadTimeout
        'port.Open()
        leftOver = ""
        'Thread.Sleep (100)
        
        Synchronize
    End Sub

    Public Sub Disconnect()
        
    End Sub

    Public Sub Synchronize()
        Dim cmd(0) As Byte
        Dim timer1 As TimerCls
        
        cmd(0) = 127

       Call WriteRawData(cmd, 0, 1)

        Dim tryCount As Integer
        tryCount = 3
        Do While tryCount > 0
            Call timer1.WasteTime(10)
            leftOver = ""
            'port.DiscardInBuffer()
            'port.DiscardOutBuffer()
            On Error GoTo ErrHandler
            Dim version As String
            version = GetVersion()
                If version <> "" And version(2) = "." And version(4) = "." Then
                    Exit Do
                End If
ErrHandler:
            tryCount = tryCount - 1
        Loop
        'port.ReadTimeout = orig
    End Sub

    Public Sub TurnEchoOff()
        If Not echo Then
            Exit Sub
        End If
        WriteCommand ("echo(0)")
        ReadRespone
        echo = False
    End Sub

    Public Function GetVersion() As String
        Dim command As String
        command = "version()"
        Call WriteCommand(command)

        Dim version As CmdRespone
        version = ReadRespone()

        If version.success Then
            If echo And InStr(1, version.respone, command) <> 0 Then
                ' echo is on => need to turn off
                Call TurnEchoOff
                'port.DiscardInBuffer()
                'port.DiscardOutBuffer()
                version.respone = Right(version.respone, version.respone.Length - command.Length)
            End If
        End If

        GetVersion = version.respone
    End Function

    Public Function RemoveEchoRespone(ByVal respone As String, ByVal cmd As String) As String
        If InStr(1, respone, cmd) <> 0 Then
            respone = Right(respone, respone.Length - cmd.Length)
            'respone.Substring(cmd.Length)
        End If

        RemoveEchoRespone = respone
    End Function

    Public Sub DiscardInBuffer()
        'port.DiscardInBuffer()
    End Sub

    Public Sub DiscardOutBuffer()
        'port.DiscardOutBuffer()
    End Sub

    Public Sub WriteCommand(ByVal command As String)
        'DiscardInBuffer()
        'DiscardOutBuffer()
        Call WriteLine(command)
    End Sub

    Private Sub WriteLine(ByVal str As String)
        str = str & vbCrLf
        Call WriteCommand(str)
        'port.Write(Encoding.UTF8.GetBytes(str), 0, str.Length)
    End Sub

    Public Function ReadRespone() As CmdRespone
        Dim str As String
        str = leftOver
        Dim startTime
        startTime = CDate(Now())

        Dim respone As New CmdRespone
        
        Dim timeNow
        timeNow = CDate(Now)
        Do While True
            Dim data As String
            data = ReadData()
    
            str = str & data

            str = str.Replace(vbCrLf, "")
            str = str.Replace(vbCr, "")
            str = str.Replace(vbLf, "")
            Dim idx1 As Integer
            idx1 = InStr(1, str, ">")
            Dim idx2 As Integer
            idx2 = InStr(1, str, "&")

            If idx1 = -1 Then
                idx1 = str.IndexOf("$")
            End If

            If idx1 = -1 And idx2 = -1 Then
                GoTo SkipWhile
                'continue do
            End If

            Dim idx As Integer
            If (idx1 = -1) Then idx 1 = idx2 Else idx1 = idx1
            

            leftOver = str.Substring(idx + 1)
            respone.success = True
            respone.respone = str.Substring(0, idx)
            Dim idx3 As Integer
            idx3 = InStr(1, str, "!")
            If idx3 <> -1 And respone.respone.Contains("error") Then
                respone.success = False
            End If

            If idx3 <> -1 And InStr(1, respone.respone, "unknown") <> 0 Then
                respone.success = False
            End If

            ReadRespone = respone
            Exit Do
             
            timeNow = CDate(Now)
            Total_Seconds = (timeNow - startTime) * 24 * 3600
            If (Total_Seconds > ReadTimeout) Then Exit Do
            SkipWhile
        Loop

        leftOver = ""

        'port.DiscardInBuffer()
        'port.DiscardOutBuffer()

        respone.success = False
        respone.respone = ""

        ReadRespone = respone
    End Function
    Private Sub WriteRawData(ByVal buffer As Byte, ByVal offset As Integer, ByVal count As Integer)
        'port.Write(buffer, offset, count)
    End Sub
    Public Function ReadRespone2() As CmdRespone
        Dim str As String
        str = leftOver
        Dim startTime
        startTime = CDate(Now())
        
        Dim timeNow
        timeNow = CDate(Now)

        Dim respone As New CmdRespone

        Do While True
            Dim data As String
            data = ReadData()

            str = str & data

            'str = str.Replace("", "")
            'str = str.Replace("\r", "")
            Dim idx1 As Integer
            idx1 = InStr(1, str, ">")
            Dim idx2 As Integer
            idx2 = isntr(1, str, "&")

            If idx1 = -1 Then
                idx1 = str.IndexOf("$")
            End If

            If idx1 = -1 And idx2 = -1 Then
                GoTo Skip
            End If

            Dim idx As Integer
            If (idx1 = -1) Then idx1 = idx2 Else idx1 = idx1

            leftOver = str.Substring(idx + 1)
            respone.success = True
            respone.respone = str.Substring(0, idx)
            Dim idx3 As Integer
            idx3 = InStr(1, str, "!")
            If idx3 <> -1 And respone.respone.Contains("error") Then
                respone.success = False
            End If

            If idx3 <> -1 And respone.respone.Contains("unknown") Then
                respone.success = False
            End If

            ReadRespone2 = respone
            
            timeNow = CDate(Now)
            Total_Seconds = (timeNow - startTime) * 24 * 3600
            If (Total_Seconds > ReadTimeout) Then Exit Do
Skip:
        Loop

        leftOver = ""

        'port.DiscardInBuffer()
        'port.DiscardOutBuffer()

        respone.success = False
        respone.respone = ""

        ReadRespone2 = respone
    
    End Function

Public Function ReadData() As String
 Dim data As String
 data = Sheets("Data In").Range("B5").Value
 ReadData = data
End Function

Public Sub WriteCommand(ByVal cmd As String)
     Sheets("Data Out").Range("A5") = cmd
End Sub
