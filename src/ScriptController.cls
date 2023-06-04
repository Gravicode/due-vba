VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScriptController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private serialPort As New SerialInterface
    Private loadscript As String
    Dim timer1 As New TimerCls
    
    Private Sub Class_Initialize()
        loadscript = ""
    End Sub

    Public Sub Run()
        Dim cmd As String
        cmd = "run"
        Call serialPort.WriteCommand(cmd)
        Call timer1.WasteTime(1)
        'Thread.Sleep (1)
    End Sub

    Public Function NewScript() As Boolean
        loadscript = ""
        Dim cmd As String
        cmd = "new"
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        NewScript = res.success
    End Function

    Public Function Load(ByVal script As String) As Boolean
        loadscript = loadscript + script
        loadscript = loadscript + vbCrLf
    End Function

    Public Function Record() As Boolean
        If loadscript = "" Then
            Err.Raise "No script for recording."
        End If

        Dim script As String
        script = loadscript
        Dim cmd As String
        cmd = "pgmstream()"

        Dim raw() As Variant
        raw = System.text.Encoding.ASCII.GetBytes(script)
        Dim data(UBound(raw)) As Byte
        data = raw

        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        If Not res.success Then
            Record = False
        End If

        Call serialPort.WriteRawData(data, 0, data.length)

        res = serialPort.ReadRespone()

        loadscript = ""
        Record = res.success
    End Function

    Private Function Load2(ByVal script As String) As Boolean
        Dim ret As Boolean
        ret = True
        Dim cmd As String
        cmd = "$"
        Call serialPort.WriteCommand(cmd)
        'Thread.Sleep (1)
        Call timer1.WasteTime(1)
        
        script = Replace(script, vbCr, "")

        Dim startIdx As Integer
        startIdx = 0
        Dim i As Integer

        For i = 0 To script.length - 1
            Dim subscript As String
            subscript = ""

            If script(i) = vbLf Then
                subscript = Mid(script, startIdx, i - startIdx)
                startIdx = i + 1
            ElseIf i = script.length - 1 Then
                subscript = Mid(script, startIdx, i - startIdx + 1)
            End If

            Call serialPort.WriteCommand(subscript)

            Dim res As CmdRespone
            res = serialPort.ReadRespone()

            If Not res.success Then
                ret = False
                Exit For
            End If
        Next

        cmd = ">"
        serialPort.WriteCommand (cmd)

        Dim res2 As CmdRespone
        res2 = serialPort.ReadRespone()

        Load2 = ret And res2.success
    End Function

    Public Function Read() As String
        Dim cmd As String
        cmd = "list"

        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone2()

        Read = res.respone
    End Function

    Public Function Execute(ByVal script As String) As Boolean
        Dim cmd As String
        cmd = script
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        Execute = res.respone
    End Function

    Public Function IsRunning() As Boolean
        Call serialPort.DiscardInBuffer

        Dim dataWrite(0) As Byte
        Dim dataRead(0) As Byte

        dataWrite(0) = &HFF
        dataRead(0) = &H0

        Call serialPort.WriteRawData(dataWrite, 0, 1)
        'Thread.Sleep (1)
        Call timer1.WasteTime(1)
        

        Dim count As Integer
        count = serialPort.ReadRawData(dataRead, 0, 1)

        If count = 0 Then
            ' if running, should received 0xff
            ' it not, need to send vbLf to clear 0xff that was sent.
            dataWrite(0) = 10
            Call serialPort.WriteRawData(dataWrite, 0, 1)

            Call serialPort.ReadRespone
        End If

        IsRunning = dataRead(0) = &HFF
    End Function

