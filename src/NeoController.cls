VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NeoController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim MAX_LED_NUM As Integer
Private serialPort As New SerialInterface
Public SupportLedNumMax As Integer
Dim timer1 As New TimerCls
Dim bit As New BitWiseCls

Private Sub Class_Initialize()
    MAX_LED_NUM = 256
    SupportLedNumMax = MAX_LED_NUM
End Sub

    Public Function Show(ByVal pin As Integer, ByVal count As Integer) As Boolean
        Dim cmd As String
        cmd = "neoshow(" & pin & ", " & count & ")"
        Call serialPort.WriteCommand(cmd)

        ' each led need 1.25us delay blocking mode
        Dim delay As Double
        delay = (MAX_LED_NUM * 3 * 8 * 1.25) / 1000000
        'Threading.Thread.Sleep (delay)
        Call timer1.WasteTime(delay)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        Show = res.success
    End Function

    Public Function Clear() As Boolean
        Dim cmd As String
        cmd = "neoclear()"
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        Clear = res.success
    End Function
    

    Public Function SetColor(ByVal id As Integer, ByVal color As Integer) As Boolean
        Dim red As Integer
        red = bit.ShiftRight(color, 16) And &HFF
        Dim green As Integer
        green = bit.ShiftRight(color, 8) And &HFF
        Dim blue As Integer
        blue = bit.ShiftRight(color, 0) And &HFF

        If id < 0 Or id > MAX_LED_NUM Then
            SetColor = False
        End If

        Dim cmd As String
        cmd = "neoset(" & id & "," & red & "," & green & "," & blue & ")"
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        SetColor = res.success
    End Function

    Public Function SetMultiple(ByVal pin As Integer, ByVal color As Object) As Boolean
        If color.length > MAX_LED_NUM Then
            SetMultiple = False
        End If

        Dim length As Integer
        length = color.length
        Dim offset As Integer
        offset = 0

        Dim data(length * 3 - 1) As Byte
        Dim i As Integer

        For i = offset To length + offset - 1
            data((i - offset) * 3 + 0) = bit.ShiftRight(color(i), 16) And &HFF
            data((i - offset) * 3 + 1) = bit.ShiftRight(color(i), 8) And &HFF
            data((i - offset) * 3 + 2) = bit.ShiftRight(color(i), 0) And &HFF
        Next

        Dim cmd As String
        cmd = "neostream(" & pin & ", " & data.length & ")"
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        If res.success Then
            Call serialPort.WriteRawData(data, 0, data.length)
            res = serialPort.ReadRespone()
        End If

        SetMultiple = res.success
    End Function
