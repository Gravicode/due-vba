VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SystemController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

    Private serialPort As New SerialInterface
    Private print_posx As Integer
    Private displayText() As String
   
    Public Enum ResetOption
        SystemReset = 0
        Bootloader = 1
    End Enum
Dim timer1 As New TimerCls

    Const DISPLAY_MAX_LINES As Integer = 8
    Const DISPLAY_MAX_CHARACTER_PER_LINE As Integer = 21
 Private Sub Class_Initialize()
     displayText = Array("", "", "", "", "", "", "", "")
     print_posx = 0
End Sub
 

    Public Sub Reset(ByVal option1 As ResetOption)
        Dim cmd As String
        Dim op
        If (option1 = ResetOption.Bootloader) Then
        op = 1
        Else
        op = 0
        End If
        cmd = "reset(" & op & ")"
        Call serialPort.WriteCommand(cmd)
        Call serialPort.Disconnect
    End Sub

    Public Function GetTickMicroseconds() As Integer
        Dim cmd As String
        cmd = "print(tickus())"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        If res.success Then
            On Error GoTo next1
            GetTickMicroseconds = CInt(res.respone)
next1:
        End If
        GetTickMicroseconds = -1
    End Function

    Public Function GetTickMilliseconds() As Integer
        Dim cmd As String
        cmd = "print(tickms())"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        If res.success Then
                  On Error GoTo next1
                GetTickMilliseconds = CInt(res.respone)
next1:
        End If
        GetTickMilliseconds = -1
    End Function

    Public Function Beep(ByVal pin As Integer, ByVal frequency As Integer, ByVal duration As Integer) As Boolean
        If frequency < 0 Or frequency > 10000 Then
            Err.Raise "Frequency is within range[0,10000] Hz"
        End If
        If duration < 0 Or duration > 1000 Then
            Err.Raise "duration is within range[0,1000] millisecond"
        End If

        Dim cmd As String
        cmd = "beep(" & pin & ", " & frequency & ", " & duration & ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        Beep = res.success
    End Function

    Private Sub PrnChar(ByVal c As Char)
        If (print_posx = SystemController.DISPLAY_MAX_CHARACTER_PER_LINE And c <> ControlChars.Cr And c <> ControlChars.Lf) Then
            Exit Sub
        End If

        If c = ControlChars.Cr Or c = ControlChars.Lf Then
            print_posx = 0
            Dim i As Integer

            For i = 1 To SystemController.DISPLAY_MAX_LINES - 1
                displayText(i - 1) = displayText(i)
            Next

            displayText(SystemController.DISPLAY_MAX_LINES - 1) = ""
        Else
            displayText(SystemController.DISPLAY_MAX_LINES - 1) = displayText(SystemController.DISPLAY_MAX_LINES - 1) & c
            print_posx = print_posx + 1
        End If
    End Sub

    Private Sub PrnText(ByVal text As String, ByVal newline As Boolean)
        Dim i As Integer
        
        For i = 0 To text.length - 1
            PrnChar (text(i))
        Next

        Dim display As New DisplayController
        
        Call display.Clear(0)

        For i = 0 To displayText.length - 1
            If displayText(i) <> "" Then
                Call display.DrawText(displayText(i), 1, 0, i * 8)
            End If
        Next

        Call display.Show

        If newline Then
            PrnChar (ControlChars.Cr)
        End If
    End Sub

    Public Function PrintText(ByVal text As Object) As Boolean
        Debug.Print text

        If TypeName(text) = "String" Then
            Call PrnText(text, False)
        Else
            Call PrnText(text.ToString(), False)
        End If

        PrintText = True
    End Function

    Public Function Println(ByVal text As Object) As Boolean
        Console.WriteLine (text)

        If TypeName(text) = "String" Then
            Call PrnText(text, True)
        Else
            Call PrnText(text.ToString(), True)
        End If

        Println = True
    End Function

    Public Function Wait(ByVal millisecond As Integer) As Boolean
        Dim cmd As String
        cmd = "wait(" & millisecond & ")"
        Call serialPort.WriteCommand(cmd)
        'Threading.Thread.Sleep (millisecond / 1000)
        Call timer1.WasteTime(millisecond / 1000)
        
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        Wait = res.success
    End Function

