VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim pinList As Variant

    Private serialPort As New SerialInterface
    
Private Sub Class_Initialize()
pinList = Array(97, 98, 65, 66, 0, 1, 2, "A", "B", "a", "b")

End Sub
   
Public Function IsContains(ByVal val1 As Integer) As Boolean
IsContains = False
For Each itm In pinList
    If itm.Value = val1 Then
        IsContains = True
        Exit For
    End If
Next
End Function
    Public Function enable(ByVal pin As Object, ByVal enable As Boolean) As Boolean
        If Not pinList.Contains(pin) Then
            Err.Raise "Invalid pin."
        End If

        Dim cmd As String
        cmd = "btnenable(" & pin & ", " & Convert.ToInt32(enable) & ")"

        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        enable = res.success
    End Function

    Public Function WasPressed(ByVal pin As Object) As Boolean
        If Not pinList.Contains(pin) Then
            Err.Raise "Invalid pin."
        End If

        Dim cmd As String
        cmd = "print(btndown(" & pin & "))"

        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        WasPressed = False
        If res.success Then
        On Error GoTo Down
                If (CInt(res.respone) = 1) Then WasPressed = True
Down:
        End If

        
    End Function

    Public Function IsReleased(ByVal pin As Object) As Boolean
        If Not pinList.Contains(pin) Then
            Err.Raise "Invalid pin."
        End If

        Dim cmd As String
        cmd = "print(btnup(" & pin & "))"

        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        WasPressed = False
        If res.success Then
             On Error GoTo Down2
             If (CInt(res.respone) = 1) Then WasPressed = True
Down2:
        End If

    End Function

