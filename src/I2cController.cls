VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "I2cController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim dataWrite() As Byte
Dim dataRead() As Byte
Dim serialPort As New SerialInterface
Public Function WriteData(ByVal address As Integer, ByRef data() As Byte, ByVal offset As Integer, ByVal length As Integer) As Boolean
    WriteData = WriteRead(address, data, offset, length, Nothing, 0, 0)
End Function

Public Function Read(ByVal address As Integer, ByRef data() As Byte, ByVal offset As Integer, ByVal length As Integer) As Boolean
    Read = WriteRead(address, Nothing, 0, 0, data, offset, length)
End Function

Public Function WriteRead(ByVal address As Integer, ByRef dataWrite As Byte, ByVal offsetWrite As Integer, ByVal countWrite As Integer, ByRef dataRead() As Byte, ByVal offsetRead As Integer, ByVal countRead As Integer) As Boolean
    If (dataWrite Is Nothing And dataRead Is Nothing) Or (countWrite = 0 And countRead = 0) Then
        Err.Raise "At least one of dataWrite or dataRead must be specified"
    End If

    If dataWrite Is Nothing And countWrite <> 0 Then
        Err.Raise "dataWrite null but countWrite not zero"
    End If

    If dataRead Is Nothing And countRead <> 0 Then
        Err.Raise "dataRead null but countRead not zero"
    End If

    If dataWrite <> Nothing And offsetWrite + countWrite > dataWrite.length Then
        Err.Raise "Invalid range for dataWrite"
    End If

    If dataRead <> Nothing And offsetRead + countRead > dataRead.length Then
        Err.Raise "Invalid range for dataRead"
    End If

    Dim cmd As String
    cmd = "i2cstream(" & address & "," & countWrite & "," & countRead & ")"
    Call serialPort.WriteCommand(cmd)

    If countWrite > 0 Then
        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        If Not res.success Then
            Err.Raise "I2c error:" & res.respone
        End If

        Call serialPort.WriteRawData(dataWrite, offsetWrite, countWrite)
    End If

    If countRead > 0 Then
        If serialPort.ReadRawData(dataRead, offsetRead, countRead) <> countRead Then
            Err.Raise "I2C read raw data error."
        End If
    End If

    Dim res As CmdRespone
    res = serialPort.ReadRespone()
    WriteRead = res.success
End Function
