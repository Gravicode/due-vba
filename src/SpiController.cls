VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpiController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


    Private serialPort As New SerialInterface


    Public Function WriteData(ByRef dataWrite() As Byte, ByVal offset As Integer, ByVal length As Integer, Optional ByVal chipselect As Integer = -1) As Boolean
        WriteData = WriteRead(dataWrite, offset, length, Nothing, 0, 0, chipselect)
    End Function

    Public Function ReadData(ByRef dataRead() As Byte, ByVal offset As Integer, ByVal length As Integer, Optional ByVal chipselect As Integer = -1) As Boolean
        ReadData = WriteRead(Nothing, 0, 0, dataRead, offset, length, chipselect)
    End Function

    Public Function WriteRead(ByRef dataWrite() As Byte, ByVal offsetWrite As Integer, ByVal countWrite As Integer, ByRef dataRead() As Byte, ByVal offsetRead As Integer, ByVal countRead As Integer, Optional ByVal chipselect As Integer = -1) As Boolean
        If chipselect >= serialPort.DeviceConfig.MaxPinIO Then
            Err.Raise "InvalidPin"
        End If

        If (dataWrite = Nothing And dataRead = Nothing) Or (countWrite = 0 And countRead = 0) Then
            Err.Raise "Invalid arguments"
        End If

        If dataWrite <> Nothing And offsetWrite + countWrite > UBound(dataWrite) Then
            Err.Raise "Invalid arguments"
        End If

        If dataRead <> Nothing And offsetRead + countRead > dataRead.length Then
            Err.Raise "Invalid arguments"
        End If

        If chipselect < 0 Then
            chipselect = 255
        End If

        Dim cmd As String
        cmd = "spistream(" + countWrite + "," + countRead + "," + chipselect + ")"
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        If Not res.success Then
            WriteRead = False
        End If

        Do While countWrite > 0 Or countRead > 0
            Dim num As Integer
            num = countRead

            If countWrite < countRead Then
                num = countWrite
            End If

            If countWrite = 0 Then
                num = countRead
            End If

            If countRead = 0 Then
                num = countWrite
            End If

            If num > serialPort.TransferBlockSizeMax Then
                num = serialPort.TransferBlockSizeMax
            End If

            If countWrite > 0 Then
                Call serialPort.WriteRawData(dataWrite, offsetWrite, num)
                offsetWrite = offsetWrite + num
                countWrite = countWrite - num
            End If

            If countRead > 0 Then
                Call serialPort.ReadRawData(dataRead, offsetRead, num)
                offsetRead = offsetRead + num
                countRead = countRead - num
            End If
        Loop

        res = serialPort.ReadRespone()
        WriteRead = res.success
    End Function

    Public Function Write4bpp(ByRef dataWrite() As Byte, ByVal offsetWrite As Integer, ByVal countWrite As Integer, Optional ByVal chipselect As Integer = -1) As Boolean
        If chipselect >= serialPort.DeviceConfig.MaxPinIO Then
            Err.Raise "InvalidPin"
        End If

        If (dataWrite Is Nothing) Or (countWrite = 0) Then
            Err.Raise "Invalid arguments"
        End If

        Dim cmd As String
        cmd = "spi4bpp(" + countWrite + "," + chipselect + ")"
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        If Not res.success Then
             Write4bpp = False
        End If

        Call serialPort.WriteRawData(dataWrite, offsetWrite, countWrite)

        res = serialPort.ReadRespone()
        Write4bpp = res.success
    End Function

    Public Function Pallete(ByVal id As Integer, ByVal color As Integer) As Boolean
        If id > 16 Then
            Err.Raise "Pallete supports 16 color index only."
        End If

        Dim cmd As String
        cmd = "palette(" + id + "," + color + ")"

        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        Pallete = res.success
    End Function

    Public Function Configuration(ByVal mode As Integer, ByVal frequencyKHz As Integer) As Boolean
        If mode > 3 Or mode < 0 Then
            Err.Raise "Mode must be in range 0...3."
        End If

        If frequencyKHz < 200 Or frequencyKHz > 20000 Then
            Err.Raise "FrequencyKHz must be in range 200KHz to 20MHz."
        End If

        Dim cmd As String
        cmd = "palette(" + mode + "," + frequencyKHz + ")"

        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        Configuration = res.success
    End Function

