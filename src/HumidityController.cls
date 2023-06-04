VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HumidityController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private serialPort As New SerialInterface
    Public Function Read(pin As Integer, sensortype As Integer) As Single
        Dim cmd As String
        cmd = "print(humidity(" + pin + "," + sensortype + "))"
        Call serialPort.WriteCommand(cmd)

        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        Read = res.success
    End Function

    Private Function get_dht11() As Integer
        get_dht11 = 11
    End Function

    Private Function get_dht12() As Integer
        get_dht12 = 12
    End Function

    Private Function get_dht21() As Integer
        get_dht21 21
    End Function

    Private Function get_dht22() As Integer
        get_dht22 = 22
    End Function



Public Property Get Dht11() As Boolean
    Dht11 = get_dht11()
End Property


Public Property Get Dht12() As Boolean
    Dht12 = get_dht12()
End Property


Public Property Get Dht21() As Boolean
    Dht21 = get_dht21()
End Property


Public Property Get Dht22() As Boolean
    Dht22 = get_dht22()
End Property

