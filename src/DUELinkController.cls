VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DUELinkController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

    Private serialPort As New SerialInterface
    Public Analog As AnalogController
    Public Digital As DigitalController
    Public I2c As I2cController
    Public Servo As ServoController
    Public Frequency As FrequencyController
    Public Spi As SpiController
    Public Infrared As InfraredController
    Public Neo As NeoController
    Public System As SystemController
    Public Uart As UartController
    Public Button As ButtonController
    Public Distance As DistanceSensorController
    Public Display As DisplayController
    Public Touch As TouchController
    Public Led As LedController
    Public Script As ScriptController
    Public Pin As PinController
    Public Temperature As TemperatureController
    Public Humidity As HudimityController
    Public Version As String
    Public DeviceConfig As DeviceConfiguration

    Private Sub Class_Initialize()
        
        If serialPort Is Nothing Then
            Err.Raise "serialPort is null"
        End If

        Set Analog = New AnalogController
        Set Digital = New DigitalController
        Set I2c = New I2cController
        Set Servo = New ServoController
        Set Frequency = New FrequencyController
        Set Spi = New SpiController
        Set Infrared = New InfraredController
        Set Neo = New NeoController
        Set System = New SystemController
        Set Uart = New UartController
        Set Button = New ButtonController
        Set Distance = New DistanceSensorController
        Set Display = New DisplayController
        Set Touch = New TouchController
        Set Led = New LedController
        Set Script = New ScriptController
        Set Pin = New PinController
        Set Temperature = New TemperatureController
        Set Humidity = New HumidityController
    End Sub

    Private Sub Connect()
        Version = serialPort.GetVersion().Split("\n")(0)

        If Version = "" Or Version.length <> 7 Then
            Err.Raise "The device is not supported."
        End If

        Set DeviceConfig = New DeviceConfiguration

        If Version(Version.length - 1) = "P" Then
            DeviceConfig.IsPulse = True
            DeviceConfig.MaxPinIO = 23
            DeviceConfig.MaxPinAnalog = 29
        ElseIf Version(Version.length - 1) = "I" Then
            DeviceConfig.IsPico = True
            DeviceConfig.MaxPinIO = 29
            DeviceConfig.MaxPinAnalog = 29
        ElseIf Version(Version.length - 1) = "F" Then
            DeviceConfig.IsFlea = True
            DeviceConfig.MaxPinIO = 11
            DeviceConfig.MaxPinAnalog = 29
        ElseIf Version(Version.length - 1) = "E" Then
            DeviceConfig.IsFlea = True
            DeviceConfig.MaxPinIO = 22
            DeviceConfig.MaxPinAnalog = 11
        End If

        serialPort.DeviceConfig = DeviceConfig
    End Sub

    Public Sub Disconnect()
        serialPort.Disconnect
    End Sub

    Public Function GetConnectionPort() As String
        GetConnectionPort = ""
    End Function


