VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DisplayController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private serialPort As New SerialInterface
    

    Public Function Show() As Boolean
        Dim cmd As String
        cmd = "lcdshow()"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        Show = res.success
    End Function

    Public Function Clear(ByVal color As Integer) As Boolean
        Dim cmd As String
        cmd = "lcdclear(" + color + ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        Clear = res.success
    End Function

    Public Function SetPixel(ByVal color As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
        Dim cmd As String
        cmd = "lcdpixel(" + color + "," + x + "," + y + ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        SetPixel = res.success
    End Function

    Public Function DrawCircle(ByVal color As Integer, ByVal x As Integer, ByVal y As Integer, ByVal radius As Integer) As Boolean
        Dim cmd As String
        cmd = "lcdcircle(" + color + "," + x + "," + y + "," + radius + ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        DrawCircle = res.success
    End Function

    Public Function DrawRectangle(ByVal color As Integer, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer) As Boolean
        Dim cmd As String
        cmd = "lcdrect(" + color + "," + x + "," + y + "," + width + "," + height + ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        DrawRectangle = res.success
    End Function

    Public Function DrawFillRect(ByVal color As Integer, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer) As Boolean
        Dim cmd As String
        cmd = "lcdfill(" + color + "," + x + "," + y + "," + width + "," + height + ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        DrawFillRect = res.success
    End Function

    Public Function DrawLine(ByVal color As Integer, ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Boolean
        Dim cmd As String
        cmd = "lcdline(" + color + "," + x1 + "," + y1 + "," + x2 + "," + y2 + ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        DrawLine = res.success
    End Function

    Public Function DrawText(ByVal text As String, ByVal color As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
        Dim cmd As String
        cmd = "lcdtext(""" + text + """," + color + "," + x + "," + y + ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        DrawText = res.success
    End Function

    Public Function DrawTextScale(ByVal text As String, ByVal color As Integer, ByVal x As Integer, ByVal y As Integer, ByVal scaleWidth As Integer, ByVal scaleHeight As Integer) As Boolean
        Dim cmd As String
        cmd = "lcdtexts(""" + text + """," + color + "," + x + "," + y + "," + scaleWidth + "," + scaleHeight + ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        DrawTextScale = res.success
    End Function

    Private Function Stream(ByVal data As Byte) As Boolean
        Dim cmd As String
        cmd = "lcdstream()"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        If res.success Then
            Call serialPort.WriteRawData(data, 0, data.length)
            ' Thread.Sleep(10)
            res = serialPort.ReadRespone()
        End If

        Stream = res.success
    End Function
Function LeftBitWise(ByVal SrcValue As Variant, ByVal Times As Byte) As Variant
Dim SrcStr As String
Dim myValue As Long
If Not IsNumeric(SrcValue) Then
Exit Function
End If

myValue = CLng("&H" & Right("00000000" & Hex(SrcValue), 8))
For a = 1 To Times
myValue = myValue And &H7FFFFFFF
If (&H40000000 And myValue) = &H40000000 Then
myValue = (&H3FFFFFFF And myValue) * 2 Or &H80000000
Else
myValue = myValue * 2
End If
Next a

LeftBitWise = Right("00000000" & Hex(myValue), 8)

If VarType(SrcValue) = vbByte Then
LeftBitWise = CByte("&H" & Right(LeftBitWise, 2))
End If
If VarType(SrcValue) = vbInteger Then
LeftBitWise = CInt("&H" & Right(LeftBitWise, 4))
End If
If VarType(SrcValue) = vbLong Then
LeftBitWise = CLng("&H" & Right(LeftBitWise, 8))
End If
End Function



    Public Function DrawBuffer(ByVal color As Integer) As Boolean
        Const width As Integer = 128
        Const height As Integer = 64

        Dim offset As Integer
        offset = 0
        Dim length As Integer
        length = color.length

        Dim data(1) As Byte
        ReDim Preserve data(width * height / 8 - 1)
        Dim i As Integer
        i = offset

        Dim y As Integer
        Dim x As Integer
        Dim index As Integer
        For y = 0 To height - 1
            For x = 0 To width - 1
                index =(y >> 3) * WIDTH + x

                If i < offset + length Then
                    If (color(i) And &HFFFFFF) <> 0 Then
                        data(index) = data(index) Or (1 << (y And 7)) And &HFF
                    Else
                        data(index) = data(index) And Not (1 << (y And 7)) And &HFF
                    End If

                    i += 1
                End If
            Next
        Next

        Return Me.__Stream(data)
    End Function

    Public Function DrawBufferBytes(ByRef color() As Byte) As Boolean
        Dim offset As Integer
        offset = 0
        Dim length As Integer
        length = color.length

        If length Mod 4 <> 0 Then
            Err.Raise "length must be multiple of 4"
        End If

        Dim data32(length / 4 - 1) As Integer
        Dim i As Integer
        
        For i = 0 To UBound(data32.length) Step 4
            data32(i) = (color((i + offset) * 4 + 0) << 0) Or (color((i + offset) * 4 + 1) << 8) Or (color((i + offset) * 4 + 2) << 16) Or (color((i + offset) * 4 + 3) << 24)
        Next

        DrawBufferBytes = DrawBuffer(data32)
    End Function

    Public Function Configuration(ByVal target As Integer, ByVal slaveAddress As Integer) As Boolean
        Dim cmd As String
        cmd = "lcdconfig(" + target + "," + slaveAddress + ")"
        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()
        Configuration = res.success
    End Function

    Public Function DrawImageScale(ByRef img() As Integer, ByVal x As Integer, ByVal y As Integer, ByVal scaleWidth As Integer, ByVal scaleHeight As Integer, ByVal transform As Integer) As Boolean
        Dim width As Integer
        width = img(0)
        Dim height As Integer
        height = img(1)

        If width <= 0 Or height <= 0 Or img.length < width * height Then
            Err.Raise "Invalid arguments"
        End If

        Dim cmd As String
        cmd = "dim a[" + UBound(img) + "]"

        Call serialPort.WriteCommand(cmd)
        Dim res As CmdRespone
        res = serialPort.ReadRespone()

        Dim i As Integer
        For i = 0 To img.length - 1
            cmd = "a[" + i + "] = " + img(i)
            Call serialPort.WriteCommand(cmd)
            res = serialPort.ReadRespone()

            If Not res.success Then
                Exit For
            End If
        Next

        If res.success Then
            cmd = "lcdimgs(a, " + x + ", " + y + ", " + scaleWidth + ", " + scaleHeight + ", " + transform + ")"

            Call serialPort.WriteCommand(cmd)
            res = serialPort.ReadRespone()
        End If

        cmd = "dim a[0]"

        Call serialPort.WriteCommand(cmd)
        res = serialPort.ReadRespone()

        DrawImageScale = res.success
    End Function

    Public Function DrawImage(ByRef img() As Integer, ByVal x As Integer, ByVal y As Integer, ByVal transform As Integer) As Boolean
        Call DrawImage(DrawImageScale(img, x, y, 1, 1, transform))
    End Function

    Private Function get_transform_none() As Integer
        get_transform_none = 0
    End Function

    Private Function get_transform_fliphorizontal() As Integer
        get_transform_fliphorizontal = 1
    End Function

    Private Function get_transform_flipvertical() As Integer
        get_transform_flipvertical = 2
    End Function

    Private Function get_transform_rotate90() As Integer
        get_transform_rotate90 = 3
    End Function

    Private Function get_transform_rotate180() As Integer
        get_transform_rotate180 = 4
    End Function

    Private Function get_transform_rotate270() As Integer
        get_transform_rotate270 = 5
    End Function

    Private Sub set_transform(ByVal value As Integer)
    End Sub
    
Public Property Get TransformNone() As Boolean
    TransformNone = get_transform_none()
End Property


Public Property Get TransformFlipHorizontal() As Boolean
    TransformFlipHorizontal = get_transform_fliphorizontal()
End Property


Public Property Get TransformFlipVertical() As Boolean
    TransformFlipVertical = get_transform_flipvertical()
End Property



Public Property Get TransformRotate90() As Boolean
    TransformRotate90 = get_transform_rotate90()
End Property



Public Property Get TransformRotate180() As Boolean
    TransformRotate180 = get_transform_rotate180()
End Property



Public Property Get TransformRotate270() As Boolean
    TransformRotate270 = get_transform_rotate270()
End Property
