VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BitWiseCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function ShiftRight(ByVal lngNumber As Long, ByVal intNumBits As Integer) As Long
'--------------
'BIT SHIFT RIGHT
'--------------

ShiftRight = lngNumber \ 2 ^ intNumBits 'note the integer division op

End Function

Public Function ShiftLeft(ByVal lngNumber As Long, ByVal intNumBits As Integer) As Long
'--------------
'BIT SHIFT LEFT
'--------------

ShiftLeft = lngNumber * 2 ^ intNumBits

End Function
