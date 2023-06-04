VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub WasteTime(Finish As Long)
 
    Dim NowTick As Long
    Dim EndTick As Long
 
    EndTick = GetTickCount + (Finish * 1000)
     
    Do
 
        NowTick = GetTickCount
        DoEvents
 
    Loop Until NowTick >= EndTick
 
End Sub
