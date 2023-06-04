VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmdRespone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

    Public p_success As Boolean
    Public p_respone As String


Public Property Get success() As Boolean
    success = p_success
End Property

Public Property Let success(ByVal val1 As Boolean)
    p_success = val1
End Property


Public Property Get respone() As String
    respone = p_respone
End Property

Public Property Let respone(ByVal val1 As String)
    p_respone = val1
End Property

