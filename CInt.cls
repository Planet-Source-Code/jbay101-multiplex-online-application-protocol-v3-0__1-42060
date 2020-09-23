VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function toByte(a)
toByte = CByte(a)
End Function

Function toBool(a)
toBool = CBool(a)
End Function

Function aDec(a)
'aDec = CDec(a)
End Function

Function toInt(a)
toInt = CInt(a)
End Function

Function toLng(a)
toLng = CLng(a)
End Function

Function toSng(a)
toSng = CSng(a)
End Function

Function aHex(a)
aHex = Hex(a)
End Function

Function aOct(a)
aOct = Oct(a)
End Function

Function aStr(a)
aStr = str(a)
End Function

Private Sub Class_Initialize()
WriteLog "CConvert::Initialize()"
End Sub

