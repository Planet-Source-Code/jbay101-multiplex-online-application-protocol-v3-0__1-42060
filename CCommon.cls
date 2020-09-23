VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function xDate()
xDate = Date
End Function

Function xTime()
xTime = Time
End Function

Function xNow()
xNow = Now
End Function
Private Sub Class_Initialize()
WriteLog "CCommon::Initialize()"
End Sub

Function vbIIf(a, b, c)
vbIIf = IIf(a, b, c)
End Function
