VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function IsString(a)
If IsNumeric(a) = False Then
    IsString = True
Else
    IsString = False
End If
End Function

Function IsNumber(a)
    IsNumber = IsNumeric(a)
End Function

Private Sub Class_Initialize()
WriteLog "CType::Initialize()"
End Sub


