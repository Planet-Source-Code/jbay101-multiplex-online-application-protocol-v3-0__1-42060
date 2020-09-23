VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function WriteKey(Subkey, Entry, Value)
SetStringValue CStr(Subkey), CStr(Entry), CStr(Value)
End Function

Function ReadKey(Subkey, Entry)
ReadKey = GetStringValue(CStr(Subkey), CStr(Entry))
End Function

Private Sub Class_Initialize()
WriteLog "CRegistry::Initialize()"
End Sub

