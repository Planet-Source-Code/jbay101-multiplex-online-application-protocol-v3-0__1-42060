VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function FileName()
FileName = gFile
End Function

Function Size()
Size = FileLen(gFile)
End Function
Private Sub Class_Initialize()
WriteLog "CScript::Initialize()"
End Sub

Function Import(file As String)
On Error GoTo errx:
Open file For Binary Access Read As #3
Dim v As String
v = String(LOF(3), "x")
Get #3, , v
Call FormatDef(Script.Language, v)
Script.AddCode v
Close #3

Exit Function
errx:
MsgBox "error - line: " & Script.Error.line & ", desc:" & Script.Error.Description
End Function
