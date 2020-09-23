VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function MessageBoxA Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Function xPrint(what)

End Function

Function xAssert(what)

End Function

Function xOut(what)
MessageBoxA 0, "CDebug: val=" & what, "MultiPlex Debug", vbInformation
End Function

Function xWriteLog(what)
WriteLog "SCRIPT: " & what
End Function
Private Sub Class_Initialize()
WriteLog "CDebug::Initialize()"
End Sub



