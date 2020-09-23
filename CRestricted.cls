VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSharedMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim SEPERATOR As String
Dim SharedVars(256) As String ' As New Collection
Dim LastAdded As Long

Function vSet(Name, Value)
WriteLog "CSharedMem::vSet()"

Dim y As String
Dim xx As Integer
Dim Found As Boolean
For xx = LBound(SharedVars) To UBound(SharedVars)
    y = SharedVars(xx)
    If Left(y, Len(Name)) = Name Then
    '    WriteLog "Variable '" & Name & "' updated"
        SharedVars(xx) = Name & SEPERATOR & Value
        Found = True
    End If
Next xx

If Found = False Then
    'WriteLog "Variable '" & Name & "' written to shared memory"
    SharedVars(LastAdded + 1) = Name & SEPERATOR & Value
    'MsgBox "last+1=" & SharedVars(LastAdded + 1)
    LastAdded = LastAdded + 1
End If

End Function

Function vGet(Name)
WriteLog "CSharedMem::vGet()"

Dim y
Dim xx As Integer
Dim tmp() As String

For xx = LBound(SharedVars) To UBound(SharedVars)
    y = SharedVars(xx)
    
    If Left(y, Len(Name)) = Name Then
        tmp = Split(y, SEPERATOR)
        vGet = tmp(1)
    End If
Next xx
End Function

Private Sub Class_Initialize()
SEPERATOR = "[" & Chr(0) & Chr(1) & Chr(1) & Chr(0) & "]"
WriteLog "CSharedMem::Initialize()"

End Sub
