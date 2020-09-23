VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'CLib

'SECTION: <STRING.H>

Function sprintf(test, Optional x1, Optional x2, Optional x3, Optional x4, Optional x5, Optional x6, Optional x7, Optional x8, Optional x9, Optional x10)
Dim pos As Integer
Dim Output As String
Dim CurX As Integer

For pos = 1 To Len(test)
If Mid(test, pos, 2) = "%s" Then
    CurX = CurX + 1
    
    If CurX = 1 Then Output = Output & x1
    If CurX = 2 Then Output = Output & x2
    If CurX = 3 Then Output = Output & x3
    If CurX = 4 Then Output = Output & x4
    If CurX = 5 Then Output = Output & x5
    If CurX = 6 Then Output = Output & x6
    If CurX = 7 Then Output = Output & x7
    If CurX = 8 Then Output = Output & x8
    If CurX = 9 Then Output = Output & x9
    If CurX = 10 Then Output = Output & x10
    
    
    pos = pos + 1
Else
    Output = Output & Mid(test, pos, 1)
End If


Next pos
Output = Replace(Output, "/n", vbCrLf)
sprintf = Output
End Function

Function strlen(what)
strlen = Len(what)
End Function

Function memcpy(from, too, size_t)
    memcpy = Left(from, size_t)
End Function

Function memcmp(a, b, size_t)
If Left(a, size_t) = (Left(b, size_t)) Then
    memcmp = True
Else
   memcmp = False
End If
End Function

Function strcpy(str, Optional char)
strcpy = str
End Function

Function strcat(a, b)
strcat = a & b
End Function
        
Function strcmp(a, b)
If a = b Then strcmp = True
If a <> b Then strcmp = False
End Function
        
Private Sub Class_Initialize()
WriteLog "CLib::Initialize()"
End Sub

