VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function strString(length, Character)
strString = String(length, Character)
End Function

Function strTrim(what)
strTrim = Trim(what)
End Function

Function strLTrim(what)
strLTrim = LTrim(what)
End Function

Function strRTrim(what)
strRTrim = RTrim(what)
End Function

Function strReplace(source, finds, replaces)
strReplace = Replace(source, finds, replaces)
End Function

Function strUCase(text)
strUCase = UCase(text)
End Function

Function strLCase(text)
strLCase = LCase(text)
End Function

Function strJoin(p1, p2)
strJoin = p1 & p2
End Function

Function strPutCh(text, pos, what)
Dim tmp As String
tmp = Mid(text, 1, pos) & what & Mid(text, pos + 1, Len(text))
strPutCh = tmp
End Function

Function strGetCh(text, pos)
Dim tmp As String
tmp = Mid(text, pos, 1)
strGetCh = tmp
End Function

Function strMid(stringg, start, length)
strMid = Mid(stringg, start, length)
End Function

Function strRight(stringg, length)
strRight = Right(stringg, length)
End Function

Function strLeft(stringg, length)
strLeft = Left(stringg, length)
End Function

Function strReverseA(stringg)
strReverseA = StrReverse(stringg)
End Function

Function strChr(CODE)
strChr = Chr(CODE)
End Function

Function strAsc(CODE)
strAsc = Asc(CODE)
End Function

Function str_sprintf(test, Optional x1, Optional x2, Optional x3, Optional x4, Optional x5, Optional x6, Optional x7, Optional x8, Optional x9, Optional x10)
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

str_sprintf = Output
End Function

Private Sub Class_Initialize()
WriteLog "CString::Initialize()"
End Sub


Function strFormat(a, b)
strFormat = Format(a, b)
End Function

