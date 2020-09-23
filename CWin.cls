VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim p1a_VALUE As Variant
Dim p2a_VALUE As Variant
Dim p3a_VALUE As Variant
Dim p4a_VALUE As Variant
Dim p5a_VALUE As Variant
Dim p6a_VALUE As Variant
Dim p7a_VALUE As Variant
Dim p8a_VALUE As Variant
Dim p9a_VALUE As Variant
Dim p10a_VALUE As Variant
Dim p11a_VALUE As Variant
Dim p12a_VALUE As Variant

Dim p1a_TYPE As String
Dim p2a_TYPE As String
Dim p3a_TYPE As String
Dim p4a_TYPE As String
Dim p5a_TYPE As String
Dim p6a_TYPE As String
Dim p7a_TYPE As String
Dim p8a_TYPE As String
Dim p9a_TYPE As String
Dim p10a_TYPE As String
Dim p11a_TYPE As String
Dim p12a_TYPE As String

Dim p1a() As Byte
Dim p2a() As Byte
Dim p3a() As Byte
Dim p4a() As Byte
Dim p5a() As Byte
Dim p6a() As Byte
Dim p7a() As Byte
Dim p8a() As Byte
Dim p9a() As Byte
Dim p10a() As Byte
Dim p11a() As Byte
Dim p12a() As Byte


Private Declare Function MessageBoxA Lib "user32" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub apiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Function MessageBox(text, caption, Optional style)
If IsMissing(style) = False Then
    MessageBoxA 0, text, caption, style
Else
    MessageBoxA 0, text, caption, vbInformation
End If
End Function

Function Sleep(ms)
apiSleep ms
End Function

Function xCall(what, Optional p1 As Variant, Optional p2 As Variant, Optional p3 As Variant, Optional p4 As Variant, Optional p5 As Variant, Optional p6 As Variant, Optional p7 As Variant, Optional p8 As Variant, Optional p9 As Variant, Optional p10 As Variant, Optional p11 As Variant, Optional p12 As Variant)
'MsgBox what
WriteLog "CWin32::Call()"

Dim Buf() As String
Dim Lib As String
Dim FFunction As String
Dim NumberOfArgs As Integer
Dim Structure As String
Dim ret
Dim s_type() As String

If InStr(1, what, "!") = 0 Then
    ret = "Not valid function!"
    xCall = ret
    Exit Function
Else
    Buf = Split(what, "!")
    Lib = Buf(0)
    FFunction = Buf(1)
    Structure = Buf(2)
    s_type = Split(Structure, ",")
    Dim i As Integer
    For i = 0 To UBound(s_type)
        s_type(i) = Trim(LCase(s_type(i)))
    Next i
    
End If

'Exit Function
If IsMissing(p12) = True Then NumberOfArgs = 11 ': p10 = 0
If IsMissing(p11) = True Then NumberOfArgs = 10 ': p10 = 0
If IsMissing(p10) = True Then NumberOfArgs = 9 ': p10 = 0
If IsMissing(p9) = True Then NumberOfArgs = 8 ': p9 = 0
If IsMissing(p8) = True Then NumberOfArgs = 7 ': p8 = 0
If IsMissing(p7) = True Then NumberOfArgs = 6 ': p7 = 0
If IsMissing(p6) = True Then NumberOfArgs = 5 ': p6 = 0
If IsMissing(p5) = True Then NumberOfArgs = 4 ': p5 = 0
If IsMissing(p4) = True Then NumberOfArgs = 3 ': p4 = 0
If IsMissing(p3) = True Then NumberOfArgs = 2 ': p3 = 0
If IsMissing(p2) = True Then NumberOfArgs = 1 ': p2 = 0
If IsMissing(p1) = True Then NumberOfArgs = 0 ': p1 = 0

If IsMissing(p12) = False Then NumberOfArgs = 12


'MsgBox "Number of args: " & NumberOfArgs



'If IsNumeric(p1) = True Then p1 = CLng(p1): p1a_TYPE = "number"
'If IsNumeric(p2) = True Then p2 = CLng(p2): p2a_TYPE = "number"
'If IsNumeric(p3) = True Then p3 = CLng(p3): p3a_TYPE = "number"
'If IsNumeric(p4) = True Then p4 = CLng(p4): p4a_TYPE = "number"
'If IsNumeric(p5) = True Then p5 = CLng(p5): p5a_TYPE = "number"
'If IsNumeric(p6) = True Then p6 = CLng(p6): p6a_TYPE = "number"
'If IsNumeric(p7) = True Then p7 = CLng(p7): p7a_TYPE = "number"
'If IsNumeric(p8) = True Then p8 = CLng(p8): p8a_TYPE = "number"
'If IsNumeric(p9) = True Then p9 = CLng(p9): p9a_TYPE = "number"
'If IsNumeric(p10) = True Then p10 = CLng(p10): p10a_TYPE = "number"
'If IsNumeric(p11) = True Then p11 = CLng(p11): p11a_TYPE = "number"
'If IsNumeric(p12) = True Then p12 = CLng(p12): p12a_TYPE = "number"

'MsgBox "P1 is " & p1a_TYPE
'MsgBox "P2 is " & p2a_TYPE
'MsgBox "P3 is " & p3a_TYPE
'MsgBox "P4 is " & p4a_TYPE


p1a_VALUE = p1
p2a_VALUE = p2
p3a_VALUE = p3
p4a_VALUE = p4
p5a_VALUE = p5
p6a_VALUE = p6
p7a_VALUE = p7
p8a_VALUE = p8
p9a_VALUE = p9
p10a_VALUE = p10
p11a_VALUE = p11
p12a_VALUE = p12

If IsMissing(p1a_VALUE) = False Then

If isnum(s_type, 0) Then
        p1a_VALUE = p1
Else
        p1a = StrConv(p1a_VALUE & Chr(0), vbFromUnicode)
        p1a_VALUE = VarPtr(p1a(0))
End If
End If

If IsMissing(p2a_VALUE) = False Then
If isnum(s_type, 1) Then
    p2a_VALUE = p2
Else
        p2a = StrConv(p2a_VALUE & Chr(0), vbFromUnicode)
        p2a_VALUE = VarPtr(p2a(0))
End If
End If

If IsMissing(p3a_VALUE) = False Then
If isnum(s_type, 2) Then
        p3a_VALUE = p3
Else
        p3a = StrConv(p3a_VALUE & Chr(0), vbFromUnicode)
        p3a_VALUE = VarPtr(p3a(0))
End If
End If

If IsMissing(p4a_VALUE) = False Then
If isnum(s_type, 3) Then
        p4a_VALUE = p4
Else
        p4a = StrConv(p4a_VALUE & Chr(0), vbFromUnicode)
        p4a_VALUE = VarPtr(p4a(0))
End If
End If



If IsMissing(p5a_VALUE) = False Then
If isnum(s_type, 4) Then
        p5a_VALUE = p5
Else
        p5a = StrConv(p5a_VALUE & Chr(0), vbFromUnicode)
        p5a_VALUE = VarPtr(p5a(0))
End If
End If

If IsMissing(p6a_VALUE) = False Then
If isnum(s_type, 5) Then
        p6a_VALUE = p6
Else
        p6a = StrConv(p6a_VALUE & Chr(0), vbFromUnicode)
        p6a_VALUE = VarPtr(p6a(0))
End If
End If

If IsMissing(p7a_VALUE) = False Then
If isnum(s_type, 6) Then
        p7a_VALUE = p7
Else
        p7a = StrConv(p7a_VALUE & Chr(0), vbFromUnicode)
        p7a_VALUE = VarPtr(p7a(0))
End If
End If


If IsMissing(p8a_VALUE) = False Then
If isnum(s_type, 7) Then
        p8a_VALUE = p8
Else
        p8a = StrConv(p8a_VALUE & Chr(0), vbFromUnicode)
        p8a_VALUE = VarPtr(p8a(0))
End If
End If


If IsMissing(p9a_VALUE) = False Then
If isnum(s_type, 8) Then
        p9a_VALUE = p9
Else
        p9a = StrConv(p9a_VALUE & Chr(0), vbFromUnicode)
        p9a_VALUE = VarPtr(p9a(0))
End If
End If


If IsMissing(p10a_VALUE) = False Then
If isnum(s_type, 9) Then
        p10a_VALUE = p10
Else
        p10a = StrConv(p10a_VALUE & Chr(0), vbFromUnicode)
        p10a_VALUE = VarPtr(p10a(0))
End If
End If

If IsMissing(p11a_VALUE) = False Then
If isnum(s_type, 10) Then
        p11a_VALUE = p11
Else
        p11a = StrConv(p11a_VALUE & Chr(0), vbFromUnicode)
        p11a_VALUE = VarPtr(p11a(0))
End If
End If

If IsMissing(p12a_VALUE) = False Then
If isnum(s_type, 11) Then
        p12a_VALUE = p12
Else
        p12a = StrConv(p12a_VALUE & Chr(0), vbFromUnicode)
        p12a_VALUE = VarPtr(p12a(0))
End If
End If




Select Case NumberOfArgs
Case 0
    ret = CallApiByName(Lib, FFunction)
Case 1
    ret = CallApiByName(Lib, FFunction, p1a_VALUE)
Case 2
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE)
Case 3
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE, p3a_VALUE)
Case 4
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE, p3a_VALUE, p4a_VALUE)
Case 5
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE, p3a_VALUE, p4a_VALUE, p5a_VALUE)
Case 6
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE, p3a_VALUE, p4a_VALUE, p5a_VALUE, p6a_VALUE)
Case 7
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE, p3a_VALUE, p4a_VALUE, p5a_VALUE, p6a_VALUE, p7a_VALUE)
Case 8
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE, p3a_VALUE, p4a_VALUE, p5a_VALUE, p6a_VALUE, p7a_VALUE, p8a_VALUE)
Case 9
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE, p3a_VALUE, p4a_VALUE, p5a_VALUE, p6a_VALUE, p7a_VALUE, p8a_VALUE, p9a_VALUE)
Case 10
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE, p3a_VALUE, p4a_VALUE, p5a_VALUE, p6a_VALUE, p7a_VALUE, p8a_VALUE, p9a_VALUE, p10a_VALUE)
Case 11
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE, p3a_VALUE, p4a_VALUE, p5a_VALUE, p6a_VALUE, p7a_VALUE, p8a_VALUE, p9a_VALUE, p10a_VALUE, p11a_VALUE)
Case 12
    ret = CallApiByName(Lib, FFunction, p1a_VALUE, p2a_VALUE, p3a_VALUE, p4a_VALUE, p5a_VALUE, p6a_VALUE, p7a_VALUE, p8a_VALUE, p9a_VALUE, p10a_VALUE, p11a_VALUE, , p12a_VALUE)
End Select
EraseVars
xCall = ret
End Function

Function CreateCall(DLL, FFunction, Structure)
WriteLog "CWin32::CreateCall()"
modApi.CheckCall CStr(DLL), CStr(FFunction)
CreateCall = DLL & "!" & FFunction & "!" & Structure
End Function

Function CheckCall(DLL, FFunction)
WriteLog "CWin32::CheckCall()"

modApi.CheckCall CStr(DLL), CStr(FFunction)
End Function

Function EraseVars()
 p1a_VALUE = 0
 p2a_VALUE = 0
 p3a_VALUE = 0
 p4a_VALUE = 0
 p5a_VALUE = 0
 p6a_VALUE = 0
 p7a_VALUE = 0
 p8a_VALUE = 0
 p9a_VALUE = 0
 p10a_VALUE = 0

 p1a_TYPE = 0
 p2a_TYPE = 0
 p3a_TYPE = 0
 p4a_TYPE = 0
 p5a_TYPE = 0
 p6a_TYPE = 0
 p7a_TYPE = 0
 p8a_TYPE = 0
 p9a_TYPE = 0
 p10a_TYPE = 0

Erase p1a()
Erase p2a()
Erase p3a()
Erase p4a()
Erase p5a()
Erase p6a()
Erase p7a()
Erase p8a()
Erase p9a()
Erase p10a()


End Function

Private Sub Class_Initialize()
WriteLog "CWin32::Initialize()"
End Sub

Private Sub yTest(s_type() As String, i As Integer, p As Variant)
Dim bak As String
bak = LCase(s_type(i))
s_type(i) = LCase(s_type(i))

If (InStr(1, bak, "as")) <> 0 Then
    s_type(i) = Trim(Mid(bak, InStr(1, bak, "as") + 2, 100))
    MsgBox s_type(i)

End If

If s_type(i) = "long" Then p = CLng(p)
If s_type(i) = "integer" Then p = CInt(p)
If s_type(i) = "byte" Then p = CByte(p)
If s_type(i) = "string" Then p = CStr(p)

If s_type(i) = "long" Then p = CLng(p)
If s_type(i) = "int" Then p = CInt(p)
If s_type(i) = "byte" Then p = CByte(p)
If s_type(i) = "str" Then p = CStr(p)

If s_type(i) = "l" Then p = CLng(p)
If s_type(i) = "i" Then p = CInt(p)
If s_type(i) = "b" Then p = CByte(p)
If s_type(i) = "s" Then p = CStr(p)
End Sub

Private Function isnum(s_type() As String, i As Integer) As Boolean
isnum = True

If s_type(i) = "string" Then isnum = False
If s_type(i) = "s" Then isnum = False
If s_type(i) = "str" Then isnum = False
End Function
