VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub apiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Dim FF As Integer
Function OpenFile(FN, Optional xType)
FF = FreeFile
If xType = "" Then xType = "output"

Select Case LCase(xType)
Case "output"
    Open FN For Output As FF
Case "input"
    Open FN For Input As FF

Case "binary"
    Open FN For Binary As FF

Case "binary/read"
    Open FN For Binary Access Read As FF

Case "binary/write"
    Open FN For Binary Access Write As FF

Case "random"
    Open FN For Random As FF

Case "random/read"
    Open FN For Random Access Read As FF
    
Case "random/write"
    Open FN For Random Access Write As FF
    
Case "append"
    Open FN For Append As FF
    
End Select
apiSleep 50
OpenFile = FF
End Function

Function xPut(number, offset, data)
Put #number, offset, data
End Function

Function xGet(number, offset, length)
Dim y As String
y = String(length, "x")

Get #number, offset, y
xGet = y
End Function

Function xSeek(number)
xSeek = Seek(number)
End Function

Function xPrint(number, what)
Print #number, what
End Function

Function xReadLine(number)
Dim T As String
Line Input #number, T
xReadLine = T
End Function

Function xReadLineEx(fnumber, NumberOfLines)
Dim y As String
Dim a As String
Dim i As Integer
For i = 1 To NumberOfLines
    If EOF(fnumber) Then Exit For
    Line Input #number, a
    y = y & vbCrLf & a
Next i

xReadLineEx = y
End Function

Function xReadLinesEx(fnumber, startline, endline)
Dim y As String
Dim a As String
Dim i As Integer
For i = 1 To endline
    
    If EOF(fnumber) Then Exit For
    Line Input #number, a
    If i >= startline Then y = y & vbCrLf & a
Next i
xReadLinesEx = y
End Function
Function xReadFile(FileName)
Dim S As String
Dim xx As Integer
xx = FreeFile
Open FileName For Binary Access Read As #xx
S = String(LOF(xx), "x")
Get #xx, , S
Close #xx

xReadFile = S
End Function

Function xReadFileA(FN)
Dim S As String
Dim T As String
Dim xx As Integer
Do While Not EOF(FN)
    Line Input #FN, T
    'MsgBox T
    S = S & vbCrLf & T
Loop
xReadFileA = S
End Function

Function xClose(number)
Close #number
End Function


Private Sub Class_Initialize()
WriteLog "CFileIO::Initialize()"
End Sub



