Attribute VB_Name = "modFormat"
Function fTime() As String
Static xyz As String
fTime = TD(Hour(Now)) & ":" & TD(Minute(Now)) & ":" & TD(Second(Now))
End Function

Function fDate() As String
Static DateA As String
Static DateB As String
Static DateC As String

DateA = TD(Day(Now))
DateB = TD(Month(Now))
DateC = Right(Year(Now), 2)

fDate = DateA & "/" & DateB & "/" & DateC

End Function

Function TD(what As String) As String
If Len(what) = 1 Then
TD = "0" & what
Else
TD = what
End If
End Function



Function FormatData()
FData = Replace(FData, "CString.String", "CString.strString")
FData = Replace(FData, "CString.Trim", "CString.strTrim")
FData = Replace(FData, "CString.LTrim", "CString.strLTrim")
FData = Replace(FData, "CString.RTrim", "CString.strRTrim")
FData = Replace(FData, "CString.Replace", "CString.strReplace")
FData = Replace(FData, "CString.UCase", "CString.strUCase")
FData = Replace(FData, "CString.LCase", "CString.strLCase")
FData = Replace(FData, "CString.Join", "CString.strJoin")
FData = Replace(FData, "CString.GetCh", "CString.strGetCh")
FData = Replace(FData, "CString.PutCh", "CString.strPutCh")
FData = Replace(FData, "CString.Mid", "CString.strMid")
FData = Replace(FData, "CString.Left", "CString.strLeft")
FData = Replace(FData, "CString.Right", "CString.strRight")
FData = Replace(FData, "CString.Reverse", "CString.strReverseA")
FData = Replace(FData, "CString.Chr", "CString.strChr")
FData = Replace(FData, "CString.Asc", "CString.strAsc")

FData = Replace(FData, "CMath.Abs", "CMath.xAbs")
FData = Replace(FData, "CMath.Atn", "CMath.xAtn")
FData = Replace(FData, "CMath.Cos", "CMath.xCos")
FData = Replace(FData, "CMath.Exp", "CMath.xExp")
FData = Replace(FData, "CMath.Log", "CMath.xLog")
FData = Replace(FData, "CMath.Randomize", "CMath.xRandomize")
FData = Replace(FData, "CMath.Rnd", "CMath.xRnd")
FData = Replace(FData, "CMath.Round", "CMath.xRound")
FData = Replace(FData, "CMath.Sgn", "CMath.xSgn")
FData = Replace(FData, "CMath.Sin", "CMath.xSin")
FData = Replace(FData, "CMath.Sqr", "CMath.xSqr")
FData = Replace(FData, "CMath.Tan", "CMath.xTan")
FData = Replace(FData, "CMath.Timer", "CMath.xTimer")


FData = Replace(FData, "CConvert.Hex", "CConvert.aHex")
FData = Replace(FData, "CConvert.Oct", "CConvert.aOct")
FData = Replace(FData, "CConvert.Str", "CConvert.aStr")

FData = Replace(FData, "CDebug.Print", "CDebug.xPrint")
FData = Replace(FData, "CDebug.Assert", "CDebug.xAssert")
FData = Replace(FData, "CDebug.Out", "CDebug.xOut")
FData = Replace(FData, "CDebug.WriteLog", "CDebug.xWriteLog")


FData = Replace(FData, "CFileSystem.ChDir", "CFileSystem.xChDir")
FData = Replace(FData, "CFileSystem.ChDrive", "CFileSystem.xChDrive")
FData = Replace(FData, "CFileSystem.CurDir", "CFileSystem.xCurDir")
FData = Replace(FData, "CFileSystem.Dir", "CFileSystem.xDir")
FData = Replace(FData, "CFileSystem.EOF", "CFileSystem.xEOF")
FData = Replace(FData, "CFileSystem.FileAttr", "CFileSystem.xFileAttr")
FData = Replace(FData, "CFileSystem.FileCopy", "CFileSystem.xFileCopy")
FData = Replace(FData, "CFileSystem.FileDateTime", "CFileSystem.xFileDateTime")
FData = Replace(FData, "CFileSystem.FileLen", "CFileSystem.xFileLen")
FData = Replace(FData, "CFileSystem.FreeFile", "CFileSystem.xFreeFile")
FData = Replace(FData, "CFileSystem.GetAttr", "CFileSystem.xGetAttr")
FData = Replace(FData, "CFileSystem.Kill", "CFileSystem.xKill")
FData = Replace(FData, "CFileSystem.MkDir", "CFileSystem.xMkDir")
FData = Replace(FData, "CFileSystem.RmDir", "CFileSystem.xRmDir")
FData = Replace(FData, "CFileSystem.SetAttr", "CFileSystem.xSetAttr")

FData = Replace(FData, "CFileIO.Put", "CFileIO.xPut")
FData = Replace(FData, "CFileIO.Get", "CFileIO.xGett")
FData = Replace(FData, "CFileIO.Seek", "CFileIO.xSeek")
FData = Replace(FData, "CFileIO.Print", "CFileIO.xPrint")
FData = Replace(FData, "CFileIO.ReadLine", "CFileIO.xReadLine")
FData = Replace(FData, "CFileIO.xReadLineEx", "CFileIO.xReadLineEx")
FData = Replace(FData, "CFileIO.ReadLinesEx", "CFileIO.xReadLinesEx")
FData = Replace(FData, "CFileIO.ReadFile", "CFileIO.xReadFile")
FData = Replace(FData, "CFileIO.xReadFileA", "CFileIO.xReadFileA")
FData = Replace(FData, "CFileIO.Close", "CFileIO.xClose")

FData = Replace(FData, "CWin32.Call", "CWin32.xCall")
FData = Replace(FData, "CString.String", "CString.strString")
FData = Replace(FData, "CString.Trim", "CString.strTrim")
FData = Replace(FData, "CString.LTrim", "CString.strLTrim")
FData = Replace(FData, "CString.RTrim", "CString.strRTrim")
FData = Replace(FData, "CString.Replace", "CString.strReplace")
FData = Replace(FData, "CString.UCase", "CString.strUCase")
FData = Replace(FData, "CString.LCase", "CString.strLCase")
FData = Replace(FData, "CString.Join", "CString.strJoin")
FData = Replace(FData, "CString.GetCh", "CString.strGetCh")
FData = Replace(FData, "CString.PutCh", "CString.strPutCh")
FData = Replace(FData, "CString.Mid", "CString.strMid")
FData = Replace(FData, "CString.Left", "CString.strLeft")
FData = Replace(FData, "CString.Right", "CString.strRight")
FData = Replace(FData, "CString.Reverse", "CString.strReverseA")
FData = Replace(FData, "CString.Chr", "CString.strChr")
FData = Replace(FData, "CString.Asc", "CString.strAsc")

FData = Replace(FData, "CMath.Abs", "CMath.xAbs")
FData = Replace(FData, "CMath.Atn", "CMath.xAtn")
FData = Replace(FData, "CMath.Cos", "CMath.xCos")
FData = Replace(FData, "CMath.Exp", "CMath.xExp")
FData = Replace(FData, "CMath.Log", "CMath.xLog")
FData = Replace(FData, "CMath.Randomize", "CMath.xRandomize")
FData = Replace(FData, "CMath.Rnd", "CMath.xRnd")
FData = Replace(FData, "CMath.Round", "CMath.xRound")
FData = Replace(FData, "CMath.Sgn", "CMath.xSgn")
FData = Replace(FData, "CMath.Sin", "CMath.xSin")
FData = Replace(FData, "CMath.Sqr", "CMath.xSqr")
FData = Replace(FData, "CMath.Tan", "CMath.xTan")
FData = Replace(FData, "CMath.Timer", "CMath.xTimer")


FData = Replace(FData, "CConvert.Hex", "CConvert.aHex")
FData = Replace(FData, "CConvert.Oct", "CConvert.aOct")
FData = Replace(FData, "CConvert.Str", "CConvert.aStr")

FData = Replace(FData, "CDebug.Print", "CDebug.xPrint")
FData = Replace(FData, "CDebug.Assert", "CDebug.xAssert")
FData = Replace(FData, "CDebug.Out", "CDebug.xOut")
FData = Replace(FData, "CDebug.WriteLog", "CDebug.xWriteLog")


FData = Replace(FData, "CFileSystem.ChDir", "CFileSystem.xChDir")
FData = Replace(FData, "CFileSystem.ChDrive", "CFileSystem.xChDrive")
FData = Replace(FData, "CFileSystem.CurDir", "CFileSystem.xCurDir")
FData = Replace(FData, "CFileSystem.Dir", "CFileSystem.xDir")
FData = Replace(FData, "CFileSystem.EOF", "CFileSystem.xEOF")
FData = Replace(FData, "CFileSystem.FileAttr", "CFileSystem.xFileAttr")
FData = Replace(FData, "CFileSystem.FileCopy", "CFileSystem.xFileCopy")
FData = Replace(FData, "CFileSystem.FileDateTime", "CFileSystem.xFileDateTime")
FData = Replace(FData, "CFileSystem.FileLen", "CFileSystem.xFileLen")
FData = Replace(FData, "CFileSystem.FreeFile", "CFileSystem.xFreeFile")
FData = Replace(FData, "CFileSystem.GetAttr", "CFileSystem.xGetAttr")
FData = Replace(FData, "CFileSystem.Kill", "CFileSystem.xKill")
FData = Replace(FData, "CFileSystem.MkDir", "CFileSystem.xMkDir")
FData = Replace(FData, "CFileSystem.RmDir", "CFileSystem.xRmDir")
FData = Replace(FData, "CFileSystem.SetAttr", "CFileSystem.xSetAttr")

FData = Replace(FData, "CFileIO.Put", "CFileIO.xPut")
FData = Replace(FData, "CFileIO.Get", "CFileIO.xGett")
FData = Replace(FData, "CFileIO.Seek", "CFileIO.xSeek")
FData = Replace(FData, "CFileIO.Print", "CFileIO.xPrint")
FData = Replace(FData, "CFileIO.ReadLine", "CFileIO.xReadLine")
FData = Replace(FData, "CFileIO.ReadLineEx", "CFileIO.xReadLineEx")
FData = Replace(FData, "CFileIO.ReadLinesEx", "CFileIO.xReadLinesEx")
FData = Replace(FData, "CFileIO.ReadFile", "CFileIO.xReadFile")
FData = Replace(FData, "CFileIO.Close", "CFileIO.xClose")

FData = Replace(FData, "CWin32.Call", "CWin32.xCall")

FData = Replace(FData, "CCommon.Time", "CCommon.xTime")
FData = Replace(FData, "CCommon.Date", "CCommon.xDate")
FData = Replace(FData, "CCommon.Now", "CCommon.xNow")

FData = Replace(FData, " {", "{")
FData = Replace(FData, "  {", "{")
FData = Replace(FData, "   {", "{")
End Function

Function ReplaceConst()
If Win32Found = True Then


End If
End Function

Function PathGetPath(path As String) As String
Dim bak As String
Dim i As Integer

For i = Len(path) To 1 Step -1
If Mid(path, i, 1) = "\" Then
bak = Mid(path, 1, i - 1)
Exit For
End If
Next i

PathGetPath = bak
End Function
