Attribute VB_Name = "modDef"
Option Explicit
Dim ffind(0 To 512) As String
Dim rreplace(0 To 512) As String

Dim lentry As Integer

Function LoadDefs(file As String)
Dim curline As String
Dim t() As String

Open file For Input As #45
Do While Not EOF(45)
    Line Input #45, curline
    curline = Trim(curline)
    If (Left(curline, 2) <> ";;") And (Left(curline, 2) <> "//") And (Left(curline, 1) <> "'") Then
        curline = Replace(curline, " ", "{enter.0}")
        
        curline = Trim(curline)
        curline = Replace(curline, ";", "")
        If curline = "" Then GoTo nextline
        t = Split(curline, "{enter.0}", 3)
        
        ffind(lentry) = Replace(t(1), "{enter.0}", " ")
        rreplace(lentry) = Replace(t(2), "{enter.0}", " ")
        
                
        If rreplace(lentry) = "__PATH" Then rreplace(lentry) = PathGetPath(g_File)
        If rreplace(lentry) = "__TIME" Then rreplace(lentry) = Time
        If rreplace(lentry) = "__DATE" Then rreplace(lentry) = Date
        
        'MsgBox ffind(lentry) & " - " & rreplace(lentry)
        
        lentry = lentry + 1
    End If
nextline:

Loop

Close #45

End Function

Function FormatDef(Language As String, Data As String)
Dim i As Integer
Dim rep As String
For i = 0 To lentry
        rep = rreplace(i)
        If LCase(Language) = "vbscript" Then
            If InStr(1, rreplace(i), "0x") <> 0 Then
                rep = Trim(Replace(rreplace(i), "0x", "&H"))
                rep = rep & "&"
            End If
        End If

        Data = Replace(Data, ffind(i), rep)
Next i
End Function
