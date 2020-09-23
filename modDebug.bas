Attribute VB_Name = "modDebug"
Dim Log2File As Boolean
Dim Log2Debug As Boolean
Function UseLog(UseIt As Boolean)
Log2File = UseIt
End Function

Function UseDebug(UseIt As Boolean)
Log2Debug = UseIt
End Function

Function OpenLog()
If Log2File = True Then
    gFile = UCase(gFile)
    Open App.Path & "\script.log" For Output As #33
    Print #33, "[MultiPlex script log @ " & fDate & " " & fTime & "]"
    Print #33, "**Start log for script: " & gFile & "**"
End If
End Function

Function WriteLog(what As String)

'"01/05/02 19:08:09:"

If Log2File = True Then Print #33, what ' fDate & " " & fTime & ": " & what
If Log2File = False Then
    If Log2Debug = True Then Debug.Print what ' fDate & " " & fTime & ": " & what
End If
End Function

Function CloseLog()
If Log2File = True Then
    Print #33, "**End log for script: " & gFile & "**"
    Close #33
End If
UnLoad
End Function

