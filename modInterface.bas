Attribute VB_Name = "modInterface"
Function CheckInterface(data As String)
Dim It

For Each It In Interface
    It = Trim(It)
    Select Case It
        Case "CString": CStringFound = True
        Case "CInt": CIntFound = True
        Case "Win32": Win32Found = True
        Case "SharedMem": SharedMemFound = True
        Case "Common": CommonFound = True
        Case "CLib": CLibFound = True
        Case "CMath": CMathFound = True
        Case "CConvert": CConvertFound = True
        Case "CFileSystem": CFileSystemFound = True
        Case "CFileIO": CFileIOFound = True
        Case "CDebug": CDebugFound = True
        Case "CScript": CScriptFound = True
        Case "CScriptEngine": CScriptEngineFound = True
        Case "CPlatform": CPlatformFound = True
        Case "CRegistry": CRegistryFound = True
        Case "CType": CTypeFound = True
        Case "CForm": CFormFound = True
        Case "CConsole": CConsoleFound = True
    End Select
    
Next

If CStringFound = False Then
    If InStr(1, data, "CString.") <> 0 Then XERROR "C001.a", "Mutiplex scripting runtime", "Interface definition missing", "[missing interface with sys.CString]"
End If

If CIntFound = False Then
    If InStr(1, data, "CInt.") <> 0 Then XERROR "C001.b", "Mutiplex scripting runtime", "Interface definition missing", "[missing interface with sys.CInt]"
End If

    
If Win32Found = False Then
    If InStr(1, data, "Win32.") <> 0 Then XERROR "C001.c", "Mutiplex scripting runtime", "Interface definition missing", "[missing interface with sys.Win32]"
End If

If SharedMemFound = False Then
    If InStr(1, data, "_vSet") <> 0 Then XERROR "C001.d", "Mutiplex scripting runtime", "Interface definition missing for _vSet", "[missing interface with sys.SharedMem]"
    If InStr(1, data, "_vGet") <> 0 Then XERROR "C001.d", "Mutiplex scripting runtime", "Interface definition missing for _vGet", "[missing interface with sys.SharedMem]"
End If

If CommonFound = False Then
    If InStr(1, data, "import") <> 0 Then XERROR "C001.e", "Mutiplex scripting runtime", "Interface definition missing", "[missing interface with sys.Common]"
End If

WriteLog "Interfaces verifyed"
End Function

