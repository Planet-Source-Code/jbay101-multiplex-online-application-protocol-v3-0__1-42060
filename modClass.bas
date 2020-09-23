Attribute VB_Name = "modClass"
Dim sInit As Boolean
Global lCConsole As CConsole
Global lCType As CType
Global lCScript As CScript
Global lCScriptEngine As CScriptEngine
Global lCSharedMem As CSharedMem
Global lCString As CString
Global lCWin As CWin
Global lCLib As CLib
Global lCMath As CMath
Global lCCommon As CCommon
Global lCConvert As CConvert
Global lCFileSystem As CFileSystem
Global lCFileIO As CFileIO
Global lCDebug As CDebug
Global lCPlatform As CPlatform
Global lCRegistry As CRegistry
Global lCForm As CForm

Global CConsoleFound As Boolean
Global CFormFound As Boolean
Global CTypeFound As Boolean
Global CFileSystemFound As Boolean
Global CDebugFound As Boolean
Global CFileIOFound As Boolean
Global CMathFound As Boolean
Global CConvertFound As Boolean
Global CStringFound As Boolean
Global CIntFound As Boolean
Global Win32Found As Boolean
Global CLibFound As Boolean
Global SharedMemFound As Boolean
Global CommonFound As Boolean
Global CScriptFound As Boolean
Global CScriptEngineFound As Boolean
Global CPlatformFound As Boolean
Global CRegistryFound As Boolean

Function LoadClasses()
If CStringFound = True Then
    Set lCString = New CString
    Script.AddObject "CString", lCString, True
End If

If CommonFound = True Then
    Set lCCommon = New CCommon
    Script.AddObject "CCommon", lCCommon, True
End If

If CLibFound = True Then
    Set lCLib = New CLib
    Script.AddObject "CLib", lCLib, True
End If

If Win32Found = True Then
    Set lCWin = New CWin
    Script.AddObject "CWin32", lCWin, True
End If

If SharedMemFound = True Then
    If sInit = False Then
        Set lCSharedMem = New CSharedMem
        sInit = True
    End If
        Script.AddObject "CSharedMem", lCSharedMem, True
End If

If CMathFound = True Then
    Set lCMath = New CMath
    Script.AddObject "CMath", lCMath, True
End If

If CConvertFound = True Then
    Set lCConvert = New CConvert
    Script.AddObject "CConvert", lCConvert, True
End If

If CFileSystemFound = True Then
    Set lCFileSystem = New CFileSystem
    Script.AddObject "CFileSystem", lCFileSystem, True
End If

If CFileIOFound = True Then
    Set lCFileIO = New CFileIO
    Script.AddObject "CFileIO", lCFileIO, True
End If

If CDebugFound = True Then
    Set lCDebug = New CDebug
    Script.AddObject "CDebug", lCDebug, True
End If

If CScriptFound = True Then
    Set lCScript = New CScript
    Script.AddObject "CScript", lCScript, True
End If

If CScriptEngineFound = True Then
    Set lCScriptEngine = New CScriptEngine
    Script.AddObject "CScriptEngine", lCScriptEngine, True
End If

If CPlatformFound = True Then
    Set lCPlatform = New CPlatform
    Script.AddObject "CPlatform", lCPlatform, True
End If

If CRegistryFound = True Then
    Set lCRegistry = New CRegistry
    Script.AddObject "CRegistry", lCRegistry, True
End If

If CTypeFound = True Then
    Set lCType = New CType
    Script.AddObject "CType", lCType, True
End If

If CFormFound = True Then
    Set lCForm = New CForm
    Script.AddObject "CForm", lCForm, True
End If

If CConsoleFound = True Then
    Set lCConsole = New CConsole
    Script.AddObject "CConsole", lCConsole, True
End If

End Function

