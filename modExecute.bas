Attribute VB_Name = "modExecute"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Global g_File As String

Option Explicit
Function Execute(file As String)
EVENT_READY = False
'Set Script = New ScriptControl
'On Error Resume Next
If FileLen(file) = 0 Then MsgBox Chr(34) & file & Chr(34) & " not found!": End
frmLoading.Show
frmLoading.Refresh

g_File = file

DecompressAPP

file = g_File
inf "Initialising DirectCall...", 0


UseLog True
SetScript

inf "Applying security policy...", 0

Script.AllowUI = False
Script.Timeout = NoTimeout
Script.UseSafeSubset = False

Dim tim As Long
Dim TotalTime As Long
OpenLog

Dim No As Boolean
Dim a As String
Dim b As String
Dim c As String
Dim D As String
TotalTime = GetTickCount

'On Error GoTo ErrS
inf "Caching code...", 0
Open file For Binary Access Read As #1
FData = String(LOF(1), "x")
Get #1, , FData
Close #1
WriteLog "Data buffer full"

Data_CPY = FData


FormatData

WriteLog "Data format complete."

xLine = Split(FData, vbCrLf)
Dim ScriptData As String
Dim cv As String
Dim Pos1 As Long
Dim Pos2 As Long
Dim LanguageX As String
Dim OLDPOS As Long
Dim tmpp() As String
Dim Entry As String

Pos2 = 1

WriteLog "Beginning preliminary check..."
inf "Beginning preliminary scan...", 0

'preliminary check
Do While UBound(xLine) <> curline
    
    inf "Scanning...", (curline / UBound(xLine)) * 100

    xLine(curline) = Trim(xLine(curline))
    If Left(Trim((xLine(curline))), Len("interface")) = "interface" Then
        Interface.Add Replace(Replace(Mid(xLine(curline), Len("interface") + 1, 100), ";", ""), "#sys.", "")
        WriteLog "Interface found: " & Trim(Replace(Replace(Mid(xLine(curline), Len("interface") + 1, 100), ";", ""), "#sys.", ""))
    End If
    
    If Left(Trim((xLine(curline))), Len("include")) = "include" Then
        tmpp() = Split(Trim(xLine(curline)), " ")
        If LCase(tmpp(1)) = "def" Then
            Dim bak As String
            bak = Replace(tmpp(2), Chr(34), "")
            bak = Replace(bak, ";", "")
            bak = PathGetPath(file) & "\" & bak
           ' MsgBox "here"
            LoadDefs bak
        End If
    End If
    
    If Left(Trim((xLine(curline))), Len("define")) = "define" Then
        tmpp() = Split(Trim(xLine(curline)), " ")
        If LCase(tmpp(1)) = "entry" Then
            Dim bak2 As String
            bak2 = Replace(tmpp(2), Chr(34), "")
            bak2 = Replace(bak2, ";", "")
            
            Entry = bak2
        End If
    End If
    If Left(Trim((xLine(curline))), Len("validate")) = "validate" Then
        No = True
        a = "C000"
        b = "MultiPlex scripting runtime"
        c = "Syntax error: 'validate' - parameter missing"
        D = "Line " & curline + 1
        
        xLine(curline) = Replace(xLine(curline), ";", "")
        xLine(curline) = Replace(xLine(curline), Chr(34), "")
        tmpp() = Split(Trim(xLine(curline)), " ")
        
      '  WriteLog "Validating " & tmpp(1)
        If LCase(tmpp(1)) = "platform" Then
            If LCase(tmpp(2)) = "win32" Then
                WriteLog "win32 validated"
            Else
                WriteLog tmpp(2) & " not validated!"
            End If
        End If
        
                c = "Unable to validate language: '" & tmpp(2) & "', this script is not present on this machine."
                
        If LCase(tmpp(1)) = "script" Then
            If LCase(tmpp(2)) = "javascript" Then
                Script.Language = "javascript"
                WriteLog tmpp(2) & " validated"
                
        
            ElseIf LCase(tmpp(2)) = "jscript" Then
                Script.Language = "javascript"
                WriteLog tmpp(2) & " validated"
                
            ElseIf LCase(tmpp(2)) = "vbscript" Then
                Script.Language = "vbscript"
                WriteLog tmpp(2) & " validated"
            
            ElseIf LCase(tmpp(2)) = "cscript" Then
                Script.Language = "javascript"
                WriteLog tmpp(2) & " validated"
            Else
                WriteLog tmpp(2) & " not validated!"
            End If
        End If
                
        If LCase(tmpp(1)) = "dll_import" Then
            If CheckCall(Replace(tmpp(2), Chr(34), ""), Replace(tmpp(3), Chr(34), "")) = False Then
                MsgBox "dll_import validation FAILED. Execution terminated!"
                End
            Else
                WriteLog "dll_import function '" & tmpp(3) & "' validated!"
            End If
            
        End If
                
                
        'If tmpp(1) <> "Win32" Then
        '    c = "Unable to validate language: '" & tmpp(1) & "', this script is not present on this machine."
        '    Script.Language = tmpp(1)
        'Else 'it is Win32
        '
        'End If
        
    End If
    
    If Left(Trim((xLine(curline))), Len("#section")) = "#section" Then
        cv = Replace(xLine(curline), "#section ", "")
        cv = Replace(cv, "{", "")
        cv = Trim(cv)
        WriteLog "Section found: " & cv
    End If
    
    curline = curline + 1
Loop

WriteLog "Beginning execution"


inf "Importing...", 0

curline = 0
Do While UBound(xLine) <> curline
    
    If Left(Trim((xLine(curline))), 6) = "import" Then
        
        
        OLDPOS = InStr(1, FData, "import " & Trim(LanguageX))
        LanguageX = Mid(Trim((xLine(curline))), 7, Len(Trim((xLine(curline)))) - 7)

        Pos1 = InStr(OLDPOS, FData, "import " & Trim(LanguageX) & "{") + Len("import ") + Len(LanguageX) + 1
        Pos2 = InStr(Pos1, FData, "}")
        
        LanguageX = Trim(LanguageX)
        WriteLog "Importing language " & LanguageX
        
        ScriptData = Mid(FData, Pos1 + 1, Pos2 - Pos1 - 2)
        'MsgBox "'" & LanguageX
        WriteLog "Verifying interface..."
        CheckInterface Data_CPY
        Script.Language = Replace(Replace(LanguageX, "#", ""), Chr(34), "")
        
        
        WriteLog "Creating interface"
        LoadClasses
        No = False
        'MsgBox ScriptData
        inf "Binding interface...", 0
       
        WriteLog "Executing..."
        
        inf "Executing...", 0
        
        tim = GetTickCount()
        FormatDef Script.Language, ScriptData
        
        If isGUI = False Then
                EVENT_READY = False
                frmLoading.Hide
        End If
        
      On Error GoTo ErrS
        If Entry = "" Then
            Script.AddCode ScriptData
        Else
            Script.AddCode Entry & vbCrLf & ScriptData
        End If
        
        WriteLog "Execution complete"
        WriteLog "Time: " & (GetTickCount - tim) / 1000 & " sec [" & GetTickCount - tim & " ms]"
        WriteLog "Import successful"

If isGUI = True Then
        EVENT_READY = True
        frmLoading.Hide

        Script.Run "Form_Load"
        Do
            DoEvents
        Loop
End If
    End If


    curline = curline + 1
Loop
WriteLog "Script execution complete."
WriteLog "Unloading classes..."
'Set c1 = Nothing
WriteLog "Unloading interfaces..."
WriteLog "Total time: " & (GetTickCount - TotalTime) / 1000 & " secs [" & GetTickCount - TotalTime & " ms]"
CloseLog

End
Exit Function

ErrS:

On Error Resume Next
If No = False Then ERRSource = Trim(xLine(Script.Error.line + curline))
If No = True Then XERROR a, b, c, D

WriteLog "Error hit:" & vbCrLf & vbCrLf & "***ERROR FOUND IN SCRIPT***" & vbCrLf & _
    "Number: " & Script.Error.number & vbCrLf & _
            "Translation layer: " & Replace(Script.Error.source, " error", "") & vbCrLf & _
            "Description: " & Script.Error.Description & vbCrLf & _
            "Line: " & curline + 1 + Script.Error.line & vbCrLf & _
            "Error source: " & ERRSource & vbCrLf & _
            "***EXECUTION ABORTED***" & vbCrLf
MsgBox "MultiPlex detected an error while executing the following file:" & vbCrLf & gFile & vbCrLf & vbCrLf & "Number: " & Script.Error.number & vbCrLf & "Translation layer: " & Replace(Script.Error.source, " error", "") & vbCrLf & "Description: " & Script.Error.Description & vbCrLf & "Line: " & curline + 1 + Script.Error.line & ", character 0" & vbCrLf & "Error source: " & Chr(34) & ERRSource & Chr(34) & vbCrLf & vbCrLf & vbCrLf & "Press OK to abort script execution.", vbExclamation Or vbMsgBoxHelpButton, "MultiPlex script error", Script.Error.HelpFile, Script.Error.HelpContext
CloseLog
End Function

Function Unload()
Set Script = Nothing
'Set lCConvert = Nothing
End
End Function
