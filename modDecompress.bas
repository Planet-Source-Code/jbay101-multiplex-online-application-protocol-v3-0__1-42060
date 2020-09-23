Attribute VB_Name = "modDecompress"
Option Explicit

Dim zip As New CGUnzipFiles
Function cache()
'On Error Resume Next
Dim x As String
    x = "x"
    Do While x <> ""
        x = Dir(App.path & "\cache\")
        'MsgBox x
        
        If x <> "" Then
                Kill App.path & "\cache\" & x
        End If
    Loop
RmDir App.path & "\cache"
MkDir App.path & "\cache"

End Function
Function DecompressAPP() As String
inf "Initiating cache...", 0


zip.HonorDirectories = False
zip.Unzip g_File, App.path & "\cache\"
 
'On Error Resume Next
Dim lines As String
Dim dat() As String

Open App.path & "\cache\index" For Input As #5
 
    Do While Not EOF(5)
        Line Input #5, lines
        dat = Split(lines, " ", 2)
        If UBound(dat) > 0 Then
        
        Select Case LCase(dat(0))
            
            Case "app_name"
                
                g_File = App.path & "\cache\" & Replace(Replace(dat(1), Chr(34), ""), ";", "")
            
            Case "app_skin"
            
                GSkin = Replace(Replace(dat(1), Chr(34), ""), ";", "")

            Case "app_type"
                Select Case UCase(Replace(dat(1), ";", ""))
                Case "WIN32_GUI"
                    EVENT_READY = False
                    isGUI = True
                Case "WIN32_CON"
                    EVENT_READY = True
                End Select
        End Select
        End If
    Loop
Close #5

End Function
