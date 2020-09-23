Attribute VB_Name = "modMain"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim browse As CommonDialog

Sub Main()
Set browse = New CommonDialog


'If Command = "-install" Then RegisterType: End
'If Command <> "" Then Execute Command: End

Dim responce As VbMsgBoxResult
responce = MsgBox("Do you wany to:" & vbCrLf & vbCrLf & "Yes: Download from the internet" & vbCrLf & "No: use a local file", vbYesNo Or vbQuestion, "MultiPlex")

If responce = vbNo Then
    browse.Filter = "MuliPlex 2.x applications (*.APP)|*.app"
    browse.InitDir = App.path
    browse.DialogTitle = "Execute script"
    browse.ShowOpen
    DoEvents
    'If Command <> "" Then
    If browse.CancelError = False Then
        Execute browse.FileName 'Command 'App.Path & "\sample\Main sample.txt"
    End If
Else
    DownloadAPP InputBox("Web address (including http://)", "Address"), App.path & "\cache\temp.app"
    Execute App.path & "\cache\temp.app"
End If
    
    End

End Sub

