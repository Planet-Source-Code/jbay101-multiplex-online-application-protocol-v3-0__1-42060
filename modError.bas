Attribute VB_Name = "modError"
Function XERROR(number, TLayer, Description, source)
WriteLog "Error hit" & vbCrLf & vbCrLf & "***ERROR FOUND IN SCRIPT***" & vbCrLf & _
    "Number: " & number & vbCrLf & _
            "Translation layer: " & TLayer & vbCrLf & _
            "Description: " & Description & vbCrLf & _
            "Error source: " & source & vbCrLf & _
            "***EXECUTION ABORTED***" & vbCrLf
Exit Function
MsgBox "MultiPlex detected an error while executing the following file:" & vbCrLf & gFile & vbCrLf & vbCrLf & _
            "Number: " & number & vbCrLf & _
            "Translation layer: " & TLayer & vbCrLf & _
            "Description: " & Description & vbCrLf & _
            "Error source: " & source & _
            vbCrLf & vbCrLf & vbCrLf & "Press OK to abort script execution.", vbExclamation Or vbMsgBoxHelpButton, "MultiPlex script error", Script.Error.HelpFile, Script.Error.HelpContext
CloseLog
End Function


