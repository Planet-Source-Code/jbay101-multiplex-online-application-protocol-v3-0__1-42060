Attribute VB_Name = "modGlobals"
Global EVENT_READY As Boolean
Global FData As String
Global Data_CPY As String
Global Form(100) As New frmTemplate
Global TRG As Object
Global IDX As Long
Global GSkin As String
Global isGUI As Boolean

Global Interface As New Collection
Global Script As ScriptControl

Global GExit As Boolean

Global curline As Integer
Global gFile As String
Global ERRSource As String
Global xLine() As String


Function SetScript()
Set Script = New ScriptControl
End Function

Function inf(text As String, value As Integer)
frmLoading.ProgressBar1.Position = value
If Len(text) = 0 Then Exit Function
If frmLoading.Label1.caption <> text Then frmLoading.Label1.caption = text: frmLoading.Label1.Refresh
End Function

