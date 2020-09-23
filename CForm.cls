VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim LastForm As Integer
Private Type Data
    ControlType As String
    ControlName As String
    ControlEvent As String
    Index As Integer
End Type

Dim Controls(1000) As Data
Dim LastControl As Integer
Dim LastCommandButton As Integer
Dim LastTimer As Integer
Dim LastLabel As Integer
Dim LastTextBox As Integer
Dim LastPictureBox As Integer
Dim LastImage As Integer
Dim LastLine As Integer

Dim LastShape As Integer
Dim LastComboBox As Integer
Dim LastListBox As Integer
Dim LastCheckBox As Integer
Dim LastOptionButton As Integer

            
Private Sub Class_Initialize()
WriteLog "CForm::Initialize()"
LastCommandButton = 1
LastTimer = 1
LastLabel = 1
End Sub

Function CreateForm()
LastForm = LastForm + 1

Set Form(LastForm) = New frmTemplate
Form(LastForm).Tag = LastForm
On Error Resume Next
Dim x As Control
For Each x In Form(LastForm).Controls
    On Error Resume Next
    If x.Tag <> "" Then x.Visible = True
Next

CreateForm = LastForm


End Function

Function frmObj(id)
Set frmObj = Form(id)
End Function
Function Unload(id)
'Unload Form(id)
End Function

Function hwnd(id)
hwnd = Form(id).hwnd
End Function

Function Show(id)
Form(id).Show
End Function

Function Hide(id)
Form(id).Hide
End Function

Function Refresh(id)
Form(id).Refresh
End Function

Function GetForm(id)
Set GetForm = Form(id)
End Function
Function BindEvent(lProc, lParent, lName, lEvent)
Dim h As Integer
For h = 0 To LastControl
    If Controls(h).ControlName = lName Then
        Controls(h).ControlEvent = lEvent
        If Controls(h).ControlType = "CommandButton" Then
            Form(lParent).Command1(Controls(h).Index).Tag = Form(lParent).Command1(Controls(h).Index).Tag & lProc & "::" & lName & "::" & lEvent & "::" & lParent & vbCrLf
        End If
    End If
Next h
End Function

Function CreateObject(lName)
Select Case LCase(lName)
Case "screen"
    Set CreateObject = Screen
Case "stdfont"
    Set CreateObject = New StdFont
Case "stdpicture"
    Set CreateObject = New StdPicture
End Select
End Function

Function CreateControl(lParent, lName, lType, Optional lTop = 0, Optional lLeft = 0, Optional lWidth = 0, Optional lHeight = 0)
'MsgBox "Control " & Name & " could not be created."
Dim xType As String

If LCase(lType) = "commandbutton" Then
    If LastCommandButton <> 0 Then Load Form(lParent).Command1(LastCommandButton)
    xType = "CommandButton"
    Controls(LastControl).Index = LastCommandButton
    Form(lParent).Command1(Controls(LastControl).Index).Tag = lName
    LastCommandButton = LastCommandButton + 1
End If

If LCase(lType) = "timer" Then
    If LastTimer <> 0 Then Load Form(lParent).Timer1(LastTimer)
    xType = "Timer"
    Controls(LastControl).Index = LastTimer
    Form(lParent).Timer1(Controls(LastControl).Index).Tag = lName
    LastTimer = LastTimer + 1
End If

If LCase(lType) = "label" Then
    If LastLabel <> 0 Then Load Form(lParent).Label1(LastLabel)
    xType = "Label"
    Controls(LastControl).Index = LastLabel
    Form(lParent).Label1(Controls(LastControl).Index).Tag = lName
    LastLabel = LastLabel + 1
End If

'textbox
If LCase(lType) = "textbox" Then
    If LastTextBox <> 0 Then Load Form(lParent).Text1(LastTextBox)
    xType = "TextBox"
    Controls(LastControl).Index = LastTextBox
    Form(lParent).Text1(Controls(LastControl).Index).Tag = lName
    LastTextBox = LastTextBox + 1
End If

'combobox
If LCase(lType) = "combobox" Then
    If LastComboBox <> 0 Then Load Form(lParent).Combo1(LastComboBox)
    xType = "ComboBox"
    Controls(LastControl).Index = LastComboBox
    Form(lParent).Combo1(Controls(LastControl).Index).Tag = lName
    LastComboBox = LastComboBox + 1
End If

'listbox
If LCase(lType) = "listbox" Then
    If LastListBox <> 0 Then Load Form(lParent).List1(LastTimer)
    xType = "ListBox"
    Controls(LastControl).Index = LastListBox
    Form(lParent).List1(Controls(LastControl).Index).Tag = lName
    LastListBox = LastListBox + 1
End If

'picturebox
If LCase(lType) = "picturebox" Then
    If LastPictureBox <> 0 Then Load Form(lParent).Picture1(LastPictureBox)
    xType = "PictureBox"
    Controls(LastControl).Index = LastPictureBox
    Form(lParent).Picture1(Controls(LastControl).Index).Tag = lName
    LastPictureBox = LastPictureBox + 1
End If

'image
If LCase(lType) = "image" Then
    If LastImage <> 0 Then Load Form(lParent).Image1(LastImage)
    xType = "Image"
    Controls(LastControl).Index = LastImage
    Form(lParent).Image1(Controls(LastControl).Index).Tag = lName
    LastImage = LastImage + 1
End If

'shape
If LCase(lType) = "shape" Then
    If LastShape <> 0 Then Load Form(lParent).Shape1(LastShape)
    xType = "Shape"
    Controls(LastControl).Index = LastShape
    Form(lParent).Shape1(Controls(LastControl).Index).Tag = lName
    LastShape = LastShape + 1
End If

'optionbutton
If LCase(lType) = "optionbutton" Then
    If LastOptionButton <> 0 Then Load Form(lParent).Option1(LastOptionButton)
    xType = "OptionButton"
    Controls(LastControl).Index = LastOptionButton
    Form(lParent).Option1(Controls(LastControl).Index).Tag = lName
    LastOptionButton = LastOptionButton + 1
    
End If

If LCase(lType) = "optionbutton(g)" Then
    If LastOptionButton <> 0 Then Load Form(lParent).Option1_G(LastOptionButton)
    xType = "OptionButton(G)"
    Controls(LastControl).Index = LastOptionButton
    Form(lParent).Option1_G(Controls(LastControl).Index).Tag = lName
    LastOptionButton = LastOptionButton + 1
    
End If

'checkbox
If LCase(lType) = "checkbox" Then
    If LastCheckBox <> 0 Then Load Form(lParent).Check1(LastCheckBox)
    xType = "CheckBox"
    Controls(LastControl).Index = LastCheckBox
    Form(lParent).Check1(Controls(LastControl).Index).Tag = lName
    LastCheckBox = LastCheckBox + 1
End If

If LCase(lType) = "line" Then
    If LastLine <> 0 Then Load Form(lParent).Line1(LastLine)
    xType = "Line"
    Controls(LastControl).Index = LastLine
    Form(lParent).Line1(Controls(LastLine).Index).Tag = lName
    LastLine = LastLine + 1
End If

    Controls(LastControl).ControlName = lName
    Controls(LastControl).ControlType = xType
    LastControl = LastControl + 1

End Function


Function Size(id, height, width)
Form(id).height = height
Form(id).width = width
End Function

Function SetCaption(id, what)
Form(id).caption = what
End Function

Function SetScalemode(id, Mode)
If LCase(Mode) = "twip" Then Form(id).ScaleMode = 1 'twip
If LCase(Mode) = "pixel" Then Form(id).ScaleMode = 3 'pixle


End Function

Function SetPicture(id, path)
Form(id).Picture = LoadPicture(path)
End Function

Function Center(id)
Form(id).Left = (Screen.width / 2) - (Form(id).width / 2)
Form(id).Top = (Screen.height / 2) - (Form(id).height / 2)
End Function


Function SizeP(id, height, width)
Form(id).height = Form(id).ScaleY(height, 3, 1)
Form(id).width = Form(id).ScaleX(width, 3, 1)
End Function

Function SetIcon(id, path)
Form(id).Icon = LoadPicture(path)
End Function

Function WaitForInput()
Do While GExit = False
    DoEvents
Loop

End Function

Function Quit()
GExit = True
End
End Function

Function CreateDrawingObject(lParent)
    Set CreateDrawingObject = Form(lParent).ClientArea
End Function

Function obj(lParent, lName)
Dim h As Integer
For h = 0 To LastControl - 1
    If Controls(h).ControlName = lName Then
        
            Select Case Controls(h).ControlType
            Case "CommandButton": Set obj = Form(lParent).Command1(Controls(h).Index)
            Case "Label": Set obj = Form(lParent).Label1(Controls(h).Index)
            Case "Timer": Set obj = Form(lParent).Timer1(Controls(h).Index)
            
            Case "TextBox": Set obj = Form(lParent).Text1(Controls(h).Index)
            Case "PictureBox": Set obj = Form(lParent).Picture1(Controls(h).Index)
            Case "Image": Set obj = Form(lParent).Image1(Controls(h).Index)
            Case "Shape": Set obj = Form(lParent).Shape1(Controls(h).Index)
            
            Case "OptionButton": Set obj = Form(lParent).Option1(Controls(h).Index)
            Case "OptionButton(G)": Set obj = Form(lParent).Option1_G(Controls(h).Index)
            Case "CheckBox": Set obj = Form(lParent).Check1(Controls(h).Index)
            
            Case "ComboBox": Set obj = Form(lParent).Combo1(Controls(h).Index)
            Case "ListBox": Set obj = Form(lParent).List1(Controls(h).Index)
            
            Case "Line": Set obj = Form(lParent).Line1(Controls(h).Index)
            
            Case Else
            MsgBox Controls(h).ControlType & " not found!"
            End Select
    End If
    If Controls(h).ControlName = "" Then MsgBox "Hit: idx = " & h
Next h
'MsgBox "X"

End Function

Function LoadImage(file As String)
Set LoadImage = LoadPicture(file)
End Function


Function LoadFromFile(file As String)
Dim Line As String
Dim Data() As String
Dim CODE As String
Dim CODE_DEF As String
Open file For Input As #1
Do While Not EOF(1)
Line Input #1, Line
Data = Split(Line, " ")

If Data(0) = "FORM" Then
    CODE = CODE & Data(2) & " = CForm.CreateForm()" & vbCrLf
    CODE = CODE & "Set " & Data(1) & " = CForm.frmObj(" & Data(2) & ")" & vbCrLf
    CODE_DEF = "Dim " & Data(1) & ", " & Data(2)
       
End If

If Data(0) = "CONTROL" Then
    CODE = CODE & "CForm.CreateControl " & Data(1) & ", " & Chr(34) & Data(2) & Chr(34) & ", " & Chr(34) & Data(3) & Chr(34) & vbCrLf
    CODE = CODE & "Set " & Data(2) & " = CForm.obj(" & Data(1) & ", " & Chr(34) & Data(2) & Chr(34) & ")" & vbCrLf
    CODE_DEF = CODE_DEF & ", " & Data(2)
End If
Loop
Close #1

CODE = CODE_DEF & CODE
Script.ExecuteStatement CODE
'Script.AddCode CODE
End Function

Function Line(obj, x1, y1, x2, y2, Optional fore, Optional b)
obj.Line (x1, y1)-(x2, y2), fore, B
End Function


Function Line2(obj, x1, y1, Optional fore, Optional b)
obj.Line -(x1, y1), fore, B
End Function

Function SetFlag(lParent, lName, lValue)
Form(lParent).SetFlag CStr(lName), lValue
End Function
