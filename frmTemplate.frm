VERSION 5.00
Begin VB.Form frmTemplate 
   BorderStyle     =   0  'None
   Caption         =   "MultiPlex - Form1"
   ClientHeight    =   2130
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6015
   Icon            =   "frmTemplate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ClientArea 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   3
      Top             =   360
      Width           =   5775
      Begin VB.OptionButton Option1_G 
         Caption         =   "Option1_G"
         Height          =   300
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   675
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   255
         Index           =   0
         Left            =   3090
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   0
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   315
         Index           =   0
         Left            =   2805
         TabIndex        =   8
         Top             =   135
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   270
         Index           =   0
         Left            =   930
         TabIndex        =   7
         Top             =   285
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.PictureBox Picture1 
         Height          =   360
         Index           =   0
         Left            =   3360
         ScaleHeight     =   300
         ScaleWidth      =   1785
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Index           =   0
         Left            =   3345
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   705
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Index           =   0
         Left            =   2715
         Top             =   630
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   300
         Index           =   0
         Left            =   930
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   1500
      End
   End
   Begin MultiPlex.SkinnedForm SkinnedForm1 
      Height          =   2130
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      _extentx        =   10610
      _extenty        =   3757
   End
   Begin VB.Line Line1 
      Index           =   0
      Visible         =   0   'False
      X1              =   825
      X2              =   3030
      Y1              =   1995
      Y2              =   2010
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   5265
      Top             =   1110
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Shape Shape1 
      Height          =   420
      Index           =   0
      Left            =   5025
      Top             =   540
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblWarning 
      Caption         =   "Warning: Applet window"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   1725
   End
End
Attribute VB_Name = "frmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flags(100) As String
Dim Values(100) As Variant
Dim LastValue As Integer

Function SetFlag(xname As String, xval As Variant)
flags(LastValue) = xname
Values(LastValue) = xval
LastValue = LastValue + 1
End Function

Function fcheck()
Dim i As Integer
Dim value As Variant
For i = 0 To LastValue - 1
    value = Values(i)
    Select Case LCase(flags(i))
        Case "minbutton"
            If value = True Then SkinnedForm1.Min True
            If value = False Then SkinnedForm1.Min False
            
        Case "maxbutton"
            If value = True Then SkinnedForm1.Max True
            If value = False Then SkinnedForm1.Max False
    
        Case "closebutton"
            If value = True Then SkinnedForm1.CloseB True
            If value = False Then SkinnedForm1.CloseB False

    End Select
    
Next i

End Function
Private Sub Check1_Click(Index As Integer)
RunEvent Check1(Index), "Click"
End Sub

Private Sub Check1_GotFocus(Index As Integer)
RunEvent Check1(Index), "GotFocus"
End Sub

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Check1(Index), "KeyDown", KeyCode, Shift
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
RunEvent Check1(Index), "KeyPress", KeyAscii
End Sub

Private Sub Check1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Check1(Index), "KeyUp", KeyCode, Shift
End Sub

Private Sub Check1_LostFocus(Index As Integer)
RunEvent Check1(Index), "LostFocus"
End Sub

Private Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Check1(Index), "MouseDown", Button, Shift, x, Y
End Sub

Private Sub Check1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Check1(Index), "MouseMove", Button, Shift, x, Y
End Sub

Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Check1(Index), "MouseUp", Button, Shift, x, Y
End Sub

Private Sub Check1_Validate(Index As Integer, Cancel As Boolean)
RunEvent Check1(Index), "Validate", Cancel
End Sub

Private Sub Combo1_Change(Index As Integer)
RunEvent Combo1(Index), "Change"
End Sub

Private Sub Combo1_Click(Index As Integer)
RunEvent Combo1(Index), "Click"
End Sub

Private Sub Combo1_DblClick(Index As Integer)
RunEvent Combo1(Index), "DblChange"
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
RunEvent Combo1(Index), "GotFocus"
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Combo1(Index), "KeyDown", KeyCode, Shift

End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
RunEvent Combo1(Index), "KeyPress", KeyAscii
End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Combo1(Index), "KeyUp", KeyCode, Shift
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
RunEvent Combo1(Index), "LostFocus"
End Sub

Private Sub Combo1_Scroll(Index As Integer)
RunEvent Combo1(Index), "Scroll"

End Sub

Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
RunEvent Combo1(Index), "Validate", Cancel
End Sub

Private Sub Command1_Click(Index As Integer)
RunEvent Command1(Index), "Click"
End Sub

Private Sub Command1_GotFocus(Index As Integer)
RunEvent Command1(Index), "GotFocus"

End Sub

Private Sub Command1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Command1(Index), "KeyDown", KeyCode, Shift

End Sub

Private Sub Command1_KeyPress(Index As Integer, KeyAscii As Integer)
RunEvent Command1(Index), "KeyPress", KeyAscii

End Sub

Private Sub Command1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Command1(Index), "KeyUp", KeyCode, Shift

End Sub

Private Sub Command1_LostFocus(Index As Integer)
RunEvent Command1(Index), "LostFocus"

End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Command1(Index), "MouseDown", Button, Shift, x, Y

End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Command1(Index), "MouseMove", Button, Shift, x, Y
End Sub

Private Sub Command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Command1(Index), "MouseUp", Button, Shift, x, Y
End Sub

Private Sub Form_Activate()
RunEventF "Form", "Activate"

On Error Resume Next
'For Each x In Me.Controls
'    If TypeOf X Is PictureBox Then
    'x.Visible = False
    'End If
    
    'MsgBox x.Tag
'    If X.Tag <> "" Then X.Visible = True
    'MsgBox x.Name
'Next

End Sub

Private Sub Form_Click()
Dim x As Control

RunEventF "Form", "Click"
End Sub

Private Sub Form_DblClick()
RunEventF "Form", "DblClick"
End Sub

Private Sub Form_Deactivate()
RunEventF "Form", "Deactivate"
End Sub

Private Sub Form_GotFocus()
'frmSplash.Hide
IDX = Me.Tag
RunEventF "Form", "GotFocus"
End Sub

Private Sub Form_Initialize()
If GSkin = "" Then GSkin = "default"

SkinnedForm1.LoadSkin App.path & "\skins\" & GSkin & "\" & "default" & ".skin", App.path & "\skins\" & GSkin
'SkinnedForm1.LoadSkin App.path & "\skins\purple\purple.skin", App.path & "\skins\purple"

RunEventF "Form", "Initialize"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
RunEventF "Form", "KeyDown", KeyCode, Shift
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
RunEventF "Form", "KeyPress", KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
RunEventF "Form", "KeyUp", KeyCode, Shift
End Sub

Private Sub Form_Load()

'ClientArea1.SetSM Me.ScaleMode
update
RunEventF "Form", "Load"
End Sub

Private Sub Form_LostFocus()
RunEventF "Form", "LostFocus"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEventF "Form", "MouseDown", Button, Shift, x, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEventF "Form", "MouseMove", Button, Shift, x, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEventF "Form", "MouseUp", Button, Shift, x, Y
End Sub

Private Sub Form_Paint()
RunEventF "Form", "Paint"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
RunEventF "Form", "QueryUnload", Cancel, UnloadMode
End Sub

Private Sub Form_Resize()
update
'Dim x As Integer
'x = Me.ScaleMode
RunEventF "Form", "Resize"
End Sub

Private Sub Form_Terminate()
RunEventF "Form", "Terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
RunEventF "Form", "Unload"
GExit = True
cache
End
End Sub

Private Sub Image1_Click(Index As Integer)
RunEvent Image1(Index), "Click"
End Sub

Private Sub Image1_DblClick(Index As Integer)
RunEvent Image1(Index), "DblClick"
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Image1(Index), "MouseDown", Button, Shift, x, Y
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Image1(Index), "MouseMove", Button, Shift, x, Y
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Image1(Index), "MouseUp", Button, Shift, x, Y
End Sub

Private Sub Label1_Click(Index As Integer)
RunEvent Label1(Index), "Click"
End Sub

Private Sub Label1_DblClick(Index As Integer)
RunEvent Label1(Index), "DblClick"
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Label1(Index), "MouseDown", Button, Shift, x, Y
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Label1(Index), "MouseMove", Button, Shift, x, Y
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Label1(Index), "MouseUp", Button, Shift, x, Y
End Sub

Private Sub List1_Click(Index As Integer)
RunEvent List1(Index), "Click"
End Sub

Private Sub List1_DblClick(Index As Integer)
RunEvent List1(Index), "DblClick"
End Sub

Private Sub List1_GotFocus(Index As Integer)
RunEvent List1(Index), "GotFocus"
End Sub

Private Sub List1_ItemCheck(Index As Integer, Item As Integer)
RunEvent List1(Index), "ItemCheck", Item
End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent List1(Index), "KeyDown", KeyCode, Shift
End Sub

Private Sub List1_KeyPress(Index As Integer, KeyAscii As Integer)
RunEvent List1(Index), "KeyPress", KeyAscii
End Sub

Private Sub List1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent List1(Index), "KeyUp", KeyCode, Shift
End Sub

Private Sub List1_LostFocus(Index As Integer)
RunEvent List1(Index), "LostFocus"
End Sub

Private Sub List1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent List1(Index), "MouseDown", Button, Shift, x, Y
End Sub

Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent List1(Index), "MouseMove", Button, Shift, x, Y
End Sub

Private Sub List1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent List1(Index), "MouseUp", Button, Shift, x, Y
End Sub

Private Sub List1_Scroll(Index As Integer)
RunEvent List1(Index), "Scroll"
End Sub

Private Sub List1_Validate(Index As Integer, Cancel As Boolean)
RunEvent List1(Index), "Validate", Cancel

End Sub

Private Sub Option1_Click(Index As Integer)
RunEvent Option1(Index), "Click"
End Sub

Private Sub Option1_DblClick(Index As Integer)
RunEvent Option1(Index), "DblClick"
End Sub

Private Sub Option1_GotFocus(Index As Integer)
RunEvent Option1(Index), "GotFocus"
End Sub

Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Option1(Index), "KeyDown", KeyCode, Shift
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
RunEvent Option1(Index), "KeyPress", KeyAscii
End Sub

Private Sub Option1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Option1(Index), "KeyUp", KeyCode, Shift
End Sub

Private Sub Option1_LostFocus(Index As Integer)
RunEvent Option1(Index), "LostFocus"
End Sub

Private Sub Option1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Option1(Index), "MouseDown", Button, Shift, x, Y
End Sub

Private Sub Option1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Option1(Index), "MouseMove", Button, Shift, x, Y
End Sub

Private Sub Option1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Option1(Index), "MouseUp", Button, Shift, x, Y
End Sub

Private Sub Option1_Validate(Index As Integer, Cancel As Boolean)
RunEvent Option1(Index), "Validate", Cancel

End Sub

Private Sub Picture1_Click(Index As Integer)
RunEvent Picture1(Index), "Click"
End Sub

Private Sub Picture1_DblClick(Index As Integer)
RunEvent Picture1(Index), "DblClick"
End Sub

Private Sub Picture1_GotFocus(Index As Integer)
RunEvent Picture1(Index), "GotFocus"
End Sub

Private Sub Picture1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Picture1(Index), "KeyDown", KeyCode, Shift
End Sub

Private Sub Picture1_KeyPress(Index As Integer, KeyAscii As Integer)
RunEvent Picture1(Index), "KeyPress", KeyAscii
End Sub

Private Sub Picture1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Picture1(Index), "KeyUp", KeyCode, Shift
End Sub

Private Sub Picture1_LostFocus(Index As Integer)
RunEvent Picture1(Index), "LostFocus"
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Picture1(Index), "MouseDown", Button, Shift, x, Y
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Picture1(Index), "MouseMove", Button, Shift, x, Y
End Sub

Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Picture1(Index), "MouseUp", Button, Shift, x, Y
End Sub

Private Sub Picture1_Paint(Index As Integer)
RunEvent Picture1(Index), "Paint"
End Sub

Private Sub Picture1_Resize(Index As Integer)
RunEvent Picture1(Index), "Resize"
End Sub

Private Sub Picture1_Validate(Index As Integer, Cancel As Boolean)
RunEvent Picture1(Index), "Validate", Cancel
End Sub


Private Sub Text1_Change(Index As Integer)
RunEvent Text1(Index), "Change"
End Sub

Private Sub Text1_Click(Index As Integer)
RunEvent Text1(Index), "Click"
End Sub

Private Sub Text1_DblClick(Index As Integer)
RunEvent Text1(Index), "DblClick"
End Sub

Private Sub Text1_GotFocus(Index As Integer)
RunEvent Text1(Index), "GotFocus"
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Text1(Index), "KeyDown", KeyCode, Shift
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
RunEvent Text1(Index), "KeyPress", KeyAscii
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Text1(Index), "KeyUp", KeyCode, Shift
End Sub

Private Sub Text1_LostFocus(Index As Integer)
RunEvent Text1(Index), "LostFocus"
End Sub

Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Text1(Index), "MouseDown", Button, Shift, x, Y
End Sub

Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Text1(Index), "MouseMove", Button, Shift, x, Y
End Sub

Private Sub Text1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Text1(Index), "MouseUp", Button, Shift, x, Y
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
RunEvent Text1(Index), "Validate", Cancel
End Sub

Private Sub Timer1_Timer(Index As Integer)
RunEvent Timer1(Index), "Timer"
End Sub


Private Sub Option1_G_Click(Index As Integer)
RunEvent Option1_G(Index), "Click"
End Sub

Private Sub Option1_G_DblClick(Index As Integer)
RunEvent Option1_G(Index), "DblClick"
End Sub

Private Sub Option1_G_GotFocus(Index As Integer)
RunEvent Option1_G(Index), "GotFocus"
End Sub

Private Sub Option1_G_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Option1_G(Index), "KeyDown", KeyCode, Shift
End Sub

Private Sub Option1_G_KeyPress(Index As Integer, KeyAscii As Integer)
RunEvent Option1_G(Index), "KeyPress", KeyAscii
End Sub

Private Sub Option1_G_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
RunEvent Option1_G(Index), "KeyUp", KeyCode, Shift
End Sub

Private Sub Option1_G_LostFocus(Index As Integer)
RunEvent Option1_G(Index), "LostFocus"
End Sub

Private Sub Option1_G_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Option1_G(Index), "MouseDown", Button, Shift, x, Y
End Sub

Private Sub Option1_G_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Option1_G(Index), "MouseMove", Button, Shift, x, Y
End Sub

Private Sub Option1_G_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEvent Option1_G(Index), "MouseUp", Button, Shift, x, Y
End Sub

Private Sub Option1_G_Validate(Index As Integer, Cancel As Boolean)
RunEvent Option1_G(Index), "Validate", Cancel

End Sub


Function RunEvent(Object As Control, EventName As String, ParamArray x() As Variant)
If EVENT_READY = False Then Exit Function
On Error Resume Next

If UBound(x) = -1 Then
    Script.Run Object.Tag & "_" & EventName
    
Else
Select Case UBound(x)
    Case 0:  Script.Run Object.Tag & "_" & EventName, x(0)
    Case 1:  Script.Run Object.Tag & "_" & EventName, x(0), x(1)
    Case 2:  Script.Run Object.Tag & "_" & EventName, x(0), x(1), x(2)
    Case 3:  Script.Run Object.Tag & "_" & EventName, x(0), x(1), x(2), x(3)
    Case 4:  Script.Run Object.Tag & "_" & EventName, x(0), x(1), x(2), x(3), x(4)
End Select

End If


End Function

Function RunEventF(Object As String, EventName As String, ParamArray x() As Variant)
If EVENT_READY = False Then Exit Function
On Error Resume Next
'Exit Function

If UBound(x) = -1 Then
    Script.Run Object & "_" & EventName
    
Else
Select Case UBound(x)
    Case 0:  Script.Run Object & "_" & EventName, x(0)
    Case 1:  Script.Run Object & "_" & EventName, x(0), x(1)
    Case 2:  Script.Run Object & "_" & EventName, x(0), x(1), x(2)
    Case 3:  Script.Run Object & "_" & EventName, x(0), x(1), x(2), x(3)
    Case 4:  Script.Run Object & "_" & EventName, x(0), x(1), x(2), x(3), x(4)
End Select
    
End If

If Err.number <> 438 And Err.number <> 0 Then

MsgBox Object & "_" & EventName & vbCrLf & vbCrLf & "Error: " & Err.number & ", Disc: " & Err.Description & ", line: " & Script.Error.Line & vbCrLf & Err.Source

End If
End Function



Function update()

SkinnedForm1.update
'ClientArea1.update
'Me.ScaleMode = 1
Dim MyWidth As Long, MyHeight As Long
Dim BorderL As Long, BorderR As Long, BorderT As Long, BorderB As Long

MyWidth = Me.ScaleX(Me.ScaleWidth, Me.ScaleMode, 1)
MyHeight = Me.ScaleY(Me.ScaleHeight, Me.ScaleMode, 1)

BorderL = 60
BorderR = 45
BorderT = 345
BorderB = 45

ClientArea.Top = Me.ScaleX(BorderT, 1, Me.ScaleMode)
ClientArea.Left = Me.ScaleY(BorderL, 1, Me.ScaleMode)
ClientArea.height = Me.ScaleY(MyHeight - (BorderT + BorderB), 1, Me.ScaleMode)
ClientArea.width = Me.ScaleX(MyWidth - (BorderL + BorderR), 1, Me.ScaleMode)
ClientArea.BackColor = Me.BackColor
'(BorderL + BorderR)
'Me.ScaleMode = x

lblWarning.Top = Me.ScaleHeight - lblWarning.height
lblWarning.Left = 0
lblWarning.width = Me.width
'Me.height = Me.height + lblWarning.height

End Function

Function Finish()
fcheck
End Function

Private Sub ClientArea_Click()

RunEventF "Form", "Click"
End Sub

Private Sub ClientArea_DblClick()
RunEventF "Form", "DblClick"
End Sub
Private Sub ClientArea_GotFocus()
'frmSplash.Hide

RunEventF "Form", "GotFocus"
End Sub

Private Sub ClientArea_KeyDown(KeyCode As Integer, Shift As Integer)
RunEventF "Form", "KeyDown", KeyCode, Shift
End Sub

Private Sub ClientArea_KeyPress(KeyAscii As Integer)
RunEventF "Form", "KeyPress", KeyAscii
End Sub

Private Sub ClientArea_KeyUp(KeyCode As Integer, Shift As Integer)
RunEventF "Form", "KeyUp", KeyCode, Shift
End Sub

Private Sub ClientArea_LostFocus()
RunEventF "Form", "LostFocus"
End Sub

Private Sub ClientArea_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEventF "Form", "MouseDown", Button, Shift, x, Y
End Sub

Private Sub ClientArea_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEventF "Form", "MouseMove", Button, Shift, x, Y
End Sub

Private Sub ClientArea_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
RunEventF "Form", "MouseUp", Button, Shift, x, Y
End Sub

Private Sub ClientArea_Paint()
RunEventF "Form", "Paint"
End Sub

Private Sub ClientArea_Resize()
update
'Dim x As Integer
'x = Me.ScaleMode
RunEventF "Form", "Resize"
End Sub

