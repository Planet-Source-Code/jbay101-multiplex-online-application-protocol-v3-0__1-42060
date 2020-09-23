VERSION 5.00
Begin VB.UserControl SkinnedForm 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7620
   ScaleHeight     =   5610
   ScaleWidth      =   7620
   ToolboxBitmap   =   "SkinForm.ctx":0000
   Begin VB.PictureBox CloseA 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6750
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   45
      Width           =   255
   End
   Begin VB.PictureBox MaxA 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6450
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   45
      Width           =   255
   End
   Begin VB.PictureBox MinA 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6210
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   45
      Width           =   240
   End
   Begin VB.Image RestoreOff 
      Height          =   255
      Left            =   4305
      Picture         =   "SkinForm.ctx":0312
      Top             =   2730
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image RestoreOn 
      Height          =   255
      Left            =   3660
      Picture         =   "SkinForm.ctx":06C8
      Top             =   2355
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image CloseDis 
      Height          =   255
      Left            =   3255
      Picture         =   "SkinForm.ctx":0A7E
      Top             =   3210
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image maxDis 
      Height          =   255
      Left            =   3090
      Picture         =   "SkinForm.ctx":0E34
      Top             =   2910
      Width           =   255
   End
   Begin VB.Image minDis 
      Height          =   255
      Left            =   1770
      Picture         =   "SkinForm.ctx":11EA
      Top             =   2220
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(Caption)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   60
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   0
      Picture         =   "SkinForm.ctx":15A0
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   90
      Picture         =   "SkinForm.ctx":191E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4725
   End
   Begin VB.Image Image3 
      Height          =   345
      Left            =   4815
      Picture         =   "SkinForm.ctx":20EC
      Top             =   0
      Width           =   2235
   End
   Begin VB.Image Image4 
      Height          =   1215
      Left            =   -15
      Picture         =   "SkinForm.ctx":496E
      Top             =   345
      Width           =   60
   End
   Begin VB.Image Image5 
      Height          =   3750
      Left            =   870
      Picture         =   "SkinForm.ctx":4D7C
      Stretch         =   -1  'True
      Top             =   165
      Width           =   60
   End
   Begin VB.Image Image6 
      Height          =   45
      Left            =   15
      Picture         =   "SkinForm.ctx":4EAE
      Top             =   5295
      Width           =   1095
   End
   Begin VB.Image Image7 
      Height          =   45
      Left            =   6435
      Picture         =   "SkinForm.ctx":5184
      Top             =   5295
      Width           =   615
   End
   Begin VB.Image Image8 
      Height          =   45
      Left            =   1335
      Picture         =   "SkinForm.ctx":533A
      Stretch         =   -1  'True
      Top             =   4245
      Width           =   5340
   End
   Begin VB.Image Image9 
      Height          =   4965
      Left            =   6165
      Picture         =   "SkinForm.ctx":5604
      Stretch         =   -1  'True
      Top             =   465
      Width           =   45
   End
   Begin VB.Image minOff 
      Height          =   255
      Left            =   1500
      Picture         =   "SkinForm.ctx":58CE
      Top             =   2235
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MaxOFF 
      Height          =   255
      Left            =   1005
      Picture         =   "SkinForm.ctx":5C40
      Top             =   1995
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image CloseOFF 
      Height          =   255
      Left            =   1755
      Picture         =   "SkinForm.ctx":5FF6
      Top             =   1755
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image MinON 
      Height          =   255
      Left            =   480
      Picture         =   "SkinForm.ctx":63AC
      Top             =   2670
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image MaxON 
      Height          =   255
      Left            =   1860
      Picture         =   "SkinForm.ctx":671E
      Top             =   2775
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image CloseOn 
      Height          =   255
      Left            =   2115
      Picture         =   "SkinForm.ctx":6AD4
      Top             =   3330
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "SkinnedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event MouseMove()
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2



Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        x As Long
        Y As Long
End Type

Private LastPoint As POINTAPI

Private lngTPPY As Long
Private lngTPPX As Long

Sub MouseMove(ctlForm As Object)
    ReleaseCapture
    SendMessage ctlForm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
Sub InitTPP()
    lngTPPX& = Screen.TwipsPerPixelX
    lngTPPY& = Screen.TwipsPerPixelY
End Sub

Private Sub CloseA_Click()
'End
Parent.Hide
Unload Parent
End Sub


Private Sub Image2_DblClick()
Swap
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If Button = 1 Then MouseDown
If Button = 2 Then SwapSkin
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then MouseMove Parent
End Sub

Private Sub Image3_DblClick()
Swap
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If Button = 1 Then MouseDown
If Button = 2 Then SwapSkin
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseMove
If Button = 1 Then MouseMove Parent
End Sub

Private Sub Label2_DblClick()
Swap
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If Button = 1 Then MouseDown
If Button = 2 Then SwapSkin
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseMove
If Button = 1 Then MouseMove Parent
End Sub

Function Max(v As Boolean)
If v = True Then
    MaxA.Picture = MaxOFF.Picture
    MaxA.Enabled = True
Else
    MaxA.Picture = maxDis.Picture
    MaxA.Enabled = False
End If

End Function

Function Min(v As Boolean)
If v = True Then
    MinA.Picture = minOff.Picture
    MinA.Enabled = True
Else
    MinA.Picture = minDis.Picture
    MinA.Enabled = False
End If
End Function

Function CloseB(v As Boolean)
If v = True Then
    CloseA.Picture = CloseOFF.Picture
    CloseA.Enabled = True
Else
    CloseA.Picture = CloseDis.Picture
    CloseA.Enabled = False
End If
End Function

Private Sub MaxA_Click()
If Parent.WindowState <> 2 Then
    Parent.WindowState = 2
Else
    Parent.WindowState = 0
End If

update
End Sub

Private Sub MinA_Click()
Parent.WindowState = 1
update
End Sub

Private Sub MinA_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MinA.Picture = MinON.Picture
End Sub

Private Sub MinA_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
MinA.Picture = minOff.Picture
End Sub


Private Sub maxA_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Parent.WindowState = 2 Then
    MaxA.Picture = RestoreOn.Picture
Else
   MaxA.Picture = MaxON.Picture
End If
End Sub

Private Sub MaxA_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Parent.WindowState = 2 Then
    MaxA.Picture = RestoreOff.Picture
Else
   MaxA.Picture = MaxOFF.Picture
End If

End Sub

Private Sub CloseA_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
CloseA.Picture = CloseOn.Picture
End Sub

Private Sub CloseA_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
CloseA.Picture = CloseOFF.Picture
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
InitTPP
Image10.Picture = frmIcon.Icon
UserControl_Resize
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseMove
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'MsgBox PropBag.Contents

End Sub

Private Sub UserControl_Resize()
'UserControl.Width = 2985
'UserControl.Height = 915
On Error Resume Next
'skin Parent

update

End Sub

Function update()
'frm.BorderStyle = 0
If Parent.WindowState = 1 Then Exit Function

'MaxA.Visible = Parent.MaxButton
'MinA.Visible = Parent.MinButton

If MaxA.Enabled = False Then
    If MinA.Enabled = False Then
        MinA.Visible = False
        MaxA.Visible = False
    
    End If
End If

If MaxA.Visible = False Then
    If MinA.Visible = True Then
        MinA.Left = MaxA.Left
    End If
End If
UserControl.BackColor = Parent.BackColor

UserControl.width = Parent.width
UserControl.height = Parent.height

If Parent.WindowState = 2 Then
    MaxA.Picture = RestoreOff.Picture
Else
   If MaxA.Enabled = True Then MaxA.Picture = MaxOFF.Picture
End If

If MinA.Enabled = True Then MinA.Picture = minOff.Picture
If CloseA.Enabled = True Then CloseA.Picture = CloseOFF.Picture

'MaxA.Visible = False

CloseA.Left = Parent.width - CloseA.width - 40
MaxA.Left = CloseA.Left - MaxA.width - 60
MinA.Left = MaxA.Left - MinA.width

Image1.Left = 0
Image1.Top = 0

Image2.Left = 75
Image2.Top = 0
Image2.width = Parent.width - Image3.width

Image3.Top = 0
Image3.Left = Parent.width - Image3.width

Image9.Top = 330
Image9.Left = Parent.width - Image9.width
Image9.height = Parent.height - 330 - Image7.height

Image7.Top = Parent.height - Image7.height
Image7.Left = Parent.width - Image7.width

Image8.Left = Image6.width
Image8.Top = Parent.height - Image8.height
Image8.width = Parent.width - Image6.width - Image7.width

Image6.Left = 0
Image6.Top = Parent.height - Image6.height

Image4.Top = Image1.height
Image4.Left = 0

Image5.Top = Image4.height + Image1.height
Image5.Left = 0
Image5.height = Parent.height - Image4.height - Image6.height

If Left(Parent.caption, Len("MultiPlex - ")) <> "MultiPlex - " Then
    Parent.caption = "MultiPlex - " & Parent.caption
End If

Label2.caption = Parent.caption

Label2.width = MinA.Left 'Parent.Width - Image3.Width
'Label2.Width = 0
'DesignGrid.Top = Image3.Height
'DesignGrid.Left = Image4.Width

'DesignGrid.Width = Parent.Width - Image4.Width - Image9.Width
'DesignGrid.Height = Parent.Height - Image3.Height - Image8.Height
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'PropBag.WriteProperty "Align", 1
End Sub

Function Swap()
If MaxA.Visible = False Then Exit Function
If Parent.WindowState = 0 Then
    Parent.WindowState = 2
Else
    Parent.WindowState = 0
End If
update
End Function

Function LoadSkin(ffile As String, curpath As String)
Dim file As String
Dim MyLine As String
Dim LineData() As String
Open ffile For Input As #2
Do While Not EOF(2)
    Line Input #2, MyLine

    If Left(Trim(MyLine), 2) = ";;" Then GoTo SkipLine
    If Trim(MyLine) = "" Then GoTo SkipLine
    LineData = Split(LCase(MyLine), " ")

    If LineData(0) = "element" Then
        file = curpath & "\" & LineData(2)
        Select Case LineData(1)
        Case "tilebar1": LoadImage Image2, file
        Case "tilebar2": LoadImage Image3, file
        Case "tilebar3": LoadImage Image1, file
        Case "border_top_left": LoadImage Image4, file
        Case "border_left": LoadImage Image5, file
        Case "border_bottom_left": LoadImage Image6, file
        Case "border_bottom": LoadImage Image8, file
        Case "border_bottom_right": LoadImage Image7, file
        Case "border_right": LoadImage Image9, file
        
        Case "button_close_normal": LoadImage CloseOFF, file
        Case "button_max_normal": LoadImage MaxOFF, file
        Case "button_min_normal": LoadImage minOff, file
        Case "button_restore_normal": LoadImage RestoreOff, file
        
        Case "button_close_disabled": LoadImage CloseDis, file
        Case "button_max_disabled": LoadImage maxDis, file
        Case "button_min_disabled": LoadImage minDis, file
        
        Case "button_close_down": LoadImage CloseOn, file
        Case "button_max_down": LoadImage MaxON, file
        Case "button_min_down": LoadImage MinON, file
        Case "button_restore_down": LoadImage RestoreOn, file
        Case Else
            MsgBox "Skin: element unknown: " & LineData(1)
        End Select
    
    End If
SkipLine:
Loop


Close #2
update
End Function

Function LoadImage(target As Image, file As String)
target.Picture = LoadPicture(file)
End Function

Function SwapSkin()
Set TRG = Me
PopupMenu frmIcon.mnuPopup

End Function
