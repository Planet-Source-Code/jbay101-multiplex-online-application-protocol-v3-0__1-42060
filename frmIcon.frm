VERSION 5.00
Begin VB.Form frmIcon 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuChangeSkin 
         Caption         =   "&Change skin"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuChangeSkin_Click()
Dim b As String
b = InputBox("Enter skin name:", "Select skins")

TRG.LoadSkin App.path & "\skins\" & b & "\" & "default" & ".skin", App.path & "\skins\" & b

End Sub

Private Sub mnuExit_Click()
End
End Sub
