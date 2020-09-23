VERSION 5.00
Begin VB.UserControl ClientArea 
   BackColor       =   &H008080FF&
   CanGetFocus     =   0   'False
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   ControlContainer=   -1  'True
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
End
Attribute VB_Name = "ClientArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Initialize()
update
End Sub


Function update()
On Error Resume Next
UserControl.BackColor = Parent.BackColor
'UserControl.ScaleMode = Parent.ScaleMode
End Function

Function SetSM(newSM As Integer)
'UserControl.ScaleMode = newSM
End Function

