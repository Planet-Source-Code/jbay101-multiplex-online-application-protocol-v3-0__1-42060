VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF7F55&
      Height          =   2160
      Left            =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "[debug build]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF7F55&
      Height          =   210
      Left            =   3915
      TabIndex        =   1
      Top             =   1095
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   2100
      Picture         =   "frmSplash.frx":0000
      Top             =   420
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF7F55&
      Caption         =   "Copyright (C) 2001-2002 Jordan Bayliss-McCulloch."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   15
      TabIndex        =   0
      Top             =   1980
      Width           =   5400
   End
   Begin VB.Image Image2 
      Height          =   1290
      Left            =   360
      Picture         =   "frmSplash.frx":B4BA
      Top             =   345
      Width           =   1530
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
Me.Hide
End Sub
