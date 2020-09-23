VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "APP Builder wizard - Step 1 of 5"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   Icon            =   "frmCompileWizard.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   405
      Left            =   5745
      TabIndex        =   2
      Top             =   4170
      Width           =   1125
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   405
      Left            =   7110
      TabIndex        =   1
      Top             =   4155
      Width           =   1125
   End
   Begin VB.PictureBox dlg2 
      BorderStyle     =   0  'None
      Height          =   2475
      Left            =   105
      ScaleHeight     =   2475
      ScaleWidth      =   8130
      TabIndex        =   0
      Top             =   1590
      Width           =   8130
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         Top             =   1305
         Width           =   4860
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse"
         Height          =   360
         Left            =   6690
         TabIndex        =   5
         Top             =   1275
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Source file:"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   1365
         Width           =   1455
      End
   End
   Begin VB.PictureBox dlg1 
      BorderStyle     =   0  'None
      Height          =   2475
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   8130
      TabIndex        =   8
      Top             =   1485
      Width           =   8130
      Begin VB.CommandButton Command2 
         Caption         =   "&Browse"
         Height          =   360
         Left            =   6690
         TabIndex        =   10
         Top             =   1275
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1620
         TabIndex        =   9
         Top             =   1305
         Width           =   4860
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Template  file:"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   1365
         Width           =   1455
      End
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   165
      X2              =   4995
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   135
      X2              =   8205
      Y1              =   4065
      Y2              =   4065
   End
   Begin VB.Label lblInfo 
      Caption         =   "If you have saved a previouse template, you can load it again. Select the template below."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   135
      TabIndex        =   4
      Top             =   765
      Width           =   8055
   End
   Begin VB.Label lblHeader 
      Caption         =   "Load previouse project"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   150
      TabIndex        =   3
      Top             =   105
      Width           =   7245
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

