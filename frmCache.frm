VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Logon"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   4005
      TabIndex        =   5
      Top             =   885
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   2745
      TabIndex        =   4
      Top             =   885
      Width           =   1200
   End
   Begin VB.TextBox Text2 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1545
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   465
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1545
      TabIndex        =   0
      Top             =   75
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   285
      Left            =   675
      TabIndex        =   2
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   330
      Left            =   690
      TabIndex        =   1
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Show
BuildDialog Me, "c:\dialog.txt"
End Sub
