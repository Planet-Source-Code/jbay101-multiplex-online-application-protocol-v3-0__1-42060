VERSION 5.00
Begin VB.Form frmLoading 
   AutoRedraw      =   -1  'True
   Caption         =   "Loading application..."
   ClientHeight    =   660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   660
   ScaleWidth      =   3420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2070
      Top             =   120
   End
   Begin MultiPlex.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   60
      Top             =   45
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   609
      Smooth          =   -1  'True
      Min             =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Initialising..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   435
      Width           =   3330
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
ProgressBar1.Position = 0

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Position = ProgressBar1.Position + 1

End Sub
