VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Load"
      Height          =   1455
      Left            =   5610
      TabIndex        =   11
      Top             =   165
      Width           =   885
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save list"
      Height          =   420
      Left            =   150
      TabIndex        =   10
      Top             =   2715
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "recurse"
      Height          =   300
      Left            =   180
      TabIndex        =   9
      Top             =   1995
      Width           =   1710
   End
   Begin VB.CommandButton Command6 
      Caption         =   "remove"
      Height          =   1395
      Left            =   5250
      TabIndex        =   8
      Top             =   195
      Width           =   210
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   300
      Left            =   4035
      TabIndex        =   7
      Top             =   2820
      Width           =   990
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Compile"
      Height          =   465
      Left            =   1545
      TabIndex        =   6
      Top             =   2685
      Width           =   2220
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Browse"
      Height          =   375
      Left            =   3945
      TabIndex        =   5
      Top             =   1665
      Width           =   1230
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   825
      TabIndex        =   4
      Text            =   "Select output .app"
      Top             =   1665
      Width           =   2940
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   75
      TabIndex        =   3
      Top             =   510
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   315
      Left            =   4005
      TabIndex        =   2
      Top             =   90
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   330
      Left            =   2865
      TabIndex        =   1
      Top             =   75
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Text            =   "Select file"
      Top             =   105
      Width           =   2685
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   -495
      Top             =   2535
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select files"
      Filter          =   "*.*"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As CGZipFiles

Private Sub Command1_Click()
dlg.ShowOpen
Text1.Text = dlg.FileName
End Sub

Private Sub Command2_Click()
List1.AddItem Text1.Text


End Sub

Private Sub Command3_Click()
dlg.ShowSave
Text2.Text = dlg.FileName
End Sub

Private Sub Command4_Click()
Dim i As Integer
Set x = New CGZipFiles
For i = 0 To List1.ListCount
    If List1.List(i) <> "" Then
        x.AddFile List1.List(i)
    End If
Next i
If LCase(Right(Text2.Text, 4)) <> ".app" Then Text2.Text = Text2.Text & ".app"

x.ZipFileName = Text2.Text
x.Encrypted = True
x.RecurseFolders = Check1.Value
x.MakeZipFile
MsgBox "APP Compiled!"

End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command7_Click()
dlg.ShowSave
Dim i As Integer
Open dlg.FileName For Output As #1

For i = 0 To List1.ListCount
    If List1.List(i) <> "" Then
        Print #1, List1.List(i)
    End If
Next i
Close 31
End Sub

Private Sub Command8_Click()
dlg.ShowOpen

Dim i As String
Open dlg.FileName For Input As #1
Do While Not EOF(1)
    Line Input #1, i
    List1.AddItem i

Loop

Close #1

End Sub
