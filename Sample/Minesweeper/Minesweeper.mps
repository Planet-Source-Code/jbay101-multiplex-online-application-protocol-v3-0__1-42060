#version 1.0;
#section .validate {
   validate platform "Win32";
   validate script "VBScript"; 
   validate script "JScript";
}

#section .link {
   include def "common.def";
   include def "win32.def";

   interface #sys.CMath;
   interface #sys.CLib;
   interface #sys.Common;
   interface #sys.Win32;
   interface #sys.CString;
   interface #sys.CInt;
   interface #sys.SharedMem;
   interface #sys.CConvert;
   interface #sys.CFileSystem;
   interface #sys.CFileIO;
   interface #sys.CDebug;
   interface #sys.CRegistry;
   interface #sys.CForm;
   interface #sys.CScript;
}

#section .runtime {
   define calling_convention cdec, stdcall;
   define param_convention VB_C_COMPATIBLE;

}

#section .code {
   import "VBScript" {
'+36
Dim frmMainID, frmMain, picFocus, picFlags, picTimeCount, optNewGame, tmrCount, img2ndIcon, imgSpecials_3, imgFaces_3, imgCounts_10, imgFaces_2, imgFaces_1, imgFaces_0, imgCounts_9, imgCounts_8, imgCounts_7, imgCounts_6, imgCounts_5, imgCounts_4, imgCounts_3, imgCounts_2, imgCounts_1, imgCounts_0, imgSpecials_0, imgSpecials_2, imgSpecials_1, imgSpecials_5, imgSpecials_4, imgNumbers_8, imgNumbers_7, imgNumbers_6, imgNumbers_5, imgNumbers_4, imgNumbers_3, imgNumbers_2, imgNumbers_1, imgNumbers_0, mnuGame, mnuNew, mnuSep1, mnuLevel_0, mnuLevel_1, mnuLevel_2, mnuLevel_3, mnuSep2, mnuMarks, mnuColor, mnuSep3, mnuBest, mnuSep4, mnuExit, mnuHelp, mnuTopics, mnuSep5, mnuAbout

Dim picArea

frmMainID = CForm.CreateForm()
Set frmMain = CForm.frmobj(frmMainID)

CForm.CreateControl frmMainID, "picFocus", "PictureBox"
Set picFocus = CForm.obj(frmMainID, "picFocus")

CForm.CreateControl frmMainID, "picFlags", "PictureBox"
Set picFlags = CForm.obj(frmMainID, "picFlags")

CForm.CreateControl frmMainID, "picTimeCount", "PictureBox"
Set picTimeCount = CForm.obj(frmMainID, "picTimeCount")

CForm.CreateControl frmMainID, "optNewGame", "OptionButton(G)"
Set optNewGame = CForm.obj(frmMainID, "optNewGame")

CForm.CreateControl frmMainID, "picArea", "PictureBox"
Set picArea = CForm.obj(frmMainID, "picArea")

CForm.CreateControl frmMainID, "tmrCount", "Timer"
Set tmrCount = CForm.obj(frmMainID, "tmrCount")

CForm.CreateControl frmMainID, "img2ndIcon", "Image"
Set img2ndIcon = CForm.obj(frmMainID, "img2ndIcon")

CForm.CreateControl frmMainID, "imgSpecials_3", "Image"
Set imgSpecials_3 = CForm.obj(frmMainID, "imgSpecials_3")

CForm.CreateControl frmMainID, "imgFaces_3", "Image"
Set imgFaces_3 = CForm.obj(frmMainID, "imgFaces_3")

CForm.CreateControl frmMainID, "imgCounts_10", "Image"
Set imgCounts_10 = CForm.obj(frmMainID, "imgCounts_10")

CForm.CreateControl frmMainID, "imgFaces_2", "Image"
Set imgFaces_2 = CForm.obj(frmMainID, "imgFaces_2")

CForm.CreateControl frmMainID, "imgFaces_1", "Image"
Set imgFaces_1 = CForm.obj(frmMainID, "imgFaces_1")

CForm.CreateControl frmMainID, "imgFaces_0", "Image"
Set imgFaces_0 = CForm.obj(frmMainID, "imgFaces_0")

CForm.CreateControl frmMainID, "imgCounts_9", "Image"
Set imgCounts_9 = CForm.obj(frmMainID, "imgCounts_9")

CForm.CreateControl frmMainID, "imgCounts_8", "Image"
Set imgCounts_8 = CForm.obj(frmMainID, "imgCounts_8")

CForm.CreateControl frmMainID, "imgCounts_7", "Image"
Set imgCounts_7 = CForm.obj(frmMainID, "imgCounts_7")

CForm.CreateControl frmMainID, "imgCounts_6", "Image"
Set imgCounts_6 = CForm.obj(frmMainID, "imgCounts_6")

CForm.CreateControl frmMainID, "imgCounts_5", "Image"
Set imgCounts_5 = CForm.obj(frmMainID, "imgCounts_5")

CForm.CreateControl frmMainID, "imgCounts_4", "Image"
Set imgCounts_4 = CForm.obj(frmMainID, "imgCounts_4")

CForm.CreateControl frmMainID, "imgCounts_3", "Image"
Set imgCounts_3 = CForm.obj(frmMainID, "imgCounts_3")

CForm.CreateControl frmMainID, "imgCounts_2", "Image"
Set imgCounts_2 = CForm.obj(frmMainID, "imgCounts_2")

CForm.CreateControl frmMainID, "imgCounts_1", "Image"
Set imgCounts_1 = CForm.obj(frmMainID, "imgCounts_1")

CForm.CreateControl frmMainID, "imgCounts_0", "Image"
Set imgCounts_0 = CForm.obj(frmMainID, "imgCounts_0")

CForm.CreateControl frmMainID, "imgSpecials_0", "Image"
Set imgSpecials_0 = CForm.obj(frmMainID, "imgSpecials_0")

CForm.CreateControl frmMainID, "imgSpecials_2", "Image"
Set imgSpecials_2 = CForm.obj(frmMainID, "imgSpecials_2")

CForm.CreateControl frmMainID, "imgSpecials_1", "Image"
Set imgSpecials_1 = CForm.obj(frmMainID, "imgSpecials_1")

CForm.CreateControl frmMainID, "imgSpecials_5", "Image"
Set imgSpecials_5 = CForm.obj(frmMainID, "imgSpecials_5")

CForm.CreateControl frmMainID, "imgSpecials_4", "Image"
Set imgSpecials_4 = CForm.obj(frmMainID, "imgSpecials_4")

CForm.CreateControl frmMainID, "imgNumbers_8", "Image"
Set imgNumbers_8 = CForm.obj(frmMainID, "imgNumbers_8")

CForm.CreateControl frmMainID, "imgNumbers_7", "Image"
Set imgNumbers_7 = CForm.obj(frmMainID, "imgNumbers_7")

CForm.CreateControl frmMainID, "imgNumbers_6", "Image"
Set imgNumbers_6 = CForm.obj(frmMainID, "imgNumbers_6")

CForm.CreateControl frmMainID, "imgNumbers_5", "Image"
Set imgNumbers_5 = CForm.obj(frmMainID, "imgNumbers_5")

CForm.CreateControl frmMainID, "imgNumbers_4", "Image"
Set imgNumbers_4 = CForm.obj(frmMainID, "imgNumbers_4")

CForm.CreateControl frmMainID, "imgNumbers_3", "Image"
Set imgNumbers_3 = CForm.obj(frmMainID, "imgNumbers_3")

CForm.CreateControl frmMainID, "imgNumbers_2", "Image"
Set imgNumbers_2 = CForm.obj(frmMainID, "imgNumbers_2")

CForm.CreateControl frmMainID, "imgNumbers_1", "Image"
Set imgNumbers_1 = CForm.obj(frmMainID, "imgNumbers_1")

CForm.CreateControl frmMainID, "imgNumbers_0", "Image"
Set imgNumbers_0 = CForm.obj(frmMainID, "imgNumbers_0")


frmMain.Caption = "Minesweeper"
frmMain.Height = 4065
frmMain.Width = 2874
frmMain.Top = 0
frmMain.Left = 0
frmMain.AutoRedraw = False
frmMain.Enabled = True
frmMain.BackColor = &H00C0C0C0&'-2147483633
frmMain.ForeColor = -2147483630
frmMain.ScaleMode = 3
Set frmMain.Icon = CForm.LoadImage("%path%\frmMain.ico")
picFocus.Left = -128
picFocus.Top = 240
picFocus.Width = 89
picFocus.Height = 17
picFocus.ScaleMode = 1
picFocus.BackColor = -2147483633
picFocus.BorderStyle = 1
picFocus.AutoRedraw = False
picFocus.Visible = True
picFocus.FillColor = 0
picFocus.FillStyle = 1
picFocus.DrawMode = 13
picFocus.Font.Bold = False
picFocus.Font.Size = 8.25
picFocus.Font.Italic = False
picFocus.Font.Name = "MS Sans Serif"

picFlags.Left = 17
picFlags.Top = 16
picFlags.Width = 39
picFlags.Height = 23
picFlags.ScaleMode = 3
picFlags.BackColor = -2147483633
picFlags.BorderStyle = 0
picFlags.AutoRedraw = True
picFlags.Visible = True
picFlags.FillColor = 0
picFlags.FillStyle = 1
picFlags.DrawMode = 13
picFlags.Font.Bold = False
picFlags.Font.Size = 8.25
picFlags.Font.Italic = False
picFlags.Font.Name = "MS Sans Serif"

picTimeCount.Left = 127
picTimeCount.Top = 16
picTimeCount.Width = 39
picTimeCount.Height = 23
picTimeCount.ScaleMode = 3
picTimeCount.BackColor = -2147483633
picTimeCount.BorderStyle = 0
picTimeCount.AutoRedraw = True
picTimeCount.Visible = True
picTimeCount.FillColor = 0
picTimeCount.FillStyle = 1
picTimeCount.DrawMode = 13
picTimeCount.Font.Bold = False
picTimeCount.Font.Size = 8.25
picTimeCount.Font.Italic = False
picTimeCount.Font.Name = "MS Sans Serif"

optNewGame.Left = 81
optNewGame.Top = 16
optNewGame.Width = 24
optNewGame.Height = 24
optNewGame.Caption = ""
optNewGame.Font.Bold = False
optNewGame.Font.Size = 8.25
optNewGame.Font.Italic = False
optNewGame.Font.Name = "MS Sans Serif"
optNewGame.Visible = True
optNewGame.BackColor = -2147483633
optNewGame.Enabled = True
optNewGame.TabIndex = 1
optNewGame.Value = False

picArea.Left = 12
picArea.Top = 55
picArea.Width = 160
picArea.Height = 160
picArea.ScaleMode = 3
picArea.BackColor = &H00C0C0C0&'-2147483633
picArea.BorderStyle = 0
picArea.AutoRedraw = True
picArea.Visible = True
picArea.FillColor = 0
picArea.FillStyle = 1
picArea.DrawMode = 13
picArea.Font.Bold = False
picArea.Font.Size = 8.25
picArea.Font.Italic = False
picArea.Font.Name = "MS Sans Serif"


tmrCount.Enabled = False
tmrCount.Interval = 50

img2ndIcon.Left = 8
img2ndIcon.Top = 261
img2ndIcon.Width = 32
img2ndIcon.Height = 32
img2ndIcon.Visible = False
Set img2ndIcon.Picture = CForm.LoadImage("%path%\img2ndIcon.bmp")

imgSpecials_3.Left = 80
imgSpecials_3.Top = 240
imgSpecials_3.Width = 16
imgSpecials_3.Height = 16
imgSpecials_3.Visible = False
Set imgSpecials_3.Picture = CForm.LoadImage("%path%\imgSpecials_3.bmp")

imgFaces_3.Left = 176
imgFaces_3.Top = 168
imgFaces_3.Width = 21
imgFaces_3.Height = 21
imgFaces_3.Visible = False
Set imgFaces_3.Picture = CForm.LoadImage("%path%\imgFaces_3.bmp")

imgCounts_10.Left = 176
imgCounts_10.Top = 104
imgCounts_10.Width = 13
imgCounts_10.Height = 23
imgCounts_10.Visible = False
Set imgCounts_10.Picture = CForm.LoadImage("%path%\imgCounts_10.bmp")

imgFaces_2.Left = 152
imgFaces_2.Top = 168
imgFaces_2.Width = 21
imgFaces_2.Height = 21
imgFaces_2.Visible = False
Set imgFaces_2.Picture = CForm.LoadImage("%path%\imgFaces_2.bmp")

imgFaces_1.Left = 176
imgFaces_1.Top = 144
imgFaces_1.Width = 21
imgFaces_1.Height = 21
imgFaces_1.Visible = False
Set imgFaces_1.Picture = CForm.LoadImage("%path%\imgFaces_1.bmp")

imgFaces_0.Left = 152
imgFaces_0.Top = 144
imgFaces_0.Width = 21
imgFaces_0.Height = 21
imgFaces_0.Visible = False
Set imgFaces_0.Picture = CForm.LoadImage("%path%\imgFaces_0.bmp")

imgCounts_9.Left = 208
imgCounts_9.Top = 80
imgCounts_9.Width = 13
imgCounts_9.Height = 23
imgCounts_9.Visible = False
Set imgCounts_9.Picture = CForm.LoadImage("%path%\imgCounts_9.bmp")

imgCounts_8.Left = 192
imgCounts_8.Top = 80
imgCounts_8.Width = 13
imgCounts_8.Height = 23
imgCounts_8.Visible = False
Set imgCounts_8.Picture = CForm.LoadImage("%path%\imgCounts_8.bmp")

imgCounts_7.Left = 176
imgCounts_7.Top = 80
imgCounts_7.Width = 13
imgCounts_7.Height = 23
imgCounts_7.Visible = False
Set imgCounts_7.Picture = CForm.LoadImage("%path%\imgCounts_7.bmp")

imgCounts_6.Left = 208
imgCounts_6.Top = 56
imgCounts_6.Width = 13
imgCounts_6.Height = 23
imgCounts_6.Visible = False
Set imgCounts_6.Picture = CForm.LoadImage("%path%\imgCounts_6.bmp")

imgCounts_5.Left = 192
imgCounts_5.Top = 56
imgCounts_5.Width = 13
imgCounts_5.Height = 23
imgCounts_5.Visible = False
Set imgCounts_5.Picture = CForm.LoadImage("%path%\imgCounts_5.bmp")

imgCounts_4.Left = 176
imgCounts_4.Top = 56
imgCounts_4.Width = 13
imgCounts_4.Height = 23
imgCounts_4.Visible = False
Set imgCounts_4.Picture = CForm.LoadImage("%path%\imgCounts_4.bmp")

imgCounts_3.Left = 208
imgCounts_3.Top = 32
imgCounts_3.Width = 13
imgCounts_3.Height = 23
imgCounts_3.Visible = False
Set imgCounts_3.Picture = CForm.LoadImage("%path%\imgCounts_3.bmp")

imgCounts_2.Left = 192
imgCounts_2.Top = 32
imgCounts_2.Width = 13
imgCounts_2.Height = 23
imgCounts_2.Visible = False
Set imgCounts_2.Picture = CForm.LoadImage("%path%\imgCounts_2.bmp")

imgCounts_1.Left = 176
imgCounts_1.Top = 32
imgCounts_1.Width = 13
imgCounts_1.Height = 23
imgCounts_1.Visible = False
Set imgCounts_1.Picture = CForm.LoadImage("%path%\imgCounts_1.bmp")

imgCounts_0.Left = 192
imgCounts_0.Top = 104
imgCounts_0.Width = 13
imgCounts_0.Height = 23
imgCounts_0.Visible = False
Set imgCounts_0.Picture = CForm.LoadImage("%path%\imgCounts_0.bmp")

imgSpecials_0.Left = 8
imgSpecials_0.Top = 240
imgSpecials_0.Width = 16
imgSpecials_0.Height = 16
imgSpecials_0.Visible = False
Set imgSpecials_0.Picture = CForm.LoadImage("%path%\imgSpecials_0.bmp")

imgSpecials_2.Left = 56
imgSpecials_2.Top = 240
imgSpecials_2.Width = 16
imgSpecials_2.Height = 16
imgSpecials_2.Visible = False
Set imgSpecials_2.Picture = CForm.LoadImage("%path%\imgSpecials_2.bmp")

imgSpecials_1.Left = 32
imgSpecials_1.Top = 240
imgSpecials_1.Width = 16
imgSpecials_1.Height = 16
imgSpecials_1.Visible = False
Set imgSpecials_1.Picture = CForm.LoadImage("%path%\imgSpecials_1.bmp")

imgSpecials_5.Left = 128
imgSpecials_5.Top = 240
imgSpecials_5.Width = 16
imgSpecials_5.Height = 16
imgSpecials_5.Visible = False
Set imgSpecials_5.Picture = CForm.LoadImage("%path%\imgSpecials_5.bmp")

imgSpecials_4.Left = 104
imgSpecials_4.Top = 240
imgSpecials_4.Width = 16
imgSpecials_4.Height = 16
imgSpecials_4.Visible = False
Set imgSpecials_4.Picture = CForm.LoadImage("%path%\imgSpecials_4.bmp")

imgNumbers_8.Left = 200
imgNumbers_8.Top = 216
imgNumbers_8.Width = 16
imgNumbers_8.Height = 16
imgNumbers_8.Visible = False
Set imgNumbers_8.Picture = CForm.LoadImage("%path%\imgNumbers_8.bmp")

imgNumbers_7.Left = 176
imgNumbers_7.Top = 216
imgNumbers_7.Width = 16
imgNumbers_7.Height = 16
imgNumbers_7.Visible = False
Set imgNumbers_7.Picture = CForm.LoadImage("%path%\imgNumbers_7.bmp")

imgNumbers_6.Left = 152
imgNumbers_6.Top = 216
imgNumbers_6.Width = 16
imgNumbers_6.Height = 16
imgNumbers_6.Visible = False
Set imgNumbers_6.Picture = CForm.LoadImage("%path%\imgNumbers_6.bmp")

imgNumbers_5.Left = 128
imgNumbers_5.Top = 216
imgNumbers_5.Width = 16
imgNumbers_5.Height = 16
imgNumbers_5.Visible = False
Set imgNumbers_5.Picture = CForm.LoadImage("%path%\imgNumbers_5.bmp")

imgNumbers_4.Left = 104
imgNumbers_4.Top = 216
imgNumbers_4.Width = 16
imgNumbers_4.Height = 16
imgNumbers_4.Visible = False
Set imgNumbers_4.Picture = CForm.LoadImage("%path%\imgNumbers_4.bmp")

imgNumbers_3.Left = 80
imgNumbers_3.Top = 216
imgNumbers_3.Width = 16
imgNumbers_3.Height = 16
imgNumbers_3.Visible = False
Set imgNumbers_3.Picture = CForm.LoadImage("%path%\imgNumbers_3.bmp")

imgNumbers_2.Left = 56
imgNumbers_2.Top = 216
imgNumbers_2.Width = 16
imgNumbers_2.Height = 16
imgNumbers_2.Visible = False
Set imgNumbers_2.Picture = CForm.LoadImage("%path%\imgNumbers_2.bmp")

imgNumbers_1.Left = 32
imgNumbers_1.Top = 216
imgNumbers_1.Width = 16
imgNumbers_1.Height = 16
imgNumbers_1.Visible = False
Set imgNumbers_1.Picture = CForm.LoadImage("%path%\imgNumbers_1.bmp")

imgNumbers_0.Left = 8
imgNumbers_0.Top = 216
imgNumbers_0.Width = 16
imgNumbers_0.Height = 16
imgNumbers_0.Visible = False
Set imgNumbers_0.Picture = CForm.LoadImage("%path%\imgNumbers_0.bmp")

MAXNAMELEN = 16
MAXCOMPLEN = 256

 Const L1ROW  = 100
 Const L1COL = 100
 Const L1MINE  = 10
 Const L2ROW  = 16
 Const L2COL  = 16
 Const L2MINE  = 40
 Const L3ROW  = 16
 Const L3COL  = 30
 Const L3MINE  = 99
 Const MINROW = 8
 Const MINCOL  = 8
 Const MAXROW  = 24
 Const MAXCOL  = 30
 Const MINMINE  = 10
 Const MAXTIME  = 999
 Const NUMBERWIDTH  = 13
 Const DARKSHADOW  = &H80000015
 Const Beginner = 0
Const Intermediate = 1
Const  Expert = 2
Const  Custom = 3

Const  Closed = 0
Const  Flaged = 1
Const  Questioned = 2
Const  Opened = 3

Const  ClearThem = 1
Const  CloseThem = 0
Const  OpenThem = -1

Const  Cross = 0
Const  Flag = 1
Const  Question = 2
Const  Quest2 = 3
Const  Hit = 4
Const  RedHit = 4


Const  Smile = 0
Const  Surprise = 1
Const  Sad = 2
Const  Win = 3

'consants
Const vbLeftButton = 1
Const vbRightButton = 2
Const vbMiddleButton = 4
Const vbButtonShadow = &H80000010



Dim LevelNow 
Dim UserName
Dim ThePosition
Dim Winner(2)

'-Area variables
    Dim FirstButtonTimeMain
    Dim  SecondButtonTimeMain
    Dim  LastButton
Dim Screen
	Set Screen  = CForm.CreateObject("Screen")

Dim Newing   '-I know there's no such word
Dim Selecting
Dim Cheating
Dim GameStarted
Dim Mines
Dim State()
Dim RowNumber, ColNumber, MineNumber
Dim FlagNumber
Dim CountNumber
Dim SmallCount
Dim EmptyNumber, OpenedNumber
Dim psw, psh
Dim Row, Col

'-Button variables
Dim OnButton
Dim OnLeftButton
Dim OnRightButton

'-Game settings
Const OnQMark = True
Const OnColor = True


Sub InitializeGame()
    
	'imgNumbers_0.Visible=true
	'CForm.obj(frmMainID, "picArea").Visible=False
	PrepareVariables
	
	PrepareControls
	
End Sub

Sub PrepareVariables()
	ReDim Mines(ColNumber - 1, RowNumber - 1)
    ReDim State(ColNumber - 1, RowNumber - 1)
    
	FlagNumber = MineNumber
    SmallCount = 0
    CountNumber = 0
    OpenedNumber = 0
    EmptyNumber = RowNumber * ColNumber - MineNumber
    GameStarted = False
    OnLeftButton = False
    OnRightButton = False
End Sub

Sub PrepareControls()
Dim a,b
a = ColNumber * 16
b = RowNumber * 16
    
	picArea.Width = a
    picArea.Height = b
	
	'--Special Case---
    psw = picArea.ScaleWidth
    psh = picArea.ScaleHeight
   
		'-----------------
    frmMain.Width = (picArea.Left + picArea.Width + 22) * Screen.TwipsPerPixelX
    frmMain.Height = (picArea.Top + picArea.Height + 40) * Screen.TwipsPerPixelY '56 with menu
    picTimeCount.Left = frmMain.ScaleWidth - picTimeCount.Width - 20
    optNewGame.Left = (frmMain.ScaleWidth - optNewGame.Width) / 2
    
    '-Doesn't need this here
    tmrCount.Enabled = False
    
    ShowCounts picFlags, FlagNumber
    optNewGame.Picture = imgFaces_0
    ShowCounts picTimeCount, 0
    
    frmMain.Cls
    DoAllBorders
    picArea.Enabled = True
    picArea.Cls
    DrawSquares

		Dim i, j
		for i = 0 to ColNumber - 1
		for j = 0 to RowNumber - 1
			State(i,j)=Closed
			'DrawState i, j, Closed
			
		next

	next

End Sub

Sub SetUpLevel(Level)
    Select Case Level
        Case Beginner
            RowNumber = L1ROW
            ColNumber = L1COL
            MineNumber = L1MINE
        Case Intermediate
            RowNumber = L2ROW
            ColNumber = L2COL
            MineNumber = L2MINE
        Case Expert
            RowNumber = L3ROW
            ColNumber = L3COL
            MineNumber = L3MINE
        Case Custom
    End Select
End Sub

Sub PutMines(X, Y)
    Dim i
    Dim randX, randY
    
    For i = 1 To MineNumber
        randY = Rnd * (RowNumber - 1)
        randX = Rnd * (ColNumber - 1)
        If Mines(randX, randY) Or (randX = X And randY = Y) Then
            i = i - 1
        Else
            Mines(randX, randY) = True
        End If
    Next
End Sub

Sub WinGame()
    tmrCount.Enabled = False
    picArea.Enabled = False
    optNewGame.Picture = imgFaces_3
    If FlagNumber > 0 Then ShowMinesLocation imgSpecials_1 'Flag
    ShowCounts picFlags, 0
    
    If LevelNow = Custom Then Exit Sub
    
    'If CountNumber < Winner(LevelNow).TimeMain Then
    '    Winner(LevelNow).TifrmMain = CountNumber
'        frmInput.Show vbModal
    'End If
End Sub

Sub LoseGame()
	ShowMinesLocation imgSpecials_4'(Hit)
    tmrCount.Enabled = False
    picArea.Enabled = False
    optNewGame.Picture = imgFaces_2
End Sub

Sub GetLimits(X, Y, LeftLimit, RightLimit, TopLimit, BottomLimit)

    LeftLimit = CCommon.vbIIf(X > 0, X - 1, 0)
    RightLimit = CCommon.vbIIf(X < ColNumber - 1, X + 1, ColNumber - 1)
    TopLimit = CCommon.vbIIf(Y > 0, Y - 1, 0)
    BottomLimit = CCommon.vbIIf(Y < RowNumber - 1, Y + 1, RowNumber - 1)

End Sub

Function CheckDanger(X, Y)
    Dim Val
	If Mines(X, Y) = True Then
        CheckDanger = -1
        Exit Function
    End If
    
    Dim LeftLim, RightLim 
    Dim TopLim, BottomLim
    GetLimits X, Y, LeftLim, RightLim, TopLim, BottomLim
    
    Dim i, j
    For j = TopLim To BottomLim
        For i = LeftLim To RightLim
            If Mines(i, j) = True Then Val = Val + 1
        Next
    Next

	if Val = "" then Val = 0
	CheckDanger = Val

End Function

Sub CheckAround(X, Y, CallX, CallY)
    Dim i, j
    Dim LeftLim, RightLim
    Dim TopLim, BottomLim
    Dim nDanger

    GetLimits X, Y, LeftLim, RightLim, TopLim, BottomLim
    
    For j = TopLim To BottomLim
        For i = LeftLim To RightLim
            If State(i, j) = Closed Then
                nDanger = CheckDanger(i, j)
				
				State(i, j) = Opened
                DoOpen
                ShowDanger i, j, nDanger
                
				If nDanger = 0 Then CheckAround2 i, j, X, Y
            End If
        Next
    Next
End Sub

Sub CheckAround2(a, b, c,d)
CheckAround a,b,c,d
End Sub

Function CheckState(X, Y, WhatState)
    Dim LeftLim, RightLim
    Dim TopLim, BottomLim
	Dim Val
    Val=0
    GetLimits X, Y, LeftLim, RightLim, TopLim, BottomLim
    
    Dim i, j
    
    For j = TopLim To BottomLim
        For i = LeftLim To RightLim
            If State(i, j) = WhatState Then
                Val = Val + 1
            End If
        Next
    Next

	CheckState = Val
End Function

Sub DoOpen()
    OpenedNumber = OpenedNumber + 1
    If OpenedNumber = EmptyNumber Then WinGame
End Sub

Sub DoFlag(Flag)   '-Flag will be 1 or -1
    FlagNumber = FlagNumber - Flag
    ShowCounts picFlags, FlagNumber
End Sub

Sub DoCount()
    CountNumber = CountNumber + 1
    ShowCounts picTimeCount, CountNumber
End Sub

Function DoMiddle(X, Y, Action)
    Dim LeftLim, RightLim
    Dim TopLim, BottomLim
    
    GetLimits X, Y, LeftLim, RightLim, TopLim, BottomLim
    
	
    Dim i, j
     For j = TopLim To BottomLim
        For i = LeftLim To RightLim
            Select Case Action
                Case ClearThem
                    If State(i, j) = Closed Then
                        
						ClearBox i, j
                    ElseIf State(i, j) = Questioned Then
						'Quest2
                        picArea.PaintPicture imgSpecials_3, i * 16, j * 16
                    End If
                Case CloseThem
                    If State(i, j) = Closed Then
                        DrawBox i, j
                    ElseIf State(i, j) = Questioned Then
						'question
                        picArea.PaintPicture imgSpecials_2, i * 16, j * 16
                    End If
                Case OpenThem
                    GoLeftUp i, j
            End Select
        Next
    Next
End Function

Function MaxMines()
    MaxMines = (RowNumber - 1) * (ColNumber - 1)
End Function

Sub ShowDanger(X, Y, nDanger)
    Dim xx, yy
	xx=x
	yy=y

	xx = xx * 16
    yy = yy * 16
    If nDanger = -1 Then
		picArea.PaintPicture imgSpecials_5, xx, yy 'redhit
    Else
		if nDanger = 0 then picArea.PaintPicture imgNumbers_0, xx, yy
		if nDanger = 1 then picArea.PaintPicture imgNumbers_1, xx, yy
		if nDanger = 2 then picArea.PaintPicture imgNumbers_2, xx, yy
		if nDanger = 3 then picArea.PaintPicture imgNumbers_3, xx, yy
		if nDanger = 4 then picArea.PaintPicture imgNumbers_4, xx, yy
		if nDanger = 5 then picArea.PaintPicture imgNumbers_5, xx, yy
		if nDanger = 6 then picArea.PaintPicture imgNumbers_6, xx, yy

    End If
End Sub

Sub DoAllBorders()
    '-Outermost border
    DrawBorder frmMain, 3, 3, frmMain.ScaleWidth - 7, _
        frmMain.ScaleHeight - 7, 3, vbWhite, vbButtonShadow
    '-Upper
    DrawBorder frmMain, 11, 11, frmMain.ScaleWidth - 25, 32, 2, vbButtonShadow, vbWhite
    '-Lower
    DrawBorder frmMain, 12, 55, frmMain.ScaleWidth - 27 _
        , frmMain.ScaleHeight - 68, 3, vbButtonShadow, vbWhite
    '-Flags
    DrawBorder frmMain, picFlags.Left, picFlags.Top, picFlags.Width, _
        picFlags.Height - 1, 1, vbButtonShadow, vbWhite
    '-NewButton
    DrawBorder frmMain, optNewGame.Left, optNewGame.Top, 22, _
        22, 1, vbButtonShadow, DARKSHADOW
    '-TifrmMainCount
    DrawBorder frmMain, picTimeCount.Left, picTimeCount.Top, picTimeCount.Width, _
        picTimeCount.Height - 1, 1, vbButtonShadow, vbWhite
End Sub

Sub DrawBorder(Bg, X, Y, W, H, BorderWidth, TopLeft_Color, BottomRight_Color)
    
    Dim b

    For b = 1 To BorderWidth
        Bg.CurrentX = X - b
        Bg.CurrentY = Y - b
        ' Top
        
	CForm.Line2 Bg, X + W + b, Y - b, TopLeft_Color
'	Bg.Line -X + W + b, Y - b, TopLeft_Color
        ' Right
        CForm.Line2 Bg, X + W + b, Y + H + b, BottomRight_Color
        ' Bottom
	CForm.Line2 Bg, X - b, Y + H + b, BottomRight_Color
        ' Left
	CForm.Line2 Bg, X - b, Y - b, TopLeft_Color
    Next
End Sub

Sub ClearBox(X, Y)
    Dim xx,yy
	xx=X
	yy=Y
	xx = xx * 16
    yy = yy * 16
    picArea.PaintPicture imgNumbers_0, xx, yy
End Sub

Sub DrawBox(X, Y)
    Dim xx,yy
	xx=x
	yy=y
	xx = xx * 16
    yy = yy * 16

    CForm.Line picArea, xx, yy,xx + 15, yy + 15, DARKSHADOW, B'picArea.Line (xx, yy)-(xx + 15, yy + 15), DARKSHADOW, B
    CForm.Line picArea,xx, yy,xx + 14, yy + 14, vbButtonShadow, B
    CForm.Line picArea,xx, yy,xx + 15, yy, vbWhite
    CForm.Line picArea,xx, yy,xx, yy + 15, vbWhite
End Sub

Sub DrawSquares()
    Dim i, j
    'Dim psw, psh
    
    'psw = picArea.ScaleWidth
    'psh = picArea.ScaleHeight
    
    For i = 0 To psw Step 16
        CForm.Line picArea, i + 14, 0,i + 14, psh, vbButtonShadow
    Next
    For j = 0 To psh Step 16
        CForm.Line picArea, 0, j + 14,psw, j + 14, vbButtonShadow
    Next
    For i = 0 To psw Step 16
        CForm.Line picArea, i, 0,i, psh, vbWhite
    Next
    For j = 0 To psh Step 16
        CForm.Line picArea, 0, j,psw, j, vbWhite
    Next
    For i = 0 To psw Step 16
        CForm.Line picArea, i + 15, 0,i + 15, psh, DARKSHADOW
    Next
    For j = 0 To psh Step 16
        CForm.Line picArea, 0, j + 15,psw, j + 15, DARKSHADOW
    Next
    
End Sub

Sub DrawState(X, Y, State)
    dim xx,yy
	xx=x
	yy=y

	If State Then
        
		xx = xx * 16
        yy = yy * 16
		
		if State = 0 then picArea.PaintPicture imgSpecials_0, xx, yy
		if State = 1 then picArea.PaintPicture imgSpecials_1, xx, yy
		if State = 2 then picArea.PaintPicture imgSpecials_2, xx, yy
		if State = 3 then picArea.PaintPicture imgSpecials_3, xx, yy


    Else
        ClearBox xx, yy
        DrawBox xx, yy
    End If
End Sub

Sub PaintCount(Bg, X, Y, Number)
    
'	CCommon.PaintPicture Bg.Picture, imgCounts_1, X, Y'imgCounts(Number)
If Number = 1 then Bg.PaintPicture imgCounts_1.Picture, X,Y
If Number = 2 then Bg.PaintPicture imgCounts_2.Picture, X,Y
If Number = 3 then Bg.PaintPicture imgCounts_3.Picture, X,Y
If Number = 4 then Bg.PaintPicture imgCounts_4.Picture, X,Y
If Number = 5 then Bg.PaintPicture imgCounts_5.Picture, X,Y
If Number = 6 then Bg.PaintPicture imgCounts_6.Picture, X,Y
If Number = 7 then Bg.PaintPicture imgCounts_7.Picture, X,Y
If Number = 8 then Bg.PaintPicture imgCounts_8.Picture, X,Y
If Number = 9 then Bg.PaintPicture imgCounts_9.Picture, X,Y
If Number = 0 then Bg.PaintPicture imgCounts_0.Picture, X,Y
End Sub

Sub ShowCounts(Bg, Numbers)
    Dim StrNum
    
    StrNum = CString.strFormat(Numbers, "0##")
    PaintCount Bg, 0, 0, CCommon.vbIIf(Left(StrNum, 1) = "-", 10, CConvert.toInt(Left(StrNum, 1)))
    PaintCount Bg, NUMBERWIDTH, 0, CConvert.toInt(Left(Right(StrNum, 2), 1))
    PaintCount Bg, NUMBERWIDTH + NUMBERWIDTH, 0, CConvert.toInt(Right(StrNum, 1))
End Sub

Sub ShowMinesLocation(imgOnTop)
    Dim i, j
    For j = 0 To RowNumber - 1
        For i = 0 To ColNumber - 1
            If Mines(i, j) Then
                If State(i, j) <> Flaged Then
                    picArea.PaintPicture imgOnTop, i * 16, j * 16
                End If
            Else
                If State(i, j) = Flaged Then
                    picArea.PaintPicture imgSpecials_0, i * 16, j * 16 'cross
                End If
            End If
        Next
    Next
End Sub

Sub Form_Load()
    frmMain.Left = (Screen.Width / 2) - (frmMain.Width / 2) 'Val(GetSettings("Position", "Left"))
    frmMain.Top = (Screen.Height / 2) - (frmMain.Height / 2)'Val(GetSettings("Position", "Top"))
    LevelNow = 2'CInt(GetSettings("Level", "LastLevel"))
'    mnuLevel(LevelNow).Checked = True

    RowNumber = 16'CInt(GetSettings("Level", "CustomHeight"))
    ColNumber = 16'CInt(GetSettings("Level", "CustomWidth"))
    MineNumber = 10'CInt(GetSettings("Level", "CustomMines"))

    'OnQMark = True'CBool(GetSettings("Others", "QMark"))
'    mnuMarks.Checked = OnQMark

	InitializeGame
	
    
    
    
    frmMain.Icon = img2ndIcon.Picture
	frmMain.Visible = True

    Randomize
End Sub

Sub Form_MouseMove(Button, Shift, X, Y)
    optNewGame.Value = False
End Sub

Sub Form_Paint()
    DoAllBorders
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

Sub Form_Resize()
    tmrCount.Enabled = CCommon.vbIIf(frmMain.WindowState = vbMinimized, False, GameStarted)
End Sub

Sub picArea_DblClick()
    picArea_MouseDown OnButton, 0, Row * 16, Col * 16
    'picArea_MouseUp OnButton, 0, Row * 16, Col * 16
End Sub

Sub picArea_KeyDown(KeyCode, Shift)
    If Shift = vbCtrlMask + vbAltMask Then
        Select Case KeyCode
            Case vbKeyS
                Cheating = True
                picArea.MousePointer = vbCustom
            Case vbKeyW
                WinGame
            Case vbKeyL
                LoseGame
        End Select
    End If
    
    If KeyCode <> vbKeyS Then
        Cheating = False
        picArea.MousePointer = vbDefault
    End If
End Sub

Sub picArea_KeyUp(KeyCode, Shift)
    Cheating = False
    picArea.MousePointer = vbDefault
End Sub

Sub picFlags_MouseMove(Button, Shift, X, Y)
    optNewGame.Value = False
End Sub

Sub picFlags_MouseUp(Button, Shift, X, Y)
    Newing = False
End Sub

Sub picTimeCount_MouseMove(Button, Shift, X, Y)
    optNewGame.Value = False
End Sub

Sub picTimeCount_MouseUp(Button, Shift, X, Y)
    Newing = False
End Sub

Sub mnuExit_Click()
    Unload frmMain
End Sub

Sub mnuLevel_Click(Index)
    If Index = LevelNow And Index <> Custom Then Exit Sub
    mnuLevel(LevelNow).Checked = False
    mnuLevel(Index).Checked = True
    LevelNow = Index
        
    SetUpLevel LevelNow
    InitializeGame
End Sub

Sub mnuNew_Click()
    InitializeGame
End Sub

Sub optNewGame_DblClick()
    optNewGame.Value = True
    mnuNew_Click
End Sub

Sub optNewGame_GotFocus()
    picFocus.SetFocus
End Sub

Sub optNewGame_MouseDown(Button, Shift, X, Y)
    Newing = (Button = vbLeftButton)
End Sub

Sub optNewGame_MouseMove(Button, Shift, X, Y)
    If Button = vbLeftButton Then optNewGame.Value = Newing
End Sub

Sub optNewGame_MouseUp(Button, Shift, X, Y)
    mnuNew_Click
    Newing = False
    optNewGame.Value = Newing
End Sub

Sub picArea_MouseDown(Button, Shift, X, Y)
    OnButton = Button
	If Button = vbLeftButton Then
        OnLeftButton = True
    ElseIf Button = vbRightButton Then
        OnRightButton = True
    End If
    If OnLeftButton And OnRightButton Then
        Button = vbMiddleButton
    End If
    
    If Button <> vbRightButton Then
        optNewGame.Picture = imgFaces_1
    End If
    
    Row = Fix(X / 16)
    Col = Fix(Y / 16)
    Selecting = True

    Select Case Button
        Case vbLeftButton
            If State(Row, Col) = Closed Then ClearBox Row, Col
        Case vbMiddleButton
            DoMiddle Row, Col, ClearThem
    End Select
End Sub

Sub picArea_MouseMove(Button, Shift, X, Y)
    optNewGame.Value = False
    If OnLeftButton And OnRightButton Then Button = vbMiddleButton
    
    Select Case Button
        Case vbLeftButton
            If State(Row, Col) <> Opened Then DrawBox Row, Col
        Case vbMiddleButton
            DoMiddle Row, Col, CloseThem
    End Select
    
    If X < 0 Or Y < 0 Or X > psw - 1 Or Y > psh - 1 Then
        Exit Sub
    End If
    'If Selecting Or True Then
        Row = Fix(X / 16)
        Col = Fix(Y / 16)
        Select Case Button
            Case vbLeftButton
                If State(Row, Col) = Closed Then ClearBox Row, Col
            Case vbMiddleButton
                DoMiddle Row, Col, ClearThem
        End Select
    'End If
    
    If Cheating Then
        frmMain.Caption = CheckDanger(Row, Col)
    Else
        'frmMain.Caption = "Minesweeper"
    End If
End Sub

Sub picArea_MouseUp(Button, Shift, X, Y)
    OnButton = Button
    If OnLeftButton And OnRightButton Then Button = vbMiddleButton
    If OnButton = vbLeftButton Then
        OnLeftButton = False
    ElseIf OnButton = vbRightButton Then
        OnRightButton = False
    End If
    
    If picArea.Enabled Then     '-Game not lost yet
        optNewGame.Picture = imgFaces_0
    End If
    Newing = False
    Selecting = False
    If X < 0 Or Y < 0 Or X > psw - 1 Or Y > psh - 1 Then
        Exit Sub
    End If
    'If Selecting Or True Then       '-Watch here
        Row = Fix(X / 16)
        Col = Fix(Y / 16)
                
        Select Case Button
            Case vbLeftButton
				
                GoLeftUp Row, Col
				
								
            Case vbMiddleButton
                GoMiddleUp Row, Col
            Case vbRightButton
                GoRightUp Row, Col
        End Select
    'End If
End Sub


Sub tmrCount_Timer()
    
    SmallCount = SmallCount + 1
    
    If SmallCount = 20 Then
        SmallCount = 0
        DoCount
    End If
End Sub

Public Sub GoLeftUp(X, Y)
    Dim nDanger

    If Not GameStarted Then
        PutMines X, Y
        GameStarted = True
        SmallCount = 0
        tmrCount.Enabled = GameStarted
        DoCount
    End If


    If State(X, Y) = Closed Or State(X, Y) = Questioned Then
        State(X, Y) = Opened
        DoOpen
        
		
		'stops here
		
		nDanger = CheckDanger(X, Y)
		
        If nDanger = 0 Then
            CheckAround X, Y, X, Y
        ElseIf nDanger = -1 Then
			LoseGame
        End If
		
        ShowDanger X, Y, nDanger
    End If
End Sub

Sub GoMiddleUp(X, Y)
    If State(X, Y) = Opened And CheckDanger(X, Y) = CheckState(X, Y, Flaged) Then
        DoMiddle X, Y, OpenThem
    Else
        DoMiddle X, Y, CloseThem
    End If
End Sub

Sub GoRightUp(X, Y)
    Select Case State(X, Y)
        Case Closed
            State(X, Y) = Flaged
            DoFlag 1
        Case Flaged
            State(X, Y) = CCommon.vbIIf(OnQMark, Questioned, Closed)
            DoFlag -1
        Case Questioned
            State(X, Y) = Closed
    End Select
    If State(X, Y) <> Opened Then
        DrawState X, Y, State(X, Y)
    End If
End Sub

'	CForm.WaitForInput


Sub Command1_Click()
End Sub

Sub Command2_Click()
CForm.Quit
End Sub


}