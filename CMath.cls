VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Math function for VB

Function xAbs(number)
xAbs (Abs(number))
End Function

Function xAtn(number)
xAtn = Atn(number)
End Function

Function xCos(number)
xCos = Cos(number)
End Function

Function xExp(number)
xExp = Exp(number)
End Function

Function xLog(number)
xLog = Log(number)
End Function

Function xRandomize(number)
Randomize number
End Function

Function xRnd(Optional number)
If IsMissing(number) = False Then xRnd = Rnd(number)
If IsMissing(number) = True Then xRnd = Rnd()

End Function

Function xRound(number, Digits)
xRound = Round(number, Digits)
End Function

Function xSgn(number)
xSgn = Sgn(number)
End Function

Function xSin(number)
xSin = Sin(number)
End Function

Function xSqr(number)
xSqr = Sqr(number)
End Function

Function xTan(number)
xTan = Tan(number)
End Function

Function xTimer()
xTimer = Timer
End Function

Private Sub Class_Initialize()
WriteLog "CMath::Initialize()"
End Sub
