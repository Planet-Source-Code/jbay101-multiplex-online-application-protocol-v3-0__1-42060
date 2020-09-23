VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function xChDir(a)
ChDir z
End Function

Function xChDrive(a)
ChDrive a
End Function

Function xCurDir(a)
xCurDir = CurDir(a)
End Function

Function xDir(Optional PathName, Optional Attributes)
If IsMissing(PathName) = False Then
    If IsMissing(Attributes) = False Then
        xDir = Dir(PathName, Attributes)
    Else
       xDir = Dir(PathName)
    End If
Else
 xDir = Dir()
End If

End Function

Function xEOF(FN)
xEOF = EOF(FN)
End Function

Function xFileAttr(FN)
xFileAttr = FileAttr(FN)
End Function

Function xFileCopy(src, dst)
FileCopy src, dst
End Function

Function xFileDateTime(Path)
xFileDateTime = FileDateTime(Path)
End Function

Function xFileLen(Path)
xFileLen = FileLen(Path)
End Function

Function xFreeFile()
xFreeFile = FreeFile
End Function

Function xGetAttr(pn)
Dim xTMP As VbFileAttribute
Dim x As String

xTMP = GetAttr(pn)
''''''''todo - format output
End Function

Function xKill(pn)
Kill pn
End Function

Function xMkDir(D)
MkDir D
End Function

Function xRmDir(D)
RmDir (D)
End Function

Function xSetAttr(pn, Value)
SetAttr pn, Value
End Function
Function Rename(a, b)
Name a As b
End Function

Private Sub Class_Initialize()
WriteLog "CFileSystem::Initialize()"
End Sub
