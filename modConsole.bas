Attribute VB_Name = "modConsole"
Option Explicit

'*' Known Bugs and/or Issues
'*'
'*' 09.27.2000 :  A console application can not be terminated by simply hitting
'*' (Open)        the terminate button on the toolbar.  The application will produce
'*'               a system error due to misread memory instructions.
'*'
'*' End of List

Public Type COORD
        X As Integer
        Y As Integer
End Type

Public Type SMALL_RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
End Type

Public Declare Function GetConsoleScreenBufferInfo Lib "kernel32" _
(ByVal hConsoleOutput As Long, _
lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long

Public Type CONSOLE_SCREEN_BUFFER_INFO
        dwSize As COORD
        dwCursorPosition As COORD
        wAttributes As Integer
        srWindow As SMALL_RECT
        dwMaximumWindowSize As COORD
End Type

Public Declare Function AllocConsole Lib "kernel32" () As Long
'*'
'*' Provided by Windows Kernel.  AllocConsole physically creates the console object
'*' that is used to create the console application in Visual Basic.

Public Declare Function FreeConsole Lib "kernel32" () As Long
'*'
'*' After the console object has been created and used, it will also need to be
'*' destroyed.  FreeConsole is used to terminate the existing console from which
'*' this function is called.

Public Declare Function GetStdHandle Lib "kernel32" _
(ByVal nStdHandle As Long) As Long
'*'
'*' In order to write or read the current console, the handle will establish a link
'*' between the source code and the object so that a means to communicate commands
'*' is established.
'*'
'*' I/O Handlers are used much like hWnd.  There are three that are involved when using
'*' consoles.
'*'
'*' I/O Handlers
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const STD_ERROR_HANDLE = -12&

Public Declare Function ReadConsole Lib "kernel32" Alias _
"ReadConsoleA" (ByVal hConsoleInput As Long, _
ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, _
lpNumberOfCharsRead As Long, lpReserved As Any) As Long
'*'
'*' ReadConsole allows for the console application to retrieve input from the user.
'*'

Public Declare Function SetConsoleMode Lib "kernel32" (ByVal _
hConsoleOutput As Long, dwMode As Long) As Long
'*'
'*' The console has several different modes for different types of console applications.
'*' The following list of constants show which types are available for developement.
Dim Console As Boolean
'*'
'*' Input Modes
Public Const ENABLE_LINE_INPUT = &H2
Public Const ENABLE_ECHO_INPUT = &H4
Public Const ENABLE_MOUSE_INPUT = &H10
Public Const ENABLE_PROCESSED_INPUT = &H1
Public Const ENABLE_WINDOW_INPUT = &H8
'*'
'*' Output Modes
Public Const ENABLE_PROCESSED_OUTPUT = &H1
Public Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

Public Declare Function SetConsoleTextAttribute Lib _
"kernel32" (ByVal hConsoleOutput As Long, ByVal _
wAttributes As Long) As Long
'*'
'*' Text in a console application has modes of output, just like in traditional
'*' console applications.  This includes the uses of colors.
'*'
'*' Attribute Constants
Public Const FOREGROUND_BLUE = &H1         '*' Blue Text
Public Const FOREGROUND_GREEN = &H2        '*' Green Text
Public Const FOREGROUND_RED = &H4          '*' Red Text
Public Const FOREGROUND_INTENSITY = &H8    '*' High Intensity Colorset
Public Const BACKGROUND_BLUE = &H10        '*' Blue Background
Public Const BACKGROUND_GREEN = &H20       '*' Green Background
Public Const BACKGROUND_RED = &H40         '*' Red Background
Public Const BACKGROUND_INTENSITY = &H80   '*' High Intensity Colorset

Public Declare Function SetConsoleTitle Lib "kernel32" Alias _
"SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
'*'
'*' In order to have control over the title bar of the console application, the
'*' SetConsoleTitle function is used.  This allows for a plain text representation
'*' to be set in the titlebar.

Public Declare Function WriteConsole Lib "kernel32" Alias _
"WriteConsoleA" (ByVal hConsoleOutput As Long, _
ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, _
lpNumberOfCharsWritten As Long, lpReserved As Any) As Long

Public Declare Function SetConsoleCursorPosition Lib "kernel32" _
(ByVal hConsoleOutput As Long, ByVal CursorPosition As Long) As Long

'*' Handles
'*'
'*' In order to communicate with the process, we need to be able to establish the
'*' I/O handlers.  Since we have three forms or I/O, we also have three handles.
'*'
'*' Input Handle
Public hConsoleIn As Long
'*'
'*' Output Handle
Public hConsoleOut As Long
'*'
'*' Error Handle
Public hConsoleErr As Long


Public Sub cPrint(strOutput As String)
'*'
'*' Purpose     : Sends the passed string to the output buffer for display to the
'*'               console windows.
'*'
'*' Inputs      : strOutput (String) -  String to be sent to the display
'*'
'*' Returns     : N/A
'*'
'*' Assumes     : Assumes that a console has been created and still exists.  Also
'*'               assumes that the output handler has been set.

    WriteConsole hConsoleOut, strOutput, Len(strOutput), vbNull, vbNull

End Sub

Public Function cInput() As String
'*'
'*' Purpose     : Sends the passed string to the output buffer for display to the
'*'               console windows.
'*'
'*' Inputs      : N/A
'*'
'*' Returns     : ReadFromConsole (String) - Waits for user input to the buffer and
'*'               returns it as a string
'*'
'*' Assumes     : Assumes that a console has been created and still exists.  Also
'*'               assumes that the input handler has been set.

If Console = False Then Exit Function
Dim sUserInput As String * 256

    Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)

    '*' Trim trailing space and vbCrLf
    '*'
    cInput = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)

End Function

Public Sub cLocate(intX As Integer, intY As Integer)

SetConsoleCursorPosition hConsoleOut, intX * &H8000& + intY

End Sub

Public Sub ProcessCommandLine(strCommand As String)
'*'
'*' Purpose     : Parse any command line parameters that were passed upon execution of
'*'               the application
'*'
'*' Inputs      : Pass the value of the Command$ variable to this sub
'*'
'*' Returns     : N/A
'*'
'*' Assumes     : Assumes that public variables have been created to handle the
'*'               arguments passed by the command line.

End Sub

Public Sub cCLS()

Dim csScreenBuffer As CONSOLE_SCREEN_BUFFER_INFO
Dim ConsoleBoundary As SMALL_RECT
Dim XPos As Integer, YPos As Integer

GetConsoleScreenBufferInfo hConsoleOut, csScreenBuffer

ConsoleBoundary = csScreenBuffer.srWindow

For XPos = 0 To ConsoleBoundary.Right
    For YPos = 0 To ConsoleBoundary.Bottom
    
    cLocate XPos, YPos
    cPrint " "
    
    Next YPos
Next XPos

cLocate 0, 0

End Sub

Public Sub cInit()
Console = True
    AllocConsole

    hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
    hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
    hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
    
End Sub

Public Sub cTitle(strConsoleTitle As String)

    SetConsoleTitle strConsoleTitle

End Sub

Public Sub cTerminate()

    FreeConsole

End Sub

Public Sub cColor(Optional ByVal intForeColor As Integer, _
    Optional ByVal intBackColor As Integer)
'*'
'*' Not Yet Available
'*'
'*' Manually set colors using the SetTextAttribute function.
'*' Will be fixed as soon as possible.

If IsNull(intForeColor) Then
    '*' User left blank [OK]
Else
    Select Case intForeColor
    
        Case 0
        
            SetConsoleTextAttribute hConsoleOut, 0
            
        Case 1
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED
            
        Case 2
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_GREEN
            
        Case 3
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_BLUE
            
        Case 4
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or _
                FOREGROUND_GREEN
            
        Case 5
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or _
                FOREGROUND_BLUE
                
        Case 6
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_GREEN Or _
                FOREGROUND_BLUE
            
        
        Case 7
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED _
                Or FOREGROUND_BLUE Or FOREGROUND_GREEN
        
        Case 8
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_INTENSITY
            
        Case 9
                
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_INTENSITY _
                Or FOREGROUND_RED
        
        Case 10
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_INTENSITY _
                Or FOREGROUND_GREEN
        
        Case 11
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_INTENSITY _
                Or FOREGROUND_BLUE
        
        Case 12
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_INTENSITY _
                Or FOREGROUND_RED Or FOREGROUND_GREEN
        
        Case 13
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_INTENSITY _
                Or FOREGROUND_RED Or FOREGROUND_BLUE
        
        Case 14
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_INTENSITY _
                Or FOREGROUND_BLUE Or FOREGROUND_GREEN
        
        Case 15
        
            SetConsoleTextAttribute hConsoleOut, FOREGROUND_INTENSITY _
                Or FOREGROUND_RED Or FOREGROUND_BLUE Or FOREGROUND_GREEN
        
        Case Else
        
            cPrint "Error: " & intForeColor & " is not a valid color."

        End Select
 
 
End If

If IsNull(intBackColor) Then
    '*' User left blank [OK]
Else
    Select Case intBackColor
    
        Case 0
        
            SetConsoleTextAttribute hConsoleOut, 0
            
        Case 1
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_RED
            
        Case 2
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_GREEN
            
        Case 3
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_BLUE
            
        Case 4
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_RED Or _
                BACKGROUND_GREEN
            
        Case 5
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_RED Or _
                BACKGROUND_BLUE
                
        Case 6
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_GREEN Or _
                BACKGROUND_BLUE
            
        
        Case 7
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_RED _
                Or BACKGROUND_BLUE Or BACKGROUND_GREEN
        
        Case 8
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_INTENSITY
            
        Case 9
                
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_INTENSITY _
                Or BACKGROUND_RED
        
        Case 10
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_INTENSITY _
                Or BACKGROUND_GREEN
        
        Case 11
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_INTENSITY _
                Or BACKGROUND_BLUE
        
        Case 12
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_INTENSITY _
                Or BACKGROUND_RED Or BACKGROUND_GREEN
        
        Case 13
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_INTENSITY _
                Or BACKGROUND_RED Or BACKGROUND_BLUE
        
        Case 14
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_INTENSITY _
                Or BACKGROUND_BLUE Or BACKGROUND_GREEN
        
        Case 15
        
            SetConsoleTextAttribute hConsoleOut, BACKGROUND_INTENSITY _
                Or BACKGROUND_RED Or BACKGROUND_BLUE Or BACKGROUND_GREEN
        
        Case Else
        
            cPrint "Error: " & intBackColor & " is not a valid color."

        End Select
 
 
End If

End Sub
