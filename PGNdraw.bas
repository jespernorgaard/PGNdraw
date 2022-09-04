Attribute VB_Name = "Module1"
Option Explicit
Public GlobalEditorPath As String
Public GlobalAcrobatPath As String
Global GlobalPath As String
Global xEditor As String
Global PrintOne As Boolean
Global ResultGiven As Boolean
Global FullFile As String
Global WordpadID As Long
Global DocxID As Long
Global Response As Long
Global ProcessAllFiles As Boolean  ' Use when analyzing all files in a directory
' Variables for handling Missing Points algorithm
Global Subdirs(10000) As String
Global ThisIx As Integer
' Options screen variables
Global xShowFile As Byte
Global GameNumber As Long
Global LineNumber As Long
Global Curminute As Integer
Global OldMinute As Integer
Global StrComment As String
Global StrBlockedBishops As String
Global MessageStack As String
Global FutureLine As String
Global PositionsChecked As Long
Global PositionsNotChecked As Long
' PGN variables
' Each string will contain the squares, unordered, separated by hyphens, for instance
' WhiteKing      'e1'
' WhiteQueens    'd1'
' WhiteRooks     'a1-h1'
' WhiteBishops   'c1-f1'
' WhiteKnights   'b1-g1'
' WhitePawns     'a2-b2-c2-d2-e2-f2-g2-h2'
' WhitePieces    'Ke1 Qd1 Ra1-h1 Bc1-f1 Nb1-g1 a2-b2-c2-d2-e2-f2-g2-h2'
'
' BlackKing      'e8'
' BlackQueens    'd8'
' BlackRooks     'a8-h8'
' BlackBishops   'c8-f8'
' BlackKnights   'b8-g8'
' BlackPawns     'a7-b7-c7-d7-e7-f7-g7-h7'
' BlackPieces    'Ke8 Qd8 Ra8-h8 Bc8-f8 Nb8-g8 a7-b7-c7-d7-e7-f7-g7-h7'

Global WhiteKing As String
Global WhiteQueens As String
Global WhiteRooks As String
Global WhiteBishops As String
Global WhiteKnights As String
Global WhitePawns As String
Global BlackKing As String
Global BlackQueens As String
Global BlackRooks As String
Global BlackBishops As String
Global BlackKnights As String
Global BlackPawns As String
Global WhitePieces As String
Global BlackPieces As String
' WhiteOnMove = 1   BlackOnMove = 2
Global PlayerOnMove As Byte
Global WhitePlayer As String
Global BlackPlayer As String
Global IsCapture As Boolean
Global PieceSelector As String   ' For the move N1f3 "1" is PieceSelector  For the move Ngf3 "g" is PieceSelector
Global PromotionPiece As String  ' For the white move e8=Q "Q" is PromotionPiece
Global Square As String
Global Xdisp As Integer
Global Ydisp As Integer
Global Ep_square As String
Global LastMove As String
Global MoveNum As Integer
Global CaptureInProgress As Boolean
Global CastlingInProgress As Boolean
Global CastlingMove As String
Global CastlingPossible As String  ' From start contains "1111", when short castle black has been lost "1101"
' When long castling black has been lost "1100", when short castling white has been lost "0100"
' When long castling white has been lost "0000"
Global CommentFound As Boolean
Global IncludeGame As Boolean
Global IncludePoint As Long
Global IncludePoint2 As Long
Global GameLines As Long
Global BracketLines As Long
Global ThisPrintStr(10000000) As String
Global FlushMe(10000) As String
Global MaxPrt As Long
Global MaxFlush As Long
Global FileVer As Long
Global ThisMove As String
Global LongMove As String
Global EatMe As String
'Global UndefendedWhitePawns As String
'Global UndefendedBlackPawns As String
'Global FirstGame As Boolean
   ' EatMe will contain all of the move that was just carried out
   ' After making that move, we want to just parse by that move completely
   ' This is done by skipping letters from EatMe
Global CheckAnnounced As Boolean
Global MateAnnounced As Boolean
Global StalemateAnnounced As Boolean
Global DeadPositionFound As Boolean
Global ThisLine As String
Global InputLine As String
Global Nextchar1 As String
Global Nextchar2 As String
Global Nextchar3 As String
Global Nextchar4 As String
Global Nextchar5 As String
Global Nextchar6 As String
Global Nextchar7 As String
Global Nextchar8 As String
Global WhitePositions_FEN(16000) As String
Global BlackPositions_FEN(16000) As String
Global MaxWhitePos As Integer
Global MaxBlackPos As Integer
Global WhitePosExtra(16000) As String
Global BlackPosExtra(16000) As String
' To make things easier, we already paste on the move number in these strings, which are only used to report
' which moves have been repeated, or from which move there was a 50 move or 75 move draw
Global WhitePositionMoves(16000) As String  ' Could be 1.d4 or 84.N1xf3
Global BlackPositionMoves(16000) As String  ' Could be 1...Nf6 or 64...Rxe4
Global MaxWhiteMoves As Integer
Global MaxBlackMoves As Integer
Global CurrentFENstring As String
Global LastFENreported As String
Global Count_50_75_limit As Integer  ' Counting the halfmoves that have passed since last {pawn move or capture}
Global Halfmoves_since_draw_claim As Integer
Global PawnMove As Boolean
Global AllowPiecesInBetween As Boolean
Global DrawDeclared As Boolean
Global Count_ReducedMaterial As Integer  ' Counting the times insufficient forcing material has been detected
Global MS_count_General As Long
Global MS_count_BlockedPos As Long
Global StartTime As Currency
Global EndTime As Currency
Global StopMeNow As Boolean
Global EventNumber As Long
Global FullGameStr As String
Global PlayerNamesStr As String

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
' Variables to handle registry reading and manipulating
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
' Reg Key Data Types...
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Binary number
Public Const REG_DWORD = 4                      ' 32-bit number

' Reg Key ROOT Types...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' Return codes from Registration functions
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_BADDB = 1009&
Public Const ERROR_BADKEY = 1010&
Public Const ERROR_CANTOPEN = 1011&
Public Const ERROR_CANTREAD = 1012&
Public Const ERROR_CANTWRITE = 1013&
Public Const ERROR_OUTOFMEMORY = 14&
Public Const ERROR_INVALID_PARAMETER = 87&
Public Const ERROR_ACCESS_DENIED = 5&

' Security Mask constants
Public Const READ_CONTROL = &H20000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or _
KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or _
KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) _
And (Not SYNCHRONIZE))
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) _
And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or _
KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
' Constants for Process handling
Public Const INFINITE = &HFFFF

' Options
Public Const REG_OPTION_VOLATILE = 0
Public Const REG_OPTION_NON_VOLATILE = 1

Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Variant
   bInheritHandle As Long
End Type
Public glPid     As Long
Public glHandle  As Long
Public colHandle As New Collection

Public Const WM_CLOSE = &H10
Public Const WM_DESTROY = &H2

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim RC As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    RC = RegOpenKey(KeyRoot, KeyName, hKey)                 ' Open Registry Key
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    tmpVal = Space$(256)              ' was String$(32, 0)  ' Allocate Variable Space -
    KeyValSize = Len(tmpVal)          ' was 32              ' Mark Variable Size
    RC = RegQueryValueEx(hKey, SubKeyRef, 0, _
             KeyValType, ByVal tmpVal, KeyValSize)          ' Get/Create Key Value
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ   ' 1                                       ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_BINARY  ' 3                                    ' Binary Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
    Case REG_DWORD  ' 4                                     ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        'KeyVal = Format$("&h" + KeyVal)                    ' Convert Double Word To String
    Case Else
        KeyVal = tmpVal
    End Select
    ' Cut off " delimiters in Windows XP    chr(34) = "
    If (Right(KeyVal, 1) = Chr(34)) And (Left(KeyVal, 1) = Chr(34)) Then
       KeyVal = Mid(KeyVal, 2, Len(KeyVal) - 2)
    End If
    GetKeyValue = True                                      ' Return Success
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    RC = RegCloseKey(hKey)                                  ' Close Registry Key
    Call AllowUserAbort
End Function

Public Sub GetEditorPath()
   Dim Value1 As String
   Dim Res As String
   Dim Pos As Integer
   If GlobalEditorPath <> "Notepad" Then
      Value1 = Dir(GlobalEditorPath)
      If Value1 = "" Then
         GlobalEditorPath = ""
      End If
   End If
   If GlobalEditorPath = "" Then
      Res = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\App Paths\WORDPAD.EXE", "", Value1)
      If Res Then
         GlobalEditorPath = Value1
      End If
   End If
   Pos = InStr(1, GlobalEditorPath, "%ProgramFiles%")
   If Pos <> 0 Then
      Res = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProgramFilesDir", Value1)
      If Res Then
         If Pos > 1 Then
            GlobalEditorPath = Mid(GlobalEditorPath, 1, Pos - 1) + Value1 + Mid(GlobalEditorPath, Pos + 14, Len(GlobalEditorPath) - Pos - 13)
         Else
            GlobalEditorPath = Value1 + Mid(GlobalEditorPath, 15, Len(GlobalEditorPath) - 14)
         End If
      Else
         GlobalEditorPath = ""
      End If
   End If
End Sub

Public Sub GetAcrobatPath()
   Dim Value1 As String
   Dim Res As String
   Dim Pos As Integer
   Dim ThisStr As String
   If Trim(GlobalAcrobatPath) = "" Then
      Res = GetKeyValue(HKEY_CLASSES_ROOT, "acrobat\shell\open\command", "", Value1)
      If Res Then
         For Pos = Len(Value1) To 1 Step -1
            If Mid(Value1, Pos, 1) = "\" Then
               ThisStr = Mid(Value1, 1, Pos - 1)
               Exit For
            End If
         Next Pos
         GlobalAcrobatPath = ThisStr + "\Acrord32.exe"
      End If
   End If
End Sub

Public Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lpData As Long) As Long
Dim lParent    As Long
Dim lThreadId  As Long
Dim lProcessId As Long
'
' This callback function is called by Windows (from the EnumWindows
' API call) for EVERY top-level window that exists.  It populates a
' collection with the handles of all parent windows owned by the
' process that we started.
'
fEnumWindowsCallBack = 1
lThreadId = GetWindowThreadProcessId(hwnd, lProcessId)

If glPid = lProcessId Then
    lParent = GetParent(hwnd)
    If lParent = 0 Then
        colHandle.Add hwnd
    End If
End If
End Function

Public Function fEnumWindows() As Boolean
Dim hwnd As Long
'
' The EnumWindows function enumerates all top-level windows
' by passing the handle of each window, in turn, to an
' application-defined callback function. EnumWindows
' continues until the last top-level window is enumerated or
' the callback function returns FALSE.
'
Call EnumWindows(AddressOf fEnumWindowsCallBack, hwnd)
End Function

Public Function KillProcess(ProcHandle As Long)
Dim i As Long
'
' Enumerate all parent windows for the process.
'
glPid = ProcHandle
Call fEnumWindows
'
' Send a close command to each parent window.
' The app may issue a close confirmation dialog
' depending on how it handles the WM_CLOSE message.
'
For i = 1 To colHandle.Count
    glHandle = colHandle.Item(i)
    Call SendMessage(glHandle, WM_CLOSE, 0&, 0&)
Next i
End Function

Sub ParseAndAnalyze(xPath As String, xFile As String)
   Dim MyNewLine As String
   Dim V As Variant ' String
   Dim sData As String
   Dim OldLine As String
   Dim xStr As String
   Dim FirstTime As Boolean
   On Error GoTo ParseErr
   
   Open xPath + "\" + xFile For Input As #1
   sData = Input(LOF(1), 1)
   Close #1
   If InStr(sData, vbCrLf) Then
      MyNewLine = vbCrLf
   Else
      MyNewLine = vbLf
   End If
   OldLine = ""
   FirstTime = True
   LineNumber = -1  ' a trick to hold back processing the first line for one iteration
   ' While Not EOF(1)
   For Each V In Split(sData, MyNewLine)
      ' was previously:
      'Line Input #1, InputLine
      InputLine = V
      InputLine = Replace(InputLine, vbTab, " ")
      FutureLine = InputLine
      InputLine = OldLine
      ' After this trick, ProccessLine can use FutureLine for looking one line ahead in reading
      ' Therefore the first read record is not processed in the first read loop, but in the second
      If LineNumber >= 0 Then
         Call ProcessLine(FirstTime)
         FirstTime = False
      End If
      OldLine = FutureLine
      If FirstTime Then LineNumber = 0
   Next V
   'Close #1
   ' Manipulate InputLine, FutureLine, etc. to return to the last physical line
   InputLine = FutureLine
   FutureLine = ""
   Call ProcessLine(FirstTime)
   Call FlushOrKeepGame
   Exit Sub
ParseErr:
   xStr = "ParseError! ParseAndAnalyze " + Err.Description + vbCrLf + "Line: (" + Trim(Str(LineNumber)) + ") " + Trim(InputLine)
   Call DisplayHexFileError(xPath, xFile, xStr)
   Call AllowUserAbort
End Sub

Sub ProcessLine(xFirstTime As Boolean)
   Dim Found As Boolean
   Dim NewNum As Long
   Dim x As Long
   Dim y As Integer
   Dim Pos1 As Integer
   Dim Pos2 As Integer
   Dim xDiff As Integer
   Dim ParseLine As String
   Dim NewMove As String
   Dim FoundPlayer As Boolean
   Dim NewSquare As String
   Dim TuborgCounter As Integer ' This counts how many {-characters found inside a single game, reset between each game
   Dim xStr As String
   On Error GoTo ProcessLineErr
   
   If xFirstTime Then
      LineNumber = 0
      IncludePoint = 0
      IncludePoint2 = 0
      FileVer = 0
      GameNumber = 0
      BracketLines = 0
      GameLines = 0
      FoundPlayer = True
      LineNumber = 0
      TuborgCounter = 0
      EventNumber = 0
      ResultGiven = False
   End If
'   If LineNumber >= 153 Then
'      x = x
'   End If
'   If EventNumber >= 82 Then
'      x = x
'   End If
   If InStr(1, InputLine, "[Event ") Then
      EventNumber = EventNumber + 1
      If EventNumber <> (GameNumber + 1) Then
         x = x
      End If
   End If
   
   LineNumber = LineNumber + 1
   EatMe = ""
   If InStr(1, InputLine, "[White ") Then
      Pos1 = InStr(1, InputLine, "[White ") + 7
      Pos2 = InStr(1, InputLine, "]")
      If Pos1 <> 0 And Pos2 <> 0 And Pos2 - 1 > Pos1 Then WhitePlayer = Mid(InputLine, Pos1 + 1, Pos2 - Pos1 - 2)
   End If
   If InStr(1, InputLine, "[Black ") Then
      Pos1 = InStr(1, InputLine, "[Black ") + 7
      Pos2 = InStr(1, InputLine, "]")
      If Pos1 <> 0 And Pos2 <> 0 And Pos2 - 1 > Pos1 Then BlackPlayer = Mid(InputLine, Pos1 + 1, Pos2 - Pos1 - 2)
   End If
   
   If LeftBracketNotInComment(FullGameStr + InputLine) Then
      If IncludeGame Then
         IncludePoint = MaxPrt + 1
         IncludePoint2 = MaxFlush + 1
      End If
      ' For the first game, we have not received bracket lines
      ' If BracketLines = 0 Then
      If InStr(1, InputLine, "[Event ") Then
         Call FlushOrKeepGame
         TuborgCounter = 0
         FullGameStr = ""
         ResultGiven = False
      End If
      MaxPrt = MaxPrt + 1
      MaxFlush = MaxFlush + 1
      ThisPrintStr(MaxPrt) = InputLine
      FlushMe(MaxFlush) = InputLine
      BracketLines = BracketLines + 1
   Else
      FullGameStr = FullGameStr + InputLine
      If (Mid(InputLine, 1, 2) = "1.") Or InStr(1, InputLine, " 1.") <> 0 Then
            MaxPrt = MaxPrt + 2
            MaxFlush = MaxFlush + 2
            TuborgCounter = 0
      End If
      ' From parseline we can perform any reasonable string operation like Mid
      ParseLine = InputLine + Space(80)
      ThisLine = ""
      GameLines = GameLines + 1
      BracketLines = 0
      ' Read in new data for a move
      For x = 1 To Len(InputLine)
         If x <= Len(InputLine) Then
            Nextchar1 = Mid(InputLine, x, 1)
            ThisLine = ThisLine + Nextchar1
         Else
            Nextchar1 = ""
         End If
         If (x + 1) <= Len(InputLine) Then
            Nextchar2 = Mid(InputLine, x + 1, 1)
         Else
            Nextchar2 = " "
         End If
         If (x + 2) <= Len(InputLine) Then
            Nextchar3 = Mid(InputLine, x + 2, 1)
         Else
            Nextchar3 = " "
         End If
         If (x + 3) <= Len(InputLine) Then
            Nextchar4 = Mid(InputLine, x + 3, 1)
         Else
            Nextchar4 = " "
         End If
         If (x + 4) <= Len(InputLine) Then
            Nextchar5 = Mid(InputLine, x + 4, 1)
         Else
            Nextchar5 = " "
         End If
         If (x + 5) <= Len(InputLine) Then
            Nextchar6 = Mid(InputLine, x + 5, 1)
         Else
            Nextchar6 = " "
         End If
         If (x + 6) <= Len(InputLine) Then
            Nextchar7 = Mid(InputLine, x + 6, 1)
         Else
            Nextchar7 = " "
         End If
         ' Just keep this character if PieceSelector is double character like 2692. Qb2c3
         If (x + 7) <= Len(InputLine) Then
            Nextchar8 = Mid(InputLine, x + 7, 1)
         Else
            Nextchar8 = " "
         End If
         If (Not CommentFound) And (Not ResultGiven) Then
            If (Right(ThisLine, 3) = "0-1" And InStr(1, ThisLine, "[Result") = 0) Or _
               (Right(ThisLine, 3) = "1-0" And InStr(1, ThisLine, "[Result") = 0) Or _
               (Right(ThisLine, 7) = "1/2-1/2" And InStr(1, ThisLine, "[Result") = 0) Then
               Call CheckIfLastMoveRep
            End If
            If Right(ThisLine, 3) = "0-1" And InStr(1, ThisLine, "[Result") = 0 Then
               ThisPrintStr(MaxPrt) = ThisPrintStr(MaxPrt) + "0-1"
               FlushMe(MaxFlush) = FlushMe(MaxFlush) + "0-1"
               MaxPrt = MaxPrt + 1
               MaxFlush = MaxFlush + 1
               FoundPlayer = False
               ResultGiven = True
            End If
            If Right(ThisLine, 3) = "1-0" And InStr(1, ThisLine, "[Result") = 0 Then
               ThisPrintStr(MaxPrt) = ThisPrintStr(MaxPrt) + "1-0"
               FlushMe(MaxFlush) = FlushMe(MaxFlush) + "1-0"
               MaxPrt = MaxPrt + 1
               MaxFlush = MaxFlush + 1
               FoundPlayer = False
               ResultGiven = True
            End If
            If Right(ThisLine, 7) = "1/2-1/2" And InStr(1, ThisLine, "[Result") = 0 Then
               ThisPrintStr(MaxPrt) = ThisPrintStr(MaxPrt) + "1/2-1/2"
               FlushMe(MaxFlush) = FlushMe(MaxFlush) + "1/2-1/2"
               MaxPrt = MaxPrt + 1
               MaxFlush = MaxFlush + 1
               FoundPlayer = False
               ResultGiven = True
            End If
         End If
         If Nextchar1 = "{" Then
            CommentFound = True
            TuborgCounter = TuborgCounter + 1
         ElseIf Nextchar1 = "}" Then
            If TuborgCounter <= 1 Then
               CommentFound = False
            End If
            TuborgCounter = TuborgCounter - 1
         End If
         If EatMe <> "" Then
            EatMe = Mid(EatMe, 2, Len(EatMe) - 1)
         ElseIf CommentFound Then
            EatMe = EatMe
         Else
            ThisMove = Nextchar1 + Nextchar2 + Nextchar3 + Nextchar4 + Nextchar5 + Nextchar6 + Nextchar7
            If LastMove = "19...exd4" Then
               x = x
            End If
            CheckAnnounced = False
            StalemateAnnounced = False
            PromotionPiece = ""
            PieceSelector = ""
            If InStr(InputLine, "[") = 0 Then
               MaxPrt = MaxPrt
            End If
            ' Pawn Move
            If IsSquare(Nextchar1, Nextchar2, NewSquare) And IsDelimiter(Nextchar3, Nextchar4, Nextchar5) Then
               Call MakePawnMove(NewSquare)
               Count_50_75_limit = 0
            ' Pawn Capture
            ElseIf InStr(1, "abcdefgh", LCase(Nextchar1)) <> 0 And LCase(Nextchar2) = "x" Then
                If IsSquare(Nextchar3, Nextchar4, NewSquare) And IsDelimiter(Nextchar5, Nextchar6, Nextchar7) Then
                   PieceSelector = Nextchar1
                   CaptureInProgress = True
                   Call MakePawnMove(NewSquare)
                   CaptureInProgress = False
                   Count_50_75_limit = 0
                   PieceSelector = ""
                End If
            End If
            If UCase(Nextchar1) = "K" Then
               If IsSquare(Nextchar2, Nextchar3, NewSquare) And IsDelimiter(Nextchar4, Nextchar5, Nextchar6) Then
                  Call MakeKingMove(NewSquare)
               ElseIf LCase(Nextchar2) = "x" And _
                  IsSquare(Nextchar3, Nextchar4, NewSquare) And IsDelimiter(Nextchar5, Nextchar6, Nextchar7) Then
                     CaptureInProgress = True
                     Call MakeKingMove(NewSquare)
                     CaptureInProgress = False
                     Count_50_75_limit = 0
               End If
            End If
            If UCase(Nextchar1) = "Q" Then
               If IsSquare(Nextchar2, Nextchar3, NewSquare) And IsDelimiter(Nextchar4, Nextchar5, Nextchar6) Then
                  Call MakeQueenMove(NewSquare)
               ElseIf LCase(Nextchar2) = "x" And _
                  IsSquare(Nextchar3, Nextchar4, NewSquare) And IsDelimiter(Nextchar5, Nextchar6, Nextchar7) Then
                     CaptureInProgress = True
                     Call MakeQueenMove(NewSquare)
                     CaptureInProgress = False
                     Count_50_75_limit = 0
               ElseIf IsPieceSelector(Nextchar2 + Nextchar3, PieceSelector) Then
                  If IsSquare(Nextchar3, Nextchar4, NewSquare) And IsDelimiter(Nextchar5, Nextchar6, Nextchar7) Then
                     Call MakeQueenMove(NewSquare)
                  ElseIf LCase(Nextchar3) = "x" Then
                     If IsSquare(Nextchar4, Nextchar5, NewSquare) And IsDelimiter(Nextchar6, Nextchar7, " ") Then
                        CaptureInProgress = True
                        Call MakeQueenMove(NewSquare)
                        CaptureInProgress = False
                        Count_50_75_limit = 0
                     End If
                  End If
               End If
               PieceSelector = ""
            End If
            If UCase(Nextchar1) = "R" Then
               If IsSquare(Nextchar2, Nextchar3, NewSquare) And IsDelimiter(Nextchar4, Nextchar5, Nextchar6) Then
                  If NewSquare = "f2" Then
                     Nextchar1 = Nextchar1
                  End If
                  Call MakeRookMove(NewSquare)
               ElseIf LCase(Nextchar2) = "x" And _
                  IsSquare(Nextchar3, Nextchar4, NewSquare) And IsDelimiter(Nextchar5, Nextchar6, Nextchar7) Then
                     CaptureInProgress = True
                     Call MakeRookMove(NewSquare)
                     CaptureInProgress = False
                     Count_50_75_limit = 0
               ElseIf IsPieceSelector(Nextchar2 + Nextchar3, PieceSelector) Then
                  If IsSquare(Nextchar3, Nextchar4, NewSquare) And IsDelimiter(Nextchar5, Nextchar6, Nextchar7) Then
                     Call MakeRookMove(NewSquare)
                  ElseIf LCase(Nextchar3) = "x" Then
                     If IsSquare(Nextchar4, Nextchar5, NewSquare) And IsDelimiter(Nextchar6, Nextchar7, " ") Then
                        CaptureInProgress = True
                        Call MakeRookMove(NewSquare)
                        CaptureInProgress = False
                        Count_50_75_limit = 0
                   End If
                  End If
               End If
               PieceSelector = ""
            End If
            If Nextchar1 = "B" Then
               If IsSquare(Nextchar2, Nextchar3, NewSquare) And IsDelimiter(Nextchar4, Nextchar5, Nextchar6) Then
                  Call MakeBishopMove(NewSquare)
               ElseIf LCase(Nextchar2) = "x" And _
                  IsSquare(Nextchar3, Nextchar4, NewSquare) And IsDelimiter(Nextchar5, Nextchar6, Nextchar7) Then
                     CaptureInProgress = True
                     Call MakeBishopMove(NewSquare)
                     CaptureInProgress = False
                     Count_50_75_limit = 0
               ElseIf IsPieceSelector(Nextchar2 + Nextchar3, PieceSelector) Then
                  If IsSquare(Nextchar3, Nextchar4, NewSquare) And IsDelimiter(Nextchar5, Nextchar6, Nextchar7) Then
                     Call MakeBishopMove(NewSquare)
                  ElseIf LCase(Nextchar3) = "x" Then
                     If IsSquare(Nextchar4, Nextchar5, NewSquare) And IsDelimiter(Nextchar6, Nextchar7, " ") Then
                        CaptureInProgress = True
                        Call MakeBishopMove(NewSquare)
                        CaptureInProgress = False
                        Count_50_75_limit = 0
                     End If
                  End If
               End If
               PieceSelector = ""
            End If
            If UCase(Nextchar1) = "N" Then
               If IsSquare(Nextchar2, Nextchar3, NewSquare) And IsDelimiter(Nextchar4, Nextchar5, Nextchar6) Then
                  Call MakeKnightMove(NewSquare)
               ElseIf LCase(Nextchar2) = "x" And _
                  IsSquare(Nextchar3, Nextchar4, NewSquare) And IsDelimiter(Nextchar5, Nextchar6, Nextchar7) Then
                     CaptureInProgress = True
                     Call MakeKnightMove(NewSquare)
                     CaptureInProgress = False
                     Count_50_75_limit = 0
               ElseIf IsPieceSelector(Nextchar2 + Nextchar3, PieceSelector) Then
                  If IsSquare(Nextchar3, Nextchar4, NewSquare) And IsDelimiter(Nextchar5, Nextchar6, Nextchar7) Then
                     Call MakeKnightMove(NewSquare)
                  ElseIf LCase(Nextchar3) = "x" Then
                     If IsSquare(Nextchar4, Nextchar5, NewSquare) And IsDelimiter(Nextchar6, Nextchar7, " ") Then
                        CaptureInProgress = True
                        Call MakeKnightMove(NewSquare)
                        CaptureInProgress = False
                        Count_50_75_limit = 0
                     End If
                  End If
               End If
               PieceSelector = ""
            End If
            If Nextchar1 = "0" And Nextchar2 = "-" And Nextchar3 = "0" And Nextchar4 <> "-" Or _
               Nextchar1 = "O" And Nextchar2 = "-" And Nextchar3 = "O" And Nextchar4 <> "-" Then
               CastlingInProgress = True
               Call MakeShortCastling
               CastlingInProgress = False
            End If
            If (Nextchar1 = "0" And Nextchar2 = "-" And Nextchar3 = "0" And _
                Nextchar4 = "-" And Nextchar5 = "0" And IsDelimiter(Nextchar6, Nextchar7, " ")) Or _
               (Nextchar1 = "O" And Nextchar2 = "-" And Nextchar3 = "O" And _
                Nextchar4 = "-" And Nextchar5 = "O" And IsDelimiter(Nextchar6, Nextchar7, " ")) Then
               CastlingInProgress = True
               Call MakeLongCastling
               CastlingInProgress = False
            End If
         End If
      Next x
   End If
   
   'Close #1
   Exit Sub
ProcessLineErr:
   xStr = "ProcessLineError! " + Err.Description + vbCrLf + "Line: (" + Trim(Str(LineNumber)) + ") " + Trim(InputLine)
   MsgBox xStr
   Call AllowUserAbort
End Sub

Sub DisplayHexFileError(xPath As String, xFile As String, xStr As String)
   Dim FileLength As Long
   Dim bytes() As Byte
   Dim FileNum As Integer
   Dim AccumulatedStr As String
   Dim x As Long
   Dim ShownMsg As Boolean
   
   On Error GoTo HexError:
   ShownMsg = False
   FileNum = FreeFile
   Open xPath + "\" + xFile For Binary As FileNum
   FileLength = LOF(FileNum) - 1
   ReDim bytes(FileLength)
   Get FileNum, , bytes
   
   For x = 1 To FileLength
      If bytes(x) > 31 Or bytes(x) = 10 Or bytes(x) = Asc(vbCr) Or bytes(x) = Asc(vbLf) Then
         If Len(AccumulatedStr) < 200 Then
            AccumulatedStr = AccumulatedStr + Chr(bytes(x))
         Else
            AccumulatedStr = Right(AccumulatedStr, 100) + Chr(bytes(x))
         End If
      Else
         MsgBox xStr
         ShownMsg = True
         MsgBox "File " + xPath + "\" + xFile + " has binary characters after " + vbCrLf + _
         AccumulatedStr
         Exit For
      End If
   Next x
   If Not ShownMsg Then
      MsgBox xStr
   End If
   Close FileNum
   Exit Sub
HexError:
   MsgBox "File " + xPath + "\" + xFile + " could not be read, probably it is a binary file"
   Close FileNum
   Call AllowUserAbort
End Sub

Sub FlushOrKeepGame()
   Dim x As Long
   On Error GoTo FlushErr
   GameNumber = GameNumber + 1
   If Not IncludeGame Then
      For x = IncludePoint To MaxPrt
         ThisPrintStr(x) = ""
      Next x
      For x = IncludePoint2 To MaxFlush
         FlushMe(x) = ""
      Next x
      MaxPrt = IncludePoint
      MaxFlush = IncludePoint2
   Else
      PlayerNamesStr = PlayerNamesStr + "{" + WhitePlayer + " - " + BlackPlayer + "}" + vbCrLf
   End If
   If OldMinute = 0 And Curminute = 0 Then
      OldMinute = Minute(Now)
   End If
   Curminute = Minute(Now)
   If Curminute >= OldMinute + 10 Or MaxFlush > 9500 Then
      FileVer = FileVer + 1
      Call FlushPGNfile
      OldMinute = Curminute
   ElseIf OldMinute > Curminute Then
      OldMinute = Curminute + OldMinute - 60
   End If
   IncludeGame = False
   IncludePoint = MaxPrt
   IncludePoint2 = MaxFlush
   Call InitializeNormalBoard
   Call CleanUpParsingVariables
   frmAnalyze.Caption = "PGN draw tool v." + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "   " + _
   Format(FileDateTime(App.Path + "\" + App.EXEName + ".exe"), "YYYY-MM-DD HH:MM:SS") + " : " + Format(GameNumber - 1, "###,###,##0")
   If GameNumber Mod 100 = 0 Then DoEvents
   Exit Sub
FlushErr:
   MsgBox "FlushOrKeepGame " + Err.Description
   Call AllowUserAbort
End Sub

Sub CleanUpParsingVariables()
   Dim x As Long
   Dim y As Long
   Dim z As Long
   On Error GoTo CleanErr
   For x = 1 To MaxWhitePos
      WhitePositions_FEN(x) = ""
      WhitePosExtra(x) = ""
   Next x
   For x = 1 To MaxBlackPos
      BlackPositions_FEN(x) = ""
      BlackPosExtra(x) = ""
   Next x
   MaxWhitePos = 0
   MaxBlackPos = 0
   For x = 1 To MaxWhiteMoves
      WhitePositionMoves(x) = ""
   Next x
   For x = 1 To MaxBlackMoves
      BlackPositionMoves(x) = ""
   Next x
   MaxWhiteMoves = 0
   MaxBlackMoves = 0
   CastlingPossible = "1111" ' all 4 castling methods are possible, white in the first two, black in the last two
   GameLines = 0  ' Count current InputLine as 1
   BracketLines = 0
   Halfmoves_since_draw_claim = 0
   PawnMove = False
   AllowPiecesInBetween = False ' This should be False for normal processing, however when checking if a king is
   ' behind a piece, check if a rook move "x-raying" a piece would be a rook move (same for queen, bishop)
   DrawDeclared = False
   Count_ReducedMaterial = 0
   DeadPositionFound = False
   MateAnnounced = False
   MessageStack = ""   ' Messagestack accumulates new messages to try and avoid repeating them
   Exit Sub
CleanErr:
   MsgBox "CleanUpParsingVariables " + Err.Description
   Call AllowUserAbort
End Sub

Function Min(x1 As Double, x2 As Double) As Double
   If x1 < x2 Then
      Min = x1
   Else
      Min = x2
   End If
End Function

Function Max(x1 As Double, x2 As Double) As Double
   If x1 > x2 Then
      Max = x1
   Else
      Max = x2
   End If
End Function

Function Reordered(xGroup As String) As String
   Dim x As Integer
   Dim y As Integer
   Dim ThisSquare As String
   Dim NewGroup As String
   NewGroup = ""
   For y = 8 To 1 Step -1
      For x = 1 To 8
         ThisSquare = Chr(x + 96) + Chr(y + 48)
         If InStr(1, xGroup, ThisSquare) <> 0 Then
            If NewGroup = "" Then
               NewGroup = ThisSquare
            Else
               NewGroup = NewGroup + "-" + ThisSquare
            End If
         End If
      Next x
   Next y
   Reordered = NewGroup
End Function

Sub CheckIfLastMoveRep()
   Dim x As Long
   Dim y As Long
   Dim ThisFEN As String
   Dim SearchMe As String
   Dim ThisMove As String
   Dim SearchMove As String
   Dim NewMove As String
   Dim CurMax As Integer
   On Error GoTo CheckIfLastErr
   If Count_50_75_limit = 99 And (Not DrawDeclared) And (Not MateAnnounced) And StrComment = "" Then
      If PawnMove Or InStr(1, LastMove, "x") <> 0 Then
         ' Skip this move, it is a pawn move or a capture
      ElseIf PlayerOnMove = 1 Then
         StrComment = "White can write a neutral move and claim a draw since that move will mean 50 moves have passed since the last pawn move or capture " + _
            "with " + WhitePositionMoves(MoveNum - 50) + " FEN=" + WhitePositions_FEN(MoveNum - 50) + _
            " " + WhitePosExtra(MoveNum - 50)
         Call InsertComment(StrComment)
      ElseIf PlayerOnMove = 2 Then
         StrComment = "Black can write a neutral move and claim a draw since that move will mean 50 moves have passed since the last pawn move or capture " + _
            "with " + BlackPositionMoves(MoveNum - 50) + " FEN=" + BlackPositions_FEN(MoveNum - 50) + _
            " " + BlackPosExtra(MoveNum - 50)
         Call InsertComment(StrComment)
      End If
   End If
   ' Note that the move has already turned, so check from the reverse outlook
   ' If 45.Ne3  1-0 we want to check if Black can make a move that will lead to a 3-rep
   If PlayerOnMove = 2 Then
      ThisFEN = WhitePositions_FEN(MaxWhitePos)
      ThisMove = WhitePositionMoves(MaxWhiteMoves)
      CurMax = MaxWhitePos
   ElseIf PlayerOnMove = 1 Then
      ThisFEN = BlackPositions_FEN(MaxBlackPos)
      ThisMove = BlackPositionMoves(MaxBlackMoves)
      CurMax = MaxBlackPos
   End If
   For x = CurMax - 1 To 1 Step -1
      If PlayerOnMove = 1 Then
         SearchMe = WhitePositions_FEN(x)
         SearchMove = WhitePositionMoves(x)
      ElseIf PlayerOnMove = 2 Then
         SearchMe = BlackPositions_FEN(x)
         SearchMove = BlackPositionMoves(x)
      End If
      If InStr(1, SearchMove, "x") Then
         ' When finding a capture, no reason to continue
         Exit Sub
      End If
      For y = x - 1 To 1 Step -1
         If PlayerOnMove = 1 Then
            If SearchMe = WhitePositions_FEN(y) Then
               ' We found a white position that was repeated twice
               NewMove = IsPermutable(SearchMe, ThisFEN, "White")
               If NewMove <> "" And (Not DrawDeclared) And (Not MateAnnounced) And _
                              Halfmoves_since_draw_claim > 5 Then
                  CurrentFENstring = SearchMe + " " + Trim(Count_50_75_limit + 1) + " " + Trim(MoveNum)
                  ThisMove = WhitePositionMoves(y)
                  Call AddCheck(NewMove, ThisMove, SearchMove)
                  StrComment = "White can write the move " + Trim(Str(MaxWhitePos + 1)) + ". " + NewMove + _
                     " on the score sheet and claim a draw," + _
                     " with the moves " + ThisMove + " and " + SearchMove + " the position occurs 3 times " + _
                     " FEN=" + CurrentFENstring
                  Call InsertComment(StrComment)
                  Exit Sub
               End If
            End If
         ElseIf PlayerOnMove = 2 Then
            If SearchMe = BlackPositions_FEN(y) Then
               ' We found a black position that was repeated twice
               NewMove = IsPermutable(SearchMe, ThisFEN, "Black")
               If NewMove <> "" And (Not DrawDeclared) And (Not MateAnnounced) And _
                              Halfmoves_since_draw_claim > 5 Then
                  CurrentFENstring = SearchMe + " " + Trim(Count_50_75_limit + 1) + " " + Trim(MoveNum + 1)
                  ThisMove = BlackPositionMoves(y)
                  Call AddCheck(NewMove, ThisMove, SearchMove)
                  StrComment = "Black can write the move " + Trim(Str(MaxWhitePos)) + "..." + NewMove + _
                     " on the score sheet and claim a draw," + _
                     " with the moves " + ThisMove + " and " + SearchMove + " the position occurs 3 times " + _
                     " FEN=" + CurrentFENstring
                  Call InsertComment(StrComment)
                  Exit Sub
               End If
            End If
         End If
      Next y
   Next x

   x = x
   Exit Sub
CheckIfLastErr:
   MsgBox "CheckIfLastMoveRep " + Err.Description
   Call AllowUserAbort
End Sub
                  
Sub AddCheck(xNewMove As String, xThisMove As String, xSearchMove As String)
   Dim x As Integer
   If InStr(1, xThisMove, "+") <> 0 Then
      If InStr(1, xThisMove, Right(xNewMove, Len(xNewMove) - 3)) Then
         xNewMove = xNewMove + "+"
      ElseIf InStr(1, xSearchMove, Right(xNewMove, Len(xNewMove) - 3)) Then
         xNewMove = xNewMove + "+"
      End If
   End If
End Sub

Function IsPermutable(OldFENstr As String, NewFENstr As String, xColor As String) As String
   ' This function compare the old and the new FEN string and finds a move that can bring new FEN string to become
   ' the old FEN string again. In that case, IsPermutable will contain the new move, else it will be an empty string
   ' There will be a piece vanishing from one square and popping up on another square for this to be successful
   ' If there are more changes than this, abort because we can never find the IsPermutable move
   Dim x As Integer
   Dim y As Integer
   Dim OldStr As String
   Dim NewStr As String
   Dim OldPiece As String
   Dim NewPiece As String
   Dim OldSquare As String
   Dim NewSquare As String
   On Error GoTo IsPermutableErr
   IsPermutable = ""
   If DifferentCastlingRights(OldFENstr, NewFENstr) Then
      Exit Function
   End If
   For x = 8 To 1 Step -1
      OldStr = GetRank(OldFENstr, x)
      NewStr = GetRank(NewFENstr, x)
      For y = 1 To 8
         If Mid(OldStr, y, 1) <> Mid(NewStr, y, 1) Then
            If Mid(OldStr, y, 1) = " " Then
               If OldSquare <> "" Then
                  Exit Function
               End If
               OldPiece = Mid(NewStr, y, 1)
               OldSquare = Chr(y + 96) + Chr(x + 48)
            ElseIf Mid(NewStr, y, 1) = " " Then
               If NewSquare <> "" Then
                  Exit Function
               End If
               NewPiece = Mid(OldStr, y, 1)
               NewSquare = Chr(y + 96) + Chr(x + 48)
            Else
               Exit Function
            End If
         End If
      Next y
   Next x
   If OldPiece <> NewPiece Then
      Exit Function
   End If
   If xColor = "Black" And UCase(NewPiece) = NewPiece Then
      Exit Function
   End If
   If xColor = "White" And LCase(NewPiece) = NewPiece Then
      Exit Function
   End If
   If NewPiece = "k" Or NewPiece = "K" Then
      If IsKingMove(OldSquare, NewSquare) Then
         IsPermutable = "K" + OldSquare + NewSquare
      End If
   ElseIf NewPiece = "q" Or NewPiece = "Q" Then
      If IsQueenMove(OldSquare, NewSquare) Then
         IsPermutable = "Q" + OldSquare + NewSquare
      End If
   ElseIf NewPiece = "r" Or NewPiece = "R" Then
      If IsRookMove(OldSquare, NewSquare) Then
         IsPermutable = "R" + OldSquare + NewSquare
      End If
   ElseIf NewPiece = "b" Or NewPiece = "B" Then
      If IsBishopMove(OldSquare, NewSquare) Then
         IsPermutable = "B" + OldSquare + NewSquare
      End If
   ElseIf NewPiece = "n" Or NewPiece = "N" Then
      If IsKnightMove(OldSquare, NewSquare) Then
         IsPermutable = "N" + OldSquare + NewSquare
      End If
   ElseIf NewPiece = "p" Then
      If IsPawnMove(OldSquare, NewSquare, "Black") Then
         IsPermutable = OldSquare + NewSquare
      End If
   ElseIf NewPiece = "P" Then
      If IsPawnMove(OldSquare, NewSquare, "White") Then
         IsPermutable = OldSquare + NewSquare
      End If
   End If
   Exit Function
IsPermutableErr:
   MsgBox "IsPermutable " + Err.Description
   Call AllowUserAbort
End Function

Function DifferentCastlingRights(Str1 As String, Str2 As String) As Boolean
   Dim ThisStr1 As String
   Dim ThisStr2 As String
   Dim x As Long
   On Error GoTo DiffCastlingErr
   ThisStr1 = GetCastlingRights(Str1)
   ThisStr2 = GetCastlingRights(Str2)
   If ThisStr1 <> ThisStr2 Then
       DifferentCastlingRights = True
   Else
       DifferentCastlingRights = False
   End If
   Exit Function
DiffCastlingErr:
   MsgBox "DifferentCaslingRights " + Err.Description
   Call AllowUserAbort
End Function

Function GetCastlingRights(Str1 As String) As String
   Dim x As Long
   Dim ThisStr As String
   Dim Pos1 As Integer
   Dim Pos2 As Integer
   For x = Len(Str1) To 1 Step -1
      If Mid(Str1, x, 1) = " " Then
         If Pos2 = 0 Then Pos2 = x Else Pos1 = x
      End If
      If Pos1 <> 0 And Pos2 <> 0 Then
         ThisStr = Mid(Str1, Pos1 + 1, Pos2 - Pos1 - 1)
         GetCastlingRights = ThisStr
         Exit Function
      End If
   Next x
End Function

Function GetRank(xFENstr As String, inx As Integer) As String
   Dim Pos(7) As Integer
   Dim InPos As Integer
   Dim ThisStr As String
   Dim NewStr As String
   Dim ThisNum As Integer
   Dim x As Long
   On Error GoTo GetRankErr
   InPos = 0
   If inx = 1 Then
      InPos = InPos
   End If
   For x = 1 To Len(xFENstr)
      If Mid(xFENstr, x, 1) = "/" Then
         InPos = InPos + 1
         Pos(InPos) = x
      End If
   Next x
   If inx = 1 Then
      ThisStr = Mid(xFENstr, Pos(7) + 1, Len(xFENstr) - Pos(7) + 1)
      InPos = InStr(1, ThisStr, " ")
      ThisStr = Mid(ThisStr, 1, InPos - 1)
   ElseIf inx = 8 Then
      ThisStr = Left(xFENstr, Pos(1) - 1)
   Else
      ThisStr = Mid(xFENstr, Pos(8 - inx) + 1, Pos(8 - inx + 1) - Pos(8 - inx) - 1)
   End If
   For x = 1 To Len(ThisStr)
      If IsNumeric(Mid(ThisStr, x, 1)) Then
         ThisNum = Val(Mid(ThisStr, x, 1))
         NewStr = NewStr + Space(ThisNum)
      Else
         NewStr = NewStr + Mid(ThisStr, x, 1)
      End If
   Next x
   GetRank = NewStr
   Exit Function
GetRankErr:
   MsgBox "GetRank " + Err.Description
   Call AllowUserAbort
End Function

Sub PrintReport(xPath As String, xFile As String)
   Dim x As Long
   On Error GoTo FileError
   For x = 1 To MaxPrt
      Print #4, ThisPrintStr(x)
   Next x
   Exit Sub
FileError:
   If PrintOne Then
      MsgBox "Unexpected error in Showing the result file. Maybe WORDPAD.EXE wasn't found." + Chr(13) + _
         "As a last resort you might include the line 'Editor=Notepad' in the TIEBREAK.INI file."
   End If
   Call AllowUserAbort
End Sub

Sub InitializeNormalBoard()
   ' White pieces
   WhiteKing = "e1"
   WhiteQueens = "d1"
   WhiteRooks = "a1-h1"
   WhiteBishops = "c1-f1"
   WhiteKnights = "b1-g1"
   WhitePawns = "a2-b2-c2-d2-e2-f2-g2-h2"
   ' Black pieces
   BlackKing = "e8"
   BlackQueens = "d8"
   BlackRooks = "a8-h8"
   BlackBishops = "c8-f8"
   BlackKnights = "b8-g8"
   BlackPawns = "a7-b7-c7-d7-e7-f7-g7-h7"
   PlayerOnMove = 1
   MoveNum = 1
   CastlingPossible = "1111"   ' Both short and long castling is possible for both players
   ' (1) short White  (2) long White  (3) short Black  (4) long Black  values are 0/1
   Count_50_75_limit = 0
   DeadPositionFound = False
End Sub

Function IsDelimiter(Char1 As String, Char2 As String, Char3 As String) As Boolean
   On Error GoTo IsDelimErr
   IsDelimiter = False
   If Char1 = " " Then
      IsDelimiter = True
   End If
   If Char1 = "?" Then
      IsDelimiter = True
   End If
   If Char1 = "!" Then
      IsDelimiter = True
   End If
   If Char1 = "=" Then
      IsDelimiter = True
   End If
   If Char1 = "-" Then
      IsDelimiter = True
   End If
   If Char1 = "+" Then
      IsDelimiter = True
      CheckAnnounced = True
   End If
   If Char1 = "#" Then
      IsDelimiter = True
      MateAnnounced = True
   End If
   If Char1 = "=" And InStr(1, "QRBN", UCase(Char2)) <> 0 Then
      IsDelimiter = True
      PromotionPiece = UCase(Char2)
      If Char3 = "+" Then
         CheckAnnounced = True
      End If
      If Char3 = "#" Then
         MateAnnounced = True
      End If
   End If
   Exit Function
IsDelimErr:
   MsgBox "IsDelimiter " + Err.Description
   Call AllowUserAbort
End Function

Function IsPieceSelector(xMove As String, xPieceSelector As String) As Boolean
   ' Note that IsPieceSelector will push all characters one position back if a double char
   ' PieceSelector is present
   Dim Char1 As String
   Dim Char2 As String
   On Error GoTo IsPieceErr
   Char1 = LCase(Mid(xMove, 1, 1))
   Char2 = ""
   If Len(xMove) > 1 Then
      Char2 = LCase(Mid(xMove, 2, 1))
   End If
   If InStr(1, "abcdefgh", Char1) <> 0 Or InStr(1, "12345678", Char1) <> 0 Then
      IsPieceSelector = True
      xPieceSelector = Char1
      If InStr(1, "abcdefgh", Char1) <> 0 And InStr(1, "12345678", Char2) <> 0 Then
         ' We have found a double character PieceSelector, now select it and push Nextchar3 and subsquent one position back
         IsPieceSelector = True
         xPieceSelector = Char1 + Char2
         Nextchar3 = Nextchar4
         Nextchar4 = Nextchar5
         Nextchar5 = Nextchar6
         Nextchar6 = Nextchar7
         Nextchar7 = Nextchar8
      End If
   Else
      IsPieceSelector = False
      xPieceSelector = ""
   End If
   Exit Function
IsPieceErr:
   MsgBox "IsPieceSelector " + Err.Description
   Call AllowUserAbort
End Function

Function IsSquare(Char1 As String, Char2 As String, xSquare As String) As Boolean
   If InStr(1, "abcdefgh", LCase(Char1)) <> 0 And InStr(1, "12345678", Char2) <> 0 Then
      IsSquare = True
      xSquare = LCase(Char1) + Char2
   Else
      IsSquare = False
      xSquare = ""
   End If
End Function

Sub ShiftTurn()
   Dim ThisFENstring As String
   Dim Repetitions As Integer
   Dim Rep(10) As String
   Dim RepIx As Integer
   Dim x As Integer
   On Error GoTo ShiftTurnErr
   StrComment = ""  ' New 2019-11-02
   If PlayerOnMove = 1 And Mid(Ep_square, 2, 1) = "6" Then
      ' White is on move and has probably just rejected to capture en passant
      ' so remove the ep_square as you can no longer capture en passant
      Ep_square = ""
   ElseIf PlayerOnMove = 2 And Mid(Ep_square, 2, 1) = "3" Then
      ' Black is on move and has probably just rejected to capture en passant
      ' so remove the ep_square as you can no longer capture en passant
      Ep_square = ""
   End If
   ' Increase MoveNum
   If PlayerOnMove = 2 Then
      PlayerOnMove = 1
      MoveNum = MoveNum + 1
      BlackPosExtra(MoveNum - 1) = Trim(Str(Count_50_75_limit)) + " " + Trim(Str(MoveNum))
   ElseIf PlayerOnMove = 1 Then
      PlayerOnMove = 2
   End If
   WhitePieces = "K" + WhiteKing + IIf(WhiteQueens <> "", " Q" + WhiteQueens, "") + _
   IIf(WhiteRooks <> "", " R" + WhiteRooks, "") + IIf(WhiteBishops <> "", " B" + WhiteBishops, "") + _
   IIf(WhiteKnights <> "", " N" + WhiteKnights, "") + IIf(WhitePawns <> "", " " + WhitePawns, "")
   BlackPieces = "K" + BlackKing + IIf(BlackQueens <> "", " Q" + BlackQueens, "") + _
   IIf(BlackRooks <> "", " R" + BlackRooks, "") + IIf(BlackBishops <> "", " B" + BlackBishops, "") + _
   IIf(BlackKnights <> "", " N" + BlackKnights, "") + IIf(BlackPawns <> "", " " + BlackPawns, "")
   ' After manipulating MoveNum, handle FEN strings and corresponding moves
   ThisFENstring = CreateFENstring(MaxWhitePos + 1)
   If PlayerOnMove = 2 Then
      MaxWhitePos = MaxWhitePos + 1
      WhitePositions_FEN(MaxWhitePos) = ThisFENstring
      MaxWhiteMoves = MaxWhiteMoves + 1
      WhitePositionMoves(MaxWhiteMoves) = LongMove
   ElseIf PlayerOnMove = 1 Then
      MaxBlackPos = MaxBlackPos + 1
      BlackPositions_FEN(MaxBlackPos) = ThisFENstring
      MaxBlackMoves = MaxBlackMoves + 1
      BlackPositionMoves(MaxBlackMoves) = LongMove
   End If
   ' Stop if there is checkmate or stalemate (are there other conditions?)
   If MoveNum >= 62 Then
      x = x
   End If
   Call CheckForDeadPosition
   If StalemateAnnounced Or MateAnnounced Then
      Exit Sub
   End If
   ' Check if there are 3 or 5 repetitions
   Repetitions = 1
   RepIx = 0
   If PlayerOnMove = 2 Then
      For x = 1 To MaxWhitePos - 1
         If WhitePositions_FEN(x) = WhitePositions_FEN(MaxWhitePos) Then
            Repetitions = Repetitions + 1
            If RepIx < 10 Then
               RepIx = RepIx + 1
               Rep(RepIx) = WhitePositionMoves(x)
            End If
         End If
      Next x
      If Repetitions > 1 And RepIx < 10 Then
         RepIx = RepIx + 1
         Rep(RepIx) = WhitePositionMoves(MaxWhiteMoves)
      End If
   ElseIf PlayerOnMove = 1 Then
      For x = 1 To MaxBlackPos - 1
         If BlackPositions_FEN(x) = BlackPositions_FEN(MaxBlackPos) Then
            Repetitions = Repetitions + 1
            If RepIx < 10 Then
               RepIx = RepIx + 1
               Rep(RepIx) = BlackPositionMoves(x)
            End If
         End If
      Next x
      If Repetitions > 1 And RepIx < 10 Then
         RepIx = RepIx + 1
         Rep(RepIx) = BlackPositionMoves(MaxBlackMoves)
      End If
   End If
   Halfmoves_since_draw_claim = Halfmoves_since_draw_claim + 1
   ' Assume we will catch 3-rep and 5-rep first, before seeing further repetitions, so only report them at the
   ' exact number of repetitions
   If Repetitions = 3 And (Not DrawDeclared) And (Not MateAnnounced) Then
      If PlayerOnMove = 1 Then
         StrComment = "White can claim a draw by 3-rep from the positions after " + Rep(1) + _
         ", " + Rep(2) + " and " + Rep(3) + " FEN=" + CurrentFENstring
      ElseIf PlayerOnMove = 2 Then
         StrComment = "Black can claim a draw by 3-rep from the positions after " + Rep(1) + _
         ", " + Rep(2) + " and " + Rep(3) + " FEN=" + CurrentFENstring
      End If
      Call InsertComment(StrComment)
   ElseIf Repetitions = 5 And (Not DrawDeclared) And (Not MateAnnounced) Then
      StrComment = "It is a draw by 5-rep based on the positions after " + Rep(1) + ", " + Rep(2) + _
      ", " + Rep(3) + ", " + Rep(4) + " and " + Rep(5) + " FEN=" + CurrentFENstring
      Call InsertComment(StrComment)
      DrawDeclared = True
   End If
   If MoveNum >= 88 Then
      x = x
   End If
   If Count_50_75_limit = 100 And (Not DrawDeclared) And (Not MateAnnounced) Then
      If PawnMove Or InStr(1, LastMove, "x") <> 0 Then
         ' Skip this move, it is a pawn move or a capture
      ElseIf PlayerOnMove = 1 Then
         StrComment = "White can claim a draw since 50 moves have passed since the last pawn move or capture " + _
            "with " + BlackPositionMoves(MoveNum - 51) + " FEN=" + BlackPositions_FEN(MoveNum - 51) + _
            " " + BlackPosExtra(MoveNum - 51)
         Call InsertComment(StrComment)
      ElseIf PlayerOnMove = 2 Then
         StrComment = "Black can claim a draw since 50 moves have passed since the last pawn move or capture " + _
            "with " + WhitePositionMoves(MoveNum - 50) + " FEN=" + WhitePositions_FEN(MoveNum - 50) + _
            " " + WhitePosExtra(MoveNum - 50)
         Call InsertComment(StrComment)
      End If
   ElseIf Count_50_75_limit = 150 And (Not DrawDeclared) And (Not MateAnnounced) Then
      If PawnMove Or InStr(1, LastMove, "x") <> 0 Then
         ' Skip this move, it is a pawn move or a capture
      ElseIf PlayerOnMove = 1 Then
         StrComment = "The game is a draw since 75 moves have passed since the last pawn move or capture " + _
            "with " + BlackPositionMoves(MoveNum - 76) + " FEN=" + BlackPositions_FEN(MoveNum - 76) + _
            " " + BlackPosExtra(MoveNum - 76)
         Call InsertComment(StrComment)
         DrawDeclared = True
      ElseIf PlayerOnMove = 2 Then
         StrComment = "The game is a draw since 75 moves have passed since the last pawn move or capture " + _
            "with " + WhitePositionMoves(MoveNum - 75) + " FEN=" + WhitePositions_FEN(MoveNum - 75) + _
            " " + WhitePosExtra(MoveNum - 75)
         Call InsertComment(StrComment)
         DrawDeclared = True
      End If
   End If
   ' Convenient breakpoint
   If PlayerOnMove = 2 And MoveNum = 16 Then
      MoveNum = MoveNum
   End If
   If MoveNum = 39 Then
      MoveNum = MoveNum
   End If
   PawnMove = False
   Exit Sub
ShiftTurnErr:
   MsgBox "ShiftTurn " + Err.Description
   Call AllowUserAbort
End Sub

Sub InsertComment(xStr As String)
   Dim x As Integer
   Dim ThisStr As String
   On Error GoTo InsCommentErr
   IncludeGame = True
   If InStr(1, MessageStack, xStr) Then
      Exit Sub
   End If
   MessageStack = MessageStack + xStr
   If (Halfmoves_since_draw_claim < 5) And (InStr(1, xStr, "by 3-rep") <> 0) Then
      ' Was < 3
      Exit Sub
      ' Try to report on 3-rep again with at least 3 moves delay
   End If
   ThisStr = xStr
   ThisPrintStr(MaxPrt) = ThisPrintStr(MaxPrt) + "{"
   FlushMe(MaxFlush) = FlushMe(MaxFlush) + "{"
   For x = 1 To Len(ThisStr)
      If Mid(ThisStr, x, 1) = " " Then
         If Len(ThisPrintStr(MaxPrt)) > 70 Or _
            Len(ThisPrintStr(MaxPrt)) > 20 And Mid(ThisStr, x + 1, 4) = "FEN=" Then
            MaxPrt = MaxPrt + 1
            MaxFlush = MaxFlush + 1
         End If
      End If
      If Mid(ThisStr, x, 1) = " " And Len(ThisPrintStr(MaxPrt)) = 0 Then
         ' Skip blanks in the start of the line
      Else
         ThisPrintStr(MaxPrt) = ThisPrintStr(MaxPrt) + Mid(ThisStr, x, 1)
         FlushMe(MaxFlush) = FlushMe(MaxFlush) + Mid(ThisStr, x, 1)
      End If
   Next x
   ThisPrintStr(MaxPrt) = ThisPrintStr(MaxPrt) + "} "
   FlushMe(MaxFlush) = FlushMe(MaxFlush) + "} "
   Halfmoves_since_draw_claim = 0
   Exit Sub
InsCommentErr:
   MsgBox "InsertComment " + Err.Description
   Call AllowUserAbort
End Sub

Function CreateFENstring(Index As Integer) As String
   ' Create the FEN string and also save the last move in long format, corresponding with
   ' the tables
   Dim RankStr As String
   Dim Result As String
   Dim x As Integer
   Dim y As Integer
   Dim ThisSquare As String
   Dim Actual_EP_square As String
   Dim EP_square_1 As String
   Dim EP_square_2 As String
   On Error GoTo FEN_error
   For x = 8 To 1 Step -1
      RankStr = "8"
      If x = 7 And WhiteKing = "b7" Then
         x = x
      End If
      ' White pieces
      If InStr(1, WhiteKing, Trim(Str(x))) <> 0 Then
         ThisSquare = WhiteKing
         If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
            Call InsertPiece("K", ThisSquare, RankStr)
         End If
      End If
      If InStr(1, WhitePawns, Trim(Str(x))) <> 0 Then
         For y = 1 To (1 + Len(WhitePawns) \ 3)
            ThisSquare = GetSquare(WhitePawns, y)
            If ThisSquare = "" Then Exit For
            If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
               Call InsertPiece("P", ThisSquare, RankStr)
            End If
         Next y
      End If
      If InStr(1, WhiteQueens, Trim(Str(x))) <> 0 Then
         For y = 1 To (1 + Len(WhiteQueens) \ 3)
            ThisSquare = GetSquare(WhiteQueens, y)
            If ThisSquare = "" Then Exit For
            If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
               Call InsertPiece("Q", ThisSquare, RankStr)
            End If
         Next y
      End If
      If InStr(1, WhiteRooks, Trim(Str(x))) <> 0 Then
         For y = 1 To (1 + Len(WhiteRooks) \ 3)
            ThisSquare = GetSquare(WhiteRooks, y)
            If ThisSquare = "" Then Exit For
            If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
               Call InsertPiece("R", ThisSquare, RankStr)
            End If
         Next y
      End If
      If InStr(1, WhiteBishops, Trim(Str(x))) <> 0 Then
         For y = 1 To (1 + Len(WhiteBishops) \ 3)
            ThisSquare = GetSquare(WhiteBishops, y)
            If ThisSquare = "" Then Exit For
            If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
               Call InsertPiece("B", ThisSquare, RankStr)
            End If
         Next y
      End If
      If InStr(1, WhiteKnights, Trim(Str(x))) <> 0 Then
         For y = 1 To (1 + Len(WhiteKnights) \ 3)
            ThisSquare = GetSquare(WhiteKnights, y)
            If ThisSquare = "" Then Exit For
            If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
               Call InsertPiece("N", ThisSquare, RankStr)
            End If
         Next y
      End If
      ' Black pieces
      If InStr(1, BlackKing, Trim(Str(x))) <> 0 Then
         ThisSquare = BlackKing
         If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
            Call InsertPiece("k", ThisSquare, RankStr)
         End If
      End If
      If InStr(1, BlackPawns, Trim(Str(x))) <> 0 Then
         For y = 1 To (1 + Len(BlackPawns) \ 3)
            ThisSquare = GetSquare(BlackPawns, y)
            If ThisSquare = "" Then Exit For
            If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
               Call InsertPiece("p", ThisSquare, RankStr)
            End If
         Next y
      End If
      If InStr(1, BlackQueens, Trim(Str(x))) <> 0 Then
         For y = 1 To (1 + Len(BlackQueens) \ 3)
            ThisSquare = GetSquare(BlackQueens, y)
            If ThisSquare = "" Then Exit For
            If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
               Call InsertPiece("q", ThisSquare, RankStr)
            End If
         Next y
      End If
      If InStr(1, BlackRooks, Trim(Str(x))) <> 0 Then
         For y = 1 To (1 + Len(BlackRooks) \ 3)
            ThisSquare = GetSquare(BlackRooks, y)
            If ThisSquare = "" Then Exit For
            If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
               Call InsertPiece("r", ThisSquare, RankStr)
            End If
         Next y
      End If
      If InStr(1, BlackBishops, Trim(Str(x))) <> 0 Then
         For y = 1 To (1 + Len(BlackBishops) \ 3)
            ThisSquare = GetSquare(BlackBishops, y)
            If ThisSquare = "" Then Exit For
            If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
               Call InsertPiece("b", ThisSquare, RankStr)
            End If
         Next y
      End If
      If InStr(1, BlackKnights, Trim(Str(x))) <> 0 Then
         For y = 1 To (1 + Len(BlackKnights) \ 3)
            ThisSquare = GetSquare(BlackKnights, y)
            If ThisSquare = "" Then Exit For
            If InStr(1, ThisSquare, Trim(Str(x))) <> 0 Then
               Call InsertPiece("n", ThisSquare, RankStr)
            End If
         Next y
      End If
      Result = Result + RankStr + "/"
      ' Remove the last "/"
      If x = 1 Then
         Result = Mid(Result, 1, Len(Result) - 1)
      End If
   Next x
   If PlayerOnMove = 1 Then
      Result = Result + " w "
   ElseIf PlayerOnMove = 2 Then
      Result = Result + " b "
   End If
   Dim FoundCastling As Boolean
   FoundCastling = False
   If Mid(CastlingPossible, 1, 1) = "1" Then
      Result = Result + "K"
      FoundCastling = True
   End If
   If Mid(CastlingPossible, 2, 1) = "1" Then
      Result = Result + "Q"
      FoundCastling = True
   End If
   If Mid(CastlingPossible, 3, 1) = "1" Then
      Result = Result + "k"
      FoundCastling = True
   End If
   If Mid(CastlingPossible, 4, 1) = "1" Then
      Result = Result + "q"
      FoundCastling = True
   End If
   If Not FoundCastling Then
      Result = Result + "-"
   End If
   If PlayerOnMove = 1 Then  ' White to move
      If Ep_square <> "" Then
         Actual_EP_square = ""
         If InStr(1, Ep_square, "6") <> 0 Then
            EP_square_1 = Chr(Asc(Mid(Ep_square, 1, 1)) - 1) + "5"
            EP_square_2 = Chr(Asc(Mid(Ep_square, 1, 1)) + 1) + "5"
            If InStr(1, WhitePawns, EP_square_1) <> 0 Or InStr(1, WhitePawns, EP_square_2) <> 0 Then
               Actual_EP_square = Ep_square
            End If
         End If
      End If
   ElseIf PlayerOnMove = 2 Then  ' Black to move
      If Ep_square <> "" Then
         Actual_EP_square = ""
         If InStr(1, Ep_square, "3") <> 0 Then
            EP_square_1 = Chr(Asc(Mid(Ep_square, 1, 1)) - 1) + "4"
            EP_square_2 = Chr(Asc(Mid(Ep_square, 1, 1)) + 1) + "4"
            If InStr(1, BlackPawns, EP_square_1) <> 0 Or InStr(1, BlackPawns, EP_square_2) <> 0 Then
               Actual_EP_square = Ep_square
            End If
         End If
      End If
   End If
   If Actual_EP_square <> "" Then
      Result = Result + " " + Actual_EP_square
   Else
      Result = Result + " -"
   End If
   CreateFENstring = Result
   ' We separate the FEN string in the first half which is checked for each move, and the second half which contains
   ' move number and number of halfmoves since last pawn move/capture, these should not be included when comparing
   ' positions to see if they are equal
   ' CurrentFENstring contains everything, and is what you would export as the real FEN string
   If PlayerOnMove = 2 Then
      WhitePosExtra(Index) = Trim(Str(Count_50_75_limit)) + " " + Trim(Str(MoveNum))
      CurrentFENstring = Result + " " + WhitePosExtra(Index)
   Else
      BlackPosExtra(Index) = Trim(Str(Count_50_75_limit)) + " " + Trim(Str(MoveNum))
      CurrentFENstring = Result + " " + BlackPosExtra(Index)
   End If
   Exit Function
FEN_error:
   MsgBox "CreateFENstring : " + Err.Description
   Call AllowUserAbort
End Function

Function GetSquare(xStr As String, ix As Integer) As String
   ' In a string with squares like "c2-d4-e2", extract the ix'th square
   ' 1 = "c2"  2 = "d4"  3 = "e2"
   Dim ThisStr As String
   On Error GoTo GetSquareErr
   If xStr = "" Then Exit Function
   ThisStr = xStr + Space(15)
   If ix = 1 Then
      GetSquare = Mid(ThisStr, 1, 2)
   Else
      GetSquare = Mid(ThisStr, 1 + (ix - 1) * 3, 2)
   End If
   If GetSquare = "  " Then
      GetSquare = ""
   End If
   Exit Function
GetSquareErr:
   MsgBox "GetSquare " + Err.Description
   Call AllowUserAbort
End Function

Sub InsertPiece(xPiece As String, xSquare As String, xRank As String)
   ' In the string representing a line in the FEN string (a rank) which starts out as "8"
   ' replace "8" with 8 blanks, put the piece on xSquare, and then revert this to the notation
   ' that FEN wants. If putting a white knight on d4, the "8" will be replaced with "3N4"
   Dim ThisStr As String
   Dim x As Integer
   Dim ThisNum As Integer
   Dim NewStr As String
   On Error GoTo InsPieceErr
   For x = 1 To Len(xRank)
      If IsNumeric(Mid(xRank, x, 1)) Then
         ThisNum = Asc(Mid(xRank, x, 1)) - 48
         ThisStr = ThisStr + Space(ThisNum)
      Else
         ThisStr = ThisStr + Mid(xRank, x, 1)
      End If
   Next x
   ThisNum = Asc(Mid(xSquare, 1, 1)) - 96
   ThisStr = Mid(ThisStr, 1, ThisNum - 1) + xPiece + Mid(ThisStr, ThisNum + 1, Len(ThisStr) - ThisNum)
   ThisNum = 0
   ThisStr = ThisStr + Space(8)
   For x = 1 To 8
      If Mid(ThisStr, x, 1) = " " Then
         ThisNum = ThisNum + 1
      Else
         If ThisNum > 0 Then
            NewStr = NewStr + Chr(ThisNum + 48) + Mid(ThisStr, x, 1)
         Else
            NewStr = NewStr + Mid(ThisStr, x, 1)
         End If
         ThisNum = 0
      End If
   Next x
   If ThisNum > 0 Then
      NewStr = NewStr + Chr(ThisNum + 48)
   End If
   xRank = NewStr
   Exit Sub
InsPieceErr:
   MsgBox "InsertPiece " + Err.Description
   Call AllowUserAbort
End Sub

Sub MakeKingMove(xNewSquare As String)
   On Error GoTo MakeKingErr
   ' Check if it is a king move, and in that case, execute the move
   If PlayerOnMove = 1 Then
      If IsKingMove(WhiteKing, xNewSquare) Then
         Call MovePiece(WhiteKing, xNewSquare, "King", "White")
         CastlingPossible = "00" + Mid(CastlingPossible, 3, 2)
      End If
   Else
      If IsKingMove(BlackKing, xNewSquare) Then
         Call MovePiece(BlackKing, xNewSquare, "King", "Black")
         CastlingPossible = Mid(CastlingPossible, 1, 2) + "00"
      End If
   End If
   Exit Sub
MakeKingErr:
   MsgBox "MakeKingMove " + Err.Description
   Call AllowUserAbort
End Sub

Sub MakeQueenMove(xNewSquare As String)
   ' Check if it is a queen move, and in that case, execute the move
   Dim OldSquare As String
   Dim x As Byte
   On Error GoTo MakeQueenErr
   x = 1
   If PlayerOnMove = 1 Then
      While x < Len(WhiteQueens)
         OldSquare = Mid(WhiteQueens, x, 2)
             If IsQueenMove(OldSquare, xNewSquare) Then
                Call MovePiece(OldSquare, xNewSquare, "Queen", "White")
                Exit Sub
             End If
         x = x + 3
      Wend
   Else
      While x < Len(BlackQueens)
         OldSquare = Mid(BlackQueens, x, 2)
             If IsQueenMove(OldSquare, xNewSquare) Then
                Call MovePiece(OldSquare, xNewSquare, "Queen", "Black")
                Exit Sub
             End If
         x = x + 3
      Wend
   End If
   Exit Sub
MakeQueenErr:
   MsgBox "MakeQueenMove " + Err.Description
   Call AllowUserAbort
End Sub

Sub MakeRookMove(xNewSquare As String)
   ' Check if it is a rook move, and in that case, execute the move
   ' Remove the possibility of castling with that rook if moving a white rook from a1 or h1,
   ' or a black rook from a8 or h8
   Dim OldSquare As String
   Dim x As Byte
   On Error GoTo MakeRookErr
   x = 1
   If PlayerOnMove = 1 Then
      While x < Len(WhiteRooks)
         OldSquare = Mid(WhiteRooks, x, 2)
             If IsRookMove(OldSquare, xNewSquare) Then
                If OldSquare = "a1" Then
                   CastlingPossible = Mid(CastlingPossible, 1, 1) + "0" + Mid(CastlingPossible, 3, 2)
                End If
                If OldSquare = "h1" Then
                   CastlingPossible = "0" + Mid(CastlingPossible, 2, 3)
                End If
                Call MovePiece(OldSquare, xNewSquare, "Rook", "White")
                Exit Sub
             End If
         x = x + 3
      Wend
   Else
      While x < Len(BlackRooks)
         OldSquare = Mid(BlackRooks, x, 2)
             If IsRookMove(OldSquare, xNewSquare) Then
                If OldSquare = "a8" Then
                   CastlingPossible = Mid(CastlingPossible, 1, 3) + "0"
                End If
                If OldSquare = "h8" Then
                   CastlingPossible = Mid(CastlingPossible, 1, 2) + "0" + Mid(CastlingPossible, 4, 1)
                End If
                Call MovePiece(OldSquare, xNewSquare, "Rook", "Black")
                Exit Sub
             End If
         x = x + 3
      Wend
   End If
   Exit Sub
MakeRookErr:
   MsgBox "MakeRookMove " + Err.Description
   Call AllowUserAbort
End Sub

Sub MakeBishopMove(xNewSquare As String)
   ' Check if it is a bishop move, and in that case, execute the move
   Dim OldSquare As String
   Dim x As Byte
   On Error GoTo MakeBishopErr
   x = 1
   If PlayerOnMove = 1 Then
      While x < Len(WhiteBishops)
         OldSquare = Mid(WhiteBishops, x, 2)
            If IsBishopMove(OldSquare, xNewSquare) Then
               Call MovePiece(OldSquare, xNewSquare, "Bishop", "White")
               Exit Sub
            End If
         x = x + 3
      Wend
   Else
      While x < Len(BlackBishops)
         OldSquare = Mid(BlackBishops, x, 2)
            If IsBishopMove(OldSquare, xNewSquare) Then
               Call MovePiece(OldSquare, xNewSquare, "Bishop", "Black")
               Exit Sub
            End If
         x = x + 3
      Wend
   End If
   Exit Sub
MakeBishopErr:
   MsgBox "MakeBishopMove " + Err.Description
   Call AllowUserAbort
End Sub

Sub MakeKnightMove(xNewSquare As String)
   ' Check if it is a knight move, and in that case, execute the move
   Dim OldSquare As String
   Dim x As Byte
   On Error GoTo MakeKnightErr
   x = 1
   If PlayerOnMove = 1 Then
      While x < Len(WhiteKnights)
         OldSquare = Mid(WhiteKnights, x, 2)
             If IsKnightMove(OldSquare, xNewSquare) Then
                Call MovePiece(OldSquare, xNewSquare, "Knight", "White")
                Exit Sub
             End If
         x = x + 3
      Wend
   Else
      While x < Len(BlackKnights)
         OldSquare = Mid(BlackKnights, x, 2)
             If IsKnightMove(OldSquare, xNewSquare) Then
                Call MovePiece(OldSquare, xNewSquare, "Knight", "Black")
                Exit Sub
             End If
         x = x + 3
      Wend
   End If
   Exit Sub
MakeKnightErr:
   MsgBox "MakeKnightMove " + Err.Description
   Call AllowUserAbort
End Sub

Sub MakePawnMove(xNewSquare As String)
   ' Check if it is a pawn move, and in that case, execute the move
   ' Also handle if double square forward, then add an en passant square behind it
   ' The en passant square is later handled with Actual_EP_square if there was an actual enemy pawn ready
   ' to capture, regardless if that pawn is pinned towards it's own king
   Dim OldSquare As String
   Dim x As Byte
   On Error GoTo MakePawnErr
   x = 1
   If PlayerOnMove = 1 Then
      While x < Len(WhitePawns)
         OldSquare = Mid(WhitePawns, x, 2)
             If IsPawnMove(OldSquare, xNewSquare, "White") Then
                If InStr(1, OldSquare, "2") <> 0 And InStr(1, xNewSquare, "4") <> 0 Then
                   Ep_square = Mid(OldSquare, 1, 1) + "3"
                End If
                Call MovePiece(OldSquare, xNewSquare, "Pawn", "White")
                Exit Sub
             End If
         x = x + 3
      Wend
   Else
      While x < Len(BlackPawns)
         OldSquare = Mid(BlackPawns, x, 2)
             If IsPawnMove(OldSquare, xNewSquare, "Black") Then
                If InStr(1, OldSquare, "7") <> 0 And InStr(1, xNewSquare, "5") <> 0 Then
                   Ep_square = Mid(OldSquare, 1, 1) + "6"
                End If
                Call MovePiece(OldSquare, xNewSquare, "Pawn", "Black")
                Exit Sub
             End If
         x = x + 3
      Wend
   End If
   Exit Sub
MakePawnErr:
   MsgBox "MakePawnMove " + Err.Description
   Call AllowUserAbort
End Sub

Sub MakeShortCastling()
   On Error GoTo MakeShortErr
   If PlayerOnMove = 1 Then
      If WhiteKing = "e1" And InStr(1, WhiteRooks, "h1") <> 0 And Mid(CastlingPossible, 1, 1) = "1" Then
         Call MovePiece("e1", "g1", "King", "White")
         Call MovePiece("h1", "f1", "Rook", "White")
         CastlingPossible = "00" + Mid(CastlingPossible, 3, 2)
      End If
   Else
      If BlackKing = "e8" And InStr(1, BlackRooks, "h8") <> 0 And Mid(CastlingPossible, 3, 1) = "1" Then
         Call MovePiece("e8", "g8", "King", "Black")
         Call MovePiece("h8", "f8", "Rook", "Black")
         CastlingPossible = Mid(CastlingPossible, 1, 2) + "00"
      End If
   End If
   Exit Sub
MakeShortErr:
   MsgBox "MakeShortCastling " + Err.Description
   Call AllowUserAbort
End Sub

Sub MakeLongCastling()
   On Error GoTo MakeLongErr
   If PlayerOnMove = 1 Then
      If WhiteKing = "e1" And InStr(1, WhiteRooks, "a1") <> 0 And Mid(CastlingPossible, 2, 1) = "1" Then
         CastlingPossible = "00" + Mid(CastlingPossible, 3, 2)
         Call MovePiece("e1", "c1", "King", "White")
         Call MovePiece("a1", "d1", "Rook", "White")
      End If
   Else
      If BlackKing = "e8" And InStr(1, BlackRooks, "a8") <> 0 And Mid(CastlingPossible, 4, 1) = "1" Then
         CastlingPossible = Mid(CastlingPossible, 1, 2) + "00"
         Call MovePiece("e8", "c8", "King", "Black")
         Call MovePiece("a8", "d8", "Rook", "Black")
      End If
   End If
   Exit Sub
MakeLongErr:
   MsgBox "MakeLongCastling " + Err.Description
   Call AllowUserAbort
End Sub

Sub MovePiece(xOldSquare As String, xNewSquare As String, xPiece As String, xColor As String)
   Dim LocPrt As String
   Dim Pos As Integer
   Dim RestOfString As String
   On Error GoTo MovePieceErr
   
   Count_50_75_limit = Count_50_75_limit + 1
   If CaptureInProgress Then
      LocPrt = "x"
      Count_50_75_limit = 0
   End If
   If xPiece = "Knight" Then
      LastMove = "N" + PieceSelector + LocPrt + xNewSquare
      LongMove = "N" + xOldSquare + LocPrt + xNewSquare
   ElseIf xPiece = "King" And CastlingInProgress And (xNewSquare = "g1" Or xNewSquare = "g8") Then
      LastMove = "O-O"
      LongMove = LastMove
   ElseIf xPiece = "King" And CastlingInProgress And (xNewSquare = "c1" Or xNewSquare = "c8") Then
      LastMove = "O-O-O"
      LongMove = LastMove
   ElseIf xPiece = "Rook" And CastlingInProgress Then
      LastMove = LastMove
   ElseIf xPiece = "Pawn" Then
      PawnMove = True
      Count_50_75_limit = 0
      If CaptureInProgress Then
         LastMove = PieceSelector + LocPrt + xNewSquare
         LongMove = xOldSquare + LocPrt + xNewSquare
      Else
         LastMove = xNewSquare
         LongMove = xOldSquare + LocPrt + xNewSquare
      End If
   Else
      LastMove = Mid(xPiece, 1, 1) + PieceSelector + LocPrt + xNewSquare
      LongMove = Mid(xPiece, 1, 1) + xOldSquare + LocPrt + xNewSquare
   End If
   If xPiece = "Rook" And CastlingInProgress Then
      LastMove = LastMove
   Else
      If xPiece = "Pawn" And PromotionPiece <> "" Then
         LastMove = LastMove + "=" + PromotionPiece
         LongMove = LongMove + "=" + PromotionPiece
      End If
      If PlayerOnMove = 2 Then
         ' ? Is this really sustainable?
         LocPrt = LastMove
         LastMove = Trim(Str(MoveNum)) + "..." + LastMove
         LongMove = Trim(Str(MoveNum)) + "..." + LongMove
      Else
         ' Put extra blank between move nunmber and White's move
         LastMove = Trim(Str(MoveNum)) + "." + LastMove
         LongMove = Trim(Str(MoveNum)) + "." + LongMove
      End If
      If CheckAnnounced Then
         LastMove = LastMove + "+"
         LongMove = LongMove + "+"
         LocPrt = LocPrt + "+"
      End If
      If MateAnnounced Then
         LastMove = LastMove + "#"
         LongMove = LongMove + "#"
         LocPrt = LocPrt + "#"
      End If
      If MaxPrt = 0 Then MaxPrt = 1
      If MaxFlush = 0 Then MaxFlush = 1
      If PlayerOnMove = 1 Then
         ThisPrintStr(MaxPrt) = ThisPrintStr(MaxPrt) + LastMove + " "
         FlushMe(MaxFlush) = FlushMe(MaxFlush) + LastMove + " "
      Else
         ThisPrintStr(MaxPrt) = ThisPrintStr(MaxPrt) + LocPrt + " "
         FlushMe(MaxFlush) = FlushMe(MaxFlush) + LocPrt + " "
      End If
      If Len(ThisPrintStr(MaxPrt)) > 70 And PlayerOnMove = 2 Then
         MaxPrt = MaxPrt + 1
         MaxFlush = MaxFlush + 1
      End If
   End If
   
   If LastMove = "44.g4+" Then
      Pos = Pos
   End If
   
   If xColor = "White" Then
      If CaptureInProgress Then
         If xPiece = "Pawn" And Ep_square = xNewSquare Then
            ' Handle en passant here
            Call RemoveOldPiece(Mid(Ep_square, 1, 1) + "5", "Black")
            Ep_square = ""
            ThisPrintStr(MaxPrt) = ThisPrintStr(MaxPrt) + "{ep} "
            FlushMe(MaxFlush) = FlushMe(MaxFlush) + "{ep} "
         Else
            Call RemoveOldPiece(xNewSquare, "Black")
         End If
      End If
      If xPiece = "Pawn" Then
         If PromotionPiece <> "" Then
            Call PromotePawn(xOldSquare, xNewSquare, xColor)
         Else
            WhitePawns = Replace(WhitePawns, xOldSquare, xNewSquare)
         End If
      ElseIf xPiece = "King" Then
         WhiteKing = Replace(WhiteKing, xOldSquare, xNewSquare)
      ElseIf xPiece = "Queen" Then
         WhiteQueens = Replace(WhiteQueens, xOldSquare, xNewSquare)
      ElseIf xPiece = "Rook" Then
         WhiteRooks = Replace(WhiteRooks, xOldSquare, xNewSquare)
      ElseIf xPiece = "Bishop" Then
         WhiteBishops = Replace(WhiteBishops, xOldSquare, xNewSquare)
      ElseIf xPiece = "Knight" Then
         WhiteKnights = Replace(WhiteKnights, xOldSquare, xNewSquare)
      End If
   ElseIf xColor = "Black" Then
      If CaptureInProgress Then
         If xPiece = "Pawn" And Ep_square = xNewSquare Then
            ' Handle en passant here
            Call RemoveOldPiece(Mid(Ep_square, 1, 1) + "4", "White")
            ThisPrintStr(MaxPrt) = ThisPrintStr(MaxPrt) + "{ep} "
            FlushMe(MaxFlush) = FlushMe(MaxFlush) + "{ep} "
         Else
            Call RemoveOldPiece(xNewSquare, "White")
         End If
      End If
      If xPiece = "Pawn" Then
         If PromotionPiece <> "" Then
            Call PromotePawn(xOldSquare, xNewSquare, xColor)
         Else
            BlackPawns = Replace(BlackPawns, xOldSquare, xNewSquare)
         End If
      ElseIf xPiece = "King" Then
         BlackKing = Replace(BlackKing, xOldSquare, xNewSquare)
      ElseIf xPiece = "Queen" Then
         BlackQueens = Replace(BlackQueens, xOldSquare, xNewSquare)
      ElseIf xPiece = "Rook" Then
         BlackRooks = Replace(BlackRooks, xOldSquare, xNewSquare)
      ElseIf xPiece = "Bishop" Then
         BlackBishops = Replace(BlackBishops, xOldSquare, xNewSquare)
      ElseIf xPiece = "Knight" Then
         BlackKnights = Replace(BlackKnights, xOldSquare, xNewSquare)
      End If
   End If
   ' For ThisMove "d4 Nf6" leave " Nf6"
   ' For ThisMove "N1xf3+" leave " "
   ThisMove = ThisMove + " "
   If CastlingInProgress Then
      Pos = InStr(1, ThisMove, "O-O-O")
      CastlingMove = "O-O-O"
      If Pos = 0 Then
         Pos = InStr(1, ThisMove, "0-0-0")
         CastlingMove = "0-0-0"
      End If
      If Pos = 0 Then
         Pos = InStr(1, ThisMove, "O-O")
         CastlingMove = "O-O"
         If Pos = 0 Then
            Pos = InStr(1, ThisMove, "0-0")
            CastlingMove = "0-0"
         End If
      End If
   Else
      Pos = InStr(1, ThisMove, xNewSquare)
   End If
   Pos = InStr(Pos, ThisMove, " ")
   EatMe = Mid(ThisMove, 1, Pos - 1)
   ' Now check if this was the last part of an input line
   If CastlingInProgress Then
      Pos = InStr(Len(ThisLine), InputLine, CastlingMove)
   Else
      Pos = InStr(Len(ThisLine), InputLine, xNewSquare)
   End If
   'Pos = InStr(Len(ThisLine), InputLine, xNewSquare)
   ' Can anybody explain me the sanity of this code? It seems to me we should always use EatMe
   'RestOfString = Mid(InputLine, Pos, Len(InputLine))
   'If Trim(RestOfString) = Trim(ThisMove) Then
   '   EatMe = ""
   'End If

   ' probably superfluous PieceSelector handled elsewhere
   ' PieceSelector = ""
   If CastlingInProgress And xPiece = "King" Then
   Else
      Call ShiftTurn
   End If
   Exit Sub
MovePieceErr:
   MsgBox "MovePiece " + Err.Description
   Call AllowUserAbort
End Sub

Sub PromotePawn(xOldSquare As String, xNewSquare As String, xColor As String)
   On Error GoTo PromPawnErr
   If xColor = "White" Then
      WhitePawns = Replace(WhitePawns, xOldSquare, "")
      WhitePawns = HyphenTrim(WhitePawns)
      If PromotionPiece = "Q" Then
         WhiteQueens = WhiteQueens + "-" + xNewSquare
         WhiteQueens = HyphenTrim(WhiteQueens)
      ElseIf PromotionPiece = "R" Then
         WhiteRooks = WhiteRooks + "-" + xNewSquare
         WhiteRooks = HyphenTrim(WhiteRooks)
      ElseIf PromotionPiece = "B" Then
         WhiteBishops = WhiteBishops + "-" + xNewSquare
         WhiteBishops = HyphenTrim(WhiteBishops)
      ElseIf PromotionPiece = "N" Then
         WhiteKnights = WhiteKnights + "-" + xNewSquare
         WhiteKnights = HyphenTrim(WhiteKnights)
      End If
   ElseIf xColor = "Black" Then
      BlackPawns = Replace(BlackPawns, xOldSquare, "")
      BlackPawns = HyphenTrim(BlackPawns)
      If PromotionPiece = "Q" Then
         BlackQueens = BlackQueens + "-" + xNewSquare
         BlackQueens = HyphenTrim(BlackQueens)
      ElseIf PromotionPiece = "R" Then
         BlackRooks = BlackRooks + "-" + xNewSquare
         BlackRooks = HyphenTrim(BlackRooks)
      ElseIf PromotionPiece = "B" Then
         BlackBishops = BlackBishops + "-" + xNewSquare
         BlackBishops = HyphenTrim(BlackBishops)
      ElseIf PromotionPiece = "N" Then
         BlackKnights = BlackKnights + "-" + xNewSquare
         BlackKnights = HyphenTrim(BlackKnights)
      End If
   End If
   Exit Sub
PromPawnErr:
   MsgBox "PromotePawn " + Err.Description
   Call AllowUserAbort
End Sub

Function IsKingMove(xOldSquare As String, xNewSquare As String) As Boolean
   Call FindDisplacement(xOldSquare, xNewSquare, Xdisp, Ydisp)
   IsKingMove = False
   If Abs(Xdisp) < 2 And Abs(Ydisp) < 2 Then
      IsKingMove = True
   End If
   If xOldSquare = xNewSquare Then IsKingMove = False
End Function

Function IsQueenMove(xOldSquare As String, xNewSquare As String) As Boolean
   On Error GoTo IsQueenErr
   Call FindDisplacement(xOldSquare, xNewSquare, Xdisp, Ydisp)
   IsQueenMove = False
   If (Xdisp = 0) Or (Ydisp = 0) Then
      ' Rook Move
      IsQueenMove = True
      If WrongPiece(xOldSquare, PieceSelector) Then
         IsQueenMove = False
      End If
   ElseIf Abs(Xdisp) = Abs(Ydisp) Then
      ' Bishop Move
      IsQueenMove = True
      If WrongPiece(xOldSquare, PieceSelector) Then
         IsQueenMove = False
      End If
   End If
   If xOldSquare = xNewSquare Then IsQueenMove = False
   If IsQueenMove And Not AllowPiecesInBetween Then
      If PiecesInBetween(xOldSquare, xNewSquare) Then
         IsQueenMove = False
      End If
   End If
   If IsQueenMove Then
      If IsPinned(xOldSquare, xNewSquare) Then
         IsQueenMove = False
      End If
   End If
   Exit Function
IsQueenErr:
   MsgBox "IsQueenMove " + Err.Description
   Call AllowUserAbort
End Function

Function IsRookMove(xOldSquare As String, xNewSquare As String)
   On Error GoTo IsRookErr
   Call FindDisplacement(xOldSquare, xNewSquare, Xdisp, Ydisp)
   IsRookMove = False
   If (Xdisp = 0) Or (Ydisp = 0) Then
      ' Rook Move
      IsRookMove = True
      If WrongPiece(xOldSquare, PieceSelector) Then
         IsRookMove = False
      End If
   End If
   If xOldSquare = xNewSquare Then IsRookMove = False
   If IsRookMove And Not AllowPiecesInBetween Then
      If PiecesInBetween(xOldSquare, xNewSquare) Then
         IsRookMove = False
      End If
   End If
   If IsRookMove Then
      If IsPinned(xOldSquare, xNewSquare) Then
         IsRookMove = False
      End If
   End If
   Exit Function
IsRookErr:
   MsgBox "IsRookMove " + Err.Description
   Call AllowUserAbort
End Function

Function IsBishopMove(xOldSquare As String, xNewSquare As String)
   On Error GoTo IsBishopErr
   Call FindDisplacement(xOldSquare, xNewSquare, Xdisp, Ydisp)
   IsBishopMove = False
   If Abs(Xdisp) = Abs(Ydisp) Then
      ' Bishop Move
      IsBishopMove = True
      If WrongPiece(xOldSquare, PieceSelector) Then
         IsBishopMove = False
      End If
   End If
   If xOldSquare = xNewSquare Then IsBishopMove = False
   If IsBishopMove And Not AllowPiecesInBetween Then
      If PiecesInBetween(xOldSquare, xNewSquare) Then
         IsBishopMove = False
      End If
   End If
   If IsBishopMove Then
      If IsPinned(xOldSquare, xNewSquare) Then
         IsBishopMove = False
      End If
   End If
   Exit Function
IsBishopErr:
   MsgBox "IsBishopMove " + Err.Description
   Call AllowUserAbort
End Function

Function IsKnightMove(xOldSquare As String, xNewSquare As String)
   On Error GoTo IsKnightErr
   Call FindDisplacement(xOldSquare, xNewSquare, Xdisp, Ydisp)
   IsKnightMove = False
   If (Abs(Xdisp) = 2 And Abs(Ydisp) = 1) Or _
      (Abs(Xdisp) = 1 And Abs(Ydisp) = 2) Then
      ' Knight Move
      IsKnightMove = True
      If WrongPiece(xOldSquare, PieceSelector) Then
         IsKnightMove = False
      End If
   End If
   If xOldSquare = xNewSquare Then IsKnightMove = False
   If IsKnightMove Then
      If IsPinned(xOldSquare, xNewSquare) Then
         IsKnightMove = False
      End If
   End If
   Exit Function
IsKnightErr:
   MsgBox "IsKnightMove " + Err.Description
   Call AllowUserAbort
End Function

Function IsPawnMove(xOldSquare As String, xNewSquare As String, xColor As String)
   On Error GoTo IsPawnerr
   Call FindDisplacement(xOldSquare, xNewSquare, Xdisp, Ydisp)
   IsPawnMove = False
   If xColor = "White" Then
      If (Abs(Xdisp) = 0 And Ydisp = 1) Or _
         (Abs(Xdisp) = 0 And Ydisp = 2 And Mid(xOldSquare, 2, 1) = "2") Then
         ' Normal Pawn Move
         IsPawnMove = True
      End If
      If (Abs(Xdisp) = 1 And Ydisp = 1) Then
         ' Normal Pawn Capture Move
         If CaptureInProgress Then
            IsPawnMove = True
         End If
      End If
   Else
      If (Abs(Xdisp) = 0 And Ydisp = -1) Or _
         (Abs(Xdisp) = 0 And Ydisp = -2 And Mid(xOldSquare, 2, 1) = "7") Then
         ' Normal Pawn Move
         IsPawnMove = True
      End If
      If (Abs(Xdisp) = 1 And Ydisp = -1) Then
         ' Normal Pawn Capture Move
         If CaptureInProgress Then
            IsPawnMove = True
         End If
      End If
   End If
   If WrongPiece(xOldSquare, PieceSelector) Then
      IsPawnMove = False
   End If
   If xOldSquare = xNewSquare Then IsPawnMove = False
   If IsPawnMove Then
      If PiecesInBetween(xOldSquare, xNewSquare) Then
         IsPawnMove = False
      End If
   End If
   If IsPawnMove Then
      If IsPinned(xOldSquare, xNewSquare) Then
         IsPawnMove = False
      End If
   End If
   Exit Function
IsPawnerr:
   MsgBox "IsPawnMove " + Err.Description
   Call AllowUserAbort
End Function

Function IsPinned(xOldSquare As String, xNewSquare As String) As Boolean
   Dim ThisSquare As String
   Dim x As Integer
   Dim CheckBishopmoves As Boolean
   Dim CheckRookmoves As Boolean
   Dim KingSquare As String
   On Error GoTo IsPinnedErr
   ' Note that we don't care which piece is on the square at this point
   If PlayerOnMove = 1 Then
      KingSquare = WhiteKing
      If JumpsLikeABishop(xOldSquare, WhiteKing) Then
         ' Check pins from black queens or black bishops
         CheckBishopmoves = True
      ElseIf JumpsLikeARook(xOldSquare, WhiteKing) Then
         ' Check pins from black queens or black rooks
         CheckRookmoves = True
      End If
      For x = 1 To (1 + Len(BlackQueens) \ 3)
         ThisSquare = GetSquare(BlackQueens, x)
         If ThisSquare = "" Then Exit For
         If CheckBishopmoves Then
            If JumpsLikeABishop(ThisSquare, xOldSquare) And JumpsLikeABishop(ThisSquare, WhiteKing) Then
               If Not PiecesInBetween(ThisSquare, xOldSquare) And Not PiecesInBetween(xOldSquare, WhiteKing) Then
                  IsPinned = True
               End If
            End If
         ElseIf CheckRookmoves Then
            ' We need to add checking if the bishop moves are in the same direction
            ' perhaps with SameDisplacements
            If JumpsLikeARook(ThisSquare, xOldSquare) And JumpsLikeARook(ThisSquare, WhiteKing) Then
               If Not PiecesInBetween(ThisSquare, xOldSquare) And Not PiecesInBetween(xOldSquare, WhiteKing) Then
                  IsPinned = True
               End If
            End If
         End If
      Next x
      If CheckBishopmoves Then
         For x = 1 To (1 + Len(BlackBishops) \ 3)
            ThisSquare = GetSquare(BlackBishops, x)
            If ThisSquare = "" Then Exit For
            If JumpsLikeABishop(ThisSquare, xOldSquare) And JumpsLikeABishop(ThisSquare, WhiteKing) Then
               If Not PiecesInBetween(ThisSquare, xOldSquare) And Not PiecesInBetween(xOldSquare, WhiteKing) Then
                  IsPinned = True
               End If
            End If
         Next x
      End If
      If CheckRookmoves Then
         For x = 1 To (1 + Len(BlackRooks) \ 3)
            ThisSquare = GetSquare(BlackRooks, x)
            If ThisSquare = "" Then Exit For
            If JumpsLikeARook(ThisSquare, xOldSquare) And JumpsLikeARook(ThisSquare, WhiteKing) Then
               If Not PiecesInBetween(ThisSquare, xOldSquare) And Not PiecesInBetween(xOldSquare, WhiteKing) Then
                  IsPinned = True
               End If
            End If
         Next x
      End If
   ElseIf PlayerOnMove = 2 Then
      KingSquare = BlackKing
      If JumpsLikeABishop(xOldSquare, BlackKing) Then
         ' Check pins from black queens or black bishops
         CheckBishopmoves = True
      ElseIf JumpsLikeARook(xOldSquare, BlackKing) Then
         ' Check pins from black queens or black rooks
         CheckRookmoves = True
      End If
      For x = 1 To (1 + Len(WhiteQueens) \ 3)
         ThisSquare = GetSquare(WhiteQueens, x)
         If ThisSquare = "" Then Exit For
         If CheckBishopmoves Then
            If JumpsLikeABishop(ThisSquare, xOldSquare) And JumpsLikeABishop(ThisSquare, BlackKing) Then
               If Not PiecesInBetween(ThisSquare, xOldSquare) And Not PiecesInBetween(xOldSquare, BlackKing) Then
                  IsPinned = True
               End If
            End If
         ElseIf CheckRookmoves Then
            ' We need to add checking if the bishop moves are in the same direction
            ' perhaps with SameDisplacements
            If JumpsLikeARook(ThisSquare, xOldSquare) And JumpsLikeARook(ThisSquare, BlackKing) Then
               If Not PiecesInBetween(ThisSquare, xOldSquare) And Not PiecesInBetween(xOldSquare, BlackKing) Then
                  IsPinned = True
               End If
            End If
         End If
      Next x
      If CheckBishopmoves Then
         For x = 1 To (1 + Len(WhiteBishops) \ 3)
            ThisSquare = GetSquare(WhiteBishops, x)
            If ThisSquare = "" Then Exit For
            If JumpsLikeABishop(ThisSquare, xOldSquare) And JumpsLikeABishop(ThisSquare, BlackKing) Then
               If Not PiecesInBetween(ThisSquare, xOldSquare) And Not PiecesInBetween(xOldSquare, BlackKing) Then
                  IsPinned = True
               End If
            End If
         Next x
      End If
      If CheckRookmoves Then
         For x = 1 To (1 + Len(WhiteRooks) \ 3)
            ThisSquare = GetSquare(WhiteRooks, x)
            If ThisSquare = "" Then Exit For
            If JumpsLikeARook(ThisSquare, xOldSquare) And JumpsLikeARook(ThisSquare, BlackKing) Then
               If Not PiecesInBetween(ThisSquare, xOldSquare) And Not PiecesInBetween(xOldSquare, BlackKing) Then
                  IsPinned = True
               End If
            End If
         Next x
      End If
   End If
   If IsPinned Then
      If CheckRookmoves Then
         If JumpsLikeARook(KingSquare, xNewSquare) And JumpsLikeARook(xOldSquare, xNewSquare) Then
            IsPinned = False
         End If
      ElseIf CheckBishopmoves Then
         If JumpsLikeABishop(KingSquare, xNewSquare) And JumpsLikeABishop(xOldSquare, xNewSquare) Then
            IsPinned = False
         End If
      End If
   End If
   Exit Function
IsPinnedErr:
   MsgBox "IsPinned " + Err.Description
   Call AllowUserAbort
End Function

Function JumpsLikeARook(xOldSquare As String, xNewSquare As String)
   Call FindDisplacement(xOldSquare, xNewSquare, Xdisp, Ydisp)
   JumpsLikeARook = False
   If (Xdisp = 0) Or (Ydisp = 0) Then
      ' Rook Move
      JumpsLikeARook = True
   End If
   If xOldSquare = xNewSquare Then JumpsLikeARook = False
End Function

Function JumpsLikeABishop(xOldSquare As String, xNewSquare As String)
   Call FindDisplacement(xOldSquare, xNewSquare, Xdisp, Ydisp)
   JumpsLikeABishop = False
   If Abs(Xdisp) = Abs(Ydisp) Then
      ' Bishop Move
      JumpsLikeABishop = True
   End If
   If xOldSquare = xNewSquare Then JumpsLikeABishop = False
End Function

Sub FindDisplacement(xOldSquare As String, xNewSquare As String, xXdisp As Integer, xYdisp As Integer)
   On Error GoTo FindDispErr
   xXdisp = Asc(Mid(xNewSquare, 1, 1)) - Asc(Mid(xOldSquare, 1, 1))
   xYdisp = Val(Mid(xNewSquare, 2, 1)) - Val(Mid(xOldSquare, 2, 1))
   Exit Sub
FindDispErr:
   MsgBox "FindDisplacement " + Err.Description
   Call AllowUserAbort
End Sub

Function SameDisplacement(xPiece As String, xSquare1 As String, xSquare2 As String)
   ' Normally xPiece would be a queen or bishop or rook to find out if going to xSquare1 (blocked piece) and
   ' xSquare2 (king) is a pin in the same direction (exception would be wBf3,bRb7,bKd1 where displacements are different)
   ' Displacements would be the same if wBf3,bRb7,bKa8 where we have a real pin with the bishop on the rook
   Dim Xdisp As Integer
   Dim Ydisp As Integer
   Dim Xstep1 As Integer
   Dim Ystep1 As Integer
   Dim Xstep2 As Integer
   Dim Ystep2 As Integer
   On Error GoTo SameDispErr
   SameDisplacement = False
   Call FindDisplacement(xPiece, xSquare1, Xdisp, Ydisp)
   If Xdisp <> 0 Then
      Xstep1 = Xdisp / Abs(Xdisp)
   Else
      Xstep1 = 0
   End If
   If Ydisp <> 0 Then
      Ystep1 = Ydisp / Abs(Ydisp)
   Else
      Ystep1 = 0
   End If
   Call FindDisplacement(xPiece, xSquare2, Xdisp, Ydisp)
   If Xdisp <> 0 Then
      Xstep2 = Xdisp / Abs(Xdisp)
   Else
      Xstep2 = 0
   End If
   If Ydisp <> 0 Then
      Ystep2 = Ydisp / Abs(Ydisp)
   Else
      Ystep2 = 0
   End If
   If Xstep1 = Xstep2 And Ystep1 = Ystep2 Then
      SameDisplacement = True
   End If
   Exit Function
SameDispErr:
   MsgBox "SameDisplacement " + Err.Description
   Call AllowUserAbort
End Function

Function PiecesInBetween(xOldSquare As String, xNewSquare As String) As Boolean
   Dim Xdisp As Integer
   Dim Ydisp As Integer
   Dim Xstep As Integer
   Dim Ystep As Integer
   Dim x As Integer
   Dim ThisSquare As String
   On Error GoTo InBetweenErr
   PiecesInBetween = False
   ' Per definition, there are no pieces on an undefined line, so abort if malformed call to function
   If Not (JumpsLikeABishop(xOldSquare, xNewSquare) Or JumpsLikeARook(xOldSquare, xNewSquare)) Then
      Exit Function
   End If
   Call FindDisplacement(xOldSquare, xNewSquare, Xdisp, Ydisp)
   If Xdisp <> 0 Then
      Xstep = Xdisp / Abs(Xdisp)
   Else
      Xstep = 0
   End If
   If Ydisp <> 0 Then
      Ystep = Ydisp / Abs(Ydisp)
   Else
      Ystep = 0
   End If
   ThisSquare = xOldSquare
   For x = 1 To 8
      ThisSquare = NextFile(Left(ThisSquare, 1), Xstep) + NextRank(Right(ThisSquare, 1), Ystep)
      If ThisSquare = xNewSquare Then
         Exit For
      End If
      If OccupiedSquare(ThisSquare) Then
         PiecesInBetween = True
         Exit For
      End If
   Next x
   Exit Function
InBetweenErr:
   MsgBox "PiecesInBetween " + Err.Description
   Call AllowUserAbort
End Function

Function NextFile(xFile As String, xJump As Integer) As String
   On Error GoTo NextFileErr
   ' If "a" return "b"
   NextFile = Chr(Asc(xFile) + xJump)
   If NextFile < "a" Or NextFile > "h" Then
      NextFile = ""
   End If
   Exit Function
NextFileErr:
   MsgBox "NextFile " + Err.Description
   Call AllowUserAbort
End Function

Function NextRank(xRank As String, xJump As Integer) As String
   On Error GoTo NextRankErr
   ' If "1" return "2"
   NextRank = Chr(Asc(xRank) + xJump)
   If NextRank < "1" Or NextRank > "8" Then
      NextRank = ""
   End If
   Exit Function
NextRankErr:
   MsgBox "NextRank " + Err.Description
   Call AllowUserAbort
End Function

Function OccupiedSquare(xThisSquare, Optional ForWhichColor As String) As Boolean
   ' Check if there is a piece on the square xThisSquare
   ' Check the white pieces
   OccupiedSquare = False
   If ForWhichColor = "" Or ForWhichColor = "White" Then
      If InStr(1, WhiteKing, xThisSquare) <> 0 Then OccupiedSquare = True
      If InStr(1, WhiteQueens, xThisSquare) <> 0 Then OccupiedSquare = True
      If InStr(1, WhiteRooks, xThisSquare) <> 0 Then OccupiedSquare = True
      If InStr(1, WhiteBishops, xThisSquare) <> 0 Then OccupiedSquare = True
      If InStr(1, WhiteKnights, xThisSquare) <> 0 Then OccupiedSquare = True
      If InStr(1, WhitePawns, xThisSquare) <> 0 Then OccupiedSquare = True
   End If
   ' Check the black pieces
   If ForWhichColor = "" Or ForWhichColor = "Black" Then
      If InStr(1, BlackKing, xThisSquare) <> 0 Then OccupiedSquare = True
      If InStr(1, BlackQueens, xThisSquare) <> 0 Then OccupiedSquare = True
      If InStr(1, BlackRooks, xThisSquare) <> 0 Then OccupiedSquare = True
      If InStr(1, BlackBishops, xThisSquare) <> 0 Then OccupiedSquare = True
      If InStr(1, BlackKnights, xThisSquare) <> 0 Then OccupiedSquare = True
      If InStr(1, BlackPawns, xThisSquare) <> 0 Then OccupiedSquare = True
   End If
End Function

Function BlockedByOwnPieces(xColor As String, xThisSquare As String) As Boolean
   Dim CheckingSquares As String
   Dim CandidateSquare As String
   Dim x As Integer
   On Error GoTo BlockedByOwnErr
   ' Check if the piece on the square is blocked by its own pieces
   ' For a queen, that would be all immediate squares around it are occupied by an own piece (see st13.pgn)
   ' For a rook, that would be all 1-square rook moves occupied
   ' For a bishop that would be all 1-square bishop moves occupied
   ' For a knight
   '    From h1(f2,g3) [2]
   '    From g1(e2,f3,h3) [3]
   '    From f1(d2,e3,g3,h2) [4]
   '    From d2(b1,b3,c4,e4,g3,g1) [6]
   '    From d3(c1,b2,b4,c5,e5,f4,f2,e1) [8]
  
   'If PlayerOnMove = 1 Then
   If xColor = "White" Then
      ' Diagonal moves like bishop or queen
      If InStr(1, WhiteQueens, xThisSquare) Or InStr(1, WhiteBishops, xThisSquare) Then
         If Left(xThisSquare, 1) < "h" And Right(xThisSquare, 1) < "8" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 1) + NextRank(Right(xThisSquare, 1), 1) + "-"
         End If
         If Left(xThisSquare, 1) < "h" And Right(xThisSquare, 1) > "1" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 1) + NextRank(Right(xThisSquare, 1), -1) + "-"
         End If
         If Left(xThisSquare, 1) > "a" And Right(xThisSquare, 1) > "1" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -1) + NextRank(Right(xThisSquare, 1), -1) + "-"
         End If
         If Left(xThisSquare, 1) > "a" And Right(xThisSquare, 1) < "8" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -1) + NextRank(Right(xThisSquare, 1), 1) + "-"
         End If
      End If
      ' Vertical moves like rook or queen
      If InStr(1, WhiteQueens, xThisSquare) Or InStr(1, WhiteRooks, xThisSquare) Then
         If Right(xThisSquare, 1) <> "8" Then
            CheckingSquares = CheckingSquares + Left(xThisSquare, 1) + NextRank(Right(xThisSquare, 1), 1) + "-"
         End If
         If Right(xThisSquare, 1) <> "1" Then
            CheckingSquares = CheckingSquares + Left(xThisSquare, 1) + NextRank(Right(xThisSquare, 1), -1) + "-"
         End If
         If Left(xThisSquare, 1) <> "h" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 1) + Right(xThisSquare, 1) + "-"
         End If
         If Left(xThisSquare, 1) <> "a" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -1) + Right(xThisSquare, 1) + "-"
         End If
      End If
      ' Knight moves
      If InStr(1, WhiteKnights, xThisSquare) Then
         If Left(xThisSquare, 1) < "h" And Right(xThisSquare, 1) < "7" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 1) + NextRank(Right(xThisSquare, 1), 2) + "-"
         End If
         If Left(xThisSquare, 1) < "h" And Right(xThisSquare, 1) > "2" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 1) + NextRank(Right(xThisSquare, 1), -2) + "-"
         End If
         If Left(xThisSquare, 1) < "g" And Right(xThisSquare, 1) < "8" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 2) + NextRank(Right(xThisSquare, 1), 1) + "-"
         End If
         If Left(xThisSquare, 1) < "g" And Right(xThisSquare, 1) > "1" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 2) + NextRank(Right(xThisSquare, 1), -1) + "-"
         End If
         
         If Left(xThisSquare, 1) > "a" And Right(xThisSquare, 1) < "7" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -1) + NextRank(Right(xThisSquare, 1), 2) + "-"
         End If
         If Left(xThisSquare, 1) > "a" And Right(xThisSquare, 1) > "2" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -1) + NextRank(Right(xThisSquare, 1), -2) + "-"
         End If
         If Left(xThisSquare, 1) > "b" And Right(xThisSquare, 1) < "8" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -2) + NextRank(Right(xThisSquare, 1), 1) + "-"
         End If
         If Left(xThisSquare, 1) > "b" And Right(xThisSquare, 1) > "1" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -2) + NextRank(Right(xThisSquare, 1), -1) + "-"
         End If
      End If
   'ElseIf PlayerOnMove = 2 Then
   ElseIf xColor = "Black" Then
      If xThisSquare = BlackQueens Then
         x = x
      End If
      ' Diagonal moves like bishop or queen
      If InStr(1, BlackQueens, xThisSquare) Or InStr(1, BlackBishops, xThisSquare) Then
         If Left(xThisSquare, 1) < "h" And Right(xThisSquare, 1) < "8" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 1) + NextRank(Right(xThisSquare, 1), 1) + "-"
         End If
         If Left(xThisSquare, 1) < "h" And Right(xThisSquare, 1) > "1" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 1) + NextRank(Right(xThisSquare, 1), -1) + "-"
         End If
         If Left(xThisSquare, 1) > "a" And Right(xThisSquare, 1) > "1" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -1) + NextRank(Right(xThisSquare, 1), -1) + "-"
         End If
         If Left(xThisSquare, 1) > "a" And Right(xThisSquare, 1) < "8" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -1) + NextRank(Right(xThisSquare, 1), 1) + "-"
         End If
      End If
      ' Vertical moves like rook or queen
      If InStr(1, BlackQueens, xThisSquare) Or InStr(1, BlackRooks, xThisSquare) Then
         If Right(xThisSquare, 1) < "8" Then
            CheckingSquares = CheckingSquares + Left(xThisSquare, 1) + NextRank(Right(xThisSquare, 1), 1) + "-"
         End If
         If Right(xThisSquare, 1) > "1" Then
            CheckingSquares = CheckingSquares + Left(xThisSquare, 1) + NextRank(Right(xThisSquare, 1), -1) + "-"
         End If
         If Left(xThisSquare, 1) < "h" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 1) + Right(xThisSquare, 1) + "-"
         End If
         If Left(xThisSquare, 1) > "a" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -1) + Right(xThisSquare, 1) + "-"
         End If
      End If
      ' Knight moves
      If InStr(1, BlackKnights, xThisSquare) Then
         If Left(xThisSquare, 1) < "h" And Right(xThisSquare, 1) < "7" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 1) + NextRank(Right(xThisSquare, 1), 2) + "-"
         End If
         If Left(xThisSquare, 1) < "h" And Right(xThisSquare, 1) > "2" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 1) + NextRank(Right(xThisSquare, 1), -2) + "-"
         End If
         If Left(xThisSquare, 1) < "g" And Right(xThisSquare, 1) < "8" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 2) + NextRank(Right(xThisSquare, 1), 1) + "-"
         End If
         If Left(xThisSquare, 1) < "g" And Right(xThisSquare, 1) > "1" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), 2) + NextRank(Right(xThisSquare, 1), -1) + "-"
         End If
         
         If Left(xThisSquare, 1) > "a" And Right(xThisSquare, 1) < "7" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -1) + NextRank(Right(xThisSquare, 1), 2) + "-"
         End If
         If Left(xThisSquare, 1) > "a" And Right(xThisSquare, 1) > "2" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -1) + NextRank(Right(xThisSquare, 1), -2) + "-"
         End If
         If Left(xThisSquare, 1) > "b" And Right(xThisSquare, 1) < "8" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -2) + NextRank(Right(xThisSquare, 1), 1) + "-"
         End If
         If Left(xThisSquare, 1) > "b" And Right(xThisSquare, 1) > "1" Then
            CheckingSquares = CheckingSquares + NextFile(Left(xThisSquare, 1), -2) + NextRank(Right(xThisSquare, 1), -1) + "-"
         End If
      End If
   End If
   BlockedByOwnPieces = True
   For x = 1 To (1 + Len(CheckingSquares) \ 3)
      CandidateSquare = GetSquare(CheckingSquares, x)
      If CandidateSquare = "" Then Exit For
      ' Each square should be blocked by an own piece to be able to declare the piece eentirely blocked
      'If PlayerOnMove = 1 Then
      If xColor = "White" Then
         If InStr(1, WhiteKing, CandidateSquare) Or InStr(1, WhiteQueens, CandidateSquare) Or _
         InStr(1, WhiteRooks, CandidateSquare) Or InStr(1, WhiteBishops, CandidateSquare) Or _
         InStr(1, WhiteKnights, CandidateSquare) Or InStr(1, WhitePawns, CandidateSquare) Then
            x = x
         Else
            BlockedByOwnPieces = False
            Exit Function
         End If
      'ElseIf PlayerOnMove = 2 Then
      ElseIf xColor = "Black" Then
         If InStr(1, BlackKing, CandidateSquare) Or InStr(1, BlackQueens, CandidateSquare) Or _
         InStr(1, BlackRooks, CandidateSquare) Or InStr(1, BlackBishops, CandidateSquare) Or _
         InStr(1, BlackKnights, CandidateSquare) Or InStr(1, BlackPawns, CandidateSquare) Then
            x = x
         Else
            BlockedByOwnPieces = False
            Exit Function
         End If
      End If
   Next x
   Exit Function
BlockedByOwnErr:
   MsgBox "BlockedByOwnPieces " + Err.Description
   Call AllowUserAbort
End Function


Function WrongPiece(xOldSquare As String, xPieceSelector) As Boolean
   On Error GoTo WrongPieceErr
   ' Check if move "N1f3" if this corresponds to a knight on g1 or on g5
   WrongPiece = False
   If PieceSelector = "" Then
      Exit Function
   End If
   If Len(xPieceSelector) > 1 Then
      If xOldSquare <> xPieceSelector Then
         WrongPiece = True
         Exit Function
      End If
   Else
      If IsNumeric(PieceSelector) Then
         If Mid(xOldSquare, 2, 1) <> PieceSelector Then
            WrongPiece = True
            Exit Function
         End If
      Else
         If Mid(xOldSquare, 1, 1) <> PieceSelector Then
            WrongPiece = True
            Exit Function
         End If
      End If
   End If
   Exit Function
WrongPieceErr:
   MsgBox "WrongPiece " + Err.Description
   Call AllowUserAbort
End Function

Sub RemoveOldPiece(xNewSquare As String, xColor As String)
   On Error GoTo RemoveOldErr
   ' Just remove the old piece when capturing it
   ' xColor if it's White or Black on the move
   If xColor = "White" Then
      If InStr(1, WhitePawns, xNewSquare) <> 0 Then
         WhitePawns = Replace(WhitePawns, xNewSquare, "")
         WhitePawns = HyphenTrim(WhitePawns)
      End If
      If InStr(1, WhiteQueens, xNewSquare) <> 0 Then
         WhiteQueens = Replace(WhiteQueens, xNewSquare, "")
         WhiteQueens = HyphenTrim(WhiteQueens)
      End If
      If InStr(1, WhiteRooks, xNewSquare) <> 0 Then
         WhiteRooks = Replace(WhiteRooks, xNewSquare, "")
         WhiteRooks = HyphenTrim(WhiteRooks)
         If xNewSquare = "a1" Then
            CastlingPossible = Mid(CastlingPossible, 1, 1) + "0" + Mid(CastlingPossible, 3, 2)
         End If
         If xNewSquare = "h1" Then
            CastlingPossible = "0" + Mid(CastlingPossible, 2, 3)
         End If
      End If
      If InStr(1, WhiteBishops, xNewSquare) <> 0 Then
         WhiteBishops = Replace(WhiteBishops, xNewSquare, "")
         WhiteBishops = HyphenTrim(WhiteBishops)
      End If
      If InStr(1, WhiteKnights, xNewSquare) <> 0 Then
         WhiteKnights = Replace(WhiteKnights, xNewSquare, "")
         WhiteKnights = HyphenTrim(WhiteKnights)
      End If
   End If
   If xColor = "Black" Then
      If InStr(1, BlackPawns, xNewSquare) <> 0 Then
         BlackPawns = Replace(BlackPawns, xNewSquare, "")
         BlackPawns = HyphenTrim(BlackPawns)
      End If
      If InStr(1, BlackQueens, xNewSquare) <> 0 Then
         BlackQueens = Replace(BlackQueens, xNewSquare, "")
         BlackQueens = HyphenTrim(BlackQueens)
      End If
      If InStr(1, BlackRooks, xNewSquare) <> 0 Then
         BlackRooks = Replace(BlackRooks, xNewSquare, "")
         BlackRooks = HyphenTrim(BlackRooks)
         If xNewSquare = "a8" Then
            CastlingPossible = Mid(CastlingPossible, 1, 3) + "0"
         End If
         If xNewSquare = "h8" Then
            CastlingPossible = Mid(CastlingPossible, 1, 2) + "0" + Mid(CastlingPossible, 4, 1)
         End If
      End If
      If InStr(1, BlackBishops, xNewSquare) <> 0 Then
         BlackBishops = Replace(BlackBishops, xNewSquare, "")
         BlackBishops = HyphenTrim(BlackBishops)
      End If
      If InStr(1, BlackKnights, xNewSquare) <> 0 Then
         BlackKnights = Replace(BlackKnights, xNewSquare, "")
         BlackKnights = HyphenTrim(BlackKnights)
      End If
   End If
   Exit Sub
RemoveOldErr:
   MsgBox "RemoveOldPiece " + Err.Description
   Call AllowUserAbort
End Sub

Sub CheckForDeadPosition()
   Dim CountBishopsKnights As Integer
   Dim CountQueens As Integer
   Dim CountRooks As Integer
   Dim CountPawns As Integer
   Dim ThisSquare As String
   Dim StaleMate As Boolean
   On Error GoTo DeadError
   ' I'm in doubt whether this statement should be here, but it is vital for the future checking of dead positions
   ' that piecemoves like IsBishopMove is not hampered by a wrong PieceSelector
   PieceSelector = ""
   Dim x As Integer
   For x = 1 To (1 + Len(WhiteBishops) \ 3)
      ThisSquare = GetSquare(WhiteBishops, x)
      If ThisSquare = "" Then Exit For
      CountBishopsKnights = CountBishopsKnights + 1
   Next x
   For x = 1 To (1 + Len(BlackBishops) \ 3)
      ThisSquare = GetSquare(BlackBishops, x)
      If ThisSquare = "" Then Exit For
      CountBishopsKnights = CountBishopsKnights + 1
   Next x
   For x = 1 To (1 + Len(WhiteKnights) \ 3)
      ThisSquare = GetSquare(WhiteKnights, x)
      If ThisSquare = "" Then Exit For
      CountBishopsKnights = CountBishopsKnights + 1
   Next x
   For x = 1 To (1 + Len(BlackKnights) \ 3)
      ThisSquare = GetSquare(BlackKnights, x)
      If ThisSquare = "" Then Exit For
      CountBishopsKnights = CountBishopsKnights + 1
   Next x
   For x = 1 To (1 + Len(WhiteQueens) \ 3)
      ThisSquare = GetSquare(WhiteQueens, x)
      If ThisSquare = "" Then Exit For
      CountQueens = CountQueens + 1
   Next x
   For x = 1 To (1 + Len(BlackQueens) \ 3)
      ThisSquare = GetSquare(BlackQueens, x)
      If ThisSquare = "" Then Exit For
      CountQueens = CountQueens + 1
   Next x
   For x = 1 To (1 + Len(WhiteRooks) \ 3)
      ThisSquare = GetSquare(WhiteRooks, x)
      If ThisSquare = "" Then Exit For
      CountRooks = CountRooks + 1
   Next x
   For x = 1 To (1 + Len(BlackRooks) \ 3)
      ThisSquare = GetSquare(BlackRooks, x)
      If ThisSquare = "" Then Exit For
      CountRooks = CountRooks + 1
   Next x
   For x = 1 To (1 + Len(WhitePawns) \ 3)
      ThisSquare = GetSquare(WhitePawns, x)
      If ThisSquare = "" Then Exit For
      CountPawns = CountPawns + 1
   Next x
   For x = 1 To (1 + Len(BlackPawns) \ 3)
      ThisSquare = GetSquare(BlackPawns, x)
      If ThisSquare = "" Then Exit For
      CountPawns = CountPawns + 1
   Next x
   If CountBishopsKnights = 1 And CountQueens = 0 And CountRooks = 0 And CountPawns = 0 Then
      StrComment = "Dead position: only one bishop or knight"
      Call InsertComment(StrComment)
      DrawDeclared = True
   End If
   If CountBishopsKnights = 0 And CountQueens = 0 And CountRooks = 0 And CountPawns = 0 Then
      StrComment = "Dead position: only kings"
      Call InsertComment(StrComment)
      DrawDeclared = True
   End If
   If CountBishopsKnights = 2 And CountQueens = 0 And CountRooks = 0 And CountPawns = 0 And _
      Len(WhiteBishops) = 2 And Len(BlackBishops) = 2 Then
      If (IsWhiteSquare(WhiteBishops) And IsWhiteSquare(BlackBishops)) Or _
         (Not IsWhiteSquare(WhiteBishops) And Not IsWhiteSquare(BlackBishops)) Then
         StrComment = "Dead position: one bishop for each player of the same colour squares"
         Call InsertComment(StrComment)
         DrawDeclared = True
      End If
   End If
   If MoveNum >= 71 Then
      x = x
   End If
   If CountBishopsKnights = 2 And CountQueens = 0 And CountRooks = 0 And CountPawns = 0 And _
      Len(WhiteBishops) = 2 And Len(BlackBishops) = 2 Then
      If (IsWhiteSquare(WhiteBishops) And Not IsWhiteSquare(BlackBishops)) Or _
         (Not IsWhiteSquare(WhiteBishops) And IsWhiteSquare(BlackBishops)) Then
         ' Different colour bishops: mate is possible but not forced mate
         StrComment = "B/N vs. B/N cannot checkmate by force - play on"
         If NoImmediateCapture Then
            Call InsertComment(StrComment)
         End If
      End If
   End If
   If CountBishopsKnights = 2 And CountQueens = 0 And CountRooks = 0 And CountPawns = 0 Then
      If (Len(WhiteBishops) = 2 And Len(BlackKnights) = 2) Or _
         (Len(WhiteKnights) = 2 And Len(BlackKnights) = 2) Or _
         (Len(WhiteKnights) = 2 And Len(BlackBishops) = 2) Then
         ' bishop or knight for each player: mate is possible but there is no forced mate
         StrComment = "B/N vs. B/N cannot checkmate by force - play on"
         If NoImmediateCapture Then
            Call InsertComment(StrComment)
         End If
      End If
      If Len(WhiteKnights) = 5 Or Len(BlackKnights) = 5 Then
         ' 2 knights for one player: mate is possible, but not forced mate
         StrComment = "2 knights vs. king cannot checkmate by force - defending player cannot demand a draw"
         If NoImmediateCapture Then
            Call InsertComment(StrComment)
         End If
      End If
   End If
   StartTime = GetMScountNow
   If EndTime > 0 Then
      MS_count_General = MS_count_General + StartTime - EndTime
   End If
   ' This procedure may issue a comment about dead position from blocked pawns
   Call CheckPawnStructure
   EndTime = GetMScountNow
   MS_count_BlockedPos = MS_count_BlockedPos + EndTime - StartTime
  
   ' Stalemate is also a type of dead draw, since no checkmate is possible after stalemate
   AllowPiecesInBetween = True
   StaleMate = False
   If PlayerOnMove = 1 Then
      If WhiteQueens = "" And WhiteRooks = "" And WhiteKnights = "" And WhiteBishops = "" And WhitePawns = "" Then
         If NoLegalSquares("White", WhiteKing) Then
            StaleMate = True
         End If
      End If
      If NoLegalNonKingMoves("White") Then
         If NoLegalSquares("White", WhiteKing) Then
            StaleMate = True
         End If
      End If
   ElseIf PlayerOnMove = 2 Then
      If BlackQueens = "" And BlackRooks = "" And BlackKnights = "" And BlackBishops = "" And BlackPawns = "" Then
         If NoLegalSquares("Black", BlackKing) Then
            StaleMate = True
         End If
      End If
'      If InStr(1, LastMove, "21.b6") <> 0 Then
'         x = x
'      End If
      If NoLegalNonKingMoves("Black") Then
         If NoLegalSquares("Black", BlackKing) Then
            StaleMate = True
         End If
      End If
   End If
   AllowPiecesInBetween = False
   If StaleMate Then
      StrComment = "Dead position: stalemate"
      Call InsertComment(StrComment)
      StalemateAnnounced = True
      DrawDeclared = True
   End If
   Exit Sub
DeadError:
   MsgBox "CheckForDeadPosition " + Err.Description
   Call AllowUserAbort
End Sub

Function NoImmediateCapture() As Boolean
   Dim CombinedInputLine As String
   On Error GoTo NoImmediateCaptureErr
   NoImmediateCapture = True
   CombinedInputLine = StripStr(InputLine, 25, ThisLine, FutureLine, LastMove)
   If InStr(1, LCase(CombinedInputLine), "x") <> 0 Then
      NoImmediateCapture = False
   End If
   Exit Function
NoImmediateCaptureErr:
   MsgBox "NoImmediateCapture " + Err.Description
   Call AllowUserAbort
End Function

Function StripStr(xStr As String, Optional Length As Double, _
   Optional CutStr As String, Optional Extra As String, Optional xLastMove As String) As String
   ' You can call this function with xStr as only parameter
   ' If Length is filled out, the final string is cut to this length
   ' If CutStr is filled out, it will be cut from the beginning of xStr first
   ' If Extra is filled out, it will be added to the resulting string
   ' If xLastMove is filled out, the last part of the move will be removed from xStr
   ' The string xStr will then be removed of comments like {ThisComment} see this example:
   '
   ' 23.Qd4 Nxe5 24.Kb1 {Dead position: blocked pawns} Kc5 27.f4 exf3 {ep}   becomes
   ' 23.Qd4 Nxe5 24.Kb1 Kc5 27.f4 exf3
   Dim x As Integer
   Dim NewStr As String
   Dim StrippedStr As String
   Dim InComment As Boolean
   Dim ThisLen As Integer
   On Error GoTo StripStrErr
   InComment = False
   ' First strip the part of xStr that has already been parsed
   If Len(xStr) > Len(CutStr) Then
      For x = 1 To Len(CutStr)
         If Mid(xStr, x, 1) = "{" Then
            InComment = True
         ElseIf Mid(xStr, x, 1) = "}" Then
            InComment = False
         End If
      Next x
   End If
   ThisLen = 0
   For x = Len(xLastMove) To 1 Step -1
      If Mid(xLastMove, x, 1) = "." Then
         Exit For
      Else
         ThisLen = ThisLen + 1
      End If
   Next x
   If Len(xStr) > (Len(CutStr) + ThisLen) Then
      NewStr = Right(xStr, Len(xStr) - ThisLen - Len(CutStr))
      NewStr = NewStr + Extra
   Else
      NewStr = xStr + Extra
   End If
   StrippedStr = ""
   For x = 1 To Len(NewStr)
      If Mid(NewStr, x, 1) = "{" Then
         InComment = True
      ElseIf Mid(NewStr, x, 1) = "}" Then
         InComment = False
      ElseIf Not InComment Then
         ' Skip double blanks
         If Len(StrippedStr) = 0 Then
            StrippedStr = StrippedStr + Mid(NewStr, x, 1)
         ElseIf Not (Right(StrippedStr, 1) = " " And Mid(NewStr, x, 1) = " ") Then
            StrippedStr = StrippedStr + Mid(NewStr, x, 1)
         End If
      End If
   Next x
   If (Length = 0) Or (Len(StrippedStr) = 0) Then
      StripStr = StrippedStr
   Else
      StripStr = Left(StrippedStr, Min(Len(StrippedStr), Length))
   End If
   Exit Function
StripStrErr:
   MsgBox "StripStr Error " + Err.Description
   Call AllowUserAbort
End Function

Sub CheckPawnStructure()
   ' Check if the pawn structure and remaining bishops warrant a claim to a dead position
   Dim x As Integer
   Dim y As Integer
   Dim Pos As Integer
   Dim ThisSquare As String
   Dim Blocker As String
   Dim LeftSquare As String
   Dim RightSquare As String
   Dim NewSquare As String
   Dim Found As Boolean
   Dim AdjacentSquares As String
   Dim WhiteAccessibleSquares As String
   Dim BlackAccessibleSquares As String
   Dim ProcessedSquares As String
   Dim Blockers As String
   Dim PawnThreat As Byte
   Dim Exonerated As Boolean
   Dim CatchExonerateStr As String
   Dim ExcludeStr As String
   ' The copies of WhitePawns, WhiteBishops, WhiteKnights so that some dead pieces can be eliminated from analysis
   ' Examples are Bb1 and Ba2 if b3 and c2 are blocked pawns. Nh1 if f2 and g3 are blocked pawns
   ' Bc1 and Bf1 if pawns on b2 and d2 and e2 and g2. In this case also black pawns on c2 or f2 are eliminated
   Dim WhitePawns_2 As String
   Dim WhiteBishops_2 As String
   Dim WhiteKnights_2 As String
   Dim BlackPawns_2 As String
   Dim BlackBishops_2 As String
   Dim BlackKnights_2 As String
   On Error GoTo CheckPawnErr
   ' First we add one to PositionsNotChecked, then if we actually enter the procedure for a real check,
   ' subtract one from PositionsNotChecked, and add one to PositionsChecked
   PositionsNotChecked = PositionsNotChecked + 1
   CatchExonerateStr = ""
   ' If the position was found to be dead, stop checking pawn structure! Helped the processing of
   ' longest_dead_position.pgn from 4061 milliseconds to 47 ms, both in average over 10 runs
   ' This was tested on a relatively slow PC
   If DeadPositionFound Then Exit Sub
   If WhitePawns = "" Or BlackPawns = "" Then
      Exit Sub
   End If
   Call EliminateWhiteDeadBishops(WhiteBishops_2, BlackPawns_2)
   Call EliminateBlackDeadBishops(BlackBishops_2, WhitePawns_2)
   Call EliminateWhiteDeadKnights(WhiteKnights_2)
   Call EliminateBlackDeadKnights(BlackKnights_2)
   If WhiteQueens <> "" Or WhiteRooks <> "" Or WhiteKnights_2 <> "" Or _
      BlackQueens <> "" Or BlackRooks <> "" Or BlackKnights_2 <> "" Then
      Exit Sub
   End If
   ' Check if in this pawn endgame, there are three blocked pawns
   ' This helped pass the heavy-processing file longest_dead_game.pgn from an average of 4061 milliseconds (ms)
   ' to 47 ms, obviously a huge gain in processing savvy, since in the 688 move game it needs to check each halfmove
   ' from move 38 to the end, e.g. 650x2 = 1300 halfmoves. The gain will turn smaller for large PGN files with
   ' maybe 164,900 games
   ' In my own 1460 games there were 62 pawn endgames, and from these many positions didn't have 3 blocked pawns, so
   ' only for 40 halfmoves in total for all 62 pawn endgames would this algorithm actually have to run.
   y = 0
   For x = 1 To (1 + Len(WhitePawns_2) \ 3)
      ThisSquare = GetSquare(WhitePawns_2, x)
      If ThisSquare = "" Then Exit For
      Blocker = Left(ThisSquare, 1) + NextRank(Right(ThisSquare, 1), 1)
      If InStr(1, BlackPawns_2, Blocker) Then
         y = y + 1
      End If
   Next x
   If y < 3 Then
      'less than 3 blocked pair of pawns found, stop checking anything more! This cannot be a blocked position
      Exit Sub
   End If
   ' Bishops can be allowed in if they have the same color squares as all the player pawns, and that
   ' the opponents pawns are all of the opposite color. To exonerate such a bishop, all pawns must be
   ' blocked by opponent pawns.
   '
   ' All pawns are blocked by an enemy pawn, and no bishop can capture a pawn is checked
   ' Shouldn't these checks be only when there are bishops?
   ' It really only makes sense to check correlation between bishops and pawns if there are bishops
   If (WhiteBishops_2 <> "" Or BlackBishops_2 <> "") Then
      If y <> (1 + Len(WhitePawns_2) \ 3) Then
         Exit Sub
      End If
      If Len(WhitePawns_2) <> Len(BlackPawns_2) Then
         Exit Sub
      End If
      If HasAllWhiteSquares(WhiteBishops_2) And HasAllWhiteSquares(WhitePawns_2) Then
         If (HasAllBlackSquares(BlackBishops_2)) And (HasAllBlackSquares(BlackPawns_2)) Then
         Else
            Exit Sub
         End If
      End If
      If HasAllWhiteSquares(BlackBishops_2) And HasAllWhiteSquares(BlackPawns_2) Then
         If (HasAllBlackSquares(WhiteBishops_2)) And (HasAllBlackSquares(WhitePawns_2)) Then
         Else
            Exit Sub
         End If
      End If
      ' if they are not all white, they must all be black - checked for each of four categories
      ' - this is so politically incorrect, I won't even comment, think apartheid
      If (Not HasAllWhiteSquares(WhiteBishops_2)) And (Not HasAllBlackSquares(WhiteBishops_2)) Then
         If WhiteBishops_2 <> "" Then Exit Sub
      End If
      If (Not HasAllWhiteSquares(WhitePawns_2)) And (Not HasAllBlackSquares(WhitePawns_2)) Then
         If WhitePawns_2 <> "" Then Exit Sub
      End If
      If (Not HasAllWhiteSquares(BlackBishops_2)) And (Not HasAllBlackSquares(BlackBishops_2)) Then
         If BlackBishops_2 <> "" Then Exit Sub
      End If
      If (Not HasAllWhiteSquares(BlackPawns_2)) And (Not HasAllBlackSquares(BlackPawns_2)) Then
         If BlackPawns_2 <> "" Then Exit Sub
      End If
      ' The last 4 if statements checked inner consistency in a category
      ' but we must also check whether all white's bishops are white-squared and all white's pawns are black-squared
      If HasAllWhiteSquares(WhiteBishops_2) And HasAllBlackSquares(WhitePawns_2) Then
         If WhiteBishops_2 <> "" And WhitePawns_2 <> "" Then Exit Sub
      End If
      If HasAllBlackSquares(WhiteBishops_2) And HasAllWhiteSquares(WhitePawns_2) Then
         If WhiteBishops_2 <> "" And WhitePawns_2 <> "" Then Exit Sub
      End If
      If HasAllWhiteSquares(BlackBishops_2) And HasAllBlackSquares(BlackPawns_2) Then
         If BlackBishops_2 <> "" And BlackPawns_2 <> "" Then Exit Sub
      End If
      If HasAllBlackSquares(BlackBishops_2) And HasAllWhiteSquares(BlackPawns_2) Then
         If BlackBishops_2 <> "" And BlackPawns_2 <> "" Then Exit Sub
      End If
      ' Check if a helpmate with a bishop to a corner square is possible
      '
      ' Check if a2 is white pawn, Bb1 and Ka1 is possible, then Bc3# or Bb2# could be possible
      ' Bb2# may be possible if there is a black pawn on c3, else Bc3# may be possible
      ' Alternative white Bb1 and Ba2 will also allow checkmate
      ' Blocking bishop(s)
      Dim B1 As Integer
      Dim G1 As Integer
      Dim B8 As Integer
      Dim G8 As Integer
      ' Attacking bishops
      Dim B2 As Integer
      Dim G2 As Integer
      Dim B7 As Integer
      Dim G7 As Integer
      B1 = BishopsHavePathTo(WhiteBishops_2, "White", "b1", 2)
      B2 = BishopsHavePathTo(BlackBishops_2, "Black", "b2", 1)
      If (InStr(1, WhitePawns_2, "a2") <> 0 And B1 >= 1) Or (B1 = 2) Then
         If B2 = 1 Then
            Exit Sub
         End If
      End If
      G1 = BishopsHavePathTo(WhiteBishops_2, "White", "g1", 2)
      G2 = BishopsHavePathTo(BlackBishops_2, "Black", "g2", 1)
      If (InStr(1, WhitePawns_2, "h2") <> 0 And G1 >= 1) Or (G1 = 2) Then
         If G2 = 1 Then
            Exit Sub
         End If
      End If
      B8 = BishopsHavePathTo(BlackBishops_2, "Black", "b8", 2)
      B7 = BishopsHavePathTo(WhiteBishops_2, "White", "b7", 1)
      If (InStr(1, BlackPawns_2, "a7") <> 0 And B8 >= 1) Or B8 = 2 Then
         If B7 = 1 Then
            Exit Sub
         End If
      End If
      G8 = BishopsHavePathTo(BlackBishops_2, "Black", "g8", 2)
      G7 = BishopsHavePathTo(WhiteBishops_2, "White", "g7", 1)
      If (InStr(1, BlackPawns, "h7") <> 0 And G8 >= 1) Or G8 = 2 Then
         If G7 = 1 Then
            Exit Sub
         End If
      End If
   End If
   If WhiteBishops <> "" Or BlackBishops <> "" Then
      If Not (InStr(1, StrBlockedBishops, WhitePlayer + "-" + BlackPlayer)) Then
         CatchExonerateStr = "Found bishop(s) in blocked position for game " + _
         WhitePlayer + "-" + BlackPlayer + vbCrLf
      End If
   End If
   If WhiteKnights <> "" Or BlackKnights <> "" Then
      If Not (InStr(1, StrBlockedBishops, WhitePlayer + "-" + BlackPlayer)) Then
         CatchExonerateStr = "Found knight(s) in blocked position for game " + _
         WhitePlayer + "-" + BlackPlayer + vbCrLf
      End If
   End If
   ' Now we know that the position will be checked fully, so modify counters
   PositionsNotChecked = PositionsNotChecked - 1
   PositionsChecked = PositionsChecked + 1
   ' Perform a full run with the white king to determine all accessible squares. For White if we can run to the 8.th.
   ' rank, we know it is possible to access any black pawns, so finish loop. In other cases we get to determine
   ' all available squares. If we don't reach the eighth rank with a full run, usually the position is closed
   ' and potentially a blocked position.
   ProcessedSquares = ""
   ExcludeStr = "KQRBN"
   ' Disregard king blocking only if the player is not on the move
   ' the opponent king will find a path to the first/last rank if there is one, in the other loop
   If PlayerOnMove = 1 And CheckAnnounced Then
      WhiteAccessibleSquares = FindUncheckedSquare(WhiteKing, "White")
   Else
      WhiteAccessibleSquares = WhiteKing
   End If
   While Len(WhiteAccessibleSquares) > Len(ProcessedSquares)
      ThisSquare = GetNextUnprocessedSquare("White", WhiteAccessibleSquares, ProcessedSquares)
      NewSquare = ThisSquare
      AdjacentSquares = GetAdjacentSquares(NewSquare, "White")
      While AdjacentSquares <> ""
         ThisSquare = GetNextSquare(AdjacentSquares) ' This will also snap off ThisSquare from AdjacentSquares
         If (InStr(1, WhiteAccessibleSquares, ThisSquare) = 0) And (InStr(1, WhitePawns, ThisSquare) = 0) Then
            If Not IsThreatenedSquare("White", ThisSquare, ExcludeStr) Then
               If WhiteAccessibleSquares = "" Then
                  WhiteAccessibleSquares = WhiteAccessibleSquares + ThisSquare
               Else
                  WhiteAccessibleSquares = WhiteAccessibleSquares + "-" + ThisSquare
               End If
            End If
         End If
         ' If we made a homerun to the eight rank, abort, this is not a blocked position
         If Right(ThisSquare, 1) = "8" Then
            Call PrintExitSubInfo(ThisSquare, _
               WhiteAccessibleSquares, BlackAccessibleSquares, Blockers, "26")
            Exit Sub
         End If
      Wend
   Wend
   ' Perform a full run with the black king to determine all accessible squares. For White if we can run to the 8.th.
   ' rank, we know it is possible to access any black pawns, so finish loop. In other cases we get to determine
   ' all available squares. If we don't reach the eighth rank with a full run, usually the position is closed
   ' and potentially a blocked position.
   ProcessedSquares = ""
   ' Disregard king blocking only if the player is not on the move
   ' the opponent king will find a path to the first/last rank if there is one, in the other loop
   If LastMove = "45.fxe3+" Then
      x = x
   End If
   If PlayerOnMove = 2 And CheckAnnounced Then
      BlackAccessibleSquares = FindUncheckedSquare(BlackKing, "Black")
   Else
      BlackAccessibleSquares = BlackKing
   End If
   While Len(BlackAccessibleSquares) > Len(ProcessedSquares)
      ThisSquare = GetNextUnprocessedSquare("Black", BlackAccessibleSquares, ProcessedSquares)
      NewSquare = ThisSquare
      AdjacentSquares = GetAdjacentSquares(NewSquare, "Black")
      While AdjacentSquares <> ""
         ThisSquare = GetNextSquare(AdjacentSquares) ' This will also snap off ThisSquare from AdjacentSquares
         If (InStr(1, BlackAccessibleSquares, ThisSquare) = 0) And (InStr(1, BlackPawns, ThisSquare) = 0) Then
            If Not IsThreatenedSquare("Black", ThisSquare, ExcludeStr) Then
               If BlackAccessibleSquares = "" Then
                  BlackAccessibleSquares = BlackAccessibleSquares + ThisSquare
               Else
                  BlackAccessibleSquares = BlackAccessibleSquares + "-" + ThisSquare
               End If
            End If
         End If
         ' If we made a homerun to the first rank, abort, this is not a blocked position
         If Right(ThisSquare, 1) = "1" Then
            Call PrintExitSubInfo(ThisSquare, _
               WhiteAccessibleSquares, BlackAccessibleSquares, Blockers, "27")
            Exit Sub
         End If
      Wend
   Wend
   ' Check if last move has allowed en passant move, if so abort
   If Ep_square <> "" Then
      If PlayerOnMove = 1 Then
         LeftSquare = NextFile(Left(Ep_square, 1), -1) + NextRank(Right(Ep_square, 1), -1)
         If Len(LeftSquare) = 2 And InStr(1, WhitePawns, LeftSquare) <> 0 Then
            Call PrintExitSubInfo(LeftSquare, _
               WhiteAccessibleSquares, BlackAccessibleSquares, Blockers, "9")
            Exit Sub
         End If
         RightSquare = NextFile(Left(Ep_square, 1), 1) + NextRank(Right(Ep_square, 1), -1)
         If Len(RightSquare) = 2 And InStr(1, WhitePawns, RightSquare) <> 0 Then
            Call PrintExitSubInfo(RightSquare, _
               WhiteAccessibleSquares, BlackAccessibleSquares, Blockers, "10")
            Exit Sub
         End If
      ElseIf PlayerOnMove = 2 Then
         LeftSquare = NextFile(Left(Ep_square, 1), -1) + NextRank(Right(Ep_square, 1), 1)
         If Len(LeftSquare) = 2 And InStr(1, BlackPawns, LeftSquare) <> 0 Then
            Call PrintExitSubInfo(LeftSquare, _
               WhiteAccessibleSquares, BlackAccessibleSquares, Blockers, "11")
            Exit Sub
         End If
         RightSquare = NextFile(Left(Ep_square, 1), 1) + NextRank(Right(Ep_square, 1), 1)
         If Len(RightSquare) = 2 And InStr(1, BlackPawns, RightSquare) <> 0 Then
            Call PrintExitSubInfo(RightSquare, _
               WhiteAccessibleSquares, BlackAccessibleSquares, Blockers, "12")
            Exit Sub
         End If
      End If
   End If
   If LastMove = "52.d5+" Or MoveNum = 52 Then
      x = x
   End If
   ' Check if a black pawn is present in the accessible path for the white king - it can be captured
   For x = 1 To (1 + Len(BlackPawns) \ 3)
      ThisSquare = GetSquare(BlackPawns, x)
      If ThisSquare = "" Then Exit For
      If InStr(1, WhiteAccessibleSquares, ThisSquare) <> 0 Then
         ' If we find another white pawn in front of this black pawn, then the position is still blocked
         Found = False
         For y = (Asc(Right(ThisSquare, 1)) + 1 - 48) To 2 Step -1
            NewSquare = Left(ThisSquare, 1) + Chr(y + 48)
            If InStr(1, WhitePawns, NewSquare) <> 0 Then
               Found = True
            End If
         Next y
         If Not Found Then
            Call PrintExitSubInfo(ThisSquare, _
               WhiteAccessibleSquares, BlackAccessibleSquares, NewSquare, "30")
            Exit Sub
         End If
      End If
      If IsPawnVulnerable(ThisSquare, WhiteAccessibleSquares, BlackAccessibleSquares) Then
         Call PrintExitSubInfo(ThisSquare, _
            WhiteAccessibleSquares, BlackAccessibleSquares, NewSquare, "40")
         Exit Sub
      End If
      ' If this pawn has opponent pawns in the path one file left or one file right of the pawn,
      ' this means that moving ahead these pawns can capture eachother
      ' such a pawn can be a problem even when the king is not able to access, if it can advance
      If Left(ThisSquare, 1) > "a" Then
         For y = (Asc(Right(ThisSquare, 1)) - 1 - 48) To 2 Step -1
            NewSquare = Chr(Asc(Left(ThisSquare, 1)) - 1) + Chr(y + 48)
            If InStr(1, BlackPawns, NewSquare) <> 0 Then
               Exit For ' Exit y-loop
            End If
            If InStr(1, WhitePawns, NewSquare) <> 0 Then
               Call PrintExitSubInfo(ThisSquare, _
                  WhiteAccessibleSquares, BlackAccessibleSquares, NewSquare, "31")
               Exit Sub
            End If
         Next y
      End If
      If Left(ThisSquare, 1) < "h" Then
         For y = (Asc(Right(ThisSquare, 1)) - 1 - 48) To 2 Step -1
            NewSquare = Chr(Asc(Left(ThisSquare, 1)) + 1) + Chr(y + 48)
            If InStr(1, BlackPawns, NewSquare) <> 0 Then
               Exit For
            End If
            If InStr(1, WhitePawns, NewSquare) <> 0 Then
               Call PrintExitSubInfo(ThisSquare, _
                  WhiteAccessibleSquares, BlackAccessibleSquares, NewSquare, "32")
               Exit Sub
            End If
         Next y
      End If
      ' Check if this pawn has a blocking pawn
      Found = False
      For y = (Asc(Right(ThisSquare, 1)) - 1 - 48) To 2 Step -1
         NewSquare = Left(ThisSquare, 1) + Chr(y + 48)
         If InStr(1, WhitePawns, NewSquare) <> 0 Then
            Found = True
         End If
      Next y
      If Not Found Then
         Call PrintExitSubInfo(ThisSquare, _
            WhiteAccessibleSquares, BlackAccessibleSquares, NewSquare, "35")
         Exit Sub
      End If
   Next x
   ' Check if a white pawn is present in the accessible path for the black king
   For x = 1 To (1 + Len(WhitePawns) \ 3)
      ThisSquare = GetSquare(WhitePawns, x)
      If ThisSquare = "" Then Exit For
      If InStr(1, BlackAccessibleSquares, ThisSquare) Then
         ' If we find another black pawn in front of this white pawn, then the position is still blocked
         Found = False
         For y = (Asc(Right(ThisSquare, 1)) + 1 - 48) To 2
            NewSquare = Left(ThisSquare, 1) + Chr(y + 48)
            If InStr(1, WhitePawns, NewSquare) <> 0 Then
               Found = True
            End If
         Next y
         If Not Found Then
            Call PrintExitSubInfo(ThisSquare, _
               WhiteAccessibleSquares, BlackAccessibleSquares, Blockers, "33")
            Exit Sub
         End If
      End If
      If ThisSquare = "f6" Then
         x = x
      End If
      If IsPawnVulnerable(ThisSquare, WhiteAccessibleSquares, BlackAccessibleSquares) Then
         Call PrintExitSubInfo(ThisSquare, _
            WhiteAccessibleSquares, BlackAccessibleSquares, NewSquare, "41")
         Exit Sub
      End If
      ' If this pawn has opponent pawns in the path one file left or one file right of the pawn,
      ' this means that moving ahead these pawns can capture eachother
      ' such a pawn can be a problem even when the king is not able to access, if it can advance
      If Left(ThisSquare, 1) > "a" Then
         For y = (Asc(Right(ThisSquare, 1)) + 1 - 48) To 7
            NewSquare = Chr(Asc(Left(ThisSquare, 1)) - 1) + Chr(y + 48)
            If InStr(1, WhitePawns, NewSquare) <> 0 Then
               Exit For ' Exit y-loop
            End If
            If InStr(1, BlackPawns, NewSquare) <> 0 Then
               Call PrintExitSubInfo(ThisSquare, _
                  WhiteAccessibleSquares, BlackAccessibleSquares, NewSquare, "34")
               Exit Sub
            End If
         Next y
      End If
      If Left(ThisSquare, 1) < "h" Then
         For y = (Asc(Right(ThisSquare, 1)) + 1 - 48) To 7
            NewSquare = Chr(Asc(Left(ThisSquare, 1)) + 1) + Chr(y + 48)
            If InStr(1, WhitePawns, NewSquare) <> 0 Then
               Exit For
            End If
            If InStr(1, BlackPawns, NewSquare) <> 0 Then
               Call PrintExitSubInfo(ThisSquare, _
                  WhiteAccessibleSquares, BlackAccessibleSquares, NewSquare, "35")
               Exit Sub
            End If
         Next y
      End If
      ' Check if this pawn has a blocking pawn
      Found = False
      For y = (Asc(Right(ThisSquare, 1)) + 1 - 48) To 7
         NewSquare = Left(ThisSquare, 1) + Chr(y + 48)
         If InStr(1, BlackPawns, NewSquare) <> 0 Then
            Found = True
         End If
      Next y
      If Not Found Then
         Call PrintExitSubInfo(ThisSquare, _
            WhiteAccessibleSquares, BlackAccessibleSquares, NewSquare, "36")
         Exit Sub
      End If
   Next x
   ' Detect stalemate and prevent commenting about blocked position
   If PlayerOnMove = 1 Then
      If NoLegalSquares("White", WhiteKing) And (WhiteBishops = "") And (BlackBishops = "") Then
         Call PrintExitSubInfo(ThisSquare, _
            WhiteAccessibleSquares, BlackAccessibleSquares, Blockers, "28 (stalemate)")
         Exit Sub
      End If
   End If
   If PlayerOnMove = 2 Then
      If NoLegalSquares("Black", BlackKing) And (WhiteBishops = "") And (BlackBishops = "") Then
         Call PrintExitSubInfo(ThisSquare, _
            WhiteAccessibleSquares, BlackAccessibleSquares, Blockers, "29 (stalemate)")
         Exit Sub
      End If
   End If
   Call PrintExitSubInfo(ThisSquare, _
      WhiteAccessibleSquares, BlackAccessibleSquares, Blockers, "OK")
   StrComment = "Dead position: blocked pawns"
   Call InsertComment(StrComment)
   StrBlockedBishops = StrBlockedBishops + CatchExonerateStr
   DeadPositionFound = True
   ' Comment the next assignment to DrawDeclared if you want all messages
   ' about 3-rep and 5-rep etc. in the continuation of the game
   DrawDeclared = True
   Exit Sub
CheckPawnErr:
   MsgBox "CheckPawnStructure " + Err.Description
   Call AllowUserAbort
End Sub

'   Call EliminateWhiteDeadBishops(WhiteBishops_2, BlackPawns_2)
'   Call EliminateBlackDeadBishops(BlackBishops_2, WhitePawns_2)
'   Call EliminateWhiteDeadKnights(WhiteKnights_2)
'   Call EliminateBlackDeadKnights(BlackKnights_2)

Sub EliminateWhiteDeadBishops(WhiteBishops_2 As String, BlackPawns_2 As String)
   WhiteBishops_2 = WhiteBishops
   BlackPawns_2 = BlackPawns
   If InStr(1, WhitePawns, "b3") <> 0 And InStr(1, WhitePawns, "c2") <> 0 Then
      If InStr(1, WhiteBishops_2, "a2") <> 0 Then
         Call EliminateSquare(WhiteBishops_2, "a2")
      End If
      If InStr(1, WhiteBishops_2, "b1") <> 0 Then
         Call EliminateSquare(WhiteBishops_2, "b1")
      End If
   End If
   If InStr(1, WhitePawns, "g3") <> 0 And InStr(1, WhitePawns, "f2") <> 0 Then
      If InStr(1, WhiteBishops_2, "h2") <> 0 Then
         Call EliminateSquare(WhiteBishops_2, "h2")
      End If
      If InStr(1, WhiteBishops_2, "g1") <> 0 Then
         Call EliminateSquare(WhiteBishops_2, "g1")
      End If
   End If
   If InStr(1, WhitePawns, "b2") <> 0 And InStr(1, WhitePawns, "d2") <> 0 Then
      If InStr(1, WhiteBishops_2, "c1") <> 0 Then
         Call EliminateSquare(WhiteBishops_2, "c1")
         If InStr(1, BlackPawns_2, "c2") <> 0 Then
            Call EliminateSquare(BlackPawns_2, "c2")
         End If
      End If
   End If
   If InStr(1, WhitePawns, "e2") <> 0 And InStr(1, WhitePawns, "g2") <> 0 Then
      If InStr(1, WhiteBishops_2, "f1") <> 0 Then
         Call EliminateSquare(WhiteBishops_2, "f1")
         If InStr(1, BlackPawns_2, "f2") <> 0 Then
            Call EliminateSquare(BlackPawns_2, "f2")
         End If
      End If
   End If
End Sub

Sub EliminateBlackDeadBishops(BlackBishops_2 As String, WhitePawns_2 As String)
   BlackBishops_2 = BlackBishops
   WhitePawns_2 = WhitePawns
   If InStr(1, BlackPawns, "b6") <> 0 And InStr(1, BlackPawns, "c7") <> 0 Then
      If InStr(1, BlackBishops_2, "a7") <> 0 Then
         Call EliminateSquare(BlackBishops_2, "a7")
      End If
      If InStr(1, BlackBishops_2, "b8") <> 0 Then
         Call EliminateSquare(BlackBishops_2, "b8")
      End If
   End If
   If InStr(1, BlackPawns, "g6") <> 0 And InStr(1, BlackPawns, "f7") <> 0 Then
      If InStr(1, BlackBishops_2, "h7") <> 0 Then
         Call EliminateSquare(BlackBishops_2, "h7")
      End If
      If InStr(1, BlackBishops_2, "g8") <> 0 Then
         Call EliminateSquare(BlackBishops_2, "g8")
      End If
   End If
   If InStr(1, BlackPawns, "b7") <> 0 And InStr(1, BlackPawns, "d7") <> 0 Then
      If InStr(1, BlackBishops_2, "c8") <> 0 Then
         Call EliminateSquare(BlackBishops_2, "c8")
         If InStr(1, WhitePawns_2, "c7") <> 0 Then
            Call EliminateSquare(WhitePawns_2, "c7")
         End If
      End If
   End If
   If InStr(1, BlackPawns, "e7") <> 0 And InStr(1, BlackPawns, "g7") <> 0 Then
      If InStr(1, BlackBishops_2, "f8") <> 0 Then
         Call EliminateSquare(BlackBishops_2, "f8")
         If InStr(1, WhitePawns_2, "f7") <> 0 Then
            Call EliminateSquare(WhitePawns_2, "f7")
         End If
      End If
   End If
End Sub

Sub EliminateWhiteDeadKnights(WhiteKnights_2 As String)
   Dim x As Integer
   Dim ThisSquare As String
   WhiteKnights_2 = WhiteKnights
   For x = 1 To 8
      ThisSquare = Chr(x + 96) + "1"
      If InStr(1, WhiteKnights_2, ThisSquare) <> 0 Then
         If BlockedByOwnPieces("White", ThisSquare) Then
            Call EliminateSquare(WhiteKnights_2, ThisSquare)
         End If
      End If
   Next x
End Sub

Sub EliminateBlackDeadKnights(BlackKnights_2 As String)
   Dim x As Integer
   Dim ThisSquare As String
   BlackKnights_2 = BlackKnights
   For x = 1 To 8
      ThisSquare = Chr(x + 96) + "8"
      If InStr(1, BlackKnights_2, ThisSquare) <> 0 Then
         If BlockedByOwnPieces("Black", ThisSquare) Then
            Call EliminateSquare(BlackKnights_2, ThisSquare)
         End If
      End If
   Next x
End Sub

Sub EliminateSquare(Group As String, xSquare As String)
   Dim Pos As Integer
   If Len(Group) = 2 Then
      Group = ""
      Exit Sub
   End If
   Pos = InStr(1, Group, xSquare)
   If Pos = 1 Then
      Group = Right(Group, Len(Group) - 3)
      Exit Sub
   End If
   If Pos = Len(Group) - 1 Then
      Group = Left(Group, Len(Group) - 3)
   Else
      Group = Left(Group, Pos - 2) + Right(Group, Len(Group) - 3 - Pos)
   End If
End Sub

Function BishopsHavePathTo(xBishops As String, xColor As String, xSquare As String, xBishopsWanted As Integer) As Integer
   ' Check how many bishops of this color that can go to a specific square. Bishopswanted is set to
   ' number of bishops you want checked, if there are more than 2 bishops with the color, max. 2 will be returned
   ' to avoid unnecessary processing time
   ' Background is to find out if the bishops can set up some kind of helpmate
   ' With a white pawn on a2, a white bishop on b1 and the white king on a1, Bc3# or Bb2# is possible depending
   ' on if there is a Black pawn on c4 or not
   ' Another scenario is two white bishops on a2 and b1 and king on a1 will still allow the same mate, but then
   ' two bishops must be able to move within pawn chain, and a black bishop outside blocked pawns
   Dim x As Integer
   Dim BishopSquares As String
   Dim ProcessedSquares As String
   Dim AdjacentSquares As String
   Dim ThisSquare As String
   Dim NewSquare As String
   Dim BishopCount As Integer
   BishopsHavePathTo = 0 'Default there are no bishops that can go to xSquare
   BishopCount = 0
   If xBishops = "" Then Exit Function
'   If xColor = "White" And xBishops = "" Then Exit Function
'   If xColor = "Black" And xBishops = "" Then Exit Function
   If xColor = "White" Then
      For x = 1 To (1 + Len(WhiteBishops) \ 3)
         ThisSquare = GetSquare(WhiteBishops, x)
         If ThisSquare = "" Then Exit For
         BishopSquares = ThisSquare
         ProcessedSquares = ""
         While Len(BishopSquares) > Len(ProcessedSquares)
            ThisSquare = GetNextUnprocessedSquare("White", BishopSquares, ProcessedSquares)
            NewSquare = ThisSquare
            AdjacentSquares = GetAdjacentSquares(NewSquare, "White")
            While AdjacentSquares <> ""
               ThisSquare = GetNextSquare(AdjacentSquares) ' This will also snap off ThisSquare from AdjacentSquares
               If IsBishopMove(NewSquare, ThisSquare) Then
                  If (InStr(1, BishopSquares, ThisSquare) = 0) And (InStr(1, WhitePawns, ThisSquare) = 0) Then
                     If BishopSquares = "" Then
                        BishopSquares = BishopSquares + ThisSquare
                     Else
                        BishopSquares = BishopSquares + "-" + ThisSquare
                     End If
                  End If
               End If
            Wend
         Wend
         If InStr(1, BishopSquares, xSquare) <> 0 Then
            BishopCount = BishopCount + 1
            If BishopCount >= xBishopsWanted Then
               BishopsHavePathTo = BishopCount
               Exit Function
            End If
         End If
      Next x
   ElseIf xColor = "Black" Then
      For x = 1 To (1 + Len(BlackBishops) \ 3)
         ThisSquare = GetSquare(BlackBishops, x)
         If ThisSquare = "" Then Exit For
         BishopSquares = ThisSquare
         ProcessedSquares = ""
         While Len(BishopSquares) > Len(ProcessedSquares)
            ThisSquare = GetNextUnprocessedSquare("Black", BishopSquares, ProcessedSquares)
            NewSquare = ThisSquare
            AdjacentSquares = GetAdjacentSquares(NewSquare, "Black")
            While AdjacentSquares <> ""
               ThisSquare = GetNextSquare(AdjacentSquares) ' This will also snap off ThisSquare from AdjacentSquares
               If IsBishopMove(NewSquare, ThisSquare) Then
                  If (InStr(1, BishopSquares, ThisSquare) = 0) And (InStr(1, BlackPawns, ThisSquare) = 0) Then
                     If BishopSquares = "" Then
                        BishopSquares = BishopSquares + ThisSquare
                     Else
                        BishopSquares = BishopSquares + "-" + ThisSquare
                     End If
                  End If
               End If
            Wend
         Wend
         If InStr(1, BishopSquares, xSquare) <> 0 Then
            BishopCount = BishopCount + 1
            If BishopCount >= xBishopsWanted Then
               BishopsHavePathTo = BishopCount
               Exit Function
            End If
         End If
      Next x
   End If
   BishopsHavePathTo = BishopCount
End Function

Sub PrintExitSubInfo(xThisSquare As String, _
   xWhiteAccessibleSquares As String, xBlackAccessibleSquares As String, _
   xBlockers As String, xExit As String)
   'Print #4, "Game " + WhitePlayer + " vs. " + BlackPlayer
   'Print #4, "WhitePieces            = " + WhitePieces
   'Print #4, "BlackPieces            = " + BlackPieces
   'Print #4, "LastMove               = " + LastMove
   'Print #4, "ThisSquare             = " + xThisSquare
   'Print #4, "WhiteAccessibleSquares = " + Reordered(xWhiteAccessibleSquares)
   'Print #4, "BlackAccessibleSquares = " + Reordered(xBlackAccessibleSquares)
   'Print #4, "Blockers               = " + xBlockers
   'Print #4, "Exit point             = " + xExit + vbCrLf
   'Call AllowUserAbort
End Sub

Function IsWhiteSquare(xThisSquare As String) As Boolean
   Dim WhiteSquares As String
   IsWhiteSquare = False
   WhiteSquares = "b1-d1-f1-h1--a2-c2-e2-g2--b3-d3-f3-h3--a4-c4-e4-g4--b5-d5-f5-h5--a6-c6-e6-g6--b7-d7-f7-h7--a8-c8-e8-g8"
   If InStr(1, WhiteSquares, xThisSquare) <> 0 Then
      IsWhiteSquare = True
   End If
End Function

Function HasAllWhiteSquares(xSquares As String) As Boolean
   Dim x As Integer
   Dim ThisSquare As String
   HasAllWhiteSquares = True
   If xSquares = "" Then Exit Function
   For x = 1 To (1 + Len(xSquares) \ 3)
      ThisSquare = GetSquare(xSquares, x)
      If ThisSquare = "" Then Exit For
      If Not IsWhiteSquare(ThisSquare) Then
         HasAllWhiteSquares = False
         Exit For
      End If
   Next x
End Function

Function HasAllBlackSquares(xSquares As String) As Boolean
   Dim x As Integer
   Dim ThisSquare As String
   HasAllBlackSquares = True
   If xSquares = "" Then Exit Function
   For x = 1 To (1 + Len(xSquares) \ 3)
      ThisSquare = GetSquare(xSquares, x)
      If ThisSquare = "" Then Exit For
      If IsWhiteSquare(ThisSquare) Then
         HasAllBlackSquares = False
         Exit For
      End If
   Next x
End Function

Function IsPawnVulnerable(xSquare As String, xWhiteAccessibleSquares As String, xBlackAccessibleSquares As String) As Boolean
   ' Vulnerable as in can be captured already or under specific circumstances, for instance runs away from
   ' the protecting pawn
   Dim NextSquare As String
   Dim NextSquare2 As String
   Dim VulnerableSquares As String
   Dim x As Integer
   Dim y As Integer
   On Error GoTo IsPawnVulnerableErr
   ' Check if this pawn can be captured by the opponent king, and if so, if it matters
   ' Good example:
   ' White: a2,c2,e2,g2,a6,c6,e6,g6   Black: a7,c7,e7,g7,a5,c5,e5,g5
   ' Here the black pawn a5 is clearly capturable, but it's demise does not matter because of white pawn a6
   ' keeps the position blocked
   ' Another example:
   ' White: b4,d4,f4,g5,h3   Black: b5,d5,f5,g6,h5
   ' Clearly pawn h5 is not capturable right away, but if Black plays h4 in the wrong moment, when White's king
   ' is on g3, the pawn can be captured, and the position is no longer blocked
   ' If there would have been a white pawn on h6 and a black pawn on h7, h5 is not considered vulnerable
   If xSquare = "f6" Then
      x = x
   End If
   If LastMove = "40.f6+" Then
      x = x
   End If
   IsPawnVulnerable = False
   If InStr(1, WhitePawns, xSquare) Then
      ' Check if protected by a white pawn
      If IsThreatenedSquare("Black", xSquare, "KQRBN") Then
         ' if this pawn is blocked by a pawn, it cannot be vulnerable
         NextSquare = Left(xSquare, 1) + NextRank(Right(xSquare, 1), 1)
         If InStr(1, BlackPawns, NextSquare) Then
            Exit Function
         End If
      End If
      VulnerableSquares = GetAdjacentSquares(xSquare, "White")
      While VulnerableSquares <> "" And (Not IsPawnVulnerable)
         NextSquare = GetNextSquare(VulnerableSquares)
         If InStr(1, xBlackAccessibleSquares, NextSquare) <> 0 Then
            IsPawnVulnerable = True
            For y = (Asc(Right(xSquare, 1)) - 1 - 48) To 2 Step -1
               NextSquare2 = Left(xSquare, 1) + Chr(y + 48)
               If InStr(1, BlackPawns, NextSquare2) <> 0 Then
                  IsPawnVulnerable = False
                  Exit For
               End If
            Next y
         End If
      Wend
   ElseIf InStr(1, BlackPawns, xSquare) Then
      ' Check if protected by a black pawn
      If IsThreatenedSquare("White", xSquare, "KQRBN") Then
         ' if this pawn is blocked by a pawn, it cannot be vulnerable
         NextSquare = Left(xSquare, 1) + NextRank(Right(xSquare, 1), -1)
         If InStr(1, WhitePawns, NextSquare) Then
            Exit Function
         End If
      End If
      VulnerableSquares = GetAdjacentSquares(xSquare, "Black")
      While VulnerableSquares <> "" And (Not IsPawnVulnerable)
         NextSquare = GetNextSquare(VulnerableSquares)
         If InStr(1, xWhiteAccessibleSquares, NextSquare) <> 0 Then
            IsPawnVulnerable = True
            For y = (Asc(Right(xSquare, 1)) + 1 - 48) To 7
               NextSquare2 = Left(xSquare, 1) + Chr(y + 48)
               If InStr(1, WhitePawns, NextSquare2) <> 0 Then
                  IsPawnVulnerable = False
                  Exit For
               End If
            Next y
         End If
      Wend
   End If
   Exit Function
IsPawnVulnerableErr:
   MsgBox "IsVulnerablePawn " + Err.Description
   Call AllowUserAbort
End Function

Function IsPawnThreatening(xSquare As String) As Byte
   ' This Function will check if any pawns are threatening the square. The result will be 1 if one or several
   ' white pawns is threatening it, and 2 if also one or several black pawns are threatening it.
   ' If only one or several black pawns are threatening it, the result is 1. When checking Borderline for holes,
   ' the expected result is 2 for all squares in the borderline
   ' Result = 1  one or several white pawns
   ' Result = 2  one or several black pawns
   ' Result = 3  both 1 and 2 are fulfilled
   Dim ThisSquare As String
   Dim Result As Byte
   Dim FoundWhite As Boolean
   Dim FoundBlack As Boolean
   On Error GoTo IsPawn2Err
   IsPawnThreatening = False
   FoundWhite = False
   FoundBlack = False
   Result = 0
   ThisSquare = Chr(Asc(Left(xSquare, 1)) + 1) + Chr(Asc(Right(xSquare, 1)) - 1)
   If Mid(ThisSquare, 1, 1) > "h" Or Mid(ThisSquare, 1, 1) < "a" Or _
      Mid(ThisSquare, 2, 1) > "8" Or Mid(ThisSquare, 2, 1) < "1" Then
            ' Skip this square
   Else
      If InStr(1, WhitePawns, ThisSquare) <> 0 Then
         FoundWhite = True
      End If
   End If
   ThisSquare = Chr(Asc(Left(xSquare, 1)) - 1) + Chr(Asc(Right(xSquare, 1)) - 1)
   If Mid(ThisSquare, 1, 1) > "h" Or Mid(ThisSquare, 1, 1) < "a" Or _
      Mid(ThisSquare, 2, 1) > "8" Or Mid(ThisSquare, 2, 1) < "1" Then
            ' Skip this square
   Else
      If InStr(1, WhitePawns, ThisSquare) <> 0 Then
         FoundWhite = True
      End If
   End If
   'ThisSquare = Chr(Asc(Mid(xSquare, 1, 1)) + 1) + Chr(Asc(Mid(xSquare, 2, 1)) + 1)
   ThisSquare = GetSquareFromSquare(xSquare, 1, 1)
   'If Mid(ThisSquare, 1, 1) > "h" Or Mid(ThisSquare, 1, 1) < "a" Or _
   '   Mid(ThisSquare, 2, 1) > "8" Or Mid(ThisSquare, 2, 1) < "1" Then
   '         ' Skip this square
   'Else
   'End If
   If InStr(1, BlackPawns, ThisSquare) <> 0 And ThisSquare <> "" Then
      FoundBlack = True
   End If
   
   ThisSquare = Chr(Asc(Left(xSquare, 1)) - 1) + Chr(Asc(Right(xSquare, 1)) + 1)
   If Mid(ThisSquare, 1, 1) > "h" Or Mid(ThisSquare, 1, 1) < "a" Or _
      Mid(ThisSquare, 2, 1) > "8" Or Mid(ThisSquare, 2, 1) < "1" Then
            ' Skip this square
   Else
      If InStr(1, BlackPawns, ThisSquare) <> 0 Then
         FoundBlack = True
      End If
   End If
   If FoundWhite And FoundBlack Then
      IsPawnThreatening = 3
   ElseIf FoundBlack Then
      IsPawnThreatening = 2
   ElseIf FoundWhite Then
      IsPawnThreatening = 1
   Else
      IsPawnThreatening = 0
   End If
   Exit Function
IsPawn2Err:
   MsgBox "IsPawnThreatening " + Err.Description
   Call AllowUserAbort
End Function

Function GetSquareFromSquare(xSquare As String, xStep_y_axis As Integer, xStep_x_axis As Integer) As String
   Dim ThisSquare As String
   On Error GoTo GetSqErr
   ThisSquare = Chr(Asc(Left(xSquare, 1)) + xStep_x_axis) + Chr(Asc(Right(xSquare, 1)) + xStep_y_axis)
   If Mid(ThisSquare, 1, 1) > "h" Or Mid(ThisSquare, 1, 1) < "a" Or _
      Mid(ThisSquare, 2, 1) > "8" Or Mid(ThisSquare, 2, 1) < "1" Then
            ' Skip this square
      GetSquareFromSquare = ""
   Else
      GetSquareFromSquare = ThisSquare
   End If
   Exit Function
GetSqErr:
   MsgBox "GetSquareFromSquare " + Err.Description
   Call AllowUserAbort
End Function

Function NoLegalSquares(xColor As String, xSquare As String) As Boolean
   Dim NewSquare As String
   Dim x As Integer
   Dim y As Integer
   Dim ThisSquare As String
   Dim AdjacentSquares As String
   On Error GoTo NoLegalErr
   NoLegalSquares = False
   If IsThreatenedSquare(xColor, xSquare) Then
      ' If this is a threatened square, then abort with NoLegalSquares as False
      ' since this is a checkmate, at least if all other squares are inaccessible
      Exit Function
   End If
   AdjacentSquares = GetAdjacentSquares(xSquare, xColor)
   While AdjacentSquares <> ""
      ThisSquare = GetNextSquare(AdjacentSquares)  ' This will also snap off ThisSquare from AdjacentSquares
      If Not OccupiedSquare(ThisSquare, xColor) Then
         If Not IsThreatenedSquare(xColor, ThisSquare) Then
            ' We found a legal king move, so exit!
            Exit Function
         End If
      End If
   Wend
   NoLegalSquares = True
   Exit Function
NoLegalErr:
   MsgBox "NoLegalSquares " + Err.Description
   Call AllowUserAbort
End Function

Function GetAdjacentSquares(xSquare As String, xColor As String) As String
   Dim x As Integer
   Dim y As Integer
   Dim Squares As String
   Dim NewSquare As String
   On Error GoTo GetSquaresErr
   Squares = ""
   For x = -1 To 1
      For y = -1 To 1
         If x = 0 And y = 0 Then
            ' Skip king's own square
         Else
            NewSquare = Chr(Asc(Mid(xSquare, 1, 1)) + x) + Chr(Asc(Mid(xSquare, 2, 1)) + y)
            If Mid(NewSquare, 1, 1) > "h" Or Mid(NewSquare, 1, 1) < "a" Or _
               Mid(NewSquare, 2, 1) > "8" Or Mid(NewSquare, 2, 1) < "1" Then
            ' Skip this square
            ElseIf InStr(1, BlackPawns, NewSquare) <> 0 And xColor = "Black" Then
            ' Skip this square
            ElseIf InStr(1, WhitePawns, NewSquare) <> 0 And xColor = "White" Then
            ' Skip this square
            Else
               If Squares = "" Then
                  Squares = Squares + NewSquare
               Else
                  Squares = Squares + "-" + NewSquare
               End If
            End If
         End If
      Next y
   Next x
   GetAdjacentSquares = Squares
   Exit Function
GetSquaresErr:
   MsgBox "GetAdjacentSquares " + Err.Description
   Call AllowUserAbort
End Function

Function GetNextSquare(xAdjacent As String) As String
   If Len(xAdjacent) <= 2 Then
      GetNextSquare = xAdjacent
      xAdjacent = ""
   Else
      GetNextSquare = Right(xAdjacent, 2)
      xAdjacent = Mid(xAdjacent, 1, Len(xAdjacent) - 3)
   End If
End Function

Function GetNextUnprocessedSquare(xColor As String, xColl As String, xProcessedSquares As String) As String
   ' Contrary to GetNextSquare this function does not manipulate the group from which the square is chosen
   ' On the other hand it adds each returned square to the xProcessedSquares string
   ' Call this function with 1 in the first call, and then 2, 3 etc. in the subsquent calls. The first n-1 squares are
   ' then ignored in the processing for each new call
   ' If called with xColor = "White" then give preference to rightmost squares on the eighth rank
   ' If called with xColor = "Black" then give preference to leftmost squares on the first rank
   Dim Score As Integer
   Dim MaxScore As Integer
   Dim ThisSquare As String
   Dim SearchSquare As String
   Dim x As Integer
   On Error GoTo GetNextUnprocErr
   If xColor = "White" Then
      MaxScore = 0
   ElseIf xColor = "Black" Then
      MaxScore = -100   ' minimum score for a square would be -88, so make it lower to get a real square in the loop
   End If
   For x = 1 To (1 + (Len(xColl) \ 3))
      ThisSquare = GetSquare(xColl, x)
      If ThisSquare = "" Then Exit For
      If InStr(1, xProcessedSquares, ThisSquare) = 0 Then
         Score = (Asc(Left(ThisSquare, 1)) - 96) + (Asc(Right(ThisSquare, 1)) - 48) * 10
         
         If xColor = "Black" Then Score = (-1 * Score)   ' reverse order for Black
         If Score > MaxScore Then
            MaxScore = Score
            SearchSquare = ThisSquare
         End If
      End If
   Next x
   GetNextUnprocessedSquare = SearchSquare
   If xProcessedSquares = "" Then
      xProcessedSquares = xProcessedSquares + SearchSquare
   Else
      xProcessedSquares = xProcessedSquares + "-" + SearchSquare
   End If
   Exit Function
GetNextUnprocErr:
   MsgBox "GetNextUnprocessedSquare " + Err.Description
   Call AllowUserAbort
End Function

Function IsThreatenedSquare(xColor As String, xSquare As String, Optional xExcludePieces As String) As Boolean
   ' xColor is the player that wants to check if the opponent is threatening a square
   ' with the ExcludePieces string you can specify which pieces you want to disregard
   ' When checking for blocked pieces, all threats we care about a from the opponent's pawns
   ' for instance a threat from a king or a bishop is irrelevant for the structure of the position
   Dim x As Integer
   Dim ThisSquare As String
   Dim NewSquare As String
   On Error GoTo IsThreatenedError
   If (InStr(1, "abcdefgh", LCase(Left(xSquare, 1))) = 0) Or _
      (InStr(1, "12345678", Right(xSquare, 1) = 0)) Then
      IsThreatenedSquare = True
      Exit Function
   End If
   IsThreatenedSquare = False
   If xColor = "White" Then
      If InStr(1, xExcludePieces, "K") = 0 Then
         If IsKingMove(BlackKing, xSquare) Then
            IsThreatenedSquare = True
            Exit Function
         End If
      End If
      If InStr(1, xExcludePieces, "Q") = 0 Then
         For x = 1 To (1 + Len(BlackQueens) \ 3)
            ThisSquare = GetSquare(BlackQueens, x)
            If ThisSquare = "" Then Exit For
            If IsQueenMove(ThisSquare, xSquare) Then
                If Not PiecesInBetween(ThisSquare, xSquare) Then
                   IsThreatenedSquare = True
                   Exit Function
                End If
            End If
         Next x
      End If
      If InStr(1, xExcludePieces, "R") = 0 Then
         For x = 1 To (1 + Len(BlackRooks) \ 3)
            ThisSquare = GetSquare(BlackRooks, x)
            If ThisSquare = "" Then Exit For
            If IsRookMove(ThisSquare, xSquare) Then
                If Not PiecesInBetween(ThisSquare, xSquare) Then
                   IsThreatenedSquare = True
                   Exit Function
                End If
            End If
         Next x
      End If
      If InStr(1, xExcludePieces, "B") = 0 Then
         For x = 1 To (1 + Len(BlackBishops) \ 3)
            ThisSquare = GetSquare(BlackBishops, x)
            If ThisSquare = "" Then Exit For
            If IsBishopMove(ThisSquare, xSquare) Then
               If Not PiecesInBetween(ThisSquare, xSquare) Then
                  IsThreatenedSquare = True
                  Exit Function
               End If
            End If
         Next x
      End If
      If InStr(1, xExcludePieces, "N") = 0 Then
         For x = 1 To (1 + Len(BlackKnights) \ 3)
            ThisSquare = GetSquare(BlackKnights, x)
            If ThisSquare = "" Then Exit For
            If IsKnightMove(ThisSquare, xSquare) Then
               IsThreatenedSquare = True
               Exit Function
            End If
         Next x
      End If
      If InStr(1, xExcludePieces, "P") = 0 Then
         For x = 1 To (1 + Len(BlackPawns) \ 3)
            ThisSquare = GetSquare(BlackPawns, x)
            If ThisSquare = "" Then Exit For
            NewSquare = NextFile(Left(ThisSquare, 1), -1) + NextRank(Right(ThisSquare, 1), -1)
            If NewSquare = xSquare Then
               IsThreatenedSquare = True
               Exit Function
            End If
            NewSquare = NextFile(Left(ThisSquare, 1), 1) + NextRank(Right(ThisSquare, 1), -1)
            If NewSquare = xSquare Then
               IsThreatenedSquare = True
               Exit Function
            End If
         Next x
      End If
   ElseIf xColor = "Black" Then
      If InStr(1, xExcludePieces, "K") = 0 Then
         If IsKingMove(WhiteKing, xSquare) Then
            IsThreatenedSquare = True
            Exit Function
         End If
      End If
      If InStr(1, xExcludePieces, "Q") = 0 Then
         For x = 1 To (1 + Len(WhiteQueens) \ 3)
            ThisSquare = GetSquare(WhiteQueens, x)
            If ThisSquare = "" Then Exit For
            If IsQueenMove(ThisSquare, xSquare) Then
                If Not PiecesInBetween(ThisSquare, xSquare) Then
                   IsThreatenedSquare = True
                   Exit Function
                End If
            End If
         Next x
      End If
      If InStr(1, xExcludePieces, "R") = 0 Then
         For x = 1 To (1 + Len(WhiteRooks) \ 3)
            ThisSquare = GetSquare(WhiteRooks, x)
            If ThisSquare = "" Then Exit For
            If IsRookMove(ThisSquare, xSquare) Then
                If Not PiecesInBetween(ThisSquare, xSquare) Then
                   IsThreatenedSquare = True
                   Exit Function
                End If
            End If
         Next x
      End If
      If InStr(1, xExcludePieces, "B") = 0 Then
         For x = 1 To (1 + Len(WhiteBishops) \ 3)
            ThisSquare = GetSquare(WhiteBishops, x)
            If ThisSquare = "" Then Exit For
            If IsBishopMove(ThisSquare, xSquare) Then
               If Not PiecesInBetween(ThisSquare, xSquare) Then
                  IsThreatenedSquare = True
                  Exit Function
               End If
            End If
         Next x
      End If
      If InStr(1, xExcludePieces, "N") = 0 Then
         For x = 1 To (1 + Len(WhiteKnights) \ 3)
            ThisSquare = GetSquare(WhiteKnights, x)
            If ThisSquare = "" Then Exit For
            If IsKnightMove(ThisSquare, xSquare) Then
               IsThreatenedSquare = True
               Exit Function
            End If
         Next x
      End If
      If InStr(1, xExcludePieces, "P") = 0 Then
         For x = 1 To (1 + Len(WhitePawns) \ 3)
            ThisSquare = GetSquare(WhitePawns, x)
            If ThisSquare = "" Then Exit For
            NewSquare = NextFile(Left(ThisSquare, 1), -1) + NextRank(Right(ThisSquare, 1), 1)
            If NewSquare = xSquare Then
               IsThreatenedSquare = True
               Exit Function
            End If
            NewSquare = NextFile(Left(ThisSquare, 1), 1) + NextRank(Right(ThisSquare, 1), 1)
            If NewSquare = xSquare Then
               IsThreatenedSquare = True
               Exit Function
            End If
         Next x
      End If
   End If
   Exit Function
IsThreatenedError:
   MsgBox "IsThreatenedSquare " + Err.Description
   Call AllowUserAbort
End Function

Function FindUncheckedSquare(xSquare As String, xColor As String) As String
   Dim AdjacentSquares As String
   Dim Score As Integer
   Dim MaxScore As Integer
   Dim ThisSquare As String
   Dim SearchSquare As String
   Dim x As Integer
   If xColor = "White" Then
      MaxScore = 0
   ElseIf xColor = "Black" Then
      MaxScore = -100   ' minimum score for a square would be -88, so make it lower to get a real square in the loop
   End If
   AdjacentSquares = GetAdjacentSquares(xSquare, xColor)
   While AdjacentSquares <> ""
      ThisSquare = GetNextSquare(AdjacentSquares)
      If Not IsThreatenedSquare(xColor, ThisSquare, "QRBN") Then
         FindUncheckedSquare = ThisSquare
         Score = (Asc(Left(ThisSquare, 1)) - 96) + (Asc(Right(ThisSquare, 1)) - 48) * 10
         If xColor = "Black" Then Score = (-1 * Score)   ' reverse order for Black
         If Score > MaxScore Then
            MaxScore = Score
            SearchSquare = ThisSquare
         End If
      End If
   Wend
   FindUncheckedSquare = SearchSquare
End Function

Function NoLegalNonKingMoves(xColor As String) As Boolean
   Dim x As Integer
   Dim y As Integer
   Dim Pinned As Boolean
   Dim ThisSquare As String
   Dim NewSquare As String
   Dim Pinner As String
   Dim Blocker As String
   Dim FirstSquare As String
   Dim SecondSquare As String
   On Error GoTo NoLegalNonErr
   NoLegalNonKingMoves = True
   If xColor = "White" Then
      For x = 1 To (1 + Len(WhiteQueens) \ 3)
         ThisSquare = GetSquare(WhiteQueens, x)
         If ThisSquare = "" Then Exit For
         If Not BlockedByOwnPieces(xColor, ThisSquare) Then
            ' We found a legal move for another piece than the king
            NoLegalNonKingMoves = False
            Exit Function
         End If
      Next x
      For x = 1 To (1 + Len(WhiteRooks) \ 3)
         ThisSquare = GetSquare(WhiteRooks, x)
         If ThisSquare = "" Then Exit For
         If Not BlockedByOwnPieces(xColor, ThisSquare) Then
            If IsPinned(ThisSquare, WhiteKing) Then
               If Not IsFrozen(ThisSquare) Then
                  ' We found a legal move for another piece than the king
                  NoLegalNonKingMoves = False
                  Exit Function
               End If
            Else
               ' We found a legal move for another piece than the king
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
      Next x
      For x = 1 To (1 + Len(WhiteBishops) \ 3)
         ThisSquare = GetSquare(WhiteBishops, x)
         If ThisSquare = "" Then Exit For
         If Not BlockedByOwnPieces(xColor, ThisSquare) Then
            If IsPinned(ThisSquare, WhiteKing) Then
               If Not IsFrozen(ThisSquare) Then
                  ' We found a legal move for another piece than the king
                  NoLegalNonKingMoves = False
                  Exit Function
               End If
            Else
               ' We found a legal move for another piece than the king
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
      Next x
      For x = 1 To (1 + Len(WhiteKnights) \ 3)
         ThisSquare = GetSquare(WhiteKnights, x)
         If ThisSquare = "" Then Exit For
         If Not BlockedByOwnPieces(xColor, ThisSquare) Then
            If IsPinned(ThisSquare, WhiteKing) Then
               If Not IsFrozen(ThisSquare) Then
                  ' We found a legal move for another piece than the king
                  NoLegalNonKingMoves = False
                  Exit Function
               End If
            Else
               ' We found a legal move for another piece than the king
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
      Next x
      For x = 1 To (1 + Len(WhitePawns) \ 3)
         ThisSquare = GetSquare(WhitePawns, x)
         If ThisSquare = "" Then Exit For
         Blocker = Left(ThisSquare, 1) + NextRank(Right(ThisSquare, 1), 1)
         If Not OccupiedSquare(Blocker) Then
            If IsPinned(ThisSquare, WhiteKing) Then
               If Not IsFrozen(ThisSquare) Then
                  ' We found a legal move for another piece than the king
                  NoLegalNonKingMoves = False
                  Exit Function
               End If
            Else
            ' We found an unblocked pawn, that can be moved
               ' We found a legal move for another piece than the king
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
         NewSquare = NextFile(Left(Blocker, 1), -1) + NextRank(Right(ThisSquare, 1), 1)
         If Len(NewSquare) = 2 Then
            If InStr(1, BlackQueens, NewSquare) <> 0 Or InStr(1, BlackRooks, NewSquare) <> 0 Or _
               InStr(1, BlackBishops, NewSquare) <> 0 Or InStr(1, BlackKnights, NewSquare) <> 0 Or _
               InStr(1, BlackPawns, NewSquare) <> 0 Then
               ' We found an opponent piece to capture
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
         NewSquare = NextFile(Left(Blocker, 1), 1) + NextRank(Right(ThisSquare, 1), 1)
         If Len(NewSquare) = 2 Then
            If InStr(1, BlackQueens, NewSquare) <> 0 Or InStr(1, BlackRooks, NewSquare) <> 0 Or _
               InStr(1, BlackBishops, NewSquare) <> 0 Or InStr(1, BlackKnights, NewSquare) <> 0 Or _
               InStr(1, BlackPawns, NewSquare) <> 0 Then
               ' We found an opponent piece to capture
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
      Next x
   ElseIf xColor = "Black" Then
      For x = 1 To (1 + Len(BlackQueens) \ 3)
         ThisSquare = GetSquare(BlackQueens, x)
         If ThisSquare = "" Then Exit For
            ' We found a legal move for another piece than the king
         If Not BlockedByOwnPieces(xColor, ThisSquare) Then
            NoLegalNonKingMoves = False
            Exit Function
         End If
      Next x
      For x = 1 To (1 + Len(BlackRooks) \ 3)
         ThisSquare = GetSquare(BlackRooks, x)
         If ThisSquare = "" Then Exit For
         If Not BlockedByOwnPieces(xColor, ThisSquare) Then
            If IsPinned(ThisSquare, BlackKing) Then
               If Not IsFrozen(ThisSquare) Then
                  ' We found a legal move for another piece than the king
                  NoLegalNonKingMoves = False
                  Exit Function
               End If
            Else
               ' We found a legal move for another piece than the king
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
      Next x
      For x = 1 To (1 + Len(BlackBishops) \ 3)
         ThisSquare = GetSquare(BlackBishops, x)
         If ThisSquare = "" Then Exit For
         If Not BlockedByOwnPieces(xColor, ThisSquare) Then
            If IsPinned(ThisSquare, BlackKing) Then
               If Not IsFrozen(ThisSquare) Then
                  ' We found a legal move for another piece than the king
                  NoLegalNonKingMoves = False
                  Exit Function
               End If
            Else
               ' We found a legal move for another piece than the king
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
      Next x
      For x = 1 To (1 + Len(BlackKnights) \ 3)
         ThisSquare = GetSquare(BlackKnights, x)
         If ThisSquare = "" Then Exit For
         If Not IsPinned(ThisSquare, BlackKing) Then
            ' We found a legal move for another piece than the king
         If Not BlockedByOwnPieces(xColor, ThisSquare) Then
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
      Next x
      For x = 1 To (1 + Len(BlackPawns) \ 3)
         ThisSquare = GetSquare(BlackPawns, x)
         If ThisSquare = "" Then Exit For
         Blocker = Left(ThisSquare, 1) + Chr(Asc(Val(Right(ThisSquare, 1))) - 1)
                     ' Perhaps we need to check if the pawn can actually capture the bishop?
         If Not OccupiedSquare(Blocker) Then
            ' We found an unblocked pawn, that can be moved
            If Not IsPinned(ThisSquare, BlackKing) Then
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
         NewSquare = NextFile(Left(Blocker, 1), -1) + NextRank(Right(ThisSquare, 1), -1)
         If Len(NewSquare) = 2 Then
            If InStr(1, WhiteQueens, NewSquare) <> 0 Or InStr(1, WhiteRooks, NewSquare) <> 0 Or _
               InStr(1, WhiteBishops, NewSquare) <> 0 Or InStr(1, WhiteKnights, NewSquare) <> 0 Or _
               InStr(1, WhitePawns, NewSquare) <> 0 Then
               ' We found an opponent piece to capture
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
         NewSquare = NextFile(Left(Blocker, 1), 1) + NextRank(Right(ThisSquare, 1), -1)
         If Len(NewSquare) = 2 Then
            If InStr(1, WhiteQueens, NewSquare) <> 0 Or InStr(1, WhiteRooks, NewSquare) <> 0 Or _
               InStr(1, WhiteBishops, NewSquare) <> 0 Or InStr(1, WhiteKnights, NewSquare) <> 0 Or _
               InStr(1, WhitePawns, NewSquare) <> 0 Then
               ' We found an opponent piece to capture
               NoLegalNonKingMoves = False
               Exit Function
            End If
         End If
      Next x
   End If
   ' Finally if ep_square is filled out, it means that en passant is possible since a double step pawn move
   ' was played. There may be no actual pawns present that can take en passant, or they are pinned
   If Ep_square <> "" Then
      ' Put additional check if the pawn that could make en passant is actually pinned
      ' but only if necessary
      ' If Black played c7-c5, we check for white pawns on b5 or d5
      ' If White played c2-c4, we check for black pawns on b4 or d4
      ' Any of these pawns only count as a contradiction if they are not pinned
      ThisSquare = Right(LastMove, 2)
      If Left(ThisSquare, 1) > "a" Then
         FirstSquare = NextFile(Left(ThisSquare, 1), -1) + Right(ThisSquare, 1)
         If PlayerOnMove = 1 Then
            If InStr(1, WhitePawns, FirstSquare) Then
               If Not IsPinned(FirstSquare, WhiteKing) Then
                  ' We found a pawn to capture en passant
                  NoLegalNonKingMoves = False
                  Exit Function
               End If
            End If
         ElseIf PlayerOnMove = 2 Then
            If InStr(1, BlackPawns, FirstSquare) Then
               If Not IsPinned(FirstSquare, BlackKing) Then
                  ' We found a pawn to capture en passant
                  NoLegalNonKingMoves = False
                  Exit Function
               End If
            End If
         End If
      End If
      If Left(ThisSquare, 1) < "h" Then
         SecondSquare = NextFile(Left(ThisSquare, 1), 1) + Right(ThisSquare, 1)
         If PlayerOnMove = 1 Then
            If InStr(1, WhitePawns, SecondSquare) Then
               If Not IsPinned(SecondSquare, WhiteKing) Then
                  ' We found a pawn to capture en passant
                  NoLegalNonKingMoves = False
                  Exit Function
               End If
            End If
         ElseIf PlayerOnMove = 2 Then
            If InStr(1, BlackPawns, SecondSquare) Then
               If Not IsPinned(SecondSquare, BlackKing) Then
                 ' We found a pawn to capture en passant
                  NoLegalNonKingMoves = False
                  Exit Function
               End If
            End If
          End If
      End If
   End If
   Exit Function
NoLegalNonErr:
   MsgBox "NoLegalNonKingMoves " + Err.Description
   Call AllowUserAbort
End Function

Function IsFrozen(Pinned As String)
   Dim x As Integer
   Dim ThisSquare As String
   On Error GoTo FrozenErr
   IsFrozen = False
   ' A knight will always be frozen in a pin
   If InStr(1, WhiteKnights, Pinned) Or InStr(1, BlackKnights, Pinned) Then
      IsFrozen = True
      Exit Function
   End If
   ' A queen will never be frozen in a pin, it can capture it's pinner or move towards it
   If InStr(1, WhiteQueens, Pinned) Or InStr(1, BlackQueens, Pinned) Then
      IsFrozen = False
      Exit Function
   End If
   If PlayerOnMove = 1 Then
      If JumpsLikeABishop(Pinned, WhiteKing) Then
         ' The pinning piece can be a bishop or a queen
         For x = 1 To (1 + Len(BlackQueens) \ 3)
            ThisSquare = GetSquare(BlackQueens, x)
            If ThisSquare = "" Then Exit For
            If Not PiecesInBetween(ThisSquare, Pinned) And Not PiecesInBetween(Pinned, WhiteKing) Then
               If InStr(1, WhiteRooks, Pinned) Or InStr(1, WhitePawns, Pinned) Then
                  IsFrozen = True
                  Exit Function
               End If
            End If
         Next x
         For x = 1 To (1 + Len(BlackBishops) \ 3)
            ThisSquare = GetSquare(BlackBishops, x)
            If ThisSquare = "" Then Exit For
            If Not PiecesInBetween(ThisSquare, Pinned) And Not PiecesInBetween(Pinned, WhiteKing) Then
               If InStr(1, WhiteRooks, Pinned) Or InStr(1, WhitePawns, Pinned) Then
                  IsFrozen = True
                  Exit Function
               End If
            End If
         Next x
      ElseIf JumpsLikeARook(Pinned, WhiteKing) Then
         ' The pinning piece can be a rook or a queen
         For x = 1 To (1 + Len(BlackQueens) \ 3)
            ThisSquare = GetSquare(BlackQueens, x)
            If ThisSquare = "" Then Exit For
            If Not PiecesInBetween(ThisSquare, Pinned) And Not PiecesInBetween(Pinned, WhiteKing) Then
               If InStr(1, WhiteBishops, Pinned) Or InStr(1, WhitePawns, Pinned) Then
                  IsFrozen = True
               End If
            End If
         Next x
         For x = 1 To (1 + Len(BlackRooks) \ 3)
            ThisSquare = GetSquare(BlackRooks, x)
            If ThisSquare = "" Then Exit For
            If Not PiecesInBetween(ThisSquare, Pinned) And Not PiecesInBetween(Pinned, WhiteKing) Then
               If InStr(1, WhiteBishops, Pinned) Or InStr(1, WhitePawns, Pinned) Then
                  IsFrozen = True
               End If
            End If
         Next x
      End If
   ElseIf PlayerOnMove = 2 Then
      If JumpsLikeABishop(Pinned, BlackKing) Then
         ' The pinning piece can be a bishop or a queen
         For x = 1 To (1 + Len(WhiteQueens) \ 3)
            ThisSquare = GetSquare(WhiteQueens, x)
            If ThisSquare = "" Then Exit For
            If Not PiecesInBetween(ThisSquare, Pinned) And Not PiecesInBetween(Pinned, BlackKing) Then
               If InStr(1, BlackRooks, Pinned) Or InStr(1, BlackPawns, Pinned) Then
                  IsFrozen = True
               End If
            End If
         Next x
         For x = 1 To (1 + Len(WhiteBishops) \ 3)
            ThisSquare = GetSquare(WhiteBishops, x)
            If ThisSquare = "" Then Exit For
            If Not PiecesInBetween(ThisSquare, Pinned) And Not PiecesInBetween(Pinned, BlackKing) Then
               If InStr(1, BlackRooks, Pinned) Or InStr(1, BlackPawns, Pinned) Then
                  IsFrozen = True
               End If
            End If
         Next x
      ElseIf JumpsLikeARook(Pinned, BlackKing) Then
         ' The pinning piece can be a rook or a queen
         For x = 1 To (1 + Len(WhiteQueens) \ 3)
            ThisSquare = GetSquare(WhiteQueens, x)
            If ThisSquare = "" Then Exit For
               If Not PiecesInBetween(ThisSquare, Pinned) And Not PiecesInBetween(Pinned, BlackKing) Then
               If InStr(1, BlackBishops, Pinned) Then
                  IsFrozen = True
               End If
               ' A pawn is not frozen from a rook pin unless the pinner or the king is in the way
               If InStr(1, BlackPawns, Pinned) Then
                  If OccupiedSquare(Left(BlackPawns, 1) + NextRank(Right(BlackPawns, 1), -1)) Then
                     IsFrozen = True
                  End If
               End If
            End If
         Next x
         For x = 1 To (1 + Len(WhiteRooks) \ 3)
            ThisSquare = GetSquare(WhiteRooks, x)
            If ThisSquare = "" Then Exit For
            If Not PiecesInBetween(ThisSquare, Pinned) And Not PiecesInBetween(Pinned, BlackKing) Then
               If InStr(1, BlackBishops, Pinned) Then
                  IsFrozen = True
               End If
               ' A pawn is not frozen from a rook pin unless the pinner or the king is in the way
               If InStr(1, BlackPawns, Pinned) Then
                  If OccupiedSquare(Left(BlackPawns, 1) + NextRank(Right(BlackPawns, 1), -1)) Then
                     IsFrozen = True
                  End If
               End If
            End If
         Next x
      End If
   End If
   Exit Function
FrozenErr:
   MsgBox "IsFrozen " + Err.Description
   Call AllowUserAbort
End Function

Function HyphenTrim(xInStr As String) As String
   ' Remove leading and trailing hyphens, as well as intermediate "--" values
   Dim x As Byte
   Dim NewString As String
   On Error GoTo HyphenTrimErr
   NewString = xInStr
   NewString = Replace(NewString, "--", "-")
   For x = 1 To Len(NewString)
      If Mid(NewString, 1, 1) = "-" Then
         NewString = Mid(NewString, 2, Len(NewString) - 1)
      Else
         Exit For
      End If
   Next x
   Do While Len(NewString) > 0
      If Right(NewString, 1) = "-" Then
         NewString = Mid(NewString, 1, Len(NewString) - 1)
      Else
         Exit Do
      End If
   Loop
   HyphenTrim = NewString
   Exit Function
HyphenTrimErr:
   MsgBox "HyphenTrim " + Err.Description
   Call AllowUserAbort
End Function

Function LeftBracketNotInComment(xStr As String) As Boolean
   ' If there is a bracket "[" in the line, we assume it is a new gameheader
   ' If the bracket is inside a comment, we will not assume it is a new gameheader
   Dim x As Long
   Dim InsideComment As Boolean
   On Error GoTo LeftBracketErr
   InsideComment = False
   LeftBracketNotInComment = False
   For x = 1 To Len(xStr)
      If Mid(xStr, x, 1) = "{" Then
         InsideComment = True
      End If
      If Mid(xStr, x, 1) = "}" Then
         InsideComment = False
      End If
      If Mid(xStr, x, 1) = "[" And Not InsideComment Then
         LeftBracketNotInComment = True
      End If
   Next x
   Exit Function
LeftBracketErr:
   MsgBox "LeftBracketNotInComment " + Err.Description
   Call AllowUserAbort
End Function

Sub FlushPGNfile()
   Dim x As Long
   Dim MyName As String
   On Error GoTo FlushPGNErr
   MyName = Dir(GlobalPath + "\PGNDRAW_" + Right("000000000" + Trim(Str(FileVer)), 10) + ".pgn")
   While Trim(MyName <> "")
      FileVer = FileVer + 1
      MyName = Dir(GlobalPath + "\PGNDRAW_" + Right("000000000" + Trim(Str(FileVer)), 10) + ".pgn")
   Wend
   Open GlobalPath + "\PGNDRAW_" + Right("000000000" + Trim(Str(FileVer)), 10) + ".pgn" For Output As #5
   For x = 1 To MaxFlush
      Print #5, FlushMe(x)
   Next x
   Close #5
   For x = 1 To MaxFlush
      FlushMe(x) = ""
   Next x
   MaxFlush = 0
   Exit Sub
FlushPGNErr:
   MsgBox "FlushPGNfile " + Err.Description
   Call AllowUserAbort
End Sub

Sub AllowUserAbort()
   Dim vbresp As Integer
   vbresp = MsgBox("An error occurred. Do you want to abort the program?", vbYesNo + vbDefaultButton2)
   If vbresp = vbYes Then
      Close #4
      End
   End If
End Sub

Function GetMScountNow() As Currency
   Dim epoch As Currency
   epoch = (DateDiff("s", "1/1/1970", Date) + Timer) * 1000
   GetMScountNow = Round(epoch)
End Function
