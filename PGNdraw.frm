VERSION 5.00
Begin VB.Form frmAnalyze 
   Caption         =   "PGN draw tool"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameOutputFiles 
      Caption         =   "Tiebreaks to select for analyzing predictiveness (optional)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   7695
      Begin VB.CheckBox chkNorg 
         Caption         =   "Show 3-rep"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkDirect 
         Caption         =   "Show 5-rep"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox chkWins 
         Caption         =   "Show 50 moves rule"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chkBlacks 
         Caption         =   "Show 75 moves rule"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CheckBox chkBerger 
         Caption         =   "Show dead position Kvs K"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CheckBox chkKoya 
         Caption         =   "Show dead position single B or N"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CheckBox chkBuchholz 
         Caption         =   "Buchholz"
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkBuchCut1 
         Caption         =   "Buchholz Cut 1"
         Height          =   255
         Left            =   3840
         TabIndex        =   24
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox chkBuchCut2 
         Caption         =   "Buchholz Cut 2"
         Height          =   255
         Left            =   3840
         TabIndex        =   23
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox chkMedBuch1 
         Caption         =   "Median Buchholz 1"
         Height          =   255
         Left            =   3840
         TabIndex        =   22
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox chkAro 
         Caption         =   "Show dead position stalemate"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   2655
      End
      Begin VB.CheckBox chkAro1 
         Caption         =   "Show dead position B vs B"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   2655
      End
      Begin VB.CheckBox chkMedBuch2 
         Caption         =   "Median Buchholz 2"
         Height          =   195
         Left            =   3840
         TabIndex        =   19
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox chkSumBuch 
         Caption         =   "Sum of opponents Buchholz"
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CheckBox chkProgressive 
         Caption         =   "Progressive"
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox chkRincewind1 
         Caption         =   "Rincewind1"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CheckBox chkNorgaard2 
         Caption         =   "Norgaard2"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   3120
         Width           =   2175
      End
      Begin VB.CheckBox chkWinbonus 
         Caption         =   "Show dead position blocked pawns"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txtRoundsToCut 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Text            =   "2"
         Top             =   4200
         Width           =   375
      End
      Begin VB.CheckBox chkUnplayedAsAverage 
         Caption         =   "Unplayed counted as average in scoregroup"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   4200
         Width           =   3615
      End
      Begin VB.CheckBox chkSmerdon 
         Caption         =   "Show 2 lone knights no mate"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3480
         Width           =   3495
      End
      Begin VB.CheckBox chkMissingPoints 
         Caption         =   "Show FEN in commentary"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   3480
         Width           =   2295
      End
      Begin VB.CheckBox chkZermelo 
         Caption         =   "Show each single piece no mate"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   3375
      End
      Begin VB.CheckBox chkAllFiles 
         Caption         =   "All files"
         Height          =   255
         Left            =   5760
         TabIndex        =   8
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CheckBox chkPepechuy 
         Caption         =   "Show {ep} for en passant"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox txtTopPct 
         Height          =   285
         Left            =   7080
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "100"
         Top             =   4200
         Width           =   375
      End
      Begin VB.Label lblWinbonus 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   36
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label lblRoundsToCut 
         Caption         =   "Rounds to cut:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label lblSmerdon 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   34
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblZermelo 
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblTopPct 
         Caption         =   "PCT"
         Height          =   255
         Left            =   7080
         TabIndex        =   32
         Top             =   3840
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "click me!"
      Top             =   4920
      Width           =   6615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "&Analyze!"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Image picHelp 
      Height          =   480
      Left            =   7080
      Top             =   5200
      Width           =   480
   End
   Begin VB.Label lblInputFile 
      Caption         =   "Input file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   855
   End
End
Attribute VB_Name = "frmAnalyze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAro_Click()
   If chkAro.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tAvRating
      'lblAro = Trim(Str(MaxTieMethods))
      'xAverageRating = MaxTieMethods
   Else
      'lblAro = ""
      'Call RemoveTiebreak(tAvRating)
   End If
End Sub

Private Sub chkAro1_Click()
   If chkAro1.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tAvRating1
      'lblAro1 = Trim(Str(MaxTieMethods))
      'xAverageRatingCut1 = MaxTieMethods
   Else
      'lblAro1 = ""
      'Call RemoveTiebreak(tAvRating1)
   End If
End Sub

Private Sub chkBerger_Click()
   If chkBerger.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'If xUseOldstyleSB Then
      '   OrderTieMethod(MaxTieMethods) = tBerger2
      'Else
      '   OrderTieMethod(MaxTieMethods) = tBerger
      'End If
      'lblBerger = Trim(Str(MaxTieMethods))
      'xBerger = MaxTieMethods
   Else
      'lblBerger = ""
      'xBerger = 0
      'Call RemoveTiebreak(tBerger)
   End If
End Sub

Private Sub chkBlacks_Click()
   If chkBlacks.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tMostBlacks
      'lblBlacks = Trim(Str(MaxTieMethods))
      'xMostBlacks = MaxTieMethods
   Else
      'lblBlacks = ""
      'xMostBlacks = 0
      'Call RemoveTiebreak(tMostBlacks)
   End If
End Sub

Private Sub chkBuchCut1_Click()
   If chkBuchCut1.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tBuchholzCut1
      'lblBuchCut1 = Trim(Str(MaxTieMethods))
      'xBuchholzCut1 = MaxTieMethods
   Else
      'lblBuchCut1 = ""
      'xBuchholzCut1 = 0
      'Call RemoveTiebreak(tBuchholzCut1)
   End If
End Sub

Private Sub chkBuchCut2_Click()
   If chkBuchCut2.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tBuchholzCut2
      'lblBuchCut2 = Trim(Str(MaxTieMethods))
      'xBuchholzCut2 = MaxTieMethods
   Else
      'lblBuchCut2 = ""
      'xBuchholzCut2 = 0
      'Call RemoveTiebreak(tBuchholzCut2)
   End If
End Sub

Private Sub chkBuchholz_Click()
   If chkBuchholz.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tBuchholz
      'lblBuch = Trim(Str(MaxTieMethods))
      'xBuchholz = MaxTieMethods
   Else
      'lblBuch = ""
      'xBuchholz = 0
      'Call RemoveTiebreak(tBuchholz)
   End If
End Sub

Private Sub chkDirect_Click()
   If chkDirect.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tDirect
      'lblDirect = Trim(Str(MaxTieMethods))
      'xDirectEncounter = MaxTieMethods
   Else
      'lblDirect = ""
      'xDirectEncounter = 0
      'Call RemoveTiebreak(tDirect)
   End If
End Sub

Private Sub chkKoya_Click()
   If chkKoya.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tKoya
      'lblKoya = Trim(Str(MaxTieMethods))
      'xKoya = MaxTieMethods
   Else
      'lblKoya = ""
      'xKoya = 0
      'Call RemoveTiebreak(tKoya)
   End If
End Sub

Private Sub chkMedBuch1_Click()
   If chkMedBuch1.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tMedBuch1
      'lblMedBuch1 = Trim(Str(MaxTieMethods))
      'xMedianBuchholz1 = MaxTieMethods
   Else
      'lblMedBuch1 = ""
      'xMedianBuchholz1 = 0
      'Call RemoveTiebreak(tMedBuch1)
   End If
End Sub

Private Sub chkMedBuch2_Click()
   If chkMedBuch2.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tMedBuch2
      'lblMedBuch2 = Trim(Str(MaxTieMethods))
      'xMedianBuchholz2 = MaxTieMethods
   Else
      'lblMedBuch2 = ""
      'xMedianBuchholz2 = 0
      'Call RemoveTiebreak(tMedBuch2)
   End If
End Sub

Private Sub chkNorg_Click()
   If chkNorg.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tNorgaard1
      'lblNorg = Trim(Str(MaxTieMethods))
      'xNorgaard = MaxTieMethods
   Else
      'lblNorg = ""
      'xNorgaard = 0
      'Call RemoveTiebreak(tNorgaard1)
   End If
End Sub

Private Sub chkNorgaard2_Click()
   If chkNorgaard2.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tNorgaard2
      'lblNorgaard2 = Trim(Str(MaxTieMethods))
      'xNorgaard2 = MaxTieMethods
   Else
      'lblNorgaard2 = ""
      'xNorgaard2 = 0
      'Call RemoveTiebreak(tNorgaard2)
   End If
End Sub

Private Sub chkProgressive_Click()
   If chkProgressive.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tProgressive
      'lblProgressive = Trim(Str(MaxTieMethods))
      'xProgressive = MaxTieMethods
   Else
      'lblProgressive = ""
      'xProgressive = 0
      'Call RemoveTiebreak(tProgressive)
   End If
End Sub

Private Sub chkRincewind1_Click()
   If chkRincewind1.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tRincewind1
      'lblRincewind1 = Trim(Str(MaxTieMethods))
      'xRincewind1 = MaxTieMethods
   Else
      'lblRincewind1 = ""
      'xRincewind1 = 0
      'Call RemoveTiebreak(tRincewind1)
   End If
End Sub

Private Sub chkMissingPoints_Click()
   If chkMissingPoints.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tMissingPoints
      'lblMissingPoints = Trim(Str(MaxTieMethods))
      'xMissingPoints = MaxTieMethods
   Else
      'lblMissingPoints = ""
      'xMissingPoints = 0
      'Call RemoveTiebreak(tMissingPoints)
   End If
End Sub

Private Sub chkZermelo_Click()
   If chkZermelo.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tZermelo
      'lblZermelo = Trim(Str(MaxTieMethods))
      'xZermelo = MaxTieMethods
   Else
      'lblZermelo = ""
      'xZermelo = 0
      'Call RemoveTiebreak(tZermelo)
   End If
End Sub

Private Sub chkSmerdon_Click()
   If chkSmerdon.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tSmerdon
      'lblSmerdon = Trim(Str(MaxTieMethods))
      'xSmerdon = MaxTieMethods
   Else
      'lblSmerdon = ""
      'xSmerdon = 0
      'Call RemoveTiebreak(tSmerdon)
   End If
End Sub

Private Sub chkSumBuch_Click()
   If chkSumBuch.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tSumBuch
      'lblSumBuch = Trim(Str(MaxTieMethods))
      'xSumBuchholz = MaxTieMethods
   Else
      'lblSumBuch = ""
      'xSumBuchholz = 0
      'Call RemoveTiebreak(tSumBuch)
   End If
End Sub

Private Sub chkWinbonus_Click()
   If chkWinbonus.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tWinbonus
      'lblWinbonus = Trim(Str(MaxTieMethods))
      'xWinbonus = MaxTieMethods
   Else
      'lblWinbonus = ""
      'xWinbonus = 0
      'Call RemoveTiebreak(tWinbonus)
   End If
End Sub

Private Sub chkWins_Click()
   If chkWins.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tMostWins
      'lblWins = Trim(Str(MaxTieMethods))
      'xMostWins = MaxTieMethods
   Else
      'lblWins = ""
      'xMostWins = 0
      'Call RemoveTiebreak(tMostWins)
   End If
End Sub

Private Sub chkPepechuy_Click()
   If chkPepechuy.Value = vbChecked Then
      'MaxTieMethods = MaxTieMethods + 1
      'OrderTieMethod(MaxTieMethods) = tPepechuy
      'lblPepechuy = Trim(Str(MaxTieMethods))
      'xPepechuy = MaxTieMethods
   Else
      'lblPepechuy = ""
      'xPepechuy = 0
      'Call RemoveTiebreak(tPepechuy)
   End If
End Sub

Public Sub RemoveTiebreak(xTie As Byte)
   Dim x As Integer
   Dim y As Integer
   Dim Keep As Integer
   'For x = 1 To MaxTieMethods
   '   If OrderTieMethod(x) = xTie Then
   '      For y = x To MaxTieMethods - 1
   '         OrderTieMethod(y) = OrderTieMethod(y + 1)
   '      Next y
   '      OrderTieMethod(MaxTieMethods) = 0
   '      MaxTieMethods = MaxTieMethods - 1
   '      Keep = x
   '      Exit For
   '   End If
   'Next x
   ' Clean up variables that represent the other selected options
   'If xNorgaard > Keep Then xNorgaard = xNorgaard - 1
   'If xNorgaard2 > Keep Then xNorgaard2 = xNorgaard2 - 1
   'If xBuchholz > Keep Then xBuchholz = xBuchholz - 1
   'If xBuchholzCut1 > Keep Then xBuchholzCut1 = xBuchholzCut1 - 1
   'If xMedianBuchholz1 > Keep Then xMedianBuchholz1 = xMedianBuchholz1 - 1
   'If xMedianBuchholz2 > Keep Then xMedianBuchholz2 = xMedianBuchholz2 - 1
   'If xSumBuchholz > Keep Then xSumBuchholz = xSumBuchholz - 1
   'If xDirectEncounter > Keep Then xDirectEncounter = xDirectEncounter - 1
   'If xMostWins > Keep Then xMostWins = xMostWins - 1
   'If xMostBlacks > Keep Then xMostBlacks = xMostBlacks - 1
   'If xBerger > Keep Then xBerger = xBerger - 1
   'If xKoya > Keep Then xKoya = xKoya - 1
   'If xAverageRating > Keep Then xAverageRating = xAverageRating - 1
   'If xAverageRatingCut1 > Keep Then xAverageRatingCut1 = xAverageRatingCut1 - 1
   'If xWinbonus > Keep Then xWinbonus = xWinbonus - 1
   'If xSmerdon > Keep Then xSmerdon = xSmerdon - 1
   'If xProgressive > Keep Then xProgressive = xProgressive - 1
   'If xRincewind1 > Keep Then xRincewind1 = xRincewind1 - 1
   'If xNorgaard2 > Keep Then xNorgaard2 = xNorgaard2 - 1
   'If xMissingPoints > Keep Then xMissingPoints = xMissingPoints - 1
   'If xZermelo > Keep Then xZermelo = xZermelo - 1
   'If xPepechuy > Keep Then xPepechuy = xPepechuy - 1
   ' Perform HouseCleaning on screen
'   If xTie = tNorgaard1 Then
'      lblNorg = ""
'   ElseIf xTie = tNorgaard2 Then
'      lblNorgaard2 = ""
'   ElseIf xTie = tBuchholz Then
'      lblBuch = ""
'   ElseIf xTie = tBuchholzCut1 Then
'      lblBuchCut1 = ""
'   ElseIf xTie = tBuchholzCut2 Then
'      lblBuchCut2 = ""
'   ElseIf xTie = tMedBuch1 Then
'      lblMedBuch1 = ""
'   ElseIf xTie = tMedBuch2 Then
'      lblMedBuch2 = ""
'   ElseIf xTie = tSumBuch Then
'      lblSumBuch = ""
'   ElseIf xTie = tDirect Then
'      lblDirect = ""
'   ElseIf xTie = tMostWins Then
'      lblWins = ""
'   ElseIf xTie = tMostBlacks Then
'      lblBlacks = ""
'   ElseIf xTie = tBerger Then
'      lblBerger = ""
'   ElseIf xTie = tKoya Then
'      lblKoya = ""
'   ElseIf xTie = tAvRating Then
'      lblAro = ""
'   ElseIf xTie = tAvRating1 Then
'      lblAro1 = ""
'   ElseIf xTie = tWinbonus Then
'      lblWinbonus = ""
'   ElseIf xTie = tRincewind1 Then
'      lblRincewind1 = ""
'   ElseIf xTie = tProgressive Then
'      lblProgressive = ""
'   ElseIf xTie = tMissingPoints Then
'      lblMissingPoints = ""
'   ElseIf xTie = tZermelo Then
'      lblZermelo = ""
'   ElseIf xTie = tSmerdon Then
'      lblSmerdon = ""
'   ElseIf xTie = tPepechuy Then
'      lblPepechuy = ""
'   End If
   Call PutScreenLabels
End Sub

Private Sub cmdAnalyze_Click()
   Dim Response As Long
   Dim OutputFiles As String
   Dim IncBlock As Integer
   Dim Treatment As Integer
   Dim StatType As Integer
   Dim Spatial As Integer
   Dim SelColumns As Integer
   Dim Editor As String
   Dim ThisFile As String
   Dim ThisPath As String
   Dim x As Integer
   Dim MyName As String
   Dim ThisDir As String
   Dim InputLine As String
   Dim FileCount As Integer
   Dim FullFilename As String
   Dim MsgTxt As String
   On Error GoTo FileError
   OutputFiles = ""
   Screen.MousePointer = vbHourglass
   MS_count_BlockedPos = 0
   MS_count_General = 0
   StartTime = 0
   StrBlockedBishops = ""
   PlayerNamesStr = ""
   PositionsChecked = 0
   PositionsNotChecked = 0
   '
   ' Also initialize variables to cater for number of rounds to cut, and if unplayed games counted as average of other
   '
   'If Not (optIncBlock(0) Or optIncBlock(1) Or optIncBlock(2)) Then
   If txtFile.Text = "click me!" Or txtFile.Text = "" Then
      Screen.MousePointer = vbDefault
      MsgBox "File name must be filled out!"
      txtFile.SetFocus
      Exit Sub
   End If
   Editor = xEditor
   Call InitializeNormalBoard
   'MsgBox "StopMeNow = " + Trim(Str(StopMeNow))
   StopMeNow = False
   '
   GlobalPath = App.Path
   ThisPath = GlobalPath
   ThisFile = txtFile.Text
   If InStr(1, ThisFile, "\") <> 0 Then
       For x = Len(ThisFile) To 1 Step -1
          If Mid(ThisFile, x, 1) = "\" Then
             ThisPath = Mid(ThisFile, 1, x - 1)
             ThisFile = Mid(ThisFile, x + 1, Len(ThisFile) - x)
             Exit For
          End If
       Next x
   End If
   If OutputFileInUse(GlobalPath + "\PGNDRAW.TXT") Then
      Screen.MousePointer = vbDefault
      MsgBox "You must avoid using output file PGNDRAW.TXT first!" + Chr(13) + _
         "Did you open it in Excel?"
      Exit Sub
   End If
   ' Check if file exists before deleting it
   MyName = Dir(GlobalPath + "\PGNDRAW.TXT")
   FullFilename = GlobalPath + "\PGNDRAW.TXT"
   If Trim$(MyName <> "") Then
      Kill GlobalPath + "\PGNDRAW.TXT"
   Else
      MsgBox "Couldn't kill the output file!"
   End If
   Open GlobalPath + "\PGNDRAW.TXT" For Output As #4
   ' First make a loop to find all the subdirectories of the current directory
   ' Make sure that C:\ will be checked as C: to make the rest of the code work
   'If PrintOne And (xProcessAllDetails = 0) Then
   If x < 10000 Then
      FullFilename = ThisPath + "\" + ThisFile
      Call ParseAndAnalyze(ThisPath, ThisFile)
      Call PrintReport(ThisPath, ThisFile)
   Else
      ProcessAllFiles = True
      ThisDir = ThisPath
      If Mid(ThisDir, 2, 2) = ":\" And Len(ThisDir) = 3 Then
         ThisDir = Left(ThisDir, 2)
      End If
      ThisIx = 0
      FileCount = 0
      Call GetSubdirs(ThisDir)
      For x = 1 To ThisIx
         ThisFile = Dir(Subdirs(x) + "\*.pgn", vbNormal)
         Do While ThisFile <> ""
            If ThisFile <> "." And ThisFile <> ".." Then
               If (GetAttr(Subdirs(x) + "\" + ThisFile) And vbDirectory) = 0 And _
                  (GetAttr(Subdirs(x) + "\" + ThisFile) And vbReadOnly) = 0 Then
                  If UCase(ThisFile) <> "TIEBREAK.TXT" Then
                     FileCount = FileCount + 1
                     If FileCount > 1 Then
                        Print #4, ""
                     End If
                     Call ParseAndAnalyze(Subdirs(x), ThisFile)
                     Call PrintReport(Subdirs(x), ThisFile)
                     If FileCount = 1 Then
                         'ProcessAllFiles = False
                         Screen.MousePointer = vbDefault
                         Response = MsgBox("The first file in the directory is not a Swiss tournament" + _
                            Chr(13) + "Are you sure you want to continue with all files?", vbOKCancel)
                         If (Response = vbCancel) Then
                            Exit Sub
                         End If
                         ProcessAllFiles = True
                         Screen.MousePointer = vbHourglass
                     End If
                  End If
                  DoEvents
               End If
            End If
            ThisFile = Dir
         Loop
      Next x
      Print #4, ""
   End If
   If Not StopMeNow Then
      MsgTxt = ThisPath + "\" + ThisFile + vbCrLf + _
         Format(GameNumber - 1, "###,###,##0") + IIf(GameNumber > 2, " games", " game") + vbCrLf + _
         "Checking pawn structure took " + Format(MS_count_BlockedPos, "###,###,##0") + " milliseconds" + vbCrLf + _
         "General program time use " + Format(MS_count_General, "###,###,##0") + " milliseconds" + vbCrLf + _
         "Percentage of pawn structure processing was " + _
         Format(MS_count_BlockedPos, "###,###,##0") + "/" + _
         Format(MS_count_BlockedPos + MS_count_General, "###,###,##0") + " = " + _
         Format(100 * (MS_count_BlockedPos / (MS_count_BlockedPos + MS_count_General)), "0.00") + " %" + vbCrLf + _
         "Percentage of positions checked " + Format(PositionsChecked, "###,###,##0") + "/" + _
         Format(PositionsNotChecked + PositionsChecked, "###,###,##0") + " = " + _
         Format(100 * (PositionsChecked / (PositionsChecked + PositionsNotChecked)), "0.00") + " %"
      If StrBlockedBishops <> "" Then
         MsgTxt = MsgTxt + vbCrLf + StrBlockedBishops
      End If
      MsgBox MsgTxt
      MsgTxt = "{" + vbCrLf + MsgTxt + IIf(Len(StrBlockedBishops) = 0, vbCrLf, "") + "}"
      Print #4, MsgTxt
      Print #4, PlayerNamesStr
   End If
   Close #4
   ' First terminate a previous version of TIEBREAK.TXT to regenerate properly
   ' If this is not done, the user will not receive the changes of TIEBREAK.TXT
   ' but would be looking at the old version
   If (WordpadID <> 0) Then
      Call KillProcess(WordpadID)
      WordpadID = 0
   End If
   ' Don't wait for this process, since we want to allow the
   ' user to prepare a new analysis while still looking at the
   ' report file of the errors in the last run as another window.
   WordpadID = Shell(GlobalEditorPath + " " + GlobalPath + "\PGNDRAW.TXT", vbNormalFocus)
   ProcessAllFiles = False
   Screen.MousePointer = vbDefault
   Exit Sub
FileError:
   ProcessAllFiles = False
   Screen.MousePointer = vbDefault
   MsgBox "Error trying to open file " + FullFilename
   Close #1
   Close #4
End Sub

Private Sub cmdExit_Click()
   Unload Me
   End
End Sub

Private Sub cmdOptions_Click()
   frmOptions.Show vbModal
End Sub

Private Sub Form_Load()
   Dim Comando As String
   Dim Curminute As Integer
   Dim Newminute As Integer
   Comando = Command
   If Comando <> "" Then
      txtFile.Text = Comando
   End If
   ' Center form
   frmAnalyze.Top = (Screen.Height - frmAnalyze.Height) / 2
   frmAnalyze.Left = (Screen.Width - frmAnalyze.Width) / 2
   frmAnalyze.Caption = "PGN draw tool v." + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "   " + _
      Format(FileDateTime(App.Path + "\" + App.EXEName + ".exe"), "YYYY-MM-DD HH:MM:SS")
   ReadIniFile
   GetEditorPath
   GetAcrobatPath
   'lblNorg = ""
   'lblDirect = ""
   'lblWins = ""
   ''lblBlacks = ""
   'lblBerger = ""
   'lblKoya = ""
   'lblAro = ""
   'lblAro1 = ""
   'lblBuch = ""
   'lblBuchCut1 = ""
   'lblBuchCut2 = ""
   'lblMedBuch1 = ""
   'lblMedBuch2 = ""
   'lblSumBuch = ""
   'lblProgressive = ""
   'lblNorgaard2 = ""
   'lblRincewind1 = ""
   'lblWinbonus = ""
   'lblSmerdon = ""
   'lblMissingPoints = ""
   'lblZermelo = ""
   'lblPepechuy = ""
   Call PutScreenOptions
   Call PutScreenLabels
End Sub

Private Sub GetSubdirs(CurrentDir As String)
   Dim ThisDir As String
   Dim ThisFile As String
   Dim ThisLine As String
   Dim NewIx As Integer
   Dim Recursive As Boolean
   Dim x As Integer

   ' First make a loop to find all the subdirectories of the current directory
   If ThisIx = 0 Then
      ThisIx = ThisIx + 1
      Subdirs(ThisIx) = CurrentDir
   End If
   NewIx = ThisIx + 1
   Recursive = True
   If Recursive Then
      ThisFile = Dir(CurrentDir + "\*.*", vbDirectory)
      Do While ThisFile <> ""
         If ThisFile <> "." And ThisFile <> ".." Then
            If (GetAttr(CurrentDir + "\" + ThisFile) And vbDirectory) = vbDirectory Then
               ThisIx = ThisIx + 1
               Subdirs(ThisIx) = CurrentDir + "\" + ThisFile
            End If
         End If
         ThisFile = Dir
      Loop
   End If
   For x = NewIx To ThisIx
      Call GetSubdirs(Subdirs(x))
   Next x
End Sub

Sub PutScreenOptions()
   Dim x As Integer
   Dim y As Integer
   'MaxTieMethods = 0
'   For x = 1 To 22
'      If x = xNorgaard Then
'         chkNorg.Value = vbChecked
'      End If
'      If x = xNorgaard2 Then
'         chkNorgaard2.Value = vbChecked
'      End If
'      If x = xBuchholz Then
'         chkBuchholz.Value = vbChecked
'      End If
'      If x = xBuchholzCut1 Then
'         chkBuchCut1.Value = vbChecked
'      End If
'      If x = xBuchholzCut2 Then
'         chkBuchCut2.Value = vbChecked
'      End If
'      If x = xMedianBuchholz1 Then
'         chkMedBuch1.Value = vbChecked
'      End If
'      If x = xMedianBuchholz2 Then
'         chkMedBuch2.Value = vbChecked
'      End If
'      If x = xSumBuchholz Then
'         chkSumBuch.Value = vbChecked
'      End If
'      If x = xDirectEncounter Then
'         chkDirect.Value = vbChecked
'      End If
'      If x = xMostWins Then
'         chkWins.Value = vbChecked
'      End If
'      If x = xMostBlacks Then
'         chkBlacks.Value = vbChecked
'      End If
'      If x = xBerger Then
'         chkBerger.Value = vbChecked
'      End If
'      If x = xKoya Then
'         chkKoya.Value = vbChecked
'      End If
'      If x = xAverageRating Then
'         chkAro.Value = vbChecked
'      End If
'      If x = xAverageRatingCut1 Then
'         chkAro1.Value = vbChecked
'      End If
'      If x = xWinbonus Then
'         chkWinbonus.Value = vbChecked
'      End If
'      If x = xRincewind1 Then
'         chkRincewind1.Value = vbChecked
'      End If
'      If x = xProgressive Then
'         chkProgressive.Value = vbChecked
'      End If
'      If x = xMissingPoints Then
'         chkMissingPoints.Value = vbChecked
'      End If
'      If x = xZermelo Then
'         chkZermelo.Value = vbChecked
'      End If
'      If x = xSmerdon Then
'         chkSmerdon.Value = vbChecked
'      End If
'      If x = xPepechuy Then
'         chkPepechuy.Value = vbChecked
'      End If
'   Next x
End Sub

Sub PutScreenLabels()
   Dim x As Integer
   ' Now put the right values to the rest of the screen labels, indicating selected tiebreak methods, in which order
'   For x = 1 To MaxTieMethods
'      If OrderTieMethod(x) = tNorgaard1 Then
'         lblNorg = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tNorgaard2 Then
'         lblNorgaard2 = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tBuchholz Then
'         lblBuch = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tBuchholzCut1 Then
'         lblBuchCut1 = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tBuchholzCut2 Then
'         lblBuchCut2 = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tMedBuch1 Then
'         lblMedBuch1 = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tMedBuch2 Then
'         lblMedBuch2 = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tSumBuch Then
'         lblSumBuch = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tDirect Then
'         lblDirect = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tMostWins Then
'         lblWins = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tMostBlacks Then
'         lblBlacks = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tBerger Then
'         lblBerger = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tKoya Then
'         lblKoya = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tAvRating Then
'         lblAro = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tAvRating1 Then
'         lblAro1 = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tWinbonus Then
'         lblWinbonus = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tRincewind1 Then
'         lblRincewind1 = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tProgressive Then
'         lblProgressive = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tMissingPoints Then
'         lblMissingPoints = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tZermelo Then
'         lblZermelo = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tPepechuy Then
'         lblPepechuy = Trim(Str(x))
'      ElseIf OrderTieMethod(x) = tSmerdon Then
'         lblSmerdon = Trim(Str(x))
'      End If
'   Next x
'   If xRoundsToCut > 0 Then
'      txtRoundsToCut = Trim(Str(xRoundsToCut))
'   End If
'   If xTopPct > 0 Then
'      txtTopPct = Trim(Str(xTopPct))
'   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call SaveIniFile
End Sub

Sub ReadIniFile()
   Dim i As Integer
   Dim InputLine As String
   Dim MyName As String
   ' Read file TIEBREAK.INI for settings with <num> a number from 1 to 18, or missing
   ' Norgaard=<num>
   ' DirectEncounter=<num>
   ' MostWins=<num>
   ' MostBlacks=<num>
   ' Berger=<num>
   ' Koya=<num>
   ' AverageRating=<num>
   ' AverageRatingCut1=<num>
   ' Winbonus=<num>
   ' Buchholz=<num>
   ' BuchholzCut1=<num>
   ' BuchholzCut2=<num>
   ' MedianBuchholz1=<num>
   ' MedianBuchholz2=<num>
   ' SumBuchholz=<num>
   ' Progressive=<num>
   ' Rincewind1=<num>
   ' Norgaard2=<num>
   ' RoundsToCut=<num>
   ' UnplayedUseAverage=<0,1>
   ' Editor=<Wordpad,Notepad>
   ' ShowFile=<0,1>
   ' NorgaardFinal=<0,1>
   ' ShowWinbonus=<0,1>
   ' OverviewBeforeLast=<0,1>
   ' BuchSummary=<0,1>
   ' BergerSummary=<0,1>
   ' NorgaardSummary=<0,1>
   ' ProgressiveSummary=<0,1>
   ' CombinedSummary=<0,1>
   ' CombinedFinal=<0,1>
   ' CombinedBeforeLast=<0,1>
   ' ShowPredictionStats=<0,1>
   ' ShowDirectPredictiveness=<0,1>
   ' ShowIndirectPredictiveness=<0,1>
   ' GeneralInfo=<0,1>
   ' ShowSwissRounds=<0,1>
   ' ShowInfoAfterRating=<0,1>
   ' RemoveDoubleSpaces=<0,1>
   ' ShowAllTiebreakValues=<0,1>
   ' UseTPRforAvgRating=<0,1>
   i = 0
   ' Put default values for the fields
   ' Check if file exists before reading it
   MyName = Dir("PGNDRAW.INI")
   If Trim$(MyName = "") Then
      Exit Sub
   End If
   Open "PGNDRAW.INI" For Input As #1
   While Not EOF(1)
      i = i + 1
      Line Input #1, InputLine
      InputLine = InputLine + Space(50)
'      If UCase(Mid$(InputLine, 1, 9)) = "NORGAARD=" Then
'         xNorgaard = CByte(Mid(InputLine, 10, Len(InputLine) - 10 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 16)) = "DIRECTENCOUNTER=" Then
'         xDirectEncounter = CByte(Mid(InputLine, 17, Len(InputLine) - 17 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 9)) = "MOSTWINS=" Then
'         xMostWins = CByte(Mid(InputLine, 10, Len(InputLine) - 10 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 11)) = "MOSTBLACKS=" Then
'         xMostBlacks = CByte(Mid(InputLine, 12, Len(InputLine) - 12 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 7)) = "BERGER=" Then
'         xBerger = CByte(Mid(InputLine, 8, Len(InputLine) - 8 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 5)) = "KOYA=" Then
'         xKoya = CByte(Mid(InputLine, 6, Len(InputLine) - 6 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 14)) = "AVERAGERATING=" Then
'         xAverageRating = CByte(Mid(InputLine, 15, Len(InputLine) - 15 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 18)) = "AVERAGERATINGCUT1=" Then
'         xAverageRatingCut1 = CByte(Mid(InputLine, 19, Len(InputLine) - 19 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 9)) = "WINBONUS=" Then
'         xWinbonus = CByte(Mid(InputLine, 10, Len(InputLine) - 10 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 9)) = "BUCHHOLZ=" Then
'         xBuchholz = CByte(Mid(InputLine, 10, Len(InputLine) - 10 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 13)) = "BUCHHOLZCUT1=" Then
'         xBuchholzCut1 = CByte(Mid(InputLine, 14, Len(InputLine) - 14 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 13)) = "BUCHHOLZCUT2=" Then
'         xBuchholzCut2 = CByte(Mid(InputLine, 14, Len(InputLine) - 14 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 16)) = "MEDIANBUCHHOLZ1=" Then
'         xMedianBuchholz1 = CByte(Mid(InputLine, 17, Len(InputLine) - 17 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 16)) = "MEDIANBUCHHOLZ2=" Then
'         xMedianBuchholz2 = CByte(Mid(InputLine, 17, Len(InputLine) - 17 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 12)) = "SUMBUCHHOLZ=" Then
'         xSumBuchholz = CByte(Mid(InputLine, 13, Len(InputLine) - 13 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 12)) = "PROGRESSIVE=" Then
'         xProgressive = CByte(Mid(InputLine, 13, Len(InputLine) - 13 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 11)) = "RINCEWIND1=" Then
'         xRincewind1 = CByte(Mid(InputLine, 12, Len(InputLine) - 12 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 10)) = "NORGAARD2=" Then
'         xNorgaard2 = CByte(Mid(InputLine, 11, Len(InputLine) - 11 + 2))
'      'ElseIf UCase(Mid$(InputLine, 1, 11)) = "RINCEWIND2=" Then
'      ElseIf UCase(Mid$(InputLine, 1, 14)) = "MISSINGPOINTS=" Then
'         xMissingPoints = CByte(Mid(InputLine, 15, Len(InputLine) - 15 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 8)) = "ZERMELO=" Then
'         xZermelo = CByte(Mid(InputLine, 9, Len(InputLine) - 9 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 8)) = "SMERDON=" Then
'         xSmerdon = CByte(Mid(InputLine, 9, Len(InputLine) - 9 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 9)) = "PEPECHUY=" Then
'         xPepechuy = CByte(Mid(InputLine, 10, Len(InputLine) - 10 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 12)) = "ROUNDSTOCUT=" Then
'         xRoundsToCut = CByte(Mid(InputLine, 13, Len(InputLine) - 13 + 2))
'      ElseIf UCase(Mid$(InputLine, 1, 19)) = "UNPLAYEDUSEAVERAGE=" Then
'         xUnplayedUseAverage = IIf(Mid(InputLine, 20, 1) = "1", True, False)
'      ElseIf UCase(Mid$(InputLine, 1, 7)) = "EDITOR=" Then
'         GlobalEditorPath = Mid(InputLine, 8, Len(InputLine) - 8 + 1)
'      ElseIf UCase(Mid$(InputLine, 1, 14)) = "ACROBATREADER=" Then
'         GlobalAcrobatPath = Mid(InputLine, 15, Len(InputLine) - 15 + 1)
'      ElseIf UCase(Mid$(InputLine, 1, 7)) = "TOPPCT=" Then
'         xTopPct = CByte(Mid(InputLine, 8, Len(InputLine) - 8 + 3))
'      ElseIf UCase(Mid$(InputLine, 1, 9)) = "SHOWFILE=" Then
'         xShowFile = CByte(Mid(InputLine, 10, Len(InputLine) - 10 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 14)) = "NORGAARDFINAL=" Then
'         xNorgaardFinal = CByte(Mid(InputLine, 15, Len(InputLine) - 15 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 13)) = "SHOWWINBONUS=" Then
'         xShowWinbonus = CByte(Mid(InputLine, 14, Len(InputLine) - 14 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 19)) = "OVERVIEWBEFORELAST=" Then
'         xOverviewBeforeLast = CByte(Mid(InputLine, 20, Len(InputLine) - 20 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 12)) = "BUCHSUMMARY=" Then
'         xBuchSummary = CByte(Mid(InputLine, 13, Len(InputLine) - 13 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 14)) = "BERGERSUMMARY=" Then
'         xBergerSummary = CByte(Mid(InputLine, 15, Len(InputLine) - 15 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 16)) = "NORGAARDSUMMARY=" Then
'         xNorgaardSummary = CByte(Mid(InputLine, 17, Len(InputLine) - 17 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 19)) = "PROGRESSIVESUMMARY=" Then
'         xProgressiveSummary = CByte(Mid(InputLine, 20, Len(InputLine) - 20 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 16)) = "COMBINEDSUMMARY=" Then
'         xCombinedSummary = CByte(Mid(InputLine, 17, Len(InputLine) - 17 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 14)) = "COMBINEDFINAL=" Then
'         xCombinedFinal = CByte(Mid(InputLine, 15, Len(InputLine) - 15 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 19)) = "COMBINEDBEFORELAST=" Then
'         xCombinedBeforeLast = CByte(Mid(InputLine, 20, Len(InputLine) - 20 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 20)) = "SHOWPREDICTIONSTATS=" Then
'         xShowPredictionStats = CByte(Mid(InputLine, 21, Len(InputLine) - 21 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 25)) = "SHOWDIRECTPREDICTIVENESS=" Then
'         xShowDirectPredictiveness = CByte(Mid(InputLine, 26, Len(InputLine) - 26 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 27)) = "SHOWINDIRECTPREDICTIVENESS=" Then
'         xShowIndirectPredictiveness = CByte(Mid(InputLine, 28, Len(InputLine) - 28 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 12)) = "GENERALINFO=" Then
'         xGeneralInfo = CByte(Mid(InputLine, 13, Len(InputLine) - 13 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 16)) = "SHOWSWISSROUNDS=" Then
'         xShowSwissRounds = CByte(Mid(InputLine, 17, Len(InputLine) - 17 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 20)) = "SHOWINFOAFTERRATING=" Then
'         xShowInfoAfterRating = CByte(Mid(InputLine, 21, Len(InputLine) - 21 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 19)) = "REMOVEDOUBLESPACES=" Then
'         xRemoveDoubleSpaces = CByte(Mid(InputLine, 20, Len(InputLine) - 20 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 22)) = "SHOWALLTIEBREAKVALUES=" Then
'         xShowAllTiebreakValues = CByte(Mid(InputLine, 23, Len(InputLine) - 23 + 1))
'      ElseIf UCase(Mid$(InputLine, 1, 19)) = "USETPRFORAVGRATING=" Then
'         xUseTPRforAvgRating = CByte(Mid(InputLine, 20, Len(InputLine) - 20 + 1))
'      End If
   Wend
   ' "Process all details" is too strong a feature to leave around for every new run of the program
   'xProcessAllDetails = False
   Close #1
End Sub

Sub SaveIniFile()
   Dim InputLine As String
   Dim MyName As String
   ' Read file PGNDRAW.INI for settings with <num> a number from 1 to 18, or missing
   ' Norgaard=<num>
   ' DirectEncounter=<num>
   ' MostWins=<num>
   ' MostBlacks=<num>
   ' Berger=<num>
   ' Koya=<num>
   ' AverageRating=<num>
   ' AverageRatingCut1=<num>
   ' Winbonus=<num>
   ' Buchholz=<num>
   ' BuchholzCut1=<num>
   ' BuchholzCut2=<num>
   ' MedianBuchholz1=<num>
   ' MedianBuchholz2=<num>
   ' SumBuchholz=<num>
   ' Progressive=<num>
   ' Rincewind1=<num>
   ' Norgaard2=<num>
   ' Rincewind2=<num>
   ' MissingPoints=<num>
   ' Zermelo=<num>
   ' Smerdon=<num>
   ' RoundsToCut=<num>
   ' UnplayedUseAverage=<0,1>
   ' Editor=<Wordpad,Notepad>
   ' ShowFile=<0,1>
   ' NorgaardFinal=<0,1>
   ' ShowWinbonus=<0,1>
   ' OverviewBeforeLast=<0,1>
   ' BuchSummary=<0,1>
   ' BergerSummary=<0,1>
   ' NorgaardSummary=<0,1>
   ' ProgressiveSummary=<0,1>
   ' CombinedSummary=<0,1>
   ' CombinedFinal=<0,1>
   ' CombinedBeforeLast=<0,1>
   ' ShowPredictionStats=<0,1>
   ' ShowDirectPredictiveness=<0,1>
   ' ShowIndirectPredictiveness=<0,1>
   ' GeneralInfo=<0,1>
   ' ShowSwissRounds=<0,1>
   ' ShowInfoAfterRating=<0,1>
   ' RemoveDoubleSpaces=<0,1>
   ' ShowAllTiebreakValues=<0,1>
   ' UseTPRforAvgRating=<0,1>
   ' Check if file exists before deleting it
   MyName = Dir("PGNDRAW.INI")
   If Trim$(MyName <> "") Then
      Kill "PGNDRAW.INI"
   End If
   Open "PGNDRAW.INI" For Output As #3
'   InputLine = "Norgaard=" + Trim(Str(xNorgaard))
'   Print #3, InputLine
'   InputLine = "DirectEncounter=" + Trim(Str(xDirectEncounter))
'   Print #3, InputLine
'   InputLine = "MostWins=" + Trim(Str(xMostWins))
'   Print #3, InputLine
'   InputLine = "MostBlacks=" + Trim(Str(xMostBlacks))
'   Print #3, InputLine
'   InputLine = "Berger=" + Trim(Str(xBerger))
'   Print #3, InputLine
'   InputLine = "Koya=" + Trim(Str(xKoya))
'   Print #3, InputLine
'   InputLine = "AverageRating=" + Trim(Str(xAverageRating))
'   Print #3, InputLine
'   InputLine = "AverageRatingCut1=" + Trim(Str(xAverageRatingCut1))
'   Print #3, InputLine
'   InputLine = "Winbonus=" + Trim(Str(xWinbonus))
'   Print #3, InputLine
'   InputLine = "Buchholz=" + Trim(Str(xBuchholz))
'   Print #3, InputLine
'   InputLine = "BuchholzCut1=" + Trim(Str(xBuchholzCut1))
'   Print #3, InputLine
'   InputLine = "BuchholzCut2=" + Trim(Str(xBuchholzCut2))
'   Print #3, InputLine
'   InputLine = "MedianBuchholz1=" + Trim(Str(xMedianBuchholz1))
'   Print #3, InputLine
'   InputLine = "MedianBuchholz2=" + Trim(Str(xMedianBuchholz2))
'   Print #3, InputLine
'   InputLine = "SumBuchholz=" + Trim(Str(xSumBuchholz))
'   Print #3, InputLine
'   InputLine = "Progressive=" + Trim(Str(xProgressive))
'   Print #3, InputLine
'   InputLine = "Rincewind1=" + Trim(Str(xRincewind1))
'   Print #3, InputLine
'   InputLine = "Norgaard2=" + Trim(Str(xNorgaard2))
'   Print #3, InputLine
'   'InputLine = "Rincewind2=" + Trim(Str(xRincewind2))
'   InputLine = "MissingPoints=" + Trim(Str(xMissingPoints))
'   Print #3, InputLine
'   InputLine = "Zermelo=" + Trim(Str(xZermelo))
'   Print #3, InputLine
'   InputLine = "Smerdon=" + Trim(Str(xSmerdon))
'   Print #3, InputLine
'   InputLine = "Pepechuy=" + Trim(Str(xPepechuy))
'   Print #3, InputLine
'   InputLine = "RoundsToCut=" + Trim(Str(xRoundsToCut))
'   Print #3, InputLine
'   InputLine = "TopPct=" + Trim(Str(xTopPct))
'   Print #3, InputLine
'   InputLine = "UnplayedUseAverage=" + IIf(xUnplayedUseAverage, "1", "0")
'   Print #3, InputLine
'   InputLine = "Editor=" + Trim(GlobalEditorPath)
'   Print #3, InputLine
'   InputLine = "AcrobatReader=" + Trim(GlobalAcrobatPath)
'   Print #3, InputLine
'   InputLine = "ShowFile=" + Trim(Str(xShowFile))
'   Print #3, InputLine
'   InputLine = "NorgaardFinal=" + Trim(Str(xNorgaardFinal))
'   Print #3, InputLine
'   InputLine = "ShowWinbonus=" + Trim(Str(xShowWinbonus))
'   Print #3, InputLine
'   InputLine = "OverviewBeforeLast=" + Trim(Str(xOverviewBeforeLast))
'   Print #3, InputLine
'   InputLine = "BuchSummary=" + Trim(Str(xBuchSummary))
'   Print #3, InputLine
'   InputLine = "BergerSummary=" + Trim(Str(xBergerSummary))
'   Print #3, InputLine
'   InputLine = "NorgaardSummary=" + Trim(Str(xNorgaardSummary))
'   Print #3, InputLine
'   InputLine = "ProgressiveSummary=" + Trim(Str(xProgressiveSummary))
'   Print #3, InputLine
'   InputLine = "CombinedSummary=" + Trim(Str(xCombinedSummary))
'   Print #3, InputLine
'   InputLine = "CombinedFinal=" + Trim(Str(xCombinedFinal))
'   Print #3, InputLine
'   InputLine = "CombinedBeforeLast=" + Trim(Str(xCombinedBeforeLast))
'   Print #3, InputLine
'   InputLine = "ShowPredictionStats=" + Trim(Str(xShowPredictionStats))
'   Print #3, InputLine
'   InputLine = "ShowDirectPredictiveness=" + Trim(Str(xShowDirectPredictiveness))
'   Print #3, InputLine
'   InputLine = "ShowIndirectPredictiveness=" + Trim(Str(xShowIndirectPredictiveness))
'   Print #3, InputLine
'   InputLine = "GeneralInfo=" + Trim(Str(xGeneralInfo))
'   Print #3, InputLine
'   InputLine = "ShowSwissRounds=" + Trim(Str(xShowSwissRounds))
'   Print #3, InputLine
'   InputLine = "ShowInfoAfterRating=" + Trim(Str(xShowInfoAfterRating))
'   Print #3, InputLine
'   InputLine = "RemoveDoubleSpaces=" + Trim(Str(xRemoveDoubleSpaces))
'   Print #3, InputLine
'   InputLine = "ShowAllTiebreakValues=" + Trim(Str(xShowAllTiebreakValues))
'   Print #3, InputLine
'   InputLine = "UseTPRforAvgRating=" + Trim(Str(xUseTPRforAvgRating))
   Print #3, InputLine
   Close #3
End Sub

Private Sub picHelp_Click()
   On Error GoTo FileError
   ' First terminate a previous version of REMLTOOL.TXT to regenerate properly
   ' If this is not done, the user will not receive the changes of REMLTOOL.TXT
   ' but would be looking at the old version
   If (DocxID <> 0) Then
      Call KillProcess(DocxID)
      DocxID = 0
   End If
   ' Don't wait for this process, since we want to allow the
   ' user to prepare a new analysis while still looking at the
   ' report file of the errors in the last run as another window.
   DocxID = Shell(GlobalAcrobatPath + " " + Chr(34) + App.Path + "\Norgaard Tiebreak tool User Manual.pdf" + Chr(34), vbNormalFocus)
   Screen.MousePointer = vbDefault
   Exit Sub
FileError:
   Screen.MousePointer = vbDefault
   MsgBox "Unexpected error in Showing the result file. Maybe Adobe PDF Reader wasn't found." + Chr(13)
   Call AllowUserAbort
End Sub

Private Sub txtFile_Click()
   GetFile.Show vbModal
   If Dir(FullFile) <> "" Then
      txtFile.Text = FullFile
   End If
End Sub

Function OutputFileInUse(xfilename As String) As Boolean
   OutputFileInUse = False
   On Error GoTo OutputFileError
   Open xfilename For Output As #9
   Close #9
   Exit Function
OutputFileError:
   OutputFileInUse = True
   Call AllowUserAbort
End Function

Private Sub txtRoundsToCut_Change()
   Dim x As Integer
   If IsNumeric(txtRoundsToCut) Then
      If Len(Trim(txtRoundsToCut)) < 3 Then
         'xRoundsToCut = Val(txtRoundsToCut)
      End If
   End If
End Sub

Private Sub txtTopPct_Change()
   Dim x As Integer
   If IsNumeric(txtTopPct) Then
      If Len(Trim(txtTopPct)) < 4 Then
         'xTopPct = Val(txtTopPct)
      End If
   End If
End Sub

