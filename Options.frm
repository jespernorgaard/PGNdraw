VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkUseTPRforAvgRating 
      Caption         =   "Use TPR for Average Rating"
      Height          =   255
      Left            =   4200
      TabIndex        =   22
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CheckBox chkUseOldstyleSB 
      Caption         =   "Apply draw against oneself in SB"
      Height          =   195
      Left            =   4200
      TabIndex        =   21
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CheckBox chkProcessAllDetails 
      Caption         =   "Process all files as a detailed report"
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CheckBox chkShowAllTiebreakValues 
      Caption         =   "Show all tiebreak values"
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CheckBox chkRemoveDoubleSpaces 
      Caption         =   "Remove double spaces in data"
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CheckBox chkShowInfoAfterRating 
      Caption         =   "Show info after rating per player"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CheckBox chkCombinedSummary 
      Caption         =   "Combined Summary"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CheckBox chkShowIndirectPredictiveness 
      Caption         =   "Show indirect predictiveness"
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   960
      Width           =   3015
   End
   Begin VB.CheckBox chkShowDirectPredictiveness 
      Caption         =   "Show direct predictiveness"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   600
      Width           =   3015
   End
   Begin VB.CheckBox chkShowSwissRounds 
      Caption         =   "Show round information for each player"
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CheckBox chkShowPredictionStats 
      Caption         =   "Show prediction statistics"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   240
      Width           =   2775
   End
   Begin VB.CheckBox chkShowWinbonus 
      Caption         =   "Show winbonus opponents as ""1101"""
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   3135
   End
   Begin VB.CheckBox chkCombinedBeforeLast 
      Caption         =   "Combined standing before last round"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   3615
   End
   Begin VB.CheckBox chkGeneralInfo 
      Caption         =   "General tournament info"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CheckBox chkCombinedFinal 
      Caption         =   "Combined tiebreak Final Standing"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3480
      Width           =   3015
   End
   Begin VB.CheckBox chkProgressiveSummary 
      Caption         =   "Progressive summary"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CheckBox chkNorgaardSummary 
      Caption         =   "Norgaard summary"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CheckBox chkBergerSummary 
      Caption         =   "Berger summary"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CheckBox chkBuchSummary 
      Caption         =   "Buchholz summary"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox chkOverviewBeforeLast 
      Caption         =   "Show Overview before last round"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CheckBox chkNorgaardFinal 
      Caption         =   "Norgaard tiebreak Final Standing"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.CheckBox chkShowFile 
      Caption         =   "Show which file was used"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   5040
      Width           =   1000
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkShowFile_Click()
   If chkShowFile.Value = vbChecked Then
      xShowFile = 1
   Else
      xShowFile = 0
   End If
End Sub

Private Sub chkNorgaardFinal_Click()
   If chkNorgaardFinal.Value = vbChecked Then
      xNorgaardFinal = 1
   Else
      xNorgaardFinal = 0
   End If
End Sub

Private Sub chkShowWinbonus_Click()
   If chkShowWinbonus.Value = vbChecked Then
      xShowWinbonus = 1
   Else
      xShowWinbonus = 0
   End If
End Sub

Private Sub chkOverviewBeforeLast_Click()
   If chkOverviewBeforeLast.Value = vbChecked Then
      xOverviewBeforeLast = 1
   Else
      xOverviewBeforeLast = 0
   End If
End Sub

Private Sub chkBuchSummary_Click()
   If chkBuchSummary.Value = vbChecked Then
      xBuchSummary = 1
   Else
      xBuchSummary = 0
   End If
End Sub

Private Sub chkBergerSummary_Click()
   If chkBergerSummary.Value = vbChecked Then
      xBergerSummary = 1
   Else
      xBergerSummary = 0
   End If
End Sub

Private Sub chkNorgaardSummary_Click()
   If chkNorgaardSummary.Value = vbChecked Then
      xNorgaardSummary = 1
   Else
      xNorgaardSummary = 0
   End If
End Sub

Private Sub chkProgressiveSummary_Click()
   If chkProgressiveSummary.Value = vbChecked Then
      xProgressiveSummary = 1
   Else
      xProgressiveSummary = 0
   End If
End Sub

Private Sub chkCombinedSummary_Click()
   If chkCombinedSummary.Value = vbChecked Then
      xCombinedSummary = 1
   Else
      xCombinedSummary = 0
   End If
End Sub

Private Sub chkCombinedFinal_Click()
   If chkCombinedFinal.Value = vbChecked Then
      xCombinedFinal = 1
   Else
      xCombinedFinal = 0
   End If
End Sub

Private Sub chkCombinedBeforeLast_Click()
   If chkCombinedBeforeLast.Value = vbChecked Then
      xCombinedBeforeLast = 1
   Else
      xCombinedBeforeLast = 0
   End If
End Sub

Private Sub chkShowPredictionStats_Click()
   If chkShowPredictionStats.Value = vbChecked Then
      xShowPredictionStats = 1
   Else
      xShowPredictionStats = 0
   End If
End Sub

Private Sub chkShowDirectPredictiveness_Click()
   If chkShowDirectPredictiveness.Value = vbChecked Then
      xShowDirectPredictiveness = 1
   Else
      xShowDirectPredictiveness = 0
   End If
End Sub

Private Sub chkShowIndirectPredictiveness_Click()
   If chkShowIndirectPredictiveness.Value = vbChecked Then
      xShowIndirectPredictiveness = 1
   Else
      xShowIndirectPredictiveness = 0
   End If
End Sub

Private Sub chkGeneralInfo_Click()
   If chkGeneralInfo.Value = vbChecked Then
      xGeneralInfo = 1
   Else
      xGeneralInfo = 0
   End If
End Sub

Private Sub chkShowSwissRounds_Click()
   If chkShowSwissRounds.Value = vbChecked Then
      xShowSwissRounds = 1
   Else
      xShowSwissRounds = 0
   End If
End Sub

Private Sub chkShowInfoAfterRating_Click()
   If chkShowInfoAfterRating.Value = vbChecked Then
      xShowInfoAfterRating = 1
   Else
      xShowInfoAfterRating = 0
   End If
End Sub

Private Sub chkRemoveDoubleSpaces_Click()
   If chkRemoveDoubleSpaces.Value = vbChecked Then
      xRemoveDoubleSpaces = 1
   Else
      xRemoveDoubleSpaces = 0
   End If
End Sub

Private Sub chkShowAllTiebreakValues_Click()
   If chkShowAllTiebreakValues.Value = vbChecked Then
      xShowAllTiebreakValues = 1
   Else
      xShowAllTiebreakValues = 0
   End If
End Sub

Private Sub chkProcessAllDetails_Click()
   If chkProcessAllDetails.Value = vbChecked Then
      xProcessAllDetails = 1
   Else
      xProcessAllDetails = 0
   End If
End Sub

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   ' Center form
   frmOptions.Top = (Screen.Height - frmOptions.Height) / 2
   frmOptions.Left = (Screen.Width - frmOptions.Width) / 2
   If xShowFile Then chkShowFile.Value = vbChecked
   If xNorgaardFinal Then chkNorgaardFinal.Value = vbChecked
   If xShowWinbonus Then chkShowWinbonus.Value = vbChecked
   If xOverviewBeforeLast Then chkOverviewBeforeLast.Value = vbChecked
   If xBuchSummary Then chkBuchSummary.Value = vbChecked
   If xBergerSummary Then chkBergerSummary.Value = vbChecked
   If xNorgaardSummary Then chkNorgaardSummary.Value = vbChecked
   If xProgressiveSummary Then chkProgressiveSummary.Value = vbChecked
   If xCombinedSummary Then chkCombinedSummary.Value = vbChecked
   If xCombinedFinal Then chkCombinedFinal.Value = vbChecked
   If xGeneralInfo Then chkGeneralInfo.Value = vbChecked
   If xCombinedBeforeLast Then chkCombinedBeforeLast.Value = vbChecked
   If xShowPredictionStats Then chkShowPredictionStats.Value = vbChecked
   If xShowDirectPredictiveness Then chkShowDirectPredictiveness.Value = vbChecked
   If xShowIndirectPredictiveness Then chkShowIndirectPredictiveness.Value = vbChecked
   If xShowSwissRounds Then chkShowSwissRounds.Value = vbChecked
   If xShowInfoAfterRating Then chkShowInfoAfterRating.Value = vbChecked
   If xRemoveDoubleSpaces Then chkRemoveDoubleSpaces.Value = vbChecked
   If xShowAllTiebreakValues Then chkShowAllTiebreakValues.Value = vbChecked
   If xProcessAllDetails Then chkProcessAllDetails.Value = vbChecked
   If xUseOldstyleSB Then chkUseOldstyleSB.Value = vbChecked
   If xUseTPRforAvgRating Then chkUseTPRforAvgRating.Value = vbChecked
End Sub
