VERSION 5.00
Begin VB.Form GetFile 
   Caption         =   "Specify File"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   4530
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox FullFilename 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   4050
      Width           =   6375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3375
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Path and File Name:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
   End
End
Attribute VB_Name = "GetFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
   If Right(FullFilename, 1) = "\" Then
      MsgBox "Please select a file or write a file name"
      Exit Sub
   End If
   FullFile = FullFilename
   Unload Me
End Sub

Private Sub Dir1_Change()
    ' Update file list box to synchronize with the
    ' directory list box.
    File1.Path = Dir1.Path
    If Right(Dir1.Path, 1) = "\" Then
       FullFilename = Dir1.Path + File1.List(File1.ListIndex)
    Else
       FullFilename = Dir1.Path + "\" + File1.List(File1.ListIndex)
    End If
End Sub
    
Private Sub Drive1_Change()
    On Error GoTo DriveHandler
    ' If new drive was selected, the Dir1 box
    ' updates its display.
    Dir1.Path = Drive1.Drive
    If Right(Dir1.Path, 1) = "\" Then
       FullFilename = Dir1.Path + File1.List(File1.ListIndex)
    Else
       FullFilename = Dir1.Path + "\" + File1.List(File1.ListIndex)
    End If
    Exit Sub
' If there is an error, reset drvList.Drive with the
' drive from dirList.Path.
DriveHandler:
    Drive1.Drive = Dir1.Path
    Call AllowUserAbort
End Sub

Private Sub File1_Click()
    If Right(Dir1.Path, 1) = "\" Then
       FullFilename = Dir1.Path + File1.List(File1.ListIndex)
    Else
       FullFilename = Dir1.Path + "\" + File1.List(File1.ListIndex)
    End If
End Sub

Private Sub File1_Change()
    If Right(Dir1.Path, 1) = "\" Then
       FullFilename = Dir1.Path + File1.List(File1.ListIndex)
    Else
       FullFilename = Dir1.Path + "\" + File1.List(File1.ListIndex)
    End If
End Sub

Private Sub Form_Load()
   Dim x As Integer
   Dim y As Integer
   Dim ThisFile As String
   GetFile.Top = (Screen.Height - GetFile.Height) / 2
   GetFile.Left = (Screen.Width - GetFile.Width) / 2
   File1.Pattern = "*.pgn"
   ' File1.Pattern could be "*.txt;*.pgn" to combine extensions
   If FullFile <> "" Then
      Drive1.Drive = Mid(FullFile, 1, 2)
      For x = Len(FullFile) To 1 Step -1
         If Mid(FullFile, x, 1) = "\" Then
            If x = 3 Then
               Dir1.Path = Mid(FullFile, 1, x)
            Else
               Dir1.Path = Mid(FullFile, 1, x - 1)
            End If
            ThisFile = Mid(FullFile, x + 1, Len(FullFile) - x)
            For y = 0 To File1.ListCount
               If LCase(File1.List(y)) = LCase(ThisFile) Then
                  File1.ListIndex = y
                  Exit For
               End If
            Next y
            Exit For
         End If
      Next x
   End If
End Sub
