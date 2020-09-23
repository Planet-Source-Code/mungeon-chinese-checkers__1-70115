VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHighScore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chinese Checker"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   6495
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3975
      Width           =   6495
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Score"
         Height          =   340
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Default         =   -1  'True
         Height          =   340
         Left            =   5160
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         X1              =   120
         X2              =   6360
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   6360
         Y1              =   120
         Y2              =   120
      End
   End
   Begin MSComctlLib.ListView lvwScore 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rank"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "No of Step"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Date/Time"
         Object.Width           =   3881
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   6360
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "HIGH SCORE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Score(1 To 10) As ScoreList

Private Sub cmdClear_Click()
    Dim z As Integer
    Dim temp As ScoreList
    If MsgBox("Are you sure you want to clear the score?", vbYesNo + vbQuestion) = vbYes Then
        For z = 1 To 10
            Score(z).Name = ""
            Score(z).NoOfMove = 0
            Score(z).DateTime = Now
        Next z
        
        Open "Score.dat" For Output As #1
        For z = 1 To 10
            Write #1, Score(z).Name, Score(z).NoOfMove, Score(z).DateTime
        Next z
        Close #1
        ShowScore
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set activeWin = Me
    Dim z As Integer
    Dim temp As ScoreList
    Open "Score.dat" For Input As #1
    For z = 1 To 10
        Input #1, Score(z).Name, Score(z).NoOfMove, Score(z).DateTime
    Next z
    Close #1
    ShowScore
End Sub

Public Function ShowScore()
    lvwScore.ListItems.Clear
    For z = 1 To 10
        With lvwScore.ListItems.Add(, , z)
        .SubItems(1) = Score(z).Name
        .SubItems(2) = Score(z).NoOfMove
        .SubItems(3) = Score(z).DateTime
        End With
    Next z
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set activeWin = frmGame
End Sub
