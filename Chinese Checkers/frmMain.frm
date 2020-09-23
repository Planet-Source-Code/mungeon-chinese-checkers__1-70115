VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chinese Checker"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboNumPlayer 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   1680
      List            =   "frmMain.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Index           =   6
      Left            =   960
      TabIndex        =   16
      Top             =   3360
      Width           =   2055
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Index           =   5
      Left            =   960
      TabIndex        =   13
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Index           =   4
      Left            =   960
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Index           =   3
      Left            =   960
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   340
      Left            =   5520
      TabIndex        =   20
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Game"
      Default         =   -1  'True
      Height          =   340
      Left            =   4320
      TabIndex        =   19
      Top             =   3960
      Width           =   1095
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Index           =   6
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Index           =   5
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Index           =   4
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Index           =   3
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Index           =   2
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox cboColor 
      Height          =   315
      Index           =   1
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Index           =   6
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3360
      Width           =   1815
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Index           =   5
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2880
      Width           =   1815
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Index           =   4
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Index           =   3
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Index           =   2
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Index           =   1
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Number of Player: "
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6720
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Color"
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
      Left            =   5280
      TabIndex        =   29
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Players"
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
      Left            =   3240
      TabIndex        =   28
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
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
      Left            =   960
      TabIndex        =   27
      Top             =   720
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   6720
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   6720
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label lblPlayer 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 6:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   26
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblPlayer 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 5:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   25
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblPlayer 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 4:"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblPlayer 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 3:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   23
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblPlayer 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   22
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblPlayer 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   960
      Width           =   1455
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   6720
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer
Dim corlorArr(1 To 6) As Integer
Dim temp As Integer

Private Sub cboColor_Click(Index As Integer)
    Dim i As Integer
    For i = cboColor.LBound To cboColor.UBound
        If cboColor(i).ListIndex >= 0 Then
            If i <> Index And cboColor(Index).ListIndex = cboColor(i).ListIndex Then
                cboColor(i).ListIndex = corlorArr(Index)
                temp = corlorArr(Index)
                corlorArr(Index) = corlorArr(i)
                corlorArr(i) = temp
                Exit Sub
            End If
        End If
    Next i
End Sub

Private Sub cboNumPlayer_Click()
    Dim i As Integer
    For i = 1 To 6
        showPlayerSet i, False
    Next
    For i = 1 To cboNumPlayer
        showPlayerSet i, True
    Next
End Sub

Public Function showPlayerSet(p As Integer, val As Boolean)
    lblPlayer(p).Visible = val
    cboName(p).Visible = val
    cboStatus(p).Visible = val
    cboColor(p).Visible = val
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
    Dim p As Integer
    Dim tempColor As ColorConstants
    p = 1
    For c = cboStatus.LBound To val(cboNumPlayer.Text)
        If cboStatus(c).Text = "User" Then
            Group(c).Auto = False
        Else: Group(c).Auto = True
        End If
        Group(c).Number = c
        Group(c).Name = cboName(c).Text
        Group(c).Status = "Playing"
        Group(c).NoOfMove = 0
        Select Case cboColor(c).Text
            Case "Blue"
                    tempColor = vbBlue
            Case "Cyan"
                    tempColor = vbCyan
            Case "Green"
                    tempColor = vbGreen
            Case "Magenta"
                    tempColor = vbMagenta
            Case "Red"
                    tempColor = vbRed
            Case "Yellow"
                    tempColor = vbYellow
        End Select
        Group(c).Color = tempColor
    Next c
    p = val(cboNumPlayer.Text) + 1
    Do While p <= cboStatus.UBound
        Group(p).Status = "None"
        Group(p).Color = vbBlack
        p = p + 1
    Loop
    Unload Me
    frmGame.Enabled = True
    frmGame.NewGame
    frmGame.SetFocus
End Sub

Private Sub Form_Load()
    Set activeWin = Me
    Dim z As Integer
    Dim temp As String
    Open "Name.dat" For Input As #1
    For z = 1 To 6
        Input #1, temp
        cboName(z).AddItem temp
        cboName(z).ListIndex = 0
    Next z
    Close #1
    For c = cboStatus.LBound To cboStatus.UBound
        If c <= 2 Then cboStatus(c).AddItem "User"
        cboStatus(c).AddItem "Computer"
        cboStatus(c).ListIndex = 0
    Next c
    For c = cboColor.LBound To cboColor.UBound
        cboColor(c).AddItem "Blue"
        cboColor(c).AddItem "Cyan"
        cboColor(c).AddItem "Green"
        cboColor(c).AddItem "Magenta"
        cboColor(c).AddItem "Red"
        cboColor(c).AddItem "Yellow"
        cboColor(c).ListIndex = c - 1
        corlorArr(c) = c - 1
    Next c
    cboNumPlayer.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set activeWin = frmGame
End Sub
