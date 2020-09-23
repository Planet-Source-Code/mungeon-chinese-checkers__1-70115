VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   Caption         =   "Chinese Checker"
   ClientHeight    =   7830
   ClientLeft      =   75
   ClientTop       =   660
   ClientWidth     =   10710
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMusic 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   1560
      Top             =   6960
   End
   Begin MSComDlg.CommonDialog cdgBGColor 
      Left            =   720
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdgSave 
      Left            =   120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrTurnBlink 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1080
      Top             =   6960
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   6960
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrTurn 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   6960
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblPlayerStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Marble 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   7440
      Width           =   375
   End
   Begin VB.Image imgTurnBlink 
      Height          =   375
      Index           =   0
      Left            =   1080
      Top             =   7440
      Width           =   375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuLine0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScore 
         Caption         =   "&High Score"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "&Setting"
      Begin VB.Menu mnuBGColor 
         Caption         =   "&Background Color"
      End
      Begin VB.Menu mnuFontColor 
         Caption         =   "&Font Color"
      End
      Begin VB.Menu mnuMusic 
         Caption         =   "&Music"
         Checked         =   -1  'True
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSpeed 
         Caption         =   "&Speed"
         Begin VB.Menu mnuSpeedValue 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu mnuSpeedValue 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu mnuSpeedValue 
            Caption         =   "3"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu mnuSpeedValue 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu mnuSpeedValue 
            Caption         =   "5"
            Index           =   5
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContent 
         Caption         =   "&Content"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const BoardWidth = 20
Const BoardHeight = 19
Const BoardMargin = 450
Const BgColor = &H8000&
Dim BoardGrid(BoardWidth, BoardHeight) As Integer
Dim AvailableMove(17) As Integer
Dim CurrentPlayer As Integer
Dim CircleRadius As Integer
Dim MarbleSelected As Boolean
Dim MarbleSelectednum As Integer
Dim GameStarted As Boolean
Dim GameSpeed As Integer
Dim Score(1 To 10) As ScoreList
Dim PlayerPos(60) As PlayerPosition
Dim WinPos(60) As PlayerPosition
Dim MoveCount As Integer
Dim CStep(100) As CompStep
Dim StepFound As Integer
Dim MoveTemp(20) As MovingStep
Dim StepTmp As Integer
Dim tmpMoveW As Integer
Dim tmpMoveH As Integer
Dim TempFromX As Integer
Dim TempFromY As Integer
Dim TempMDownX As Integer
Dim TempMDownY As Integer
Dim TempTurnBool As Boolean
Dim TempBlinkBool As Boolean
Dim w As Integer
Dim h As Integer
Dim p As Integer
Dim i As Integer
Dim j As Integer
Dim c As Integer

Private Sub Form_Activate()
    tmrTurn.Enabled = TempTurnBool
    tmrTurnBlink.Enabled = TempBlinkBool
    On Error Resume Next
    activeWin.Show
    activeWin.SetFocus
End Sub

Private Sub Form_Deactivate()
    TempTurnBool = tmrTurn.Enabled
    TempBlinkBool = tmrTurnBlink.Enabled
    tmrTurn.Enabled = False
    tmrTurnBlink.Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 49 And KeyAscii <= 53 Then
        mnuSpeedValue_Click (KeyAscii - 48)
    End If
End Sub

Private Sub Form_Load()
    Set activeWin = frmGame
    Me.Caption = App.Title
    Me.ClipControls = True
    Me.KeyPreview = True
    MusicBool = True
    If MusicBool Then
        mnuMusic.Checked = True
    Else
        mnuMusic.Checked = True
    End If
    PlayMidi
    tmrMusic.Enabled = True
    AvailableMove(0) = 1:    AvailableMove(1) = 2
    AvailableMove(2) = 3:    AvailableMove(3) = 4
    AvailableMove(4) = 13:   AvailableMove(5) = 12
    AvailableMove(6) = 11:   AvailableMove(7) = 10
    AvailableMove(8) = 9:    AvailableMove(9) = 10
    AvailableMove(10) = 11:  AvailableMove(11) = 12
    AvailableMove(12) = 13:  AvailableMove(13) = 4
    AvailableMove(14) = 3:    AvailableMove(15) = 2
    AvailableMove(16) = 1
    
    Me.BackColor = BgColor
    GameStarted = False
    mnuSave.Enabled = False
    CurrentPlayer = -1
    MarbleSelected = False
    tmrTurn.Enabled = False
    tmrTurnBlink.Enabled = False
    TempTurnBool = tmrTurn.Enabled
    TempBlinkBool = tmrTurnBlink.Enabled
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.WindowState = 2
    For p = 1 To 60
        Load Marble(p)
    Next
    For p = 1 To 6
        Load imgTurnBlink(p)
        Load lblPlayerStatus(p)
    Next
    mnuSpeedValue_Click (3)
    mnuNew_Click
End Sub

Public Function setDefaultPos()
    Dim CenterPos As Integer
    CenterPos = ((BoardWidth) / 2)
    If BoardWidth Mod 2 = 0 Then
        CenterPos = CenterPos - 1
    End If

    setPlayerPos 1, CenterPos, 16
    setPlayerPos 2, CenterPos - 1, 15
    setPlayerPos 3, CenterPos, 15
    setPlayerPos 4, CenterPos - 1, 14
    setPlayerPos 5, CenterPos, 14
    setPlayerPos 6, CenterPos + 1, 14
    setPlayerPos 7, CenterPos - 2, 13
    setPlayerPos 8, CenterPos - 1, 13
    setPlayerPos 9, CenterPos, 13
    setPlayerPos 10, CenterPos + 1, 13
    
    setPlayerPos 11, CenterPos - 6, 12
    setPlayerPos 12, CenterPos - 6, 11
    setPlayerPos 13, CenterPos - 5, 12
    setPlayerPos 14, CenterPos - 5, 10
    setPlayerPos 15, CenterPos - 5, 11
    setPlayerPos 16, CenterPos - 4, 12
    setPlayerPos 17, CenterPos - 5, 9
    setPlayerPos 18, CenterPos - 4, 10
    setPlayerPos 19, CenterPos - 4, 11
    setPlayerPos 20, CenterPos - 3, 12

    setPlayerPos 21, CenterPos - 6, 4
    setPlayerPos 22, CenterPos - 5, 4
    setPlayerPos 23, CenterPos - 6, 5
    setPlayerPos 24, CenterPos - 4, 4
    setPlayerPos 25, CenterPos - 5, 5
    setPlayerPos 26, CenterPos - 5, 6
    setPlayerPos 27, CenterPos - 3, 4
    setPlayerPos 28, CenterPos - 4, 5
    setPlayerPos 29, CenterPos - 4, 6
    setPlayerPos 30, CenterPos - 5, 7
    
    setPlayerPos 31, CenterPos, 0
    setPlayerPos 32, CenterPos - 1, 1
    setPlayerPos 33, CenterPos, 1
    setPlayerPos 34, CenterPos - 1, 2
    setPlayerPos 35, CenterPos, 2
    setPlayerPos 36, CenterPos + 1, 2
    setPlayerPos 37, CenterPos - 2, 3
    setPlayerPos 38, CenterPos - 1, 3
    setPlayerPos 39, CenterPos, 3
    setPlayerPos 40, CenterPos + 1, 3
    
    setPlayerPos 41, CenterPos + 6, 4
    setPlayerPos 42, CenterPos + 5, 4
    setPlayerPos 43, CenterPos + 5, 5
    setPlayerPos 44, CenterPos + 4, 4
    setPlayerPos 45, CenterPos + 4, 5
    setPlayerPos 46, CenterPos + 5, 6
    setPlayerPos 47, CenterPos + 3, 4
    setPlayerPos 48, CenterPos + 3, 5
    setPlayerPos 49, CenterPos + 4, 6
    setPlayerPos 50, CenterPos + 4, 7
    
    setPlayerPos 51, CenterPos + 6, 12
    setPlayerPos 52, CenterPos + 5, 11
    setPlayerPos 53, CenterPos + 5, 12
    setPlayerPos 54, CenterPos + 5, 10
    setPlayerPos 55, CenterPos + 4, 11
    setPlayerPos 56, CenterPos + 4, 12
    setPlayerPos 57, CenterPos + 4, 9
    setPlayerPos 58, CenterPos + 4, 10
    setPlayerPos 59, CenterPos + 3, 11
    setPlayerPos 60, CenterPos + 3, 12
End Function

Public Function setPlayerPos(count1 As Integer, x1 As Integer, y1 As Integer)
    PlayerPos(count1).X = x1
    PlayerPos(count1).Y = y1
    WinPos(count1).X = x1
    WinPos(count1).Y = y1
End Function

Public Function NewGame()
    Dim p As Integer
    Dim NumPlayer As Integer
    Dim CircleCount As Integer
    setDefaultPos
    CircleCount = 1
    For h = 0 To BoardHeight - 1
        For w = 0 To BoardWidth - 1
            If CheckMove(w, h) Then
                BoardGrid(w, h) = 0
            Else
                BoardGrid(w, h) = -1
            End If
            CircleCount = CircleCount + 1
        Next w
    Next h
    NumPlayer = 0
    For p = 1 To 6
        If Group(p).Status = "Playing" Then
            NumPlayer = NumPlayer + 1
        End If
    Next
    Select Case NumPlayer
        Case 2
            SetPlayerStartPos 4, 2
            SetPlayerStartPos 1, 1
        Case 3
            SetPlayerStartPos 5, 3
            SetPlayerStartPos 3, 2
            SetPlayerStartPos 1, 1
        Case 4
            SetPlayerStartPos 5, 4
            SetPlayerStartPos 4, 3
            SetPlayerStartPos 2, 2
            SetPlayerStartPos 1, 1
        Case 6
            SetPlayerStartPos 6, 6
            SetPlayerStartPos 5, 5
            SetPlayerStartPos 4, 4
            SetPlayerStartPos 3, 3
            SetPlayerStartPos 2, 2
            SetPlayerStartPos 1, 1
    End Select
    GameStarted = True
    DrawBoard
    mnuSave.Enabled = True
    CurrentPlayer = 1
    MarbleSelected = False
    tmrTurn.Enabled = True
    tmrTurnBlink.Enabled = True
    TempTurnBool = tmrTurn.Enabled
    TempBlinkBool = tmrTurnBlink.Enabled
    For p = 1 To 60
        LoadMarble (p)
    Next
    For p = 1 To 6
        LoadTurnBlink (p)
    Next
    ShowStatus
End Function

Public Function ShowStatus()
    For p = 1 To 6
        If Group(p).Status = "Playing" Or Group(p).Status = "Won" Then
            lblPlayerStatus(p).Caption = Group(p).Name & vbCrLf & Group(p).Status & " - " & Group(p).NoOfMove & " move(s)"
        Else
            lblPlayerStatus(p).Visible = False
        End If
    Next
End Function

Public Function LoadTurnBlink(Num As Integer)
    Dim colorName As String
    Dim PointX As Integer
    Dim PointY As Integer
    Dim StatusX As Long
    Dim StatusY As Long
    Dim SpaceLen As Integer
    SpaceLen = 100
    If Num >= 1 And Num <= 6 Then
        lblPlayerStatus(Num).Width = 2415
        lblPlayerStatus(Num).Height = 735
        Select Case Num
            Case 1
                    PointX = WinPos(Num * 10 - 9).X - 2
                    PointY = WinPos(Num * 10 - 9).Y - 1
                    StatusX = -SpaceLen - lblPlayerStatus(Num).Width
                    StatusY = 0
            Case 2
                    PointX = WinPos(Num * 10 - 9).X
                    PointY = WinPos(Num * 10 - 9).Y - 2
                    StatusX = -lblPlayerStatus(Num).Width + CircleRadius
                    StatusY = -SpaceLen - lblPlayerStatus(Num).Height
            Case 3
                    PointX = WinPos(Num * 10 - 9).X + 1
                    PointY = WinPos(Num * 10 - 9).Y - 1
                    StatusX = -(lblPlayerStatus(Num).Width / 2) + CircleRadius
                    StatusY = -SpaceLen - lblPlayerStatus(Num).Height
            Case 4
                    PointX = WinPos(Num * 10 - 9).X + 1
                    PointY = WinPos(Num * 10 - 9).Y + 1
                    StatusX = SpaceLen + (CircleRadius * 2)
                    StatusY = 0
            Case 5
                    PointX = WinPos(Num * 10 - 9).X
                    PointY = WinPos(Num * 10 - 9).Y + 2
                    StatusX = (CircleRadius * 2) - CircleRadius
                    StatusY = SpaceLen + lblPlayerStatus(Num).Height
            Case 6
                    PointX = WinPos(Num * 10 - 9).X - 2
                    PointY = WinPos(Num * 10 - 9).Y + 1
                    StatusX = -(lblPlayerStatus(Num).Width / 2) + CircleRadius
                    StatusY = SpaceLen + lblPlayerStatus(Num).Height
        End Select
        If PointY Mod 2 = 0 Then
            CurrentX = (PointX * (Me.Width - 140) / BoardWidth) + CircleRadius
            CurrentY = (PointY * (Me.Height - 650) / BoardHeight) + CircleRadius
        Else
            CurrentX = (PointX * (Me.Width - 140) / BoardWidth) + CircleRadius + (CircleRadius * 2)
            CurrentY = (PointY * (Me.Height - 650) / BoardHeight) + CircleRadius
        End If
        If Group(Num).Status = "Playing" Or Group(Num).Status = "Won" Then
            Select Case Group(Num).Color
                Case vbBlue
                    colorName = "Blue"
                Case vbCyan
                    colorName = "Cyan"
                Case vbGreen
                    colorName = "Green"
                Case vbMagenta
                    colorName = "Magenta"
                Case vbRed
                    colorName = "Red"
                Case vbYellow
                    colorName = "Yellow"
            End Select
            imgTurnBlink(Num).Stretch = True
            imgTurnBlink(Num).Picture = LoadPicture(App.Path & "\" & colorName & ".gif")
            imgTurnBlink(Num).Width = CircleRadius * 2
            imgTurnBlink(Num).Height = CircleRadius * 2
            imgTurnBlink(Num).Visible = True
            imgTurnBlink(Num).Move CurrentX + BoardMargin - CircleRadius + 10, CurrentY + BoardMargin - CircleRadius + 10
            
            lblPlayerStatus(Num).Left = StatusX + CurrentX + BoardMargin - CircleRadius + 10
            lblPlayerStatus(Num).Top = StatusY + CurrentY + BoardMargin - CircleRadius + 10
            lblPlayerStatus(Num).Visible = True
        Else
            imgTurnBlink(Num).Visible = False
        End If
    End If
End Function

Public Function SetPlayerStartPos(p As Integer, g As Integer)
    Dim Count As Integer
    If p <> g Then
        Group(p).Auto = Group(g).Auto
        Group(p).Name = Group(g).Name
        Group(p).Number = Group(g).Number
        Group(p).Status = Group(g).Status
        Group(p).NoOfMove = Group(g).NoOfMove
        Group(p).Color = Group(g).Color
        ClearGroup (g)
    End If
    For Count = p * 10 - 9 To p * 10
        BoardGrid(PlayerPos(Count).X, PlayerPos(Count).Y) = p
    Next
End Function

Public Function CheckMove(w As Integer, h As Integer) As Boolean
    Dim LeftPoint As Integer
    If h < 17 And h >= 0 Then
        LeftPoint = Int(((BoardWidth) / 2) - AvailableMove(h) / 2)
        If BoardWidth Mod 2 = 1 And h Mod 2 = 1 Then LeftPoint = LeftPoint + 1
        If h Mod 2 = 1 Then LeftPoint = LeftPoint - 1
        For j = LeftPoint To LeftPoint + AvailableMove(h) - 1
            If j = w Then
                CheckMove = True
                Exit Function
            End If
        Next j
    End If
    CheckMove = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    CloseMidi
    End
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuBGColor_Click()
    With cdgBGColor
        ' Prevent display of the custom color section
        ' of the dialog.
        .flags = cdlCCPreventFullOpen Or cdlCCRGBInit
        .Color = Me.BackColor
            .CancelError = False
            .ShowColor
            Me.BackColor = .Color
    End With
    DrawBoard
End Sub

Private Sub mnuContent_Click()
    frmHelp.Show
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuFontColor_Click()
    With cdgBGColor
        ' Prevent display of the custom color section
        ' of the dialog.
        .flags = cdlCCPreventFullOpen Or cdlCCRGBInit
        .Color = lblPlayerStatus(0).ForeColor
            .CancelError = False
            .ShowColor
            For i = 0 To 6
                lblPlayerStatus(i).ForeColor = .Color
            Next
    End With
End Sub

Private Sub mnuLoad_Click()
    Dim p As Integer
    On Error GoTo err
    cdgSave.FileName = ""
    cdgSave.DialogTitle = "Load Game..."
    cdgSave.Filter = "Chinese Checker (*.ccs) | *.ccs"
    cdgSave.ShowOpen
    If cdgSave.FileName <> "" Then
        setDefaultPos
        MarbleSelected = False
        tmrTurn.Enabled = True
        tmrTurnBlink.Enabled = True
        TempTurnBool = tmrTurn.Enabled
        TempBlinkBool = tmrTurnBlink.Enabled
    
        Open cdgSave.FileName For Input As #1
        Input #1, CurrentPlayer
        Input #1, GameStarted
        Input #1, GameSpeed
        For p = 1 To 6
            Input #1, Group(p).Auto, Group(p).Color, Group(p).Name, Group(p).Number, Group(p).NoOfMove, Group(p).Status
        Next p
        For h = 0 To BoardHeight - 1
            For w = 0 To BoardWidth - 1
                Input #1, BoardGrid(w, h)
            Next w
        Next h
        For p = 1 To 60
            Input #1, PlayerPos(p).X, PlayerPos(p).Y
            Input #1, WinPos(p).X, WinPos(p).Y
        Next p
        Close #1
        DrawBoard
        For p = 1 To 60
            LoadMarble (p)
        Next
        For p = 1 To 6
            LoadTurnBlink (p)
        Next
        ShowStatus
        If GameStarted Then
            mnuSave.Enabled = True
        End If
    End If
    Exit Sub
err:
    MsgBox "Error loading game..." & err.Description & ". ", vbExclamation
    Close #1
End Sub

Private Sub mnuMusic_Click()
    If mnuMusic.Checked Then
        mnuMusic.Checked = False
    Else
        mnuMusic.Checked = True
    End If
    MusicBool = mnuMusic.Checked
    If Not MusicBool Then
        CloseMidi
    Else
        PlayMidi
    End If
End Sub

Private Sub mnuNew_Click()
    frmMain.Show
End Sub

Private Sub mnuSave_Click()
    Dim p As Integer
    On Error GoTo err
    cdgSave.FileName = ""
    cdgSave.DialogTitle = "Save Game..."
    cdgSave.Filter = "Chinese Checker (*.ccs) | *.ccs"
    cdgSave.ShowSave
    If cdgSave.FileName <> "" Then
        If Right$(cdgSave.FileName, 4) <> ".ccs" Then
            cdgSave.FileName = cdgSave.FileName & ".ccs"
        End If
        Open cdgSave.FileName For Output As #1
        Write #1, CurrentPlayer
        Write #1, GameStarted
        Write #1, GameSpeed
        For p = 1 To 6
            Write #1, Group(p).Auto, Group(p).Color, Group(p).Name, Group(p).Number, Group(p).NoOfMove, Group(p).Status
        Next p
        For h = 0 To BoardHeight - 1
            For w = 0 To BoardWidth - 1
                Write #1, BoardGrid(w, h)
            Next w
        Next h
        For p = 1 To 60
            Write #1, PlayerPos(p).X, PlayerPos(p).Y
            Write #1, WinPos(p).X, WinPos(p).Y
        Next p
        Close #1
        MsgBox "Game saved...", vbInformation
    End If
    Exit Sub
err:
    MsgBox "Save game failed...." & err.Description & ". ", vbExclamation
    Close #1
End Sub

Private Sub mnuScore_Click()
    frmHighScore.Show
End Sub

Private Sub mnuSpeedValue_Click(Index As Integer)
    Dim s As Integer
    For s = mnuSpeedValue.LBound To mnuSpeedValue.UBound
        mnuSpeedValue(s).Checked = False
    Next s
    mnuSpeedValue(Index).Checked = True
    s = 0
    Do
        s = s + 1
    Loop Until mnuSpeedValue(s).Checked Or s = mnuSpeedValue.UBound
    GameSpeed = 11 - (Index * 2)
    tmrTurn.Interval = 200 * GameSpeed
    tmrMove.Interval = 100 * GameSpeed
End Sub

Private Sub tmrMusic_Timer()
    On Error Resume Next
    If MusicBool Then
        If v_dmss.GetSeek >= v_dms.GetLength Then
            Set v_dmss = v_dmp.PlaySegment(v_dms, 0, 0)
        End If
    End If
End Sub

'########################################################################

Private Sub tmrTurn_Timer()
    'On Error GoTo ExitSub
    If GameStarted And Group(CurrentPlayer).Auto Then
        ComputerAI
    End If
ExitSub:
End Sub

Public Function ComputerAI()
    Dim RandNum As Integer
    Dim StepNum As Integer
    Dim Dest As Integer
    Dim TargetPos As Integer
    StepFound = 0
    List1.Clear
    tmrTurn.Enabled = False
    TargetPos = CurrentPlayer + 3
    If TargetPos > 6 Then TargetPos = TargetPos - 6
    For i = CurrentPlayer * 10 - 9 To CurrentPlayer * 10
        CompFindMove i, PlayerPos(i).X, PlayerPos(i).Y, TargetPos * 10 - 9, -1, -1, 0
        CompFindNearMove i, PlayerPos(i).X, PlayerPos(i).Y, TargetPos * 10 - 9
    Next i
    StepNum = 0
    If StepFound = 0 Then
        For i = CurrentPlayer * 10 - 9 To CurrentPlayer * 10
            CompFindAltMove i, PlayerPos(i).X, PlayerPos(i).Y, TargetPos * 10 - 9, -1, -1, 0
            CompFindNearAltMove i, PlayerPos(i).X, PlayerPos(i).Y, TargetPos * 10 - 9
        Next i
    End If
    If StepFound > 0 Then
        StepNum = 0
        For c = 1 To StepFound
            If CStep(c).Dist > CStep(StepNum).Dist Then
                StepNum = c
            End If
        Next
        If CStep(StepNum).Dist < 4 Then
            Randomize
            RandNum = Int(Rnd * 100) Mod StepFound
            RandNum = RandNum + 1
            For c = 1 To RandNum
                If CStep(c).Dist >= StepNum Then
                    StepNum = c
                End If
            Next
            StepNum = RandNum
        End If
        MoveMarble CStep(StepNum).Num, CStep(StepNum).DestX, CStep(StepNum).DestY
        LoadMarble (CStep(StepNum).Num)
    Else
        tmrTurn.Enabled = True
    End If
End Function

Public Function CompFindNearAltMove(Index As Integer, x1 As Integer, y1 As Integer, TargetDest As Integer)
    Dim DestX As Integer
    Dim DestY As Integer
        Select Case TargetDest
        Case 1 To 10, 31 To 40
            DestX = x1 - 1
            DestY = y1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
            
            DestX = x1 + 1
            DestY = y1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
        Case 11 To 20, 41 To 50
            DestX = x1
            If y1 Mod 2 = 0 Then DestX = x1 - 1
            DestY = y1 - 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
            
            DestX = x1
            If y1 Mod 2 = 1 Then DestX = x1 + 1
            DestY = y1 + 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
        Case 21 To 30, 51 To 60
            DestX = x1
            If y1 Mod 2 = 1 Then DestX = x1 + 1
            DestY = y1 - 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
            DestX = x1
            If y1 Mod 2 = 0 Then DestX = x1 - 1
            DestY = y1 + 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
            
    End Select
End Function

Public Function CompFindNearMove(Index As Integer, x1 As Integer, y1 As Integer, TargetDest As Integer)
    Dim DestX As Integer
    Dim DestY As Integer
        Select Case TargetDest
        Case 1 To 10
            DestX = x1
            If y1 Mod 2 = 0 Then DestX = x1 - 1
            DestY = y1 + 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1

            DestX = x1
            If y1 Mod 2 = 1 Then DestX = x1 + 1
            DestY = y1 + 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
        Case 11 To 20
            DestX = x1 - 1
            DestY = y1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
            
            DestX = x1
            If y1 Mod 2 = 0 Then DestX = x1 - 1
            DestY = y1 + 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
        Case 21 To 30
            DestX = x1
            If y1 Mod 2 = 0 Then DestX = x1 - 1
            DestY = y1 - 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
            
            DestX = x1 - 1
            DestY = y1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
        Case 31 To 40
            DestX = x1
            If y1 Mod 2 = 1 Then DestX = x1 + 1
            DestY = y1 - 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
            
            DestX = x1
            If y1 Mod 2 = 0 Then DestX = x1 - 1
            DestY = y1 - 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
        Case 41 To 50
            DestX = x1 + 1
            DestY = y1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1

            DestX = x1
            If y1 Mod 2 = 1 Then DestX = x1 + 1
            DestY = y1 - 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
        Case 51 To 60
            DestX = x1
            If y1 Mod 2 = 1 Then DestX = x1 + 1
            DestY = y1 + 1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
            
            DestX = x1 + 1
            DestY = y1
            If ValidateNearMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then StoreStep Index, DestX, DestY, 1
    End Select
End Function

Public Function StoreStep(MarbleNum As Integer, x1 As Integer, y1 As Integer, cs As Integer)
    StepFound = StepFound + 1
    CStep(StepFound).Num = MarbleNum
    CStep(StepFound).DestX = x1
    CStep(StepFound).DestY = y1
    CStep(StepFound).Dist = cs
    List1.AddItem cs & " " & x1 & " " & y1
End Function

Public Function CheckValidMove(x1 As Integer, y1 As Integer, p1 As Integer) As Boolean
    Dim p As Integer, i As Integer
    Dim p2 As Integer
    Dim boolPlay As Boolean
    Dim val As Boolean
    p2 = p1 + 3
    If p2 > 6 Then p2 = p2 - 6
    val = True
    For p = 1 To 6
        If Not (p = p1 Or p = p2) Then
            boolPlay = False
            If Group(p).Status = "Playing" Or Group(p).Status = "Won" Then boolPlay = True
            For i = p * 10 - 9 To p * 10
                If boolPlay Then
                    If WinPos(i).X = x1 And WinPos(i).Y = y1 Then
                        val = False
                    End If
                Else
                    If PlayerPos(i).X = x1 And PlayerPos(i).Y = y1 Then
                        val = False
                    End If
                End If
            Next i
        End If
    Next p
    CheckValidMove = val
End Function

Public Function CompFindAltMove(Index As Integer, x1 As Integer, y1 As Integer, TargetDest As Integer, PreX As Integer, PreY As Integer, NumLoop As Integer)
    Dim DestX As Integer
    Dim DestY As Integer
    Dim bool As Boolean
    bool = False
    Select Case TargetDest
        Case 1 To 10, 31 To 40
            DestX = x1 - 2
            DestY = y1
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
            DestX = x1 + 2
            DestY = y1
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
        Case 11 To 20, 41 To 50
            DestX = x1 - 1
            DestY = y1 - 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
            DestX = x1 + 1
            DestY = y1 + 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
        Case 21 To 30, 51 To 60
            DestX = x1 + 1
            DestY = y1 - 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
            DestX = x1 - 1
            DestY = y1 + 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
    End Select
    If Not bool Then
        If CStep(StepFound).Dist <= NumLoop And NumLoop <> 0 Then
            StepFound = StepFound + 1
            CStep(StepFound).Num = Index
            CStep(StepFound).DestX = x1
            CStep(StepFound).DestY = y1
            CStep(StepFound).Dist = NumLoop
            List1.AddItem NumLoop & " " & x1 & " " & y1 & " " & PlayerPos(Index).X & " " & PlayerPos(Index).Y
        End If
    End If
End Function

Public Function CompFindMove(Index As Integer, x1 As Integer, y1 As Integer, TargetDest As Integer, PreX As Integer, PreY As Integer, NumLoop As Integer)
    Dim DestX As Integer
    Dim DestY As Integer
    Dim bool As Boolean
    bool = False
    Select Case TargetDest
        Case 1 To 10
            DestX = x1 - 1
            DestY = y1 + 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
            DestX = x1 + 1
            DestY = y1 + 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
        Case 11 To 20
            DestX = x1 - 2
            DestY = y1
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
            DestX = x1 - 1
            DestY = y1 + 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
        Case 21 To 30
            DestX = x1 - 2
            DestY = y1
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
            DestX = x1 - 1
            DestY = y1 - 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
        Case 31 To 40
            DestX = x1 - 1
            DestY = y1 - 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
            DestX = x1 + 1
            DestY = y1 - 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
        Case 41 To 50
            DestX = x1 + 1
            DestY = y1 - 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
            DestX = x1 + 2
            DestY = y1
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
        Case 51 To 60
            DestX = x1 + 2
            DestY = y1
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
            DestX = x1 + 1
            DestY = y1 + 2
            bool = CompFindMoveBool(Index, DestX, DestY, x1, y1, TargetDest, PreX, PreY, NumLoop)
    End Select
    If Not bool Then
        If CStep(StepFound).Dist <= NumLoop And NumLoop <> 0 Then
            StepFound = StepFound + 1
            CStep(StepFound).Num = Index
            CStep(StepFound).DestX = x1
            CStep(StepFound).DestY = y1
            CStep(StepFound).Dist = NumLoop
            List1.AddItem NumLoop & " " & x1 & " " & y1 & " " & PlayerPos(Index).X & " " & PlayerPos(Index).Y
        End If
    End If
End Function

Public Function CompFindMoveBool(Index As Integer, DestX As Integer, DestY As Integer, x1 As Integer, y1 As Integer, TargetDest As Integer, PreX As Integer, PreY As Integer, NumLoop As Integer) As Boolean
    Dim bool As Boolean
    bool = False
    If Not (DestX = PreX And DestY = PreY) Then
        If ValidateNearJumpMove(x1, y1, DestX, DestY) And CheckValidMove(DestX, DestY, CurrentPlayer) Then
            CompFindMove Index, DestX, DestY, TargetDest, x1, y1, NumLoop + 2
            bool = True
        End If
    End If
    CompFindMoveBool = bool
End Function
            
'########################################################################


Private Sub Marble_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MarbleSelected Then
        Marble(MarbleSelectednum).Move Marble(MarbleSelectednum).Left + X - TempMDownX, Marble(MarbleSelectednum).Top + Y - TempMDownY
    End If
End Sub

Private Sub Marble_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GameStarted And Not Group(CurrentPlayer).Auto Then
        If Int((Index - 1) / 10) + 1 = CurrentPlayer Then
            MoveMarble Int(MarbleSelectednum), Int((Marble(MarbleSelectednum).Left + X - BoardMargin) / ((Me.Width - 140) / BoardWidth)), Int((Marble(MarbleSelectednum).Top + Y - BoardMargin) / ((Me.Height - 650) / BoardHeight))
            LoadMarble (Index)
            MarbleSelected = False
        End If
    End If
End Sub

Private Sub Marble_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GameStarted And Not Group(CurrentPlayer).Auto Then
        If Int((Index - 1) / 10) + 1 = CurrentPlayer Then
            MarbleSelectednum = Index
            TempMDownX = X
            TempMDownY = Y
            MarbleSelected = True
        End If
    End If
End Sub

Public Sub MoveMarble(Index As Integer, w As Integer, h As Integer)
    Dim found As Boolean
    If BoardGrid(w, h) = 0 Then
        found = False
        MoveCount = 0
        List2.Clear
        If ValidateNearMove(PlayerPos(Index).X, PlayerPos(Index).Y, w, h) Then
            found = True
        Else
            found = ValidateFarMove(PlayerPos(Index).X, PlayerPos(Index).Y, w, h, -1, -1, 0)
            ClearVisited
        End If
        If found Then
            If MoveCount > 1 Then
                StepTmp = MoveCount - 2
            Else
                StepTmp = -1
            End If
            tmpMoveW = w
            tmpMoveH = h
            BoardGrid(w, h) = CurrentPlayer
            BoardGrid(PlayerPos(Index).X, PlayerPos(Index).Y) = 0
            PlayerPos(Index).X = w
            PlayerPos(Index).Y = h
            MarbleSelectednum = Index
            tmrMove.Enabled = True
            tmrMove_Timer
        End If
    End If
End Sub

Private Sub tmrMove_Timer()
    If StepTmp = -1 Then
        tmrMove.Enabled = False
        If MoveCount >= 0 Then
            If MarbleSelectednum >= 1 And MarbleSelectednum <= 60 Then
                PlayerPos(MarbleSelectednum).X = tmpMoveW
                PlayerPos(MarbleSelectednum).Y = tmpMoveH
                LoadMarble (MarbleSelectednum)
            End If
        End If
        
        Dim i As Integer, j As Integer
        Dim WinPlayerPos As Integer
        Dim win As Boolean
        Dim MarbleIn As Boolean
        WinPlayerPos = CurrentPlayer + 3
        If WinPlayerPos > 6 Then WinPlayerPos = WinPlayerPos - 6
        i = 1
        win = True
        Do
            MarbleIn = False
            For j = 1 To 10
                If PlayerPos(CurrentPlayer * 10 - 10 + i).X = WinPos(WinPlayerPos * 10 - 10 + j).X And PlayerPos(CurrentPlayer * 10 - 10 + i).Y = WinPos(WinPlayerPos * 10 - 10 + j).Y Then
                    MarbleIn = True
                    Exit For
                End If
            Next j
            If Not MarbleIn Then win = False
            i = i + 1
        Loop Until i > 10 Or (i > 10 And win = True)
        Group(CurrentPlayer).NoOfMove = Group(CurrentPlayer).NoOfMove + 1
        If win Then
            Group(CurrentPlayer).Status = "Won"
            If Group(CurrentPlayer).Auto = False Then
                SaveScore (CurrentPlayer)
            End If
        End If
        'Find next player
        imgTurnBlink(CurrentPlayer).Visible = True
        i = 1
        Do
            CurrentPlayer = CurrentPlayer + 1
            If CurrentPlayer > 6 Then CurrentPlayer = 1
            i = i + 1
        Loop Until Group(CurrentPlayer).Status = "Playing" Or i > 6
        If i > 6 Then
            frmHighScore.Show
            GameStarted = False
            mnuSave.Enabled = False
        End If
        ClearVisited
        StepTmp = 0
        MoveCount = 0
        tmrTurn.Enabled = True
        ShowStatus
    Else
        If MarbleSelectednum >= 1 And MarbleSelectednum <= 60 Then
            PlayerPos(MarbleSelectednum).X = MoveTemp(StepTmp).X
            PlayerPos(MarbleSelectednum).Y = MoveTemp(StepTmp).Y
            LoadMarble (MarbleSelectednum)
        End If
        StepTmp = StepTmp - 1
    End If
End Sub

Public Function SaveScore(p As Integer)
    Dim z As Integer
    Dim a As Integer
    Dim temp As ScoreList
    Open "Score.dat" For Input As #1
    For z = 1 To 10
        Input #1, Score(z).Name, Score(z).NoOfMove, Score(z).DateTime
    Next z
    Close #1
    
    For z = 1 To 10
        If Group(CurrentPlayer).NoOfMove < Score(z).NoOfMove Or Score(z).NoOfMove = 0 Then
            temp = Score(z)
            Score(z).Name = Group(p).Name
            Score(z).NoOfMove = Group(CurrentPlayer).NoOfMove
            Score(z).DateTime = Now
            For a = 10 To z + 2 Step -1
                Score(a) = Score(a - 1)
            Next a
            If z < 10 Then Score(z + 1) = temp
            Exit For
        End If
    Next z
    
    Open "Score.dat" For Output As #1
    For z = 1 To 10
        Write #1, Score(z).Name, Score(z).NoOfMove, Score(z).DateTime
    Next z
    Close #1
End Function

Public Function LocateMarble(x1 As Integer, y1 As Integer) As Integer
    For i = 1 To 60
        If PlayerPos(i).X = w And PlayerPos(i).Y = h Then Exit For
    Next
    If i <= 60 Then
        LocateMarble = i
    Else
        LocateMarble = -1
    End If
End Function

Public Function DrawMarble(w As Integer, h As Integer)
    Dim i As Integer
    If h Mod 2 = 0 Then
        CurrentX = (w * (Me.Width - 140) / BoardWidth) + CircleRadius
        CurrentY = (h * (Me.Height - 650) / BoardHeight) + CircleRadius
    Else
        CurrentX = (w * (Me.Width - 140) / BoardWidth) + CircleRadius + (CircleRadius * 2)
        CurrentY = (h * (Me.Height - 650) / BoardHeight) + CircleRadius
    End If
    FillColor = vbBlack
    Circle (CurrentX + BoardMargin, CurrentY + BoardMargin), CircleRadius, vbBlack
End Function

Public Function ValidateNearMove(W1 As Integer, H1 As Integer, W2 As Integer, H2 As Integer) As Boolean
    Dim valid As Boolean
    valid = False
    If W1 >= 0 And H1 >= 0 And W2 >= 0 And H2 >= 0 And H2 <= 16 Then
        If BoardGrid(W2, H2) = 0 Then
            If (W2 = W1 + 1 Or W2 = W1 - 1) And H1 = H2 Then
                valid = True
            ElseIf H1 Mod 2 = 1 Then
                If (H2 = H1 + 1 Or H2 = H1 - 1) And (W2 = W1 + 1 Or W2 = W1) Then valid = True
            ElseIf H1 Mod 2 = 0 Then
                If (H2 = H1 + 1 Or H2 = H1 - 1) And (W2 = W1 - 1 Or W2 = W1) Then valid = True
            End If
            If Not valid Then valid = ValidateNearJumpMove(W1, H1, W2, H2)
        Else
            valid = False
        End If
    Else
        valid = False
    End If
    ValidateNearMove = valid
End Function

Public Function ValidateNearJumpMove(W1 As Integer, H1 As Integer, W2 As Integer, H2 As Integer) As Boolean
    Dim valid As Boolean
    valid = False
    If W1 >= 0 And H1 >= 0 And W2 >= 0 And H2 >= 0 And H2 <= 16 Then
        If BoardGrid(W2, H2) = 0 Then
            If (W2 = W1 + 1 Or W2 = W1 - 1) And H1 = H2 Then
                valid = True
            ElseIf H1 Mod 2 = 1 Then
                If (H2 = H1 + 1 Or H2 = H1 - 1) And (W2 = W1 + 1 Or W2 = W1) Then valid = True
            ElseIf H1 Mod 2 = 0 Then
                If (H2 = H1 + 1 Or H2 = H1 - 1) And (W2 = W1 - 1 Or W2 = W1) Then valid = True
            End If
            If (W2 = W1 + 2 Or W2 = W1 - 2) And H1 = H2 Then
                If W2 > W1 Then
                    If BoardGrid(W2 - 1, H2) <> 0 Then valid = True
                Else
                    If BoardGrid(W2 + 1, H2) <> 0 Then valid = True
                End If
            ElseIf (H2 = H1 + 2 Or H2 = H1 - 2) And (W2 = W1 + 1 Or W2 = W1 - 1) Then
                If H2 > H1 Then
                    If H1 Mod 2 = 1 Then
                        If W2 > W1 Then
                            If BoardGrid(W2, H2 - 1) <> 0 Then valid = True
                        Else
                            If BoardGrid(W2 + 1, H2 - 1) <> 0 Then valid = True
                        End If
                    ElseIf H1 Mod 2 = 0 Then
                        If W2 > W1 Then
                            If BoardGrid(W2 - 1, H2 - 1) <> 0 Then valid = True
                        Else
                            If BoardGrid(W2, H2 - 1) <> 0 Then valid = True
                        End If
                    End If
                Else
                    If H1 Mod 2 = 1 Then
                        If W2 > W1 Then
                            If BoardGrid(W2, H2 + 1) <> 0 Then valid = True
                        Else
                            If BoardGrid(W2 + 1, H2 + 1) <> 0 Then valid = True
                        End If
                    ElseIf H1 Mod 2 = 0 Then
                        If W2 > W1 Then
                            If BoardGrid(W2 - 1, H2 + 1) <> 0 Then valid = True
                        Else
                            If BoardGrid(W2, H2 + 1) <> 0 Then valid = True
                        End If
                    End If
                End If
            End If
        Else
            valid = False
        End If
    Else
        valid = False
    End If
    ValidateNearJumpMove = valid
End Function

Public Function ValidateFarMove(W1 As Integer, H1 As Integer, W2 As Integer, H2 As Integer, W3 As Integer, H3 As Integer, NumLoop As Integer) As Boolean
    Dim val As Boolean
    If W1 = W2 And H1 = H2 Then
        MoveTemp(MoveCount).X = W1
        MoveTemp(MoveCount).Y = H1
        List2.AddItem W1 & " " & H1 & " " & MoveCount
        val = True
    ElseIf W1 >= 0 And H1 >= 0 And BoardGrid(W1, H1) <> 7 Then
        If BoardGrid(W1, H1) = 0 Then BoardGrid(W1, H1) = 7
        val = False
        If Not (W3 = (W1 - 1) And H3 = H1 - 2) And ValidateNearMove(W1, H1, W1 - 1, H1 - 2) Then
            val = ValidateFarMove(W1 - 1, H1 - 2, W2, H2, W1, H1, NumLoop + 1)
        End If
        If val = False Then
            If Not (W3 = (W1 + 1) And H3 = H1 - 2) And ValidateNearMove(W1, H1, W1 + 1, H1 - 2) Then
                val = ValidateFarMove(W1 + 1, H1 - 2, W2, H2, W1, H1, NumLoop + 1)
            End If
        End If
        If val = False Then
            If Not (W3 = (W1 + 2) And H3 = H1) And ValidateNearMove(W1, H1, W1 + 2, H1) Then
                val = ValidateFarMove(W1 + 2, H1, W2, H2, W1, H1, NumLoop + 1)
            End If
        End If
        If val = False Then
            If Not (W3 = (W1 + 1) And H3 = H1 + 2) And ValidateNearMove(W1, H1, W1 + 1, H1 + 2) Then
                val = ValidateFarMove(W1 + 1, H1 + 2, W2, H2, W1, H1, NumLoop + 1)
            End If
        End If
        If val = False Then
            If Not (W3 = (W1 - 1) And H3 = H1 + 2) And ValidateNearMove(W1, H1, W1 - 1, H1 + 2) Then
                val = ValidateFarMove(W1 - 1, H1 + 2, W2, H2, W1, H1, NumLoop + 1)
            End If
        End If
        If val = False Then
            If Not (W3 = (W1 - 2) And H3 = H1) And ValidateNearMove(W1, H1, W1 - 2, H1) Then
                val = ValidateFarMove(W1 - 2, H1, W2, H2, W1, H1, NumLoop + 1)
            End If
        End If
        If val Then
            MoveTemp(MoveCount).X = W1
            MoveTemp(MoveCount).Y = H1
            MoveCount = MoveCount + 1
            List2.AddItem W1 & " " & H1 & " " & MoveCount
        End If
    Else
        val = False
    End If
    ValidateFarMove = val
End Function

Public Function ClearVisited()
    For h = 0 To BoardHeight - 1
        For w = 0 To BoardWidth - 1
            If BoardGrid(w, h) = 7 Then BoardGrid(w, h) = 0
        Next w
    Next h
End Function

Private Sub Form_Resize()
    DrawBoard
    If GameStarted Then
        For p = 1 To 60
            LoadMarble (p)
        Next
    End If
End Sub

Private Sub DrawBoard()
    If GameStarted Then
        Dim BlockW As Integer
        Dim BlockH As Integer
        CircleRadius = ((Me.Width - 140) / BoardWidth) / 4
        BlockW = (Me.Width - 140) / BoardWidth
        BlockH = (Me.Height - 650) / BoardHeight
        
        Cls
        FillStyle = vbFSTransparent
        FillStyle = vbSolid
        
        Dim point1(2) As Long
        Dim point2(2) As Long
        Dim point3(2) As Long
        DrawWidth = 4
        ' Player #1
        h = 0
        w = 9
        point1(0) = (w * BlockW) + BoardMargin + CircleRadius
        point1(1) = (h * BlockH) + BoardMargin - (CircleRadius * 2)
        point2(0) = ((w - 1) * BlockW) + BoardMargin - (CircleRadius * 4)
        point2(1) = ((h + 3) * BlockH) + BoardMargin + (CircleRadius * 2.5)
        point3(0) = ((w + 2) * BlockW) + BoardMargin + (CircleRadius * 2)
        point3(1) = ((h + 3) * BlockH) + BoardMargin + (CircleRadius * 2.5)
        Line (point1(0), point1(1))-(point2(0), point2(1)), Group(1).Color
        Line (point1(0), point1(1))-(point3(0), point3(1)), Group(1).Color
        Line (point2(0), point2(1))-(point3(0), point3(1)), Group(1).Color
        ' Player #2
        h = 4
        w = 15
        point1(0) = (w * BlockW) + BoardMargin + (CircleRadius * 4)
        point1(1) = (h * BlockH) + BoardMargin - (CircleRadius * 0.5)
        point2(0) = ((w - 3) * BlockW) + BoardMargin - (CircleRadius * 2)
        point2(1) = (h * BlockH) + BoardMargin - (CircleRadius * 0.5)
        point3(0) = ((w - 1) * BlockW) + BoardMargin - (CircleRadius)
        point3(1) = ((h + 3) * BlockH) + BoardMargin + (CircleRadius * 4)
        Line (point1(0), point1(1))-(point2(0), point2(1)), Group(2).Color
        Line (point1(0), point1(1))-(point3(0), point3(1)), Group(2).Color
        Line (point2(0), point2(1))-(point3(0), point3(1)), Group(2).Color
        ' Player #3
        h = 12
        w = 15
        point1(0) = (w * BlockW) + BoardMargin + (CircleRadius * 4)
        point1(1) = (h * BlockH) + BoardMargin + (CircleRadius * 2.5)
        point2(0) = ((w - 3) * BlockW) + BoardMargin - (CircleRadius * 2)
        point2(1) = (h * BlockH) + BoardMargin + (CircleRadius * 2.5)
        point3(0) = ((w - 1) * BlockW) + BoardMargin - (CircleRadius)
        point3(1) = ((h - 3) * BlockH) + BoardMargin - (CircleRadius * 2)
        Line (point1(0), point1(1))-(point2(0), point2(1)), Group(3).Color
        Line (point1(0), point1(1))-(point3(0), point3(1)), Group(3).Color
        Line (point2(0), point2(1))-(point3(0), point3(1)), Group(3).Color
        ' Player #4
        h = 16
        w = 9
        point1(0) = (w * BlockW) + BoardMargin + CircleRadius
        point1(1) = (h * BlockH) + BoardMargin + (CircleRadius * 4)
        point2(0) = ((w - 1) * BlockW) + BoardMargin - (CircleRadius * 4)
        point2(1) = ((h - 3) * BlockH) + BoardMargin - (CircleRadius / 2)
        point3(0) = ((w + 2) * BlockW) + BoardMargin + (CircleRadius * 2)
        point3(1) = ((h - 3) * BlockH) + BoardMargin - (CircleRadius / 2)
        Line (point1(0), point1(1))-(point2(0), point2(1)), Group(4).Color
        Line (point1(0), point1(1))-(point3(0), point3(1)), Group(4).Color
        Line (point2(0), point2(1))-(point3(0), point3(1)), Group(4).Color
        ' Player #5
        h = 12
        w = 6
        point1(0) = (w * BlockW) + BoardMargin + (CircleRadius * 4)
        point1(1) = (h * BlockH) + BoardMargin + (CircleRadius * 2.5)
        point2(0) = ((w - 3) * BlockW) + BoardMargin - (CircleRadius * 2)
        point2(1) = (h * BlockH) + BoardMargin + (CircleRadius * 2.5)
        point3(0) = ((w - 1) * BlockW) + BoardMargin - (CircleRadius)
        point3(1) = ((h - 3) * BlockH) + BoardMargin - (CircleRadius * 2)
        Line (point1(0), point1(1))-(point2(0), point2(1)), Group(5).Color
        Line (point1(0), point1(1))-(point3(0), point3(1)), Group(5).Color
        Line (point2(0), point2(1))-(point3(0), point3(1)), Group(5).Color
        ' Player #6
        h = 4
        w = 6
        point1(0) = (w * BlockW) + BoardMargin + (CircleRadius * 4)
        point1(1) = (h * BlockH) + BoardMargin - (CircleRadius * 0.5)
        point2(0) = ((w - 3) * BlockW) + BoardMargin - (CircleRadius * 2)
        point2(1) = (h * BlockH) + BoardMargin - (CircleRadius * 0.5)
        point3(0) = ((w - 1) * BlockW) + BoardMargin - (CircleRadius)
        point3(1) = ((h + 3) * BlockH) + BoardMargin + (CircleRadius * 4)
        Line (point1(0), point1(1))-(point2(0), point2(1)), Group(6).Color
        Line (point1(0), point1(1))-(point3(0), point3(1)), Group(6).Color
        Line (point2(0), point2(1))-(point3(0), point3(1)), Group(6).Color
        
        DrawWidth = 2
        For h = 0 To BoardHeight - 1
            For w = 0 To BoardWidth - 1
                If BoardGrid(w, h) <> -1 Then
                    If h Mod 2 = 0 Then
                        If w > 1 And h > 1 Then
                            If BoardGrid(w - 1, h - 1) <> -1 Then Line ((w * BlockW) + BoardMargin - CircleRadius, (h * BlockH) + BoardMargin - BlockH + CircleRadius)-((w * BlockW) + BoardMargin + CircleRadius, (h * BlockH) + BoardMargin + CircleRadius)
                            If BoardGrid(w, h - 1) <> -1 Then Line ((w * BlockW) + BoardMargin + BlockW - CircleRadius, (h * BlockH) + BoardMargin - BlockH + CircleRadius)-((w * BlockW) + BoardMargin + CircleRadius, (h * BlockH) + BoardMargin + CircleRadius)
                        End If
                        If BoardGrid(w + 1, h) <> -1 Then Line ((w * BlockW) + BoardMargin + BlockW + CircleRadius, (h * BlockH) + BoardMargin + CircleRadius)-((w * BlockW) + BoardMargin + CircleRadius, (h * BlockH) + BoardMargin + CircleRadius)
                        If BoardGrid(w, h + 1) <> -1 Then Line ((w * BlockW) + BoardMargin + BlockW - CircleRadius, (h * BlockH) + BoardMargin + BlockH + CircleRadius)-((w * BlockW) + BoardMargin + CircleRadius, (h * BlockH) + BoardMargin + CircleRadius)
                        If w > 1 Then
                            If BoardGrid(w - 1, h + 1) <> -1 Then Line ((w * BlockW) + BoardMargin - CircleRadius, (h * BlockH) + BoardMargin + BlockH + CircleRadius)-((w * BlockW) + BoardMargin + CircleRadius, (h * BlockH) + BoardMargin + CircleRadius)
                        End If
                    Else
                        If BoardGrid(w + 1, h) <> -1 Then Line (((w + 1) * BlockW) + BoardMargin - CircleRadius, (h * BlockH) + BoardMargin + CircleRadius)-(((w + 2) * BlockW) + BoardMargin - CircleRadius, (h * BlockH) + BoardMargin + CircleRadius)
                    End If
                End If
            Next w
        Next h
        
        For w = 0 To BoardWidth - 1
            For h = 0 To BoardHeight - 1
                If BoardGrid(w, h) <> -1 Then DrawMarble w, h
            Next h
        Next w
    End If
End Sub

Public Function LoadMarble(Num As Integer)
    Dim colorName As String
    If Num >= 1 And Num <= 60 Then
        If PlayerPos(Num).Y Mod 2 = 0 Then
            CurrentX = (PlayerPos(Num).X * (Me.Width - 140) / BoardWidth) + CircleRadius
            CurrentY = (PlayerPos(Num).Y * (Me.Height - 650) / BoardHeight) + CircleRadius
        Else
            CurrentX = (PlayerPos(Num).X * (Me.Width - 140) / BoardWidth) + CircleRadius + (CircleRadius * 2)
            CurrentY = (PlayerPos(Num).Y * (Me.Height - 650) / BoardHeight) + CircleRadius
        End If
        If Group(Int((Num - 1) / 10) + 1).Status = "Playing" Or Group(Int((Num - 1) / 10) + 1).Status = "Won" Then
            Select Case Group(Int((Num - 1) / 10) + 1).Color
                Case vbBlue
                    colorName = "Blue"
                Case vbCyan
                    colorName = "Cyan"
                Case vbGreen
                    colorName = "Green"
                Case vbMagenta
                    colorName = "Magenta"
                Case vbRed
                    colorName = "Red"
                Case vbYellow
                    colorName = "Yellow"
            End Select
            Marble(Num).Stretch = True
            Marble(Num).Picture = LoadPicture(App.Path & "\" & colorName & ".gif")
            Marble(Num).Width = CircleRadius * 2
            Marble(Num).Height = CircleRadius * 2
            Marble(Num).Visible = True
            Marble(Num).Move CurrentX + BoardMargin - CircleRadius + 10, CurrentY + BoardMargin - CircleRadius + 10
            If GameStarted Then
                i = PlaySound(App.Path & "/Move.wav", 0, SND_FILENAME Or SND_ASYNC)
            End If
        Else
            Marble(Num).Visible = False
        End If
    End If
End Function

Private Sub tmrTurnBlink_Timer()
    If GameStarted Then
        If imgTurnBlink(CurrentPlayer).Visible Then
            imgTurnBlink(CurrentPlayer).Visible = False
        Else
            imgTurnBlink(CurrentPlayer).Visible = True
        End If
    End If
End Sub
