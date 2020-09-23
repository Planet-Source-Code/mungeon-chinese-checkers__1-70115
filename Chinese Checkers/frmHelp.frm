VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   Caption         =   "Help Content"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9390
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9390
      TabIndex        =   1
      Top             =   6345
      Width           =   9390
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   340
         Left            =   3480
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
   End
   Begin RichTextLib.RichTextBox rtbHelp 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmHelp.frx":0442
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set activeWin = Me
    rtbHelp.Text = "About Chinese Checker" & vbCrLf & _
                    vbTab & "Chinese Checkers is descended from an earlier game called'Halma', which is played at a square game board. Halma was invented by the Victorians in around 1880.  The objective is to move all your pieces from your corner into the opposing corner." & vbCrLf & _
                    "Chinese checker was first patented in the West by Ravensburger, the famous German games company, under the name Stern-Halma in Germany a few years after Halma appeared.  It was later launched in the USA under the catchier name of Chinese Checkers, and this is the form that is most well-known today." & vbCrLf & _
                    vbCrLf & _
                    "How to play" & vbCrLf & _
                    vbTab & "Chinese Checkers is played on a star-shaped game board. Each player uses markers of a different color placed within one of the points of the star. The object is to move your markers across the board to occupy the star point directly opposite. The player getting all markers across first wins." & vbCrLf & _
                    "The game is started by anyone and the play continues to the left of the starter. One can move or jump in any direction as long as one follows the lines. As in checkers, move only one hole or jump only one marble, although successive jumps are permissible wherever they can be made in any direction."
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    rtbHelp.Width = Me.Width - 435
    rtbHelp.Height = Me.Height - 1310
    cmdClose.Left = Me.Width - 1350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set activeWin = frmGame
End Sub
