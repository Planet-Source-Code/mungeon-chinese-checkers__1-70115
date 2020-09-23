VERSION 5.00
Begin VB.Form frmTaskbarWindow 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   90
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   90
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   90
   ScaleMode       =   0  'User
   ScaleWidth      =   90
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmTaskbarWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    On Error Resume Next
    activeWin.Show
    activeWin.SetFocus
End Sub

Private Sub Form_Load()
    Set activeWin = frmGame
    Me.Caption = App.Title
    'On Error Resume Next
    activeWin.Show
    activeWin.SetFocus
End Sub
