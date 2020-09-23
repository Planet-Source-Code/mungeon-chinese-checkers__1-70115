Attribute VB_Name = "modSetting"
Option Explicit

Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_ALIAS = &H10000 'lpszName is a string identifying the name of the system event sound to play.
Public Const SND_ALIAS_ID = &H110000 'lpszName is a string identifying the name of the predefined sound identifier to play.
Public Const SND_APPLICATION = &H80 'lpszName is a string identifying the application-specific event association sound to play.
Public Const SND_ASYNC = &H1 'Play the sound asynchronously -- return immediately after beginning to play the sound and have it play in the background.
Public Const SND_FILENAME = &H20000 'lpszName is a string identifying the filename of the .wav file to play.
Public Const SND_LOOP = &H8 'Continue looping the sound until this function is called again ordering the looped playback to stop. SND_ASYNC must also be specified.
Public Const SND_MEMORY = &H4 'lpszName is a numeric pointer refering to the memory address of the image of the waveform sound loaded into RAM.
Public Const SND_NODEFAULT = &H2 'If the specified sound cannot be found, terminate the function with failure instead of playing the SystemDefault sound. If this flag is not specified, the SystemDefault sound will play if the specified sound cannot be located and the function will return with success.
Public Const SND_NOSTOP = &H10 'If a sound is already playing, do not prematurely stop that sound from playing and instead return with failure. If this flag is not specified, the playing sound will be terminated and the sound specified by the function will play instead.
Public Const SND_NOWAIT = &H2000 'If a sound is already playing, do not wait for the currently playing sound to stop and instead return with failure.
Public Const SND_PURGE = &H40 'Stop playback of any waveform sound. lpszName must be an empty string.
Public Const SND_RESOURCE = &H4004 'lpszName is the numeric resource identifier of the sound stored in an application. hModule must be specified as that application's module handle.
Public Const SND_SYNC = &H0 'Play the sound synchronously -- do not return until the sound has finished playing.

Public v_dx As New DirectX7
Public v_dmp As DirectMusicPerformance
Public v_dml As DirectMusicLoader
Public v_dms As DirectMusicSegment
Public v_dmss As DirectMusicSegmentState
Dim vs_filename As String
Public vl_volume As Long

Public Type Player
    Number As Integer
    Name As String
    Auto As Boolean
    Status As String
    NoOfMove As Integer
    Color As ColorConstants
End Type

Public Type ScoreList
    Name As String
    NoOfMove As Integer
    DateTime As Date
End Type

Public Type PlayerPosition
    X As Integer
    Y As Integer
End Type

Public Type MovingStep
    X As Integer
    Y As Integer
End Type

Public Type CompStep
    Num As Integer
    DestX As Integer
    DestY As Integer
    Dist As Integer
End Type

Public MusicBool As Boolean
Public activeWin As Form
Public Group(1 To 6) As Player

Public Function ClearGroup(g As Integer)
    Group(g).Auto = False
    Group(g).Name = ""
    Group(g).Number = -1
    Group(g).Status = "None"
    Group(g).NoOfMove = -1
    Group(g).Color = vbBlack
End Function

Public Sub PlayMidi()
    Set v_dml = v_dx.DirectMusicLoaderCreate
    Set v_dmp = v_dx.DirectMusicPerformanceCreate
    
    Call v_dmp.Init(Nothing, frmGame.hWnd)
    Call v_dmp.SetPort(-1, 1)
    
    vs_filename = "music.mid"
    v_dml.SetSearchDirectory (App.Path)
    Set v_dms = v_dml.LoadSegment(vs_filename)
    If StrConv(Right(vs_filename, 4), vbLowerCase) = ".mid" Then
        v_dms.SetStandardMidiFile
    End If
    
    Call v_dmp.SetMasterAutoDownload(True)
    Call v_dms.Download(v_dmp)
    Call v_dmp.SetMasterVolume(0)
    Set v_dmss = v_dmp.PlaySegment(v_dms, 0, 0)
    frmGame.tmrMusic.Enabled = True
End Sub

Public Sub CloseMidi()
    If v_dms Is Nothing Then Exit Sub
    Call v_dmp.Stop(v_dms, v_dmss, 0, 0)
    Call v_dms.Unload(v_dmp)
    frmGame.tmrMusic.Enabled = False
End Sub

