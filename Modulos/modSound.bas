Attribute VB_Name = "modSound"
Option Explicit

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Sub PlaySound(strSound As String)

    Dim wFlags%
    
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    sndPlaySound strSound, wFlags%

End Sub




