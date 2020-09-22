Attribute VB_Name = "Common_SoundModule"
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0 ' play synchronously (default)
Private Const SND_ASYNC = &H1 ' play asynchronously
Private Const SND_LOOP = &H8 ' loop the sound until next
Private Const SND_MEMORY = &H4

Public Enum ePlaySoundModes
    [Sync] = 0
    [Async] = 1
    [Loop] = 8
End Enum

Public Function PlaySound(FilePath As String, Optional Mode As ePlaySoundModes = [Async])
On Error GoTo Error
    
    sndPlaySound ByVal FilePath, Mode

Exit Function
Error:
    Assert , "SoundModule.PlaySound", Err.Number, Err.Description, "FilePath: '" & FilePath & "'"
End Function

Public Function PlaySoundFromMemory(SoundData As String, Optional Mode As ePlaySoundModes = [Async])
On Error GoTo Error

    sndPlaySound SoundData, Mode Or SND_MEMORY

Exit Function
Error:
    Assert , "SoundModule.PlaySoundFromMemory", Err.Number, Err.Description
End Function
