Attribute VB_Name = "MCI_Control"
Enum Mp3
    Mp3Play = 1
    Mp3Pause = 2
    Mp3Stop = 3
End Enum

Public Const MCI_STATUS_LENGTH = &H1&
Public Const MCI_STATUS_POSITION = &H2&
Public Const MCI_PLAY = &H806

Public Filename As String
Public MpegFilename As String
Public MpegSongTitle As String
Public MpegArtist As String
Public MpegAblum As String
Public MpegYear As String
Public MpegComments As String
Public DisplayInfo As Boolean
Public Const MCI_ANIM_PLAY_FAST = &H40000

Dim OnOff As Boolean
Dim TagInfo As String

Global Length As Double
Global Position As Double

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Declare Sub ReleaseCapture Lib "user32" ()


Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long


Public Sub FormMove(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Private Function PlayMP3(Mp3Command As Mp3)
    Select Case Mp3Command
    Case 1
    Call mciSendString("Open " & Filename & " Alias MM", 0, 0, 0)
    Call mciSendString("Play MM", 0, 0, MCI_ANIM_PLAY_FAST)
    Case 2
    Call mciSendString("Stop MM", 0, 0, 0)
    Case 3
    Call mciSendString("Stop MM", 0, 0, 0)
    Call mciSendString("Close MM", 0, 0, 0)
    End Select
End Function

Sub mPlay()
     PlayMP3 Mp3Play
End Sub
 
Sub mStop()
    PlayMP3 Mp3Stop
End Sub

Sub mPause()
    Select Case OnOff
    Case True
    OnOff = False
    PlayMP3 Mp3Pause
    Case False
    OnOff = True
    PlayMP3 Mp3Play
    End Select
End Sub

Public Function GetShortPath(strFileName As String) As String
Dim lngRes As Long, strPath As String
strPath = String$(165, 0)
lngRes = GetShortPathName(strFileName, strPath, 164)
GetShortPath = Left$(strPath, lngRes)
End Function

Public Function mGetLength() As Double
Dim Ret As String * 256
Ret = String(256, " ")
Call mciSendString("Status MM length", Ret, 256, 0)
mGetLength = Val(Ret)
End Function

Public Function mSetTimeFormat() As String
Call mciSendString("Set MM Time Format milliseconds", 0, 0, 0)
End Function

Public Function mGetPosition() As Double
Dim Ret As String * 256
Ret = String(256, " ")
Call mciSendString("Status MM Position", Ret, 256, 0)
mGetPosition = Val(Ret)
End Function
