Attribute VB_Name = "Module1"
'All variables must be declared
Option Explicit

Dim retVal As Long

'Declare a user-defined type that corresponds to a
'record in the file BK.DAT
Type Info
    BkTitle     As String * 40
    BkNumber    As String * 40
    StName      As String * 40
    StNumber    As String * 40
    StTaName    As String * 40
    StTaNumber  As String * 40
    StSbCode    As String * 40
End Type

'Declare a user-defined type that corresponds to a
'record in the file LOGIN.DAT
Type Login
    LgName      As String * 40
    LgPassword  As String * 40
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
   Const SND_ASYNC = &H1
   Const SND_SYNC = &H0

