VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Electronic Book Keeper"
   ClientHeight    =   4200
   ClientLeft      =   2535
   ClientTop       =   2970
   ClientWidth     =   6960
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4200
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   330
      Left            =   2520
      TabIndex        =   13
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   330
      Left            =   1320
      TabIndex        =   12
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton Command13 
      Caption         =   "print now!"
      Height          =   330
      Left            =   5640
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Add"
      Height          =   330
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   1065
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Previous"
      Height          =   330
      Left            =   5640
      TabIndex        =   9
      Top             =   3120
      Width           =   1065
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      Height          =   330
      Left            =   5640
      TabIndex        =   8
      Top             =   2640
      Width           =   1065
   End
   Begin VB.TextBox Text12 
      Height          =   330
      Left            =   2040
      TabIndex        =   7
      Top             =   3225
      Width           =   2010
   End
   Begin VB.TextBox Text11 
      Height          =   330
      Left            =   2040
      TabIndex        =   6
      Top             =   2805
      Width           =   750
   End
   Begin VB.TextBox Text10 
      Height          =   330
      Left            =   2040
      TabIndex        =   5
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      Height          =   330
      Left            =   2040
      TabIndex        =   4
      Top             =   1965
      Width           =   1275
   End
   Begin VB.TextBox Text8 
      Height          =   330
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Width           =   4485
   End
   Begin VB.TextBox Text7 
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Top             =   810
      Width           =   1485
   End
   Begin VB.TextBox Text6 
      Height          =   330
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   4485
   End
   Begin VB.Timer Timer1 
      Interval        =   15000
      Left            =   480
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   960
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject Code:"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   20
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TA Number:"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TA Name:"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Number:"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name::"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book Number:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book Title:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   765
   End
   Begin VB.Shape Shape2 
      Height          =   2325
      Left            =   120
      Top             =   1440
      Width           =   6750
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   120
      Top             =   240
      Width           =   6735
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4440
      TabIndex        =   0
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup"
      End
      Begin VB.Menu Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu Separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "P&rinter Setup"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Now!"
      End
      Begin VB.Menu Seperator 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuText 
         Caption         =   "&Text"
         Begin VB.Menu mnuTColour 
            Caption         =   "&Colour"
            Begin VB.Menu mnuWhite 
               Caption         =   "&White"
            End
            Begin VB.Menu mnuBlack 
               Caption         =   "&Black"
            End
            Begin VB.Menu mnuRed 
               Caption         =   "&Red"
            End
            Begin VB.Menu mnuYellow 
               Caption         =   "&Yellow"
            End
            Begin VB.Menu mnuGreen 
               Caption         =   "&Green"
            End
            Begin VB.Menu mnuBlue 
               Caption         =   "B&lue"
            End
            Begin VB.Menu mnuPurple 
               Caption         =   "&Purple"
            End
         End
         Begin VB.Menu mnuTFont 
            Caption         =   "&Font"
         End
      End
      Begin VB.Menu mnuBckcolor 
         Caption         =   "&Background"
         Begin VB.Menu mnuBWhite 
            Caption         =   "&White"
         End
         Begin VB.Menu mnuBBlack 
            Caption         =   "&Black"
         End
         Begin VB.Menu mnuBRed 
            Caption         =   "&Red"
         End
         Begin VB.Menu mnuBYellow 
            Caption         =   "&Yellow"
         End
         Begin VB.Menu mnuBGreen 
            Caption         =   "&Green"
         End
         Begin VB.Menu mnuBBlue 
            Caption         =   "B&lue"
         End
         Begin VB.Menu mnuBPurple 
            Caption         =   "&Purple"
         End
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset Defult Colors"
      End
      Begin VB.Menu mnuSound 
         Caption         =   "&Sound"
         Begin VB.Menu mnuSndon 
            Caption         =   "&Sound On"
            Checked         =   -1  'True
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSndoff 
            Caption         =   "&Sound Off"
         End
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuBrowser 
         Caption         =   "&Browser"
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "&Email"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "&Notepad"
      End
      Begin VB.Menu mnuCalendar 
         Caption         =   "&Calendar"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuSHelp 
         Caption         =   "&Search For Help On.."
      End
      Begin VB.Menu mnuOTS 
         Caption         =   "&Obtaining Technical Support..."
      End
      Begin VB.Menu SEPERATOR2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Dim retVal As Long

'All variables must be declared
Option Explicit
'Declare variables that should be visible in all
'The Procedures of the forms.
Dim gInfo As Info
Dim gFileNum As Integer
Dim gRecordLen As Long
Dim gCurrentRecord As Long
Dim gLastRecord As Long
Dim counter As Integer

'Sound stuff
Dim WaveFiles() As String
Dim StopMe As Integer
Dim Flag As Integer
Dim SoundIsOn As Integer

Dim retVal As Long
'Another stuff
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
   Const SND_ASYNC = &H1
   Const SND_SYNC = &H0

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Public Sub SaveCurrentRecord()

    'Fill gPerson with the currecty displayed data.
    gInfo.BkTitle = Text6.Text
    gInfo.BkNumber = Text7.Text
    gInfo.StName = Text8.Text
    gInfo.StNumber = Text9.Text
    gInfo.StTaName = Text10.Text
    gInfo.StTaNumber = Text11.Text
    gInfo.StSbCode = Text12.Text
    
    'save gPerson to the current record.
    Put #gFileNum, gCurrentRecord, gInfo
    
End Sub


' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Public Sub ShowCurrentRecord()

    'Gets gPerson from the current record.
    Get #gFileNum, gCurrentRecord, gInfo
    
    'Fill gPerson with the currecty displayed data.
    Text6.Text = Trim(gInfo.BkTitle)
    Text7.Text = Trim(gInfo.BkNumber)
    Text8.Text = Trim(gInfo.StName)
    Text9.Text = Trim(gInfo.StNumber)
    Text10.Text = Trim(gInfo.StTaName)
    Text11.Text = Trim(gInfo.StTaNumber)
    Text12.Text = Trim(gInfo.StSbCode)
    
    'Display the current record number in the
    'caption of the form.
    Form1.Caption = "Electornic Book Keeper Record " + _
                    Str(gCurrentRecord) + "/" + _
                    Str(gLastRecord)

End Sub



Private Sub Command1_Click()
    
    'Declare variables that should be visible in all
    Dim DirResult
    Dim TmpFileNum
    Dim TmpInfo As Info
    Dim RecNum As Long
    Dim TmpRecNum As Long

If SoundIsOn = True Then
    
    Command1.Caption = "deleting..."
    Command1.Enabled = False
    Call sndPlaySound(App.Path & "\" & "Sounds\Auth.wav", SND_SYNC)
    Command1.Caption = "Delete"
    Command1.Enabled = True
    
Else
End If

    'Sets the Form2 caption to "Student Locker Administration Utility: Delete" when command4 is executed
    Form1.Caption = "Electronic Book Keeper Utility: Delete"
    
    'Before deleting get a confirmation from the user
    If MsgBox("Do you want to delete this record?", vbYesNo) = vbNo Then
        'Gives the focus to the Text1 field
        Text6.SetFocus
        'Exit the procedure without deleting
        Exit Sub
    End If

    'To physically delete the current record of STUDENT.DAT,
    'all the record of STUDENT.DAT, except the current record, are
    'copied into a temporary file (STUDENT.TMP) and then the file STUDENT.TMP is
    'copied into the STUDENT.DAT
    If Dir(App.Path & "\" & "bk.tmp") = App.Path & "\" & "bk.tmp" Then
        Kill App.Path & "\" & "bk.tmp" 'Delete Student.tmp if there is any at the time of the run
    End If

    'Create STUDENT.TMP with the same format as STUDENT.DAT
    TmpFileNum = FreeFile
    Open App.Path & "\" & "bk.tmp" For Random As TmpFileNum Len = gRecordLen

    'Copy all the records from STUDENT.TMP to STUDENT.DAT, except current record
    RecNum = 1
    TmpRecNum = 1
    Do While RecNum < gLastRecord + 1
        If RecNum <> gCurrentRecord Then
            Get #gFileNum, RecNum, TmpInfo
            Put #TmpFileNum, TmpRecNum, TmpInfo
            TmpRecNum = TmpRecNum + 1
        End If
        RecNum = RecNum + 1
    Loop

    'Delete STUDENT.DAT
    Close gFileNum
    Kill App.Path & "\" & "bk.dat"

    'Rename STUDENT.TMP to STUDENT.DAT
    Close TmpFileNum
    Name App.Path & "\" & "bk.tmp" As App.Path & "\" & "bk.dat"

    'Re-open the updated STUDENT.DAT file
    gFileNum = FreeFile
    Open App.Path & "\" & "bk.dat" For Random As gFileNum Len = gRecordLen

    'Update the value of LastRecord
    gLastRecord = gLastRecord - 1

    'Make sure that the gLastRecord is not 0
    If gLastRecord = 0 Then gLastRecord = 1

    'make sure gCurrentRecord is not out of range
    If gCurrentRecord > gLastRecord Then
        gCurrentRecord = gLastRecord
    End If

    'Show the current record
    ShowCurrentRecord

    'Give the focus to the Text1 field
    Text6.SetFocus

End Sub

Private Sub Command13_Click()
If SoundIsOn = True Then
    
    Command2.Caption = "searching..."
    Command2.Enabled = False
    Call sndPlaySound(App.Path & "\" & "Sounds\Auth.wav", SND_SYNC)
    Command2.Caption = "Search"
    Command2.Enabled = True
Else
End If
    Form1.PrintForm
    Printer.EndDoc

End Sub

Private Sub Command2_Click()

    'Declare variables that should be visible in all
    Dim NameToSearch As String
    Dim Found As Integer
    Dim RecNum As Long
    Dim TmpInfo As Info

If SoundIsOn = True Then
    
    Command2.Caption = "searching..."
    Command2.Enabled = False
    Call sndPlaySound(App.Path & "\" & "Sounds\Auth.wav", SND_SYNC)
    Command2.Caption = "Search"
    Command2.Enabled = True
Else
End If
    
    'Sets the Form2 caption to "Student Locker Administration Utility: Search" when command5 is executed
    Form2.Caption = "Electronic Book Keeper Utility: Search"
    
    'Get the name to search from the user
    NameToSearch = InputBox("Search for:", "Search")

    'If the user did not enter a name, exit from this procedure
    If NameToSearch = " " Then
        'Gives the focus to the Text1 field
        Text6.SetFocus
        'Exit the procedure without deleting
        Exit Sub
    End If

    'Convert the name to be searched to upper case
    NameToSearch = UCase(NameToSearch)

    'initialize the Found flag to False
    Found = False

    'Search for the name that the user entered
    For RecNum = 1 To gLastRecord
        Get #gFileNum, RecNum, TmpInfo
        If NameToSearch = UCase(Trim(TmpInfo.StNumber)) Then
            Found = True
            Exit For
        End If

    Next

    'if the name was found, then display the record of the found name
    If Found = True Then
        SaveCurrentRecord
        gCurrentRecord = RecNum
        ShowCurrentRecord
    Else
        'Popups a message box that tells that use what the name he has search is not found on the list/database
        MsgBox "Name " + NameToSearch + " not found!"
    End If

    'Give the focus to the Text1 field
    Text6.SetFocus

End Sub


Private Sub Command5_Click()
If SoundIsOn = True Then
    
    Command2.Caption = "searching..."
    Command2.Enabled = False
    Call sndPlaySound(App.Path & "\" & "Sounds\Auth.wav", SND_SYNC)
    Command2.Caption = "Search"
    Command2.Enabled = True
Else
End If
    'If the current record is the last record, beep and display and error message.
    'Otherwise, save the current record and skip to the next record.
    If gCurrentRecord = gLastRecord Then
        Beep
        MsgBox "End Of file!", vbExclamation
    Else
        SaveCurrentRecord
        gCurrentRecord = gCurrentRecord + 1
        ShowCurrentRecord
    End If

    'Give the focus to the Text1 field
    Text6.SetFocus
 
End Sub

Private Sub Command6_Click()
If SoundIsOn = True Then
    
    Command2.Caption = "searching..."
    Command2.Enabled = False
    Call sndPlaySound(App.Path & "\" & "Sounds\Auth.wav", SND_SYNC)
    Command2.Caption = "Search"
    Command2.Enabled = True
Else
End If
    'If the current record is the last record, beep and display and error message.
    'Otherwise, save the current record and go to the previous record.
    If gCurrentRecord = 1 Then
        Beep
        MsgBox "Beginning Of file!", vbExclamation
    Else
        SaveCurrentRecord
        gCurrentRecord = gCurrentRecord - 1
        ShowCurrentRecord
    End If

    'Give the focus to the Text1 field
    Text6.SetFocus

End Sub

Private Sub Command7_Click()
If SoundIsOn = True Then
    
    Command7.Caption = "adding..."
    Command7.Enabled = False
    Call sndPlaySound(App.Path & "\" & "Sounds\Auth.wav", SND_SYNC)
    Command7.Caption = "Add"
    Command7.Enabled = True
    
Else
End If

    SaveCurrentRecord 'Save the current record

    'Add a new blank record
    gLastRecord = gLastRecord + 1
    gInfo.BkTitle = " "
    gInfo.BkNumber = " "
    gInfo.StName = " "
    gInfo.StNumber = " "
    gInfo.StTaName = " "
    gInfo.StTaNumber = " "
    gInfo.StSbCode = " "
    
    Put #gFileNum, gLastRecord, gInfo

    'Update gCurrentRecord
    gCurrentRecord = gLastRecord
    
    'Display the record that was just created
    ShowCurrentRecord

    'Give the focus to the Text1 field
    Text6.SetFocus

End Sub


Private Sub Form_Load()

'ChDir App.Path
If Dir(App.Path & "\" & "bk.tmp") = App.Path & "\" & "bk.tmp" Then
Kill App.Path & "\" & "bk.dat"
End If

'Sound
SoundIsOn = True

ReDim WaveFiles(3)   'Put here the number of the WAV files
WaveFiles(1) = App.Path & "\" & "Sounds\Startup.wav"   'The name of the 1st. WAV file
WaveFiles(2) = App.Path & "\" & "Sounds\Email.wav"   '2nd....
WaveFiles(3) = App.Path & "\" & "Sounds\Auth.wav"
              
If SoundIsOn = True Then
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.FileName = WaveFiles(1)
    Flag = 1
    StopMe = 0
    MMControl1.Command = "open"
    MMControl1.Command = "play"
Else
End If
             
'Main

    Label15 = Time$ & " " & Date
    
    'Calculate the leght of a record
    gRecordLen = Len(gInfo)
    gCurrentRecord = 1
    
    'Get the next available file number.
    gFileNum = FreeFile
    
    'Open the file for random-access. if the file
    'does not exsist then it is created.
    Open App.Path & "\" & "bk.dat" For Random As gFileNum Len = gRecordLen
    
    'Update gCurrentRecord.
    gLastRecord = 1
    
    'Find what is the last record number of
    'the file.
    gLastRecord = FileLen(App.Path & "\" & "bk.dat") / gRecordLen
    
    'If the file was just created
    '(i.e. gLastRecord=0) then update gLastRecord
    'to 1.
    If gLastRecord = 0 Then
        gLastRecord = 1
    End If
    
    'Display the current Record.
    ShowCurrentRecord

    'Centers the form to the middle screen
    Form1.Top = (Screen.Height - Form1.Height) / 2
    Form1.Left = (Screen.Width - Form1.Width) / 2
    
End Sub




Private Sub Form_Unload(Cancel As Integer)
'Sound
    StopMe = 1
             
'Main
    SaveCurrentRecord 'Call SaveCurrentRecord Function
'    Close #gFileNum 'Close #gFileNum
'    Form2.Show
    End
'    Form2.Text3.Text = ""
'    Form2.Text2.Text = ""
'    Form2.Text2.SetFocus
    
End Sub


Private Sub MMControl1_Done(NotifyCode As Integer)
              MMControl1.Command = "close"
              If StopMe = 1 Then Exit Sub
              Flag = Flag + 1
              If Flag = 4 Then Flag = 1   'Change the (4) to the total number of WAV files + 1
              MMControl1.FileName = WaveFiles(Flag)
              MMControl1.Command = "open"
              MMControl1.Command = "play"
End Sub


Private Sub mnuAbout_Click()

If SoundIsOn = True Then
    Call sndPlaySound("D:\VB\Sounds\Auth.wav", SND_SYNC)
    Form3.Show 1
Else
    Form3.Show 1
End If

End Sub

Private Sub mnuAdd_Click()

    Command7_Click

End Sub

Private Sub mnuBackup_Click()

    'Declare variables that should be visible in all
    Dim DirResult

    'Before deleting get a confirmation from the user
    If MsgBox("Do you want to backup the current database?", vbYesNo) = vbNo Then
        'Exit the procedure without deleting
        Exit Sub
    End If

    'To physically delete the current record of STUDENT.DAT,
    'all the record of STUDENT.DAT, except the current record, are
    'copied into a temporary file (STUDENT.TMP) and then the file STUDENT.TMP is
    'copied into the STUDENT.DAT
    If Dir(App.Path & "\" & "login.dat") = App.Path & "\" & "login.dat" Then
        Name App.Path & "\" & "login.dat" As "a:\login.dat"
        Name App.Path & "\" & "bk.dat" As "a:\bk.dat"
    Else
        MsgBox "You have no database to backup..."
    End If

End Sub

Private Sub mnuBBlack_Click()
    
    Form1.BackColor = vbBlack

End Sub

Private Sub mnuBBlue_Click()
 Form1.BackColor = vbBlue

End Sub

Private Sub mnuBGreen_Click()
 Form1.BackColor = vbGreen
End Sub

Private Sub mnuBlack_Click()
   
    For counter = 0 To 6
    Label1(counter).ForeColor = vbBlack
  Next

End Sub

Private Sub mnuBlue_Click()
For counter = 0 To 6
    Label1(counter).ForeColor = vbBlue
  Next
End Sub

Private Sub mnuBPurple_Click()
    Form1.BackColor = &HC000C0
End Sub

Private Sub mnuBRed_Click()
    
    Form1.BackColor = vbRed

End Sub

Private Sub mnuBrowser_Click()
If SoundIsOn = True Then
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.FileName = WaveFiles(2)
    Flag = 1
    StopMe = 0
    MMControl1.Command = "open"
    MMControl1.Command = "play"
Else
End If
    
    Dim tmplong As Long
    Dim SW_SHOWNORMAL As Long
    Dim UrlAddress As String
    
    ' In Microsoft Windows:
    ' Specifying 1 as the second argument opens the application in
    ' normal size and gives it the focus.
    ' I will be adding new features to this one. Such as active application opening. No more adding of the browser path.
    UrlAddress = InputBox("Enter the Url address you want to go to:", "Email Now!")
    retVal = ShellExecute(tmplong, vbNullString, UrlAddress, vbNullString, vbNullString, SW_SHOWNORMAL)

End Sub

Private Sub mnuBWhite_Click()
    
    Form1.BackColor = vbWhite

End Sub

Private Sub mnuBYellow_Click()
 Form1.BackColor = vbYellow
End Sub

Private Sub mnuDelete_Click()

    Command1_Click

End Sub

Private Sub mnuEmail_Click()

If SoundIsOn = True Then
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.FileName = WaveFiles(2)
    Flag = 1
    StopMe = 0
    MMControl1.Command = "open"
    MMControl1.Command = "play"
Else
End If
    
    Dim tmplong As Long
    Dim SW_SHOWNORMAL As Long
    Dim EmailAddress As String
    
    ' In Microsoft Windows:
    ' Specifying 1 as the second argument opens the application in
    ' normal size and gives it the focus.
    ' I will be adding new features to this one. Such as active application opening. No more adding of the browser path.
    EmailAddress = InputBox("Enter the recipients email address:", "Email Now!")
    retVal = ShellExecute(tmplong, vbNullString, "mailto:" & EmailAddress, vbNullString, vbNullString, SW_SHOWNORMAL)

End Sub

Private Sub mnuExit_Click()
    
    SaveCurrentRecord 'Call SaveCurrentRecord Function
'    Close #gFileNum 'Close #gFileNum
    End
'    Form2.Show
'    Form2.Text3.Text = ""
'    Form2.Text2.Text = ""
'    Form2.Text2.SetFocus
 
End Sub

Private Sub mnuGreen_Click()
For counter = 0 To 6
    Label1(counter).ForeColor = vbGreen
  Next
End Sub

Private Sub mnuNotepad_Click()
    ' In Microsoft Windows:
    ' Specifying 1 as the second argument opens the application in
    ' normal size and gives it the focus.
    retVal = Shell("NOTEPAD.EXE NOTES.TXT", 1)  ' Run Calculator.

End Sub

Private Sub mnuPrint_Click()

    Command13_Click

End Sub

Private Sub mnuPrintSetup_Click()
' Call the printer setup procedure in the common dialog control.
On Error Resume Next
    CommonDialog1.Flags = &H40  ' Printer setup dialog box only.
    CommonDialog1.ShowPrinter
End Sub

Private Sub mnuPurple_Click()
For counter = 0 To 6
    Label1(counter).ForeColor = &HC000C0
Next
End Sub

Private Sub mnuRed_Click()

 For counter = 0 To 6
    Label1(counter).ForeColor = vbRed
  Next
  End Sub

Private Sub mnuReset_Click()
For counter = 0 To 6
    Label1(counter).ForeColor = vbBlack
    Next
    
    Form1.BackColor = &HC0C0C0
  
End Sub

Private Sub mnuRestore_Click()
    
    
    'Declare variables that should be visible in all
    Dim DirResult
    Dim TmpFileNum
    Dim TmpInfo As Info
    Dim RecNum As Long
    Dim TmpRecNum As Long

If SoundIsOn = True Then
    
    Call sndPlaySound(App.Path & "\" & "Sounds\Auth.wav", SND_SYNC)
    
Else
End If

    
    'Before deleting get a confirmation from the user
    If MsgBox("Do you want to restore the old database?", vbYesNo) = vbNo Then
        'Exit the procedure without deleting
        MsgBox "Restore has been cancelled..."
        Exit Sub
    End If


Shell ("MOVE A:\login.dat" & App.Path & "\")
    'To physically delete the current record of STUDENT.DAT,
    'all the record of STUDENT.DAT, except the current record, are
    'copied into a temporary file (STUDENT.TMP) and then the file STUDENT.TMP is
    'copied into the STUDENT.DAT
'    If Dir("login.dat") = "login.dat" Then
'        Name "a:\login.dat" As App.Path & "\" & "login.dat"
'        Name "a:\bk.dat" As App.Path & "\" & "bk.dat"
'    Else
'        MsgBox "You have no database to restore..."
'    End If
    
    'Create STUDENT.TMP with the same format as STUDENT.DAT
'    TmpFileNum = FreeFile
'    Open App.Path & "\" & "bk.tmp" For Random As TmpFileNum Len = gRecordLen

    'Copy all the records from STUDENT.TMP to STUDENT.DAT, except current record
'    RecNum = 1
'    TmpRecNum = 1
'    Do While RecNum < gLastRecord + 1
'        If RecNum <> gCurrentRecord Then
'            Get #gFileNum, RecNum, TmpInfo
'            Put #TmpFileNum, TmpRecNum, TmpInfo
'            TmpRecNum = TmpRecNum + 1
'        End If
'        RecNum = RecNum + 1
'    Loop

    'Delete STUDENT.DAT
'    Close gFileNum
'    Kill App.Path & "\" & "bk.dat"

    'Rename STUDENT.TMP to STUDENT.DAT
'    Close TmpFileNum
'    Name App.Path & "\" & "bk.tmp" As App.Path & "\" & "bk.dat"

    'Re-open the updated STUDENT.DAT file
 '   gFileNum = FreeFile
 '   Open App.Path & "\" & "bk.dat" For Random As gFileNum Len = gRecordLen



End Sub

Private Sub mnuSearch_Click()

    Command2_Click

End Sub

Private Sub mnuSndoff_Click()
    
    SoundIsOn = False
    StopMe = 1
    Call MMControl1_Done(1)
    mnuSndoff.Checked = Not mnuSndoff.Checked
    mnuSndon.Checked = Not mnuSndon.Checked
    mnuSndoff.Enabled = Not mnuSndoff.Enabled
    mnuSndon.Enabled = Not mnuSndon.Enabled

End Sub

Private Sub mnuSndon_Click()

    SoundIsOn = True
    mnuSndoff.Checked = Not mnuSndoff.Checked
    mnuSndon.Checked = Not mnuSndon.Checked
    mnuSndoff.Enabled = Not mnuSndoff.Enabled
    mnuSndon.Enabled = Not mnuSndon.Enabled

End Sub

Private Sub mnuTFont_Click()

    Form6.Show 1

End Sub

Private Sub mnuWhite_Click()
   
    For counter = 0 To 6
        
        Label1(counter).ForeColor = vbWhite
    
    Next
    
End Sub

Private Sub mnuYellow_Click()

    For counter = 0 To 6
        Label1(counter).ForeColor = vbYellow
    Next
    
End Sub

Private Sub Timer1_Timer()

    Label15 = Time$ & " " & Date

End Sub


