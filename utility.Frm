VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "EBK Utility"
   ClientHeight    =   1770
   ClientLeft      =   3405
   ClientTop       =   2955
   ClientWidth     =   3645
   LinkTopic       =   "Form4"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1770
   ScaleWidth      =   3645
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Previous"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Login:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'All variables must be declared
Option Explicit
'Declare variables that should be visible in all
'The Procedures of the forms.
Dim gLogin As Login
Dim gFileNum As Integer
Dim gRecordLen As Long
Dim gCurrentRecord As Long
Dim gLastRecord As Long

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Public Sub ShowCurrentRecord()

    'Gets gPerson from the current record.
    Get #gFileNum, gCurrentRecord, gLogin
    
    'Fill gPerson with the currecty displayed data.
    Text1.Text = Trim(gLogin.LgName)
    Text2.Text = Trim(gLogin.LgPassword)
   
    'Display the current record number in the
    'caption of the form.
    Caption = "Electornic Book Keeper Record " + _
                    Str(gCurrentRecord) + "/" + _
                    Str(gLastRecord)

End Sub

' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Public Sub SaveCurrentRecord()

    'Fill gPerson with the currecty displayed data.
    gLogin.LgName = Text1.Text
    gLogin.LgPassword = Text2.Text
    
    'save gPerson to the current record.
    Put #gFileNum, gCurrentRecord, gLogin
    
End Sub


Private Sub Command1_Click()

    SaveCurrentRecord 'Save the current record

    'Add a new blank record
    gLastRecord = gLastRecord + 1
    gLogin.LgName = " "
    gLogin.LgPassword = " "
    
    Put #gFileNum, gLastRecord, gLogin

    'Update gCurrentRecord
    gCurrentRecord = gLastRecord
    
    'Display the record that was just created
    ShowCurrentRecord

    'Give the focus to the Text1 field
    Text1.SetFocus

End Sub


Private Sub Command2_Click()

    'Declare variables that should be visible in all
    Dim DirResult
    Dim TmpFileNum
    Dim Tmplogin As Login
    Dim RecNum As Long
    Dim TmpRecNum As Long

    'Sets the Form2 caption to "Student Locker Administration Utility: Delete" when command4 is executed
    Form4.Caption = "Electronic Book Keeper Utility"
    
    'Before deleting get a confirmation from the user
    If MsgBox("Do you want to delete this record?", vbYesNo) = vbNo Then
        'Gives the focus to the Text1 field
        Text1.SetFocus
        'Exit the procedure without deleting
        Exit Sub
    End If

    'To physically delete the current record of STUDENT.DAT,
    'all the record of STUDENT.DAT, except the current record, are
    'copied into a temporary file (STUDENT.TMP) and then the file STUDENT.TMP is
    'copied into the STUDENT.DAT
    If Dir("login.tmp") = "login.tmp" Then
        Kill "login.tmp" 'Delete Student.tmp if there is any at the time of the run
    End If

    'Create STUDENT.TMP with the same format as STUDENT.DAT
    TmpFileNum = FreeFile
    Open "login.tmp" For Random As TmpFileNum Len = gRecordLen

    'Copy all the records from STUDENT.TMP to STUDENT.DAT, except current record
    RecNum = 1
    TmpRecNum = 1
    Do While RecNum < gLastRecord + 1
        If RecNum <> gCurrentRecord Then
            Get #gFileNum, RecNum, Tmplogin
            Put #TmpFileNum, TmpRecNum, Tmplogin
            TmpRecNum = TmpRecNum + 1
        End If
        RecNum = RecNum + 1
    Loop

    'Delete STUDENT.DAT
    Close gFileNum
    If Dir("login.dat") = "login.dat" Then
        Kill "login.dat"
    End If

    'Rename STUDENT.TMP to STUDENT.DAT
    Close TmpFileNum
    Name "login.tmp" As "login.dat"

    'Re-open the updated STUDENT.DAT file
    gFileNum = FreeFile
    Open "login.dat" For Random As gFileNum Len = gRecordLen

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
    Text1.SetFocus
    
End Sub

Private Sub Command3_Click()

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
    Text1.SetFocus
 
End Sub


Private Sub Command4_Click()
        
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
    Text1.SetFocus

End Sub

Private Sub Command5_Click()
    
    SaveCurrentRecord 'Save the current record

End Sub

Private Sub Command6_Click()

    Unload Me
    ShowCurrentRecord
End Sub


Private Sub Form_Load()

    'Calculate the leght of a record
    gRecordLen = Len(gLogin)
    gCurrentRecord = 1
    
    'Get the next available file number.
    gFileNum = FreeFile
    
    'Open the file for random-access. if the file
    'does not exsist then it is created.
    Open "login.dat" For Random As gFileNum Len = gRecordLen
    
    'Update gCurrentRecord.
    gLastRecord = 1
    
    'Find what is the last record number of
    'the file.
    gLastRecord = FileLen("login.dat") / gRecordLen
    
    'If the file was just created
    '(i.e. gLastRecord=0) then update gLastRecord
    'to 1.
    If gLastRecord = 0 Then
        gLastRecord = 1
    End If
    
    'Display the current Record.
    ShowCurrentRecord

    Form4.Top = (Screen.Height - Form1.Height) / 2
    Form4.Left = (Screen.Width - Form1.Width) / 2

End Sub


