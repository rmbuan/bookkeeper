VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Protection"
   ClientHeight    =   2865
   ClientLeft      =   3150
   ClientTop       =   1725
   ClientWidth     =   3675
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   330
      Left            =   1920
      TabIndex        =   7
      Top             =   2160
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New Account"
      Height          =   330
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   330
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1590
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "login.frx":0000
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Label2 
      Caption         =   "Please Enter Your Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Choose your Login Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   1410
   End
End
Attribute VB_Name = "Form2"
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

Private Sub Command1_Click()
    
    Dim Found As Integer
    Dim RecNum As Long
    Dim Tmplogin As Login
    

    Found = False
    
    'Search for the name that the user entered
    For RecNum = 1 To gLastRecord
        Get #gFileNum, RecNum, Tmplogin
        If UCase$(Text1.Text) = UCase$(Trim(Tmplogin.LgName)) Then
            If UCase$(Text2.Text) = UCase$(Trim(Tmplogin.LgPassword)) Then
                Found = True
                Exit For
            End If
        End If
    Next

    'if the name was found, then display the record of the found name
    If Found = True Then
        'Close gFileNum
        Form2.Hide
        Form1.Show 1
   Else
        'Popups a message box that tells that use what the name he has search is not found on the list/database
        MsgBox "password incorrect", , "Login"
    End If
 
End Sub

Private Sub Command2_Click()
    
    Dim LoginPassword As String
    
    LoginPassword = InputBox("Please Enter Your Admin Password:", "Admin Login")

    'If the user did not enter a name, exit from this procedure
    If UCase$(LoginPassword) = "SECRETPASS" Then
        'Gives the focus to the Text1 field
        Close gFileNum
        Form4.Show
        'Exit the procedure without deleting
    Else
        MsgBox "YOUR ATTEMPT HAS BEEN DENIED", , "Admin Login"
    End If


End Sub

Private Sub Command3_Click()

    End
    
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
       
    'Centers the form to the middle screen
    Form2.Top = (Screen.Height - Form1.Height) / 2
    Form2.Left = (Screen.Width - Form1.Width) / 2
 
'    Form2.Text2.SetFocus

End Sub


Private Sub Form_Unload(Cancel As Integer)

    End

End Sub


