VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   Caption         =   "About"
   ClientHeight    =   2460
   ClientLeft      =   2940
   ClientTop       =   2055
   ClientWidth     =   4980
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2460
   ScaleWidth      =   4980
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Stephanie Lui"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Rodito M. Buan II"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   0
      Picture         =   "about.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim retVal As Long

Private Sub Command1_Click()
    
    Unload Form3
    
End Sub

Private Sub Form_Load()

    'Centers the form to the middle screen
    Form3.Top = (Screen.Height - Form1.Height) / 2
    Form3.Left = (Screen.Width - Form1.Width) / 2
    

End Sub


Private Sub Image1_Click()

    retVal = ShellExecute(tmplong, vbNullString, "http://telanis.hypermart.net/bk/", vbNullString, vbNullString, SW_SHOWNORMAL)

End Sub

Private Sub Label2_Click()
    
retVal = ShellExecute(tmplong, vbNullString, "mailto:jeff@kaibigan.com", vbNullString, vbNullString, SW_SHOWNORMAL)

End Sub


