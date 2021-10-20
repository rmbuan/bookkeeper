VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Fonts Utility"
   ClientHeight    =   1545
   ClientLeft      =   1875
   ClientTop       =   1785
   ClientWidth     =   5115
   LinkTopic       =   "Form6"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1545
   ScaleWidth      =   5115
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Font Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gNumOfScreenFonts
Dim gNumOfPrinterFonts

Private Sub Command1_Click()

    Label1.Font = Combo1.Text

End Sub

Private Sub Command2_Click()

Dim apply

    For apply = 0 To 6
        
        Form1.Label1(apply).Font = Combo1.Text
        Form1.Text6.Font = Combo1.Text
    
    Next
        
   
    
    
End Sub

Private Sub Command3_Click()
 Form6.Hide
End Sub

Private Sub Form_Load()

Dim I

gNumOfScreenFonts = Screen.FontCount - 1

For I = 0 To gNumOfScreenFonts - 1 Step 1
    Combo1.AddItem Printer.Fonts(I)
Next

Combo1.ListIndex = 0

'Label1.FontName = Combo1.Text

End Sub
