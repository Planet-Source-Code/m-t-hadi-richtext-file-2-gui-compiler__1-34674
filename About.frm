VERSION 5.00
Begin VB.Form about 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2880
   ClientLeft      =   2970
   ClientTop       =   2385
   ClientWidth     =   3660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2880
   ScaleWidth      =   3660
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "© 2002- Take IT Eazy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   3000
      Picture         =   "About.frx":030A
      Tag             =   "Exit"
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   3960
      Picture         =   "About.frx":0614
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   3960
      Picture         =   "About.frx":091E
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   8052
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   240
      Picture         =   "About.frx":0C28
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Internet. http://www.what-should-i-do.co.uk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "About.frx":14F2
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail.   ez@what-should-i-do.co.uk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "About.frx":17FC
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Create Standalone RichText EXE files for Windows, fast"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   380
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   2895
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Created With RichText Compile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    about.Hide
End Sub

Private Sub Command3d1_Click(Index As Integer)
    Unload Me
    Form1.SetFocus
End Sub

Private Sub Command3d1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command3d1(1).Picture = Picture2(1).Picture
End Sub

Private Sub Command3d1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label6.Caption = "Exit"
End Sub

Private Sub Command3d1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command3d1(1).Picture = Picture1(1).Picture
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label6.Caption = "Information"
End Sub

Private Sub Label1_Click()
    a = Shell("explorer http://www.what-should-i-do.co.uk", vbNormalFocus)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label6.Caption = "Information"
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label6.Caption = "Information"
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label6.Caption = "Information"
End Sub

Private Sub Label4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label6.Caption = "Information"
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label6.Caption = "Status Bar"
End Sub

Private Sub Label9_Click()
    Shell "explorer mailto:ez@what-should-i-do.co.uk"
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label6.Caption = "Information"
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)

End Sub
