VERSION 5.00
Begin VB.Form about 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2880
   ClientLeft      =   3375
   ClientTop       =   2505
   ClientWidth     =   3615
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
   Icon            =   "About3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2880
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   2880
      Picture         =   "About3.frx":030A
      Tag             =   "Exit"
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   3960
      Picture         =   "About3.frx":0614
      Top             =   3240
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   3960
      Picture         =   "About3.frx":091E
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
      TabIndex        =   4
      Top             =   2520
      Width           =   8052
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   360
      Picture         =   "About3.frx":0C28
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
      Left            =   0
      MouseIcon       =   "About3.frx":14F2
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail.   ez@what-should-i-do.co.uk"
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
      Left            =   0
      MouseIcon       =   "About3.frx":17FC
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
      Caption         =   "Create self running RTF EXE files for Windows, fast"
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
      Height          =   612
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   2892
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RichText Compiler v1.0"
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
      Height          =   492
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   3012
   End
   Begin VB.Label Label2 
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
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
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

