VERSION 5.00
Begin VB.Form frmFind 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FindME"
   ClientHeight    =   1800
   ClientLeft      =   1470
   ClientTop       =   3345
   ClientWidth     =   3810
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
   Icon            =   "Find.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1800
   ScaleWidth      =   3810
   Tag             =   "Find Text"
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   408
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      Tag             =   "Find This"
      Text            =   "Text1"
      Top             =   120
      Width           =   2412
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   4920
      Picture         =   "Find.frx":030A
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   4920
      Picture         =   "Find.frx":0614
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   4440
      Picture         =   "Find.frx":091E
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   3240
      Picture         =   "Find.frx":0C28
      Tag             =   "Exit"
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "Find.frx":0F32
      Tag             =   "Start Search"
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   4440
      Picture         =   "Find.frx":123C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "FindME"
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
      TabIndex        =   2
      Tag             =   "Status Bar"
      Top             =   1440
      Width           =   4332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Find What"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Tag             =   "Find Text"
      Top             =   240
      Width           =   1092
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub SetWindowPos Lib "User32" (ByVal H%, ByVal HB%, ByVal X%, ByVal Y%, ByVal Cx%, ByVal Cy%, ByVal F%)

Private Sub Command3d1_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
           Case 0
                gFindString = Text1.Text
                Form1.SetFocus
                FindIt
           Case 1
                For aa = 0 To 5
                    Form1.Command3d1(aa).Enabled = True
                    Form1.Command3d1(aa).Picture = Form1.Picture1(aa).Picture
                Next
                gFindString = Text1.Text
                gFindCase = 0
                Search = False
                Unload frmFind
                Form1.SetFocus
           Case 3
    End Select
End Sub

Private Sub Command3d1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command3d1(Index).Picture = Picture2(Index).Picture
End Sub

Private Sub Command3d1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = Command3d1(Index).Tag
End Sub

Private Sub Command3d1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command3d1(Index).Picture = Picture1(Index).Picture
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Left = Lefty - frmFind.Width '+ 200
    Top = Righty - frmFind.Height '+ 200
    Call SetWindowPos(frmFind.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2)
    Search = True
    Text1.Text = ""
    gFindCase = 0
    gFindDirection = 1
    Command3d1(0).Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = frmFind.Tag
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = "Status Bar"
End Sub

Private Sub Text1_Change()
    FirstTime = True
    If Text1.Text = "" Then
       Command3d1(0).Picture = Picture2(0).Picture
       Command3d1(0).Enabled = False
    Else
       Command3d1(0).Picture = Picture1(0).Picture
       Command3d1(0).Enabled = True
    End If
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = Text1.Tag
End Sub

