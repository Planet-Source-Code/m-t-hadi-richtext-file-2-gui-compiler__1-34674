VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ReadME Compiler"
   ClientHeight    =   5505
   ClientLeft      =   1410
   ClientTop       =   1830
   ClientWidth     =   6300
   ClipControls    =   0   'False
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
   Icon            =   "Part1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Part1.frx":08CA
   ScaleHeight     =   5505
   ScaleWidth      =   6300
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin VB.TextBox TWindow 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3552
      HideSelection   =   0   'False
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Part1.frx":0D10
      Top             =   3000
      Width           =   3312
   End
   Begin VB.FileListBox filList 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   3240
      TabIndex        =   6
      Tag             =   "Select File"
      Top             =   1680
      Width           =   2892
   End
   Begin VB.DirListBox dirList 
      Appearance      =   0  'Flat
      Height          =   1752
      Left            =   120
      TabIndex        =   5
      Tag             =   "Select Directory"
      Top             =   1680
      Width           =   2892
   End
   Begin VB.DriveListBox drvList 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   120
      TabIndex        =   4
      Tag             =   "Select Drive"
      Top             =   1320
      Width           =   2892
   End
   Begin VB.TextBox txtSearchSpec 
      Appearance      =   0  'Flat
      Height          =   288
      HideSelection   =   0   'False
      Left            =   3240
      TabIndex        =   3
      Tag             =   "File Specification"
      Text            =   "*.rtf"
      Top             =   1320
      Width           =   2892
   End
   Begin VB.Label Label3 
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
      Left            =   720
      TabIndex        =   11
      Top             =   410
      Width           =   3615
   End
   Begin VB.Label Label6 
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
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "Part1.frx":0D82
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   8040
      Picture         =   "Part1.frx":164C
      Top             =   660
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   7320
      Picture         =   "Part1.frx":1956
      Top             =   660
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   4200
      Picture         =   "Part1.frx":1C60
      Tag             =   "Test Compilation"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   7320
      Picture         =   "Part1.frx":1F6A
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   7320
      Picture         =   "Part1.frx":2274
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   7320
      Picture         =   "Part1.frx":257E
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   7320
      Picture         =   "Part1.frx":2888
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   8040
      Picture         =   "Part1.frx":2B92
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   8040
      Picture         =   "Part1.frx":2E9C
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   8040
      Picture         =   "Part1.frx":31A6
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   8040
      Picture         =   "Part1.frx":34B0
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "Part1.frx":37BA
      Tag             =   "Create ReadME EXE"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   6840
      Picture         =   "Part1.frx":3AC4
      Tag             =   "Print Order Form"
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   4920
      Picture         =   "Part1.frx":3DCE
      Tag             =   "About Information"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   5640
      Picture         =   "Part1.frx":40D8
      Tag             =   "Exit"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Label aGetLineCount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   7200
      TabIndex        =   9
      Top             =   4800
      Width           =   1332
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Compile"
      ForeColor       =   &H80000008&
      Height          =   372
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   4560
      Width           =   1692
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drive"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   612
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Text File"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   2892
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "ReadME Compiler"
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
      TabIndex        =   0
      Tag             =   "Status Bar"
      Top             =   5160
      Width           =   6492
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage% Lib "user32" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam As Any)
Dim Buffer As String
Dim resizing As Integer
Const EM_GETLINE = &H400 + 20
Const EM_GETLINECOUNT = &H400 + 10
Const MAX_CHAR_PER_LINE = 68

Private Sub Command3d1_Click(Index As Integer)
    Dim i$, i2$, msg As String, X As Integer, filedata As String, msg2 As String, Y As Integer
    Dim iFreeFile As Integer
    Dim iFreeFile2 As Integer
    Dim sBuffer As String
    Dim sBefore As String
    String2File = False
    If waga_waga Then
       Index = 0
    End If
    On Error Resume Next
    Select Case Index
           Case 0
                waga_waga = False
                mmm1 = App.Path + "\RTF-VIEW.EXE"
                mmm2 = App.Path + "\RTF-VIEW.DLL"
                mmm3 = FxName
                If File(mmm1) Then
                   CR$ = Chr$(13) + Chr$(10)
                   TheMessage$ = "Overwrite Existing RTF-VIEW.EXE File?"
                   TheStyle = 276
                   TheTitle$ = "File Found"
                   TheAnswer = MsgBox(TheMessage$, TheStyle, TheTitle$)
                   If TheAnswer = 6 Then
                      Kill mmm1
                   Else
                      GoTo Finale
                   End If
                End If
                FileCopy mmm2, mmm1
                yyy = "z0RbA1 " + ReadFile(mmm3)
                i$ = mmm1
                i2$ = mmm2
                Open i$ For Output As #1
                Open i2$ For Binary As #2
                Do While Not EOF(2)
                   filedata = Input$(2000, #2)
                   msg = filedata
                   msg2 = msg2 + msg
                   Print #1, msg2;
                   msg2 = ""
                   If Len(msg) > 2000 Then
                      msg = ""
                   End If
                Loop
                Print #1, yyy
                Close #2
                Close #1
'**********************************
                CR$ = Chr$(13) + Chr$(10)
                TheMessage$ = "File Compiled Successfully."
                TheStyle = 64
                TheTitle$ = "Finished"
                TheAnswer = MsgBox(TheMessage$, TheStyle, TheTitle$)
                mmm1 = App.Path + "\RTF-VIEW.EXE"
                If File(mmm1) Then
                   Command3d1(1).Enabled = True
                   Command3d1(1).Picture = Picture1(1).Picture
                Else
                   Command3d1(1).Enabled = False
                   Command3d1(1).Picture = Picture2(1).Picture
                End If
'**********************************
           Case 1
                mmm1 = App.Path + "\RTF-VIEW.EXE"
                If File(mmm1) Then
                   If Not Shell(mmm1, 1) Then
                      Form1.WindowState = 1
                      Timer1.Enabled = True
                   Else
                      ' Error
                   End If
                End If
           Case 3 ' I can't remember what I was doing then, but this suppose to print something
                Dim ndx&, N&
                ndx& = fGetLineCount()
                For N& = 1 To ndx&
                    On Error Resume Next
                    Buffer = fGetLine(N& - 1)
                    Printer.Print Buffer
                Next N&
                Printer.EndDoc
           Case 4
                about.Show 1
           Case 5
                CR$ = Chr$(13) + Chr$(10)
                TheMessage$ = "Exit Programme?"
                TheStyle = 292
                TheTitle$ = "Exit"
                TheAnswer = MsgBox(TheMessage$, TheStyle, TheTitle$)
                If TheAnswer = 6 Then
                   End
                End If
           Case 6
    End Select
Finale:
waga_waga = False
End Sub

Private Sub Command3d1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command3d1(Index).Picture = Picture2(Index).Picture
End Sub

Private Sub Command3d1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2 = Command3d1(Index).Tag
End Sub

Private Sub Command3d1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command3d1(Index).Picture = Picture1(Index).Picture
End Sub

Private Sub dirList_Change()
    filList.Path = dirList.Path
End Sub

Private Sub dirList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2 = dirList.Tag
End Sub

Private Sub drvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub
DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub
Private Function fGetLine$(Linenumber As Long)
    Dim byteLo%, byteHi%, X%
    Dim Buffer$
    byteLo% = MAX_CHAR_PER_LINE And (255)
    byteHi% = Int(MAX_CHAR_PER_LINE / 256)
    Buffer$ = Chr$(byteLo%) + Chr$(byteHi%) + Space$(MAX_CHAR_PER_LINE - 2)
    X% = SendMessage(TWindow.hwnd, EM_GETLINE, Linenumber, Buffer$)
    fGetLine$ = Left$(Buffer$, X%)
End Function

Private Function fGetLineCount%()
    Dim lcount%
    lcount% = SendMessage(TWindow.hwnd, EM_GETLINECOUNT, 0&, 0&)
    aGetLineCount.Caption = "GetLineCount = " + Str$(lcount%)
    fGetLineCount% = lcount%
End Function

Private Function File(ByVal FileName As String) As Integer
    Dim fileFile As Integer
    fileFile = FreeFile
    On Error Resume Next
    Open FileName For Input As fileFile
    If Err Then
        File = False
    Else
        Close fileFile
        File = True
    End If
End Function

Private Sub filList_Click()
    On Error Resume Next
    Dim msg, TimeStamp
    If Right$(filList.Path, 1) <> "\" Then
       txtSearchSpec = filList.Path + "\" + filList.FileName
    Else
       txtSearchSpec = filList.Path + filList.FileName
    End If
    FxName = txtSearchSpec
    MAD_Size = FileLen(FxName)
    If MAD_Size > 30000 Then
       Command3d1(0).Enabled = False
       Command3d1(0).Picture = Picture2(0).Picture
       CR$ = Chr$(13) + Chr$(10)
       TheMessage$ = "File too large to embed, Please reduce size." '+ CR$
       TheStyle = 16
       TheTitle$ = "Error! - File Size  " & MAD_Size & "  Max 30Kb"
       MsgBox TheMessage$, TheStyle, TheTitle$
    Else
       Command3d1(0).Enabled = True
       Command3d1(0).Picture = Picture1(0).Picture
    End If
End Sub

Private Sub filList_DblClick()
    On Error Resume Next
    Dim msg, TimeStamp
    If Right$(filList.Path, 1) <> "\" Then
       txtSearchSpec = filList.Path + "\" + filList.FileName
    Else
       txtSearchSpec = filList.Path + filList.FileName
    End If
    FxName = txtSearchSpec
    MAD_Size = FileLen(FxName)
    'we don't need this with a richtextbox
    'If MAD_Size > 30000 Then
    '   Command3d1(0).Enabled = False
    '   Command3d1(0).Picture = Picture2(0).Picture
    '   CR$ = Chr$(13) + Chr$(10)
    '   TheMessage$ = "File too large to embed, Please reduce size." '+ CR$
    '   TheStyle = 16
    '   TheTitle$ = "Error! - File Size  " & MAD_Size & "  Max 30Kb"
    '   MsgBox TheMessage$, TheStyle, TheTitle$
    'Else
       Command3d1(0).Enabled = True
       Command3d1(0).Picture = Picture1(0).Picture
       waga_waga = True
       Command3d1_Click (0)
    'End If
End Sub

Private Sub filList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2 = filList.Tag
End Sub

Private Sub Form_Load()
    waga_waga = False
    If App.PrevInstance Then
       End
    End If
    ChDir App.Path
    ChDrive App.Path
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    FxName = ""
    TWindow.Text = ""
    TWindow.Text = TWindow.Text + ""
    Call drvList_Change
    filList.Path = dirList.Path
    txtSearchSpec.Text = "*.rtf"
    filList.Pattern = txtSearchSpec.Text
    mmm2 = App.Path + "\RTF-VIEW.DLL"
    If Not File(mmm2) Then
       CR$ = Chr$(13) + Chr$(10)
       TheMessage$ = "Missing  RTF-VIEW.DLL  File!"
       TheStyle = 16
       TheTitle$ = "Error!"
       TheAnswer = MsgBox(TheMessage$, TheStyle, TheTitle$)
       End
    End If
    mmm1 = App.Path + "\RTF-VIEW.EXE"
    If File(mmm1) Then
       Command3d1(1).Enabled = True
       Command3d1(1).Picture = Picture1(1).Picture
    Else
       Command3d1(1).Enabled = False
       Command3d1(1).Picture = Picture2(1).Picture
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2 = "RichText Compiler"
End Sub

Private Function ReadFile(ByVal sFileName As String) As String
    Dim fhFile As Integer
    fhFile = FreeFile
    Open sFileName For Binary As #fhFile
    ReadFile = Input$(LOF(fhFile), fhFile)
    Close #fhFile
End Function

Private Sub Timer1_Timer()
   Dim hwnd&
   hwnd = FindWindowA(vbNullString, "RichText View")
   If hwnd = 0 Then
      Form1.WindowState = 0
      Form1.SetFocus
      Timer1.Enabled = False
   End If
End Sub

Private Sub txtSearchSpec_Change()
    On Error GoTo ErrHandler
    If txtSearchSpec = "" Then
       txtSearchSpec = "*.rtf"
    End If
    txtSearchSpec.Text = Trim(LCase(txtSearchSpec.Text))
    filList.Pattern = txtSearchSpec.Text
ErrHandler:
    Resume Next
End Sub

Private Sub txtSearchSpec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2 = txtSearchSpec.Tag
End Sub

