VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RichText View"
   ClientHeight    =   6315
   ClientLeft      =   1095
   ClientTop       =   2010
   ClientWidth     =   7530
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
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   36.875
   ScaleMode       =   0  'User
   ScaleWidth      =   1.169
   Tag             =   "Exit to Windows"
   Begin RichTextLib.RichTextBox TWindow 
      CausesValidation=   0   'False
      Height          =   5535
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   9763
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0D10
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   6840
      Picture         =   "Form1.frx":0DE7
      Tag             =   "Exit"
      Top             =   5160
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   6840
      Picture         =   "Form1.frx":10F1
      Tag             =   "About Information"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   8520
      Picture         =   "Form1.frx":13FB
      Tag             =   "Print Text"
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   6840
      Picture         =   "Form1.frx":1705
      Tag             =   "Save to File"
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   6840
      Picture         =   "Form1.frx":1A0F
      Tag             =   "Copy Text"
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Command3d1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   8640
      Picture         =   "Form1.frx":1D19
      Tag             =   "Search Text"
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   9120
      Picture         =   "Form1.frx":2023
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   9120
      Picture         =   "Form1.frx":232D
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   9120
      Picture         =   "Form1.frx":2637
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   9120
      Picture         =   "Form1.frx":2941
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   9120
      Picture         =   "Form1.frx":2C4B
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Picture2 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   9120
      Picture         =   "Form1.frx":2F55
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   8400
      Picture         =   "Form1.frx":325F
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   8400
      Picture         =   "Form1.frx":3569
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   8400
      Picture         =   "Form1.frx":3873
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   8400
      Picture         =   "Form1.frx":3B7D
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   8400
      Picture         =   "Form1.frx":3E87
      Top             =   840
      Width           =   480
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   8400
      Picture         =   "Form1.frx":4191
      Top             =   360
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
      Left            =   8400
      TabIndex        =   1
      Top             =   5520
      Width           =   1332
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "RichText View"
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
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp
Private Declare Function SendMessage% Lib "User32" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam As Any)
Private Declare Function GetDriveType Lib "Kernel32" (ByVal nDrive As Integer) As Integer
'int nDrive: 0 = A, 1 = B, and so on
'GetDriveType return values
Dim Buffer As String
Dim resizing As Integer
Const EM_GETLINE = &H400 + 20
Const EM_GETLINECOUNT = &H400 + 10
Const MAX_CHAR_PER_LINE = 68

Private Sub Command3d1_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
           Case 0
                For aa = 0 To 5
                    Command3d1(aa).Picture = Picture2(aa).Picture
                    Command3d1(aa).Enabled = False
                Next
                Search = True
                frmFind.Show
           Case 1
                On Error Resume Next
                Clipboard.Clear
                Clipboard.SetText TWindow.Text
           Case 2
                MSave2 = "READ-ME.RTF"
                TWindow.SaveFile MSave2
                TheMessage$ = "File Saved as  " & MSave2
                TheStyle = 64
                TheTitle$ = "Save"
                MsgBox TheMessage$, TheStyle, TheTitle$
           Case 3
                Dim ndx&, N&
                ndx& = fGetLineCount()
                Printer.FontName = "Arial"
                Printer.FontSize = 14
                For N& = 1 To ndx&
                    On Error Resume Next
                    Buffer = fGetLine(N& - 1)
                    Printer.Print Buffer
                Next N&
                Printer.EndDoc
           Case 4
                Search = False
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

Private Function fGetLine$(Linenumber As Long)
    Dim byteLo%, byteHi%, X%
    Dim Buffer$
    byteLo% = MAX_CHAR_PER_LINE And (255)
    byteHi% = Int(MAX_CHAR_PER_LINE / 256)
    Buffer$ = Chr$(byteLo%) + Chr$(byteHi%) + Space$(MAX_CHAR_PER_LINE - 2)
    X% = SendMessage(TWindow.hWnd, EM_GETLINE, Linenumber, Buffer$)
    fGetLine$ = Left$(Buffer$, X%)
End Function

Private Function fGetLineCount%()
    Dim lcount%
    lcount% = SendMessage(TWindow.hWnd, EM_GETLINECOUNT, 0&, 0&)
    aGetLineCount.Caption = "GetLineCount = " + Str$(lcount%)
    fGetLineCount% = lcount%
End Function

Private Sub Form_Load()
    On Error Resume Next
    If App.PrevInstance Then
       End
    End If
    ChDir App.Path
    ChDrive App.Path
    StartPosition = 0
    Clipboard.Clear
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    Lefty = Left + Width
    Righty = Top + Height
    Search = False
    first_srch = ""
    If Command <> "" Then
       MOption = Val(Command)
    End If
    Text2 = Chr(123) & Chr(92) & Chr(114) & Chr(116) & Chr(102)
    Text1 = App.Path + "\" + App.EXEName + ".EXE"
    ' I have to bypass the file error handler to make it work? Very weird...
    On Error Resume Next
    tempfile = App.Path
    HeaderExists = False
    FileLocation = 1
    filechunk$ = ""
    DoEvents
    Open Text1 For Binary As #1
    FileSize = LOF(1)
    filechunk$ = String(FileSize, 0)
    Get #1, FileLocation, filechunk$
    HeaderInFile = InStr(filechunk$, Text2)
    If HeaderInFile > 0 Then
       HeaderExists = True
       TempStartData$ = Mid(filechunk$, HeaderInFile)
       Mtempfile = "5£TEMP£5.$$$"
       Kill Mtempfile
       Open Mtempfile For Binary As #6
       Put #6, , TempStartData$
       Close #6
       Found = True
       TWindow.LoadFile Mtempfile
       Kill Mtempfile
    End If
SkipNext:
    If Not HeaderExists Then
       CR$ = Chr$(13) + Chr$(10)
       TheMessage$ = "No embeded text object found!"
       TheStyle = 16
       TheTitle$ = "Error!"
       TheAnswer = MsgBox(TheMessage$, TheStyle, TheTitle$)
       End
    End If
    Exit Sub
FileError:
       MsgBox "Error " + Trim$(Str$(Err)) + ": " + Error$(Err), 16
       Exit Sub
       Resume Abort
Abort:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = "RichText View"
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.Caption = "Status Bar"
End Sub

Private Sub TWindow_KeyDown(keycode As Integer, Shift As Integer)
    keycode = 0
End Sub

Private Sub TWindow_KeyPress(keyascii As Integer)
    keyascii = 0
End Sub

Private Sub TWindow_KeyUp(keycode As Integer, Shift As Integer)
    keycode = 0
End Sub

Private Function Validate_Drive(ByVal strDrive As String)
    On Error GoTo BAD2
    Dim strOldDrive As String
    'strOldDrive = Get_Drive_Name(CurDir$)
    ChDrive (strDrive)
    ChDrive (strOldDrive)
    On Error GoTo 0
    Validate_Drive = True
    Exit Function
BAD2:
    Validate_Drive = False
    Resume Exit2
Exit2:
    Exit Function
End Function

Function TextFile(ByRef pStrFileName As String)
    Dim llngFileNum As Long
    Dim llngFileLen As Long
    On Error Resume Next
    llngFileLen = 0
    llngFileLen = FileLen(pStrFileName)
    If llngFileLen = 0 Then Exit Function
    On Error GoTo 0
    llngFileNum = FreeFile
    Open pStrFileName For Input As #llngFileNum
    TextFile = Input(llngFileLen, #llngFileNum)
    Close #lngFileNum
End Function
