Attribute VB_Name = "Module1"
Global gFindString, gFindCase As Integer, gFindDirection As Integer
Global gCurPos As Integer, gFirstTime As Integer
Global YES
Global tempfile
Global Search
Global first_srch
Global Const DRIVE_REMOVABLE = 2
Global Const DRIVE_FIXED = 3
Global Const DRIVE_REMOTE = 4
Global Lefty
Global Righty

Sub FindIt()
    Dim start, pos, findstring, sourcestring, msg, response, Offset
    Form1.SetFocus
    If (gCurPos = Form1.ActiveControl.SelStart) Then
        Offset = 1
    Else
        Offset = 0
    End If
    If gFirstTime Then Offset = 0
    start = Form1.ActiveControl.SelStart + Offset
    If gFindCase Then
        findstring = gFindString
        sourcestring = Form1.ActiveControl.Text
    Else
        findstring = UCase(gFindString)
        sourcestring = UCase(Form1.ActiveControl.Text)
    End If
    If gFindDirection = 1 Then
        pos = InStr(start + 1, sourcestring, findstring)
    Else
        For pos = start - 1 To 0 Step -1
            If pos = 0 Then Exit For
            If Mid(sourcestring, pos, Len(findstring)) = findstring Then Exit For
        Next
    End If
    If pos Then
        Form1.ActiveControl.SelStart = pos - 1
        Form1.ActiveControl.SelLength = Len(findstring)
    Else
        CR$ = Chr$(13) + Chr$(10)
        TheMessage$ = "Text Not Found."
        TheStyle = 16
        TheTitle$ = "FindME"
        TheAnswer = MsgBox(TheMessage$, TheStyle, TheTitle$)
    End If
    gCurPos = Form1.ActiveControl.SelStart
    gFirstTime = False
End Sub

Function MyDriveType(ByVal DR As String) As String
    nDrive% = Asc(UCase(DR)) - 65 'A=0, B=1, ...    x% = GetDriveType(nDrive%)
    Select Case X%
    Case DRIVE_REMOVABLE '= 2
         MyDriveType = "disk can be removed from the drive"
    Case DRIVE_FIXED '= 3
         MyDriveType = "disk cannot be removed from the drive"
    Case DRIVE_REMOTE '= 4
         MyDriveType = "drive is a remote, or network, drive"
    Case Else
         MyDriveType = "can't determine drive type"
    End Select
End Function

Function Validate_Drive(ByVal strDrive As String)
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

