Attribute VB_Name = "modJStrip"
Option Explicit

' ++++++ Added this lot to reset the original file's create time ++++++

Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Currency, lpLastAccessTime As Currency, lpLastWriteTime As Currency) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Currency, lpLastAccessTime As Currency, lpLastWriteTime As Currency) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const OPEN_EXISTING = 3

' ++++++++++++++++ See GetCreateTime and SetCreateTime ++++++++++++++++

Public sFiles() As String
Public fCancelFlag As Boolean

Dim bIn() As Byte
Dim bOut() As Byte
Dim lPos As Long
Dim sLogMsg As String
Dim iLogFile As Integer
Dim lFileSize As Long
Dim lFileOutSize As Long
Dim fErrorFlag As Boolean

Const OKAY As Long = 0
Const ERROR As Long = -1
Const DONE As Long = -2
Const PROBLEM As Long = -3

Private Function DoAFile(sFileIn As String) As Long
    Dim lret As Long
    Dim fProbFlag As Boolean
    Dim sFileOut As String
    Dim sFileRename As String
    Dim dtFileCreated As Currency
    Dim fReadOnly As Boolean
    Dim idx As Integer
    For idx = Len(sFileIn) To 1 Step -1
        If Mid(sFileIn, idx, 1) = "." Then
            sFileOut = Left(sFileIn, idx) + "tmp"
            sFileRename = Left(sFileIn, idx - 1) + "_OLD.jpg"
            idx = 0
        End If
    Next idx
    lFileOutSize = 0
    fProbFlag = False
    lPos = 1
    lret = ReadFileSize(sFileIn)   ' returns filesize or ERROR
    If lret = ERROR Then
        sLogMsg = "Couldn't access input file."
        GoTo UhOh
    End If
    lFileSize = lret
    ReDim bIn(1 To lFileSize + 10) ' dim variables, 1 based, with some extra space
    ReDim bOut(1 To lFileSize + 10)
    lret = ReadFile(sFileIn)       ' read the file into bIn
    If lret = ERROR Then
        sLogMsg = "Couldn't open input file."
        GoTo UhOh
    End If
    lret = FindJpgHeader()         ' find the jpg header
    If lret = ERROR Then
        sLogMsg = "Not a valid JPEG file."
        GoTo UhOh
    End If
    lret = 0
    Do Until lret = DONE Or lret = ERROR Or lret = PROBLEM
        lret = GetMarkers()        ' copy needed data
    Loop
    If lret = ERROR Then
        sLogMsg = "Problem parsing file."
        GoTo UhOh
    End If
    If lret = PROBLEM Then fProbFlag = True
    DoEvents
    lret = WriteOutFile(sFileOut)  ' write output file
    If lret = ERROR Then
        sLogMsg = "Could not write output file"
        GoTo UhOh
    End If                         ' delete or rename input file
    lret = KillInFile(sFileIn, sFileRename, fProbFlag, dtFileCreated, fReadOnly)
    If lret = ERROR Then
        If frmMain.chkBack.Value = 0 Then
            sLogMsg = "Could not delete original file."
        Else
            sLogMsg = "Could not backup original file. This file is possibly" _
            & vbCrLf & "already stripped and a backup file may already exist."
        End If
    End If                         ' rename output file to original name
    lret = ReNameOutFile(sFileOut, sFileIn, dtFileCreated, fReadOnly)
    If lret = ERROR Then
        sLogMsg = sLogMsg & vbCrLf & _
            "Could not rename stripped file to original filename." _
            & vbCrLf & "Look for the new file named " & sFileOut
        GoTo UhOh
    End If
    If fProbFlag = True Then
        sLogMsg = "Processed successfully, but problems encountered." _
            & vbCrLf & "A backup file named " & sFileRename & " has been created," _
            & vbCrLf & "check the new " & sFileIn & " before deleting this backup."
        DoAFile = PROBLEM
    Else
        sLogMsg = "Processed successfully."
        DoAFile = OKAY
    End If
    Exit Function
UhOh:                              ' Houston, we have a problem
    MsgBox "Error processing file " & sFileIn & vbCrLf _
        & "See js.log for details.", vbOKOnly, "JpgStrip Alert"
    If iLogFile = -4 Then
        iLogFile = FreeFile
        Open "js.log" For Output As iLogFile
    End If
    fErrorFlag = True
    DoAFile = ERROR
End Function

Private Function ReadFileSize(sFileName As String) As Long
    On Error GoTo HandleIt
    ReadFileSize = FileLen(sFileName)
    Exit Function
HandleIt:
    ReadFileSize = ERROR
End Function

Private Function ReadFile(sFileName As String) As Long
    Dim iFN As Integer
    On Error GoTo HandleIt
    iFN = FreeFile
    Open sFileName For Binary As iFN
    Get #iFN, 1, bIn()
    Close iFN
    ReadFile = OKAY
    Exit Function
HandleIt:
    ReadFile = ERROR
End Function

Private Function FindJpgHeader() As Long
    Do
        If bIn(lPos) = &HFF And bIn(lPos + 1) = &HD8 And bIn(lPos + 2) = &HFF Then
            FindJpgHeader = OKAY
            Exit Do
        End If
        If lPos >= lFileSize Then
            FindJpgHeader = ERROR
            Exit Do
        End If
        lPos = lPos + 1
    Loop
    lPos = lPos + 1
End Function

Private Function GetMarkers() As Long
    Dim lSkip As Long
    Dim lTemp As Long
    Dim bFlag As Byte
    Select Case bIn(lPos)
        Case &HD8
            WriteArray &HFF
            WriteArray &HD8
            WriteArray &HFF
            lSkip = 2
            ' Okay
        Case &HE0, &HDB, &HC0 To &HCB, &HDD
            lSkip = Mult(bIn(lPos + 2), bIn(lPos + 1)) + 1
            If lSkip + lPos >= lFileSize Then GoTo Oops
            For lTemp = lPos To lPos + lSkip
                WriteArray bIn(lTemp)
            Next lTemp
            ' Okay
        Case &HDA
            bFlag = 1 ' End of file
            Do
                WriteArray bIn(lPos)
                If bIn(lPos + 1) = &HFF And bIn(lPos + 2) = &HD9 Then Exit Do
                lPos = lPos + 1
                If lPos > lFileSize Then
                    bFlag = 2 ' Problem
                    Exit Do
                End If
            Loop
            WriteArray &HFF
            WriteArray &HD9
            ' Done
        Case Else
            lSkip = Mult(bIn(lPos + 2), bIn(lPos + 1)) + 1
            If lSkip + lPos > lFileSize Then GoTo Oops
            ' Okay
    End Select
    lPos = lPos + lSkip
    Do
        If bIn(lPos) <> &HFF Then Exit Do
        lPos = lPos + 1
        If lPos > lFileSize Then GoTo Oops
        ' Okay
    Loop
    If bFlag = 0 Then GetMarkers = OKAY
    If bFlag = 1 Then GetMarkers = DONE
    If bFlag = 2 Then GetMarkers = PROBLEM
    Exit Function
Oops:
    GetMarkers = ERROR
End Function

Private Function Mult(lsb As Byte, msb As Byte) As Long
    Mult = CLng(lsb) + (CLng(msb) * 256&)
End Function

Private Sub WriteArray(bData As Byte)
    lFileOutSize = lFileOutSize + 1
    bOut(lFileOutSize) = bData
End Sub

Private Function WriteOutFile(sFileName As String) As Long
    Dim iFN As Integer
    On Error GoTo NoOpen
    iFN = FreeFile
    ReDim Preserve bOut(1 To lFileOutSize)
    Open sFileName For Binary As iFN
    On Error GoTo Opened
    Put #iFN, , bOut()
    Close iFN
    WriteOutFile = OKAY
    Exit Function
NoOpen:
    WriteOutFile = ERROR
    Exit Function
Opened:
    Close iFN
    WriteOutFile = ERROR
End Function

Private Function KillInFile(sFileName As String, sFileRename As String, fProbFlag As Boolean, dtFileCreated As Currency, fReadOnly As Boolean) As Long
    fReadOnly = (FileSystem.GetAttr(sFileName) And vbReadOnly = vbReadOnly)
    If fReadOnly Then FileSystem.SetAttr sFileName, vbNormal 'needed
    dtFileCreated = GetCreateTime(sFileName)
    If frmMain.chkBack.Value = 0 And fProbFlag = False Then
        On Error GoTo HandleIt
        Kill sFileName
        KillInFile = OKAY
        Exit Function
    Else
        On Error GoTo HandleIt
        Name sFileName As sFileRename
        If fReadOnly Then FileSystem.SetAttr sFileRename, vbReadOnly
        KillInFile = OKAY
        Exit Function
    End If
HandleIt:
    KillInFile = ERROR
End Function

Private Function ReNameOutFile(sFileOld As String, sFileNew As String, ByVal dtFileCreated As Currency, fReadOnly As Boolean) As Long
    Dim Flags As Long
    On Error GoTo HandleIt
    Name sFileOld As sFileNew
    SetCreateTime sFileNew, dtFileCreated
    Flags = IIf(fReadOnly, vbReadOnly, vbNormal) Or vbArchive
    FileSystem.SetAttr sFileNew, Flags
    ReNameOutFile = OKAY
    Exit Function
HandleIt:
    ReNameOutFile = ERROR
End Function

Private Function GetCreateTime(sFileSpec As String) As Currency
    On Error GoTo HandleIt
    ' Gets the creation time for the specified file
    Dim junk1 As Currency, junk2 As Currency
    Dim hFile As Long, dtCreationTime As Currency

    hFile = CreateFile(sFileSpec, GENERIC_READ, 0, 0, _
                     OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    GetFileTime hFile, dtCreationTime, junk1, junk2
    CloseHandle hFile
    GetCreateTime = dtCreationTime
HandleIt:
End Function

Private Sub SetCreateTime(sFileSpec As String, ByVal dtFileCreated As Currency)
    On Error GoTo HandleIt
    ' Updates the date/time for the specified file
    Dim hFile As Long, junk As Currency
    Dim dtAccTime As Currency, dtModTime As Currency

    hFile = CreateFile(sFileSpec, GENERIC_WRITE Or GENERIC_READ, _
                     0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    GetFileTime hFile, junk, dtAccTime, dtModTime
    SetFileTime hFile, dtFileCreated, dtAccTime, dtModTime
    DoEvents
    CloseHandle hFile
HandleIt:
End Sub

Public Sub DoIt(lNumber As Long)
    On Error Resume Next
    Dim lCount As Long
    Dim lret As Long
    Dim lBefore As Long
    Dim lAfter As Long
    Dim lDiff As Long
    Dim lTotal As Long
    Dim sDone As String
    Dim sAlertMsg As String
    iLogFile = -4
    fErrorFlag = False
    If frmMain.chkLog.Value = 1 Then
        iLogFile = FreeFile
        Open "js.log" For Output As iLogFile
    End If
    For lCount = 0 To lNumber - 1
        frmMain.lblMessage.Caption = sFiles(lCount)
        lBefore = ReadFileSize(sFiles(lCount))
        lret = DoAFile(sFiles(lCount))
        lAfter = ReadFileSize(sFiles(lCount))
        lDiff = lBefore - lAfter
        lTotal = lTotal + lDiff
        If frmMain.chkLog.Value = 1 Or fErrorFlag = True Then
            Dim sFileString As String
            sFileString = "**************" & vbCrLf & sFiles(lCount) & vbCrLf _
                 & sLogMsg & vbCrLf & lBefore & vbCrLf & lAfter & vbCrLf & lDiff & " bytes saved"
            Print #iLogFile, sFileString
        End If
        frmMain.pbUpdate lCount + 1, lNumber
        DoEvents
        fErrorFlag = False
        If fCancelFlag = True Then Exit For
    Next lCount
    frmMain.cmdCancel.Enabled = False
    fCancelFlag = False
    sDone = "-----------------------------" _
        & vbCrLf & " Files processed: " & lCount _
        & vbCrLf & " Total bytes saved: " & lTotal _
        & vbCrLf & "-----------------------------" & vbCrLf
    If frmMain.chkBack.Value = 1 Then
        sDone = sDone & " Backup file(s) have been created." _
            & vbCrLf & " You should delete all OLD.jpg files" _
            & vbCrLf & " after test viewing the new files." _
            & vbCrLf & "-----------------------------" & vbCrLf
    End If
    If iLogFile <> -4 Then
        Print #iLogFile, vbCrLf & sDone
        Close iLogFile
        sDone = sDone & " A file named js.log has been" _
            & vbCrLf & " saved in the current folder." _
            & vbCrLf & "-----------------------------"
    End If
    MsgBox sDone, vbOKOnly, "JpgStrip Done"
End Sub
