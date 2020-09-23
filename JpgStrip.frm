VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "JPGstrip"
   ClientHeight    =   4095
   ClientLeft      =   2595
   ClientTop       =   4185
   ClientWidth     =   4830
   Icon            =   "JpgStrip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      Height          =   805
      Left            =   3570
      TabIndex        =   8
      Top             =   2880
      Width           =   1180
      Begin VB.CheckBox chkBack 
         Caption         =   "Backup"
         Height          =   225
         Left            =   130
         TabIndex        =   10
         Top             =   480
         Width           =   960
      End
      Begin VB.CheckBox chkLog 
         Caption         =   "Logging"
         Height          =   225
         Left            =   130
         TabIndex        =   9
         Top             =   200
         Value           =   1  'Checked
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   1285
      TabIndex        =   4
      Top             =   3315
      Width           =   945
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   330
      Left            =   2320
      TabIndex        =   5
      Top             =   3315
      Width           =   945
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   330
      Left            =   250
      TabIndex        =   3
      Top             =   3315
      Width           =   945
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   85
      Width           =   2310
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   60
      TabIndex        =   1
      Top             =   520
      Width           =   2310
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   2460
      MultiSelect     =   2  'Extended
      Pattern         =   "*.jp*"
      TabIndex        =   2
      Top             =   60
      Width           =   2310
   End
   Begin VB.PictureBox pbProgress 
      Align           =   2  'Align Bottom
      FillColor       =   &H8000000D&
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   4770
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3795
      Width           =   4830
   End
   Begin VB.Label lblMessage 
      Height          =   225
      Left            =   105
      TabIndex        =   7
      Top             =   3000
      Width           =   4650
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim lFormHeight As Long
Dim lFormWidth As Long
Dim lDirHeight As Long
Dim lFileHeight As Long
Dim lBorderWidth As Long
Dim lDiff As Long
Dim lLabelTop As Long
Dim lButtonTop As Long
Dim lOptionsTop As Long
Dim lOptionsLeft As Long
Dim iNumSelected As Integer
Dim bWorking As Boolean

Option Explicit

Private Sub Form_Load()
    frmMain.Visible = False
    SetResize
    cmdCancel.Enabled = False
    frmMain.Visible = True
    CountSelectedFiles
End Sub

Private Sub Dir1_Change()
    On Error GoTo er
    ChDir Dir1.Path
    RefreshFiles
    Exit Sub
er:
    NewPath
End Sub

Private Sub Drive1_Change()
    On Error GoTo er
    ChDrive Drive1.Drive
    RefreshFiles
    Exit Sub
er:
    NewPath
End Sub

Private Sub NewPath()
    Drive1.Drive = App.Path
    Dir1.Path = App.Path
    RefreshFiles
End Sub

Sub RefreshFiles()
    Drive1.Drive = CurDir
    Dir1.Path = CurDir
    File1.Path = CurDir
    Drive1.Refresh
    Dir1.Refresh
    File1.Refresh
    CountSelectedFiles
End Sub

Private Sub File1_Click()
    CountSelectedFiles
End Sub

Private Sub File1_DblClick()
    cmdStart_Click
End Sub

Private Sub CountSelectedFiles()
    Dim iCount As Integer
    iNumSelected = 0
    For iCount = 0 To File1.ListCount - 1
        If File1.Selected(iCount) = True Then
            iNumSelected = iNumSelected + 1
        End If
    Next iCount
    lblMessage.Caption = iNumSelected & "  selected,    " & File1.ListCount & "  total"
    If iNumSelected < 1 Then
        cmdStart.Enabled = False
    Else
        cmdStart.Enabled = True
    End If
End Sub

Private Sub SetResize()
    lDirHeight = frmMain.Height - Dir1.Height
    lFileHeight = frmMain.Height - File1.Height
    lBorderWidth = frmMain.Width - (Dir1.Width + File1.Width)
    lLabelTop = frmMain.Height - lblMessage.Top
    lButtonTop = frmMain.Height - cmdStart.Top
    lOptionsTop = frmMain.Height - fraOptions.Top
    lOptionsLeft = frmMain.Width - fraOptions.Left
    frmMain.Width = 6000
End Sub

Sub Form_Resize()
    ' Do only if the form is not minimized
    If Me.WindowState <> vbMinimized Then
        If frmMain.Height < 4000 Then
            frmMain.Height = 4000
        End If
        If frmMain.Width < 4950 Then
            frmMain.Width = 4950
        End If
        If lFormHeight <> frmMain.Height Then
            lFormHeight = frmMain.Height
            Dir1.Height = lFormHeight - lDirHeight
            File1.Height = lFormHeight - lFileHeight
            lblMessage.Top = lFormHeight - lLabelTop
            cmdStart.Top = lFormHeight - lButtonTop
            cmdCancel.Top = lFormHeight - lButtonTop
            cmdQuit.Top = lFormHeight - lButtonTop
            fraOptions.Top = lFormHeight - lOptionsTop
        End If
        If lFormWidth <> frmMain.Width Then
            lFormWidth = frmMain.Width
            lDiff = lFormWidth - (Dir1.Width + File1.Width + lBorderWidth)
            lDiff = lDiff / 3    ' share the difference
            Dir1.Width = Dir1.Width + lDiff
            Drive1.Width = Drive1.Width + lDiff
            File1.Left = File1.Left + lDiff
            File1.Width = File1.Width + lDiff * 2
            fraOptions.Left = lFormWidth - lOptionsLeft
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If bWorking = True Then
        Cancel = vbCancel
    End If
End Sub

Private Sub cmdCancel_Click()
    fCancelFlag = True
End Sub

Private Sub cmdStart_Click()
    bWorking = True
    MousePointer = 13
    ChDrive Drive1.Drive
    ChDir Dir1.Path
    Dir1.Enabled = False
    Drive1.Enabled = False
    File1.Enabled = False
    cmdQuit.Enabled = False
    cmdStart.Enabled = False
    cmdCancel.Enabled = True
    chkLog.Enabled = False
    chkBack.Enabled = False
    Dim iCount As Integer
    Dim iFiles As Integer
    ReDim sFiles(iNumSelected)
    For iCount = 0 To File1.ListCount - 1
        If File1.Selected(iCount) = True Then
            sFiles(iFiles) = File1.List(iCount)
            iFiles = iFiles + 1
        End If
    Next iCount
    DoIt (iFiles)
    cmdQuit.Enabled = True
    cmdCancel.Enabled = False
    chkLog.Enabled = True
    chkBack.Enabled = True
    Dir1.Enabled = True
    Drive1.Enabled = True
    File1.Enabled = True
    MousePointer = vbDefault
    bWorking = False
    pbProgress.Cls
    RefreshFiles
End Sub

Public Sub pbUpdate(progress As Long, total As Long)
    pbProgress.ScaleWidth = total
    pbProgress.Line (0, 0)-(progress, pbProgress.ScaleHeight), pbProgress.ForeColor, BF
    DoEvents
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

