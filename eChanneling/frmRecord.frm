VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRecord 
   Caption         =   "Record Announcement"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   Picture         =   "frmRecord.frx":0000
   ScaleHeight     =   1605
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   1200
   End
   Begin VB.TextBox txtAnnouncement 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin btButtonEx.ButtonEx bttnRecord 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Record"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnPause 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Pause"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnResume 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Resume"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnStop 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Stop"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    Dim Cmd As String
    Private Declare Function mciSendStringA Lib "WinMM" _
        (ByVal MCIcommand As String, ByVal returnStr As String, _
        ByVal returnLength As Integer, ByVal callBack As Integer) As Long
    
    Private Declare Function mciGetErrorStringA Lib "WinMM" _
        (ByVal error As Long, ByVal Buffer As String, _
        ByVal length As Integer) As Integer
    Dim errorCode As Integer
    Dim returnStr As String * 255
    Dim returnCode As Integer
    Dim errorStr As String * 255
    Dim EllapsedSeconds As Long

Private Sub bttnPause_Click()
    Dim TR As Integer
    
    bttnRecord.Enabled = False
    bttnPause.Enabled = False
    bttnResume.Enabled = True
    bttnStop.Enabled = False
    bttnSave.Enabled = False
    
    Cmd = "Pause NewAnnouncement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    If errorCode <> 0 Then
        TR = MsgBox(errorStr, vbCritical, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
End Sub

Private Sub bttnRecord_Click()
    Dim TR As String
    Timer1.Interval = 1000
    bttnRecord.Enabled = False
    bttnPause.Enabled = True
    bttnResume.Enabled = False
    bttnStop.Enabled = True
    bttnSave.Enabled = False

    Cmd = "Close NewAnnouncement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    
    Cmd = "OPEN New Type waveaudio alias NewAnnouncement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    If errorCode <> 0 Then
        TR = MsgBox(errorStr, vbCritical, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Cmd = "set NewAnnouncement bitspersample 16"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    If errorCode <> 0 Then
        TR = MsgBox(errorStr, vbCritical, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Cmd = "set NewAnnouncement samplespersec 44100"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    If errorCode <> 0 Then
        TR = MsgBox(errorStr, vbCritical, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    Cmd = "Record NewAnnouncement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    If errorCode <> 0 Then
        TR = MsgBox(errorStr, vbCritical, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
End Sub

Private Sub bttnResume_Click()
    Dim TR As Integer
    
    bttnRecord.Enabled = False
    bttnPause.Enabled = True
    bttnResume.Enabled = False
    bttnStop.Enabled = True
    bttnSave.Enabled = False
    
    Cmd = "Resume NewAnnouncement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    If errorCode <> 0 Then
        TR = MsgBox(errorStr, vbCritical, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
End Sub

Private Sub bttnSave_Click()
On Error GoTo EH
Dim TR As Integer

If Trim(txtAnnouncement.Text) = "" Then
    TR = MsgBox("You have not enter a name for the sound file", vbCritical, "Name?")
    txtAnnouncement.SetFocus
    Exit Sub
End If
    Dim TemPath As String
    bttnRecord.Enabled = False
    bttnPause.Enabled = False
    bttnResume.Enabled = False
    bttnStop.Enabled = False
    bttnSave.Enabled = True

    TemPath = FSys.GetParentFolderName(DatabasePath)
    With DataEnvironment1.rssqlTem
            If .State = 1 Then .Close
            .Source = "SELECT tblAnnouncement.* FROM tblAnnouncement"
            .Open
            .AddNew
            !doctor_ID = Val(frmAnnouncements.ListConsultantIDs.Text)
            !announcement = Trim(txtAnnouncement.Text)
            .Update
        Cmd = "Save NewAnnouncement " & TemPath & "\" & !Announcement_ID & ".wav"
        !File = !Announcement_ID & ".wav"
        .Update
        .Close
    End With
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    
    If errorCode <> 0 Then
        TR = MsgBox(errorStr, vbCritical, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    bttnSave.Enabled = False
    Unload Me
    Exit Sub
EH:
    TR = MsgBox("Error " & vbNewLine & Err.Description, vbCritical, "Error")
    Exit Sub

End Sub

Private Sub bttnStop_Click()
    Timer1.Interval = 0
    Timer1.Enabled = False
    ProgressBar1.Value = 100
    Dim TR As Integer
    bttnRecord.Enabled = False
    bttnPause.Enabled = False
    bttnResume.Enabled = False
    bttnStop.Enabled = False
    bttnSave.Enabled = True
    Cmd = "Stop NewAnnouncement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    If errorCode <> 0 Then
        TR = MsgBox(errorStr, vbCritical, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    bttnRecord.Enabled = True
    bttnPause.Enabled = False
    bttnResume.Enabled = False
    bttnStop.Enabled = False
    bttnSave.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim TR As Integer
If bttnSave.Enabled = True Or bttnResume.Enabled = True Or bttnStop.Enabled = True Then
    TR = MsgBox("The recording is not saved. Are you sure you want to Exit with out saving?", vbQuestion + vbYesNo, "Quit without saving")
    If TR = vbNo Then
        Cancel = True
        Exit Sub
    Else
        Cancel = False
    End If
Else
    Cmd = "Close NewAnnouncement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    frmAnnouncements.FillAnnouncements
End If
End Sub

Private Sub Timer1_Timer()
EllapsedSeconds = EllapsedSeconds + 1
ProgressBar1.Value = (EllapsedSeconds / (60 * 2)) * 100
If EllapsedSeconds > 60 * 2 Then
    Timer1.Interval = 0
    bttnStop_Click
End If
End Sub
