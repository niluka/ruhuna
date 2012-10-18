VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAnnouncements 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Announcements"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   10335
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnPlay 
      Default         =   -1  'True
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Play"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox ListAnnouncementIDs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   8280
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox ListAnnouncements 
      Height          =   4140
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.ListBox ListSpecialityIDs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox ListConsultantIDs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox ListSpecialities 
      Height          =   4140
      ItemData        =   "frmAnnouncements.frx":0000
      Left            =   120
      List            =   "frmAnnouncements.frx":0002
      TabIndex        =   1
      ToolTipText     =   "List of Specialities"
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox ListConsultants 
      Height          =   4140
      ItemData        =   "frmAnnouncements.frx":0004
      Left            =   2880
      List            =   "frmAnnouncements.frx":0006
      TabIndex        =   0
      ToolTipText     =   "List of Consultants of selected speciality"
      Top             =   120
      Width           =   2895
   End
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnNew 
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Record"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAnnouncements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    
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

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnNew_Click()
    frmRecord.Show 1
End Sub

Private Sub Form_Load()
    Call FormatGridSpeciality
    Call FormatGridConsultants
    Call FillSpeciality
    bttnPlay.Enabled = False
    bttnDelete.Enabled = False
    bttnNew.Enabled = False
End Sub
Private Sub FormatGridSpeciality()
    ListSpecialities.Clear
    ListSpecialityIDs.Clear
End Sub

Private Sub FormatGridConsultants()
    ListConsultants.Clear
    ListConsultantIDs.Clear
End Sub
Private Sub FillSpeciality()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblspeciality order by speciality "
    .Open
    If NoAllNames = False Then
        ListSpecialities.AddItem "All"
        ListSpecialityIDs.AddItem "All"
    End If
    If .RecordCount <> 0 Then
        While Not .EOF
            ListSpecialities.AddItem !Speciality
            ListSpecialityIDs.AddItem !speciality_ID
            .MoveNext
        Wend
    End If
    .Close
End With
End Sub

Private Sub ListAnnouncements_DblClick()
    bttnPlay_Click
End Sub

Private Sub ListSpecialities_Click()
    ListSpecialityIDs.ListIndex = ListSpecialities.ListIndex
    ListConsultantIDs.Clear
    ListConsultants.Clear
    ListAnnouncementIDs.Clear
    ListAnnouncements.Clear
    If ListSpecialities.Text = "All" Then
        ListAllConsultants
    ElseIf ListSpecialities.Text <> "All" And IsNumeric(ListSpecialityIDs.Text) = True Then
        ListSelectedConsultants
    Else
        FormatGridConsultants
    End If
    bttnNew.Enabled = False
    bttnDelete.Enabled = False
    bttnPlay.Enabled = False
End Sub
Private Sub ListAllConsultants()
Call FormatGridConsultants
With DataEnvironment1.rssqlTem1
    If .State = 1 Then .Close
    If SurnameFirst = True Then
        .Source = "SELECT  tbldoctor.*  FROM  tbldoctor  order by doctorlistedname"
    Else
        .Source = "SELECT  tbldoctor.*  FROM  tbldoctor  order by doctorname"
    End If
    .Open
    If .RecordCount = 0 Then Exit Sub
    While Not .EOF
            If SurnameFirst = True Then
                ListConsultants.AddItem !doctorlistedname
            Else
                ListConsultants.AddItem !doctorname
            End If
        ListConsultantIDs.AddItem !doctor_ID
        .MoveNext
    Wend
    .Close
End With
End Sub
Private Sub ListSelectedConsultants()
    Call FormatGridConsultants
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        If SurnameFirst = True Then
            .Source = "SELECT tbldoctor.* FROM tbldoctor where  doctorspeciality_ID = " & Val(ListSpecialityIDs.Text) & " order by doctorlistedname"
        Else
            .Source = "SELECT tbldoctor.* FROM tbldoctor where  doctorspeciality_ID = " & Val(ListSpecialityIDs.Text) & " order by doctorname"
        End If
        .Open
        If .RecordCount = 0 Then Exit Sub
        While Not .EOF
            If SurnameFirst = True Then
                ListConsultants.AddItem !doctorlistedname
            Else
                ListConsultants.AddItem !doctorname
            End If
            ListConsultantIDs.AddItem !doctor_ID
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub ListSpecialities_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    If ListSpecialities.ListIndex < 0 And ListSpecialities.ListCount > 0 Then ListSpecialities.ListIndex = 0
    ListConsultants.SetFocus
    KeyCode = Empty
Else

End If
End Sub

Private Sub ListConsultants_Click()
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    Call FillAnnouncements
End Sub
Private Sub ClearAnnouncements()
    ListAnnouncementIDs.Clear
    ListAnnouncements.Clear
End Sub

Public Sub FillAnnouncements()
    Call ClearAnnouncements
    If Not IsNumeric(ListConsultantIDs.Text) Then
        bttnPlay.Enabled = False
        bttnDelete.Enabled = False
        bttnNew.Enabled = False
        Exit Sub
    End If
    bttnNew.Enabled = True
    bttnDelete.Enabled = False
    bttnPlay.Enabled = False
    ListAnnouncements.Visible = False
    Me.MousePointer = vbHourglass
    With DataEnvironment1.rssqlAnnouncement
        If .State = 1 Then .Close
        .Source = "SELECT tblAnnouncement.* FROM tblAnnouncement WHERE (Doctor_ID = " & Val(ListConsultantIDs.Text) & ") ORDER BY Announcement_ID"
        .Open
        If .RecordCount = 0 Then ListAnnouncements.Visible = True: Me.MousePointer = vbDefault: Exit Sub
        While .EOF = False
            ListAnnouncements.AddItem !announcement
            ListAnnouncementIDs.AddItem !Announcement_ID
            .MoveNext
        Wend
    If .State = 1 Then .Close
    ListAnnouncements.Visible = True
    End With
    Me.MousePointer = vbDefault
End Sub

Private Sub ListAnnouncements_Click()
    If ListAnnouncementIDs.ListCount < 1 Or ListAnnouncements.ListIndex < 1 Then
        bttnDelete.Enabled = False
        bttnPlay.Enabled = False
    End If
    ListAnnouncementIDs.ListIndex = ListAnnouncements.ListIndex
    If Not IsNumeric(ListAnnouncementIDs.Text) Then
        bttnDelete.Enabled = False
        bttnPlay.Enabled = False
    End If
    bttnDelete.Enabled = True
    bttnPlay.Enabled = True
End Sub

Private Sub ListAnnouncements_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If ListAnnouncementIDs.ListIndex < 0 And ListAnnouncementIDs.ListCount > 1 Then ListAnnouncements.ListIndex = 0
    bttnPlay_Click
    KeyCode = Empty
ElseIf KeyCode = vbKeyRight Then
    If ListAnnouncements.ListIndex < 0 And ListAnnouncementIDs.ListCount > 1 Then ListAnnouncements.ListIndex = 0
    bttnPlay.SetFocus
    KeyCode = Empty
ElseIf KeyCode = vbKeyLeft Then
    ListConsultants.SetFocus
    KeyCode = Empty
End If
End Sub

Private Sub bttnPlay_Click()
    Dim TemFile As String
    Dim TR As Long
    Dim Cmd As String
'    On Error Resume Next
    If Not IsNumeric(ListAnnouncementIDs.Text) Then Exit Sub
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblAnnouncement.* FROM tblAnnouncement WHERE (Announcement_ID = " & Val(ListAnnouncementIDs.Text) & ") ORDER BY Announcement_ID"
        .Open
        If .RecordCount = 0 Then Exit Sub
        TemFile = FSys.GetParentFolderName(DatabasePath)
        TemFile = TemFile & "\" & !File
        If FSys.FileExists(TemFile) = False Then Exit Sub
    End With
    
    Me.MousePointer = vbHourglass
    
    Cmd = "Close Announcement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    
' To Delete

'CommonDialog1.ShowOpen
'TemFile = CommonDialog1.FileName

' Delete end
    
    Cmd = "OPEN " & Chr(34) & TemFile & " " & Chr(34) & " type waveaudio alias Announcement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    If errorCode <> 0 Then
        TR = MsgBox(errorStr, vbCritical, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Cmd = "Play Announcement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)
    returnCode = mciGetErrorStringA(errorCode, errorStr, 255)
    If errorCode <> 0 Then
        TR = MsgBox(errorStr, vbCritical, "Error")
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.MousePointer = vbDefault
    
End Sub


Private Sub bttnDelete_Click()
    On Error GoTo EH
    Dim TemFile As String
    Dim TR As Long
    Dim Cmd As String
        Cmd = "Close Announcement"
    errorCode = mciSendStringA(Cmd, returnStr, 255, 0)

    If Not IsNumeric(ListAnnouncementIDs.Text) Then Exit Sub
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblAnnouncement.* FROM tblAnnouncement WHERE (Announcement_ID = " & Val(ListAnnouncementIDs.Text) & ") ORDER BY Announcement_ID"
        .Open
        If .RecordCount = 0 Then Exit Sub
        TemFile = FSys.GetParentFolderName(DatabasePath) & "\" & !File
        If FSys.FileExists(TemFile) = False Then Exit Sub
        FSys.DeleteFile (TemFile)
        .Delete adAffectCurrent
        .Close
        Call FillAnnouncements
        Exit Sub
EH:
    If .State = 1 Then
        .CancelUpdate
        .Close
    End If
    TR = MsgBox("An error occured", vbCritical, "Error")
    Exit Sub

    End With
End Sub



