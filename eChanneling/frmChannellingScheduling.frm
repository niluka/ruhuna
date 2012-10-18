VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChannelingScheduling 
   Caption         =   "Channelling Scheduling"
   ClientHeight    =   8625
   ClientLeft      =   3735
   ClientTop       =   1770
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChannellingScheduling.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   15240
   Begin VB.ListBox ListSecessionStart 
      Height          =   300
      Left            =   4440
      TabIndex        =   51
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   15
   End
   Begin VB.ListBox ListSecessionIDs 
      Height          =   300
      Left            =   4440
      TabIndex        =   30
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DataComboDoctors 
      Bindings        =   "frmChannellingScheduling.frx":0442
      Height          =   7620
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   13441
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      ListField       =   "DoctorListedName"
      BoundColumn     =   "Doctor_ID"
      Text            =   ""
      Object.DataMember      =   "sqlDoctor"
   End
   Begin VB.Frame FrameSecessions 
      Height          =   6375
      Left            =   4320
      TabIndex        =   18
      Top             =   1560
      Width           =   10455
      Begin VB.ComboBox cmbSecession 
         Height          =   360
         ItemData        =   "frmChannellingScheduling.frx":0461
         Left            =   1560
         List            =   "frmChannellingScheduling.frx":046B
         TabIndex        =   55
         Top             =   240
         Width           =   2775
      End
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   3000
         TabIndex        =   43
         Top             =   3480
         Width           =   135
      End
      Begin VB.Frame Frame9 
         Height          =   2895
         Left            =   4080
         TabIndex        =   50
         Top             =   3480
         Width           =   135
      End
      Begin VB.Frame Frame8 
         Height          =   2895
         Left            =   7080
         TabIndex        =   49
         Top             =   3480
         Width           =   135
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   2655
         Left            =   6240
         TabIndex        =   48
         Top             =   3720
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame6"
         Height          =   2775
         Left            =   9360
         TabIndex        =   47
         Top             =   3600
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   2655
         Left            =   8280
         TabIndex        =   46
         Top             =   3720
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   2655
         Left            =   5160
         TabIndex        =   45
         Top             =   3720
         Width           =   15
      End
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   1800
         TabIndex        =   42
         Top             =   3480
         Width           =   135
      End
      Begin VB.TextBox txtComments 
         Height          =   1200
         Left            =   6000
         TabIndex        =   15
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox txtUsualDuration 
         Height          =   360
         Left            =   7920
         MaxLength       =   250
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox ChkCalculateTime 
         Caption         =   "Calculate Appointment Time"
         Height          =   255
         Left            =   6000
         TabIndex        =   12
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox chkBypassOrder 
         Caption         =   "Can bypass order"
         Height          =   255
         Left            =   6000
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtAgentHospitalFee 
         Height          =   360
         Left            =   3960
         MaxLength       =   250
         TabIndex        =   9
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtAgentDoctorFee 
         Height          =   360
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtFogrignerHospitalFee 
         Height          =   360
         Left            =   3960
         MaxLength       =   250
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtForeginerDoctorFee 
         Height          =   360
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtLocalHospitalFee 
         Height          =   360
         Left            =   3960
         MaxLength       =   250
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtLocalDoctorFee 
         Height          =   360
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtMaximum 
         Height          =   360
         Left            =   1560
         MaxLength       =   250
         TabIndex        =   3
         Top             =   1200
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   360
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Format          =   56623106
         CurrentDate     =   39401
      End
      Begin btButtonEx.ButtonEx bttnSecessionDelete 
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
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
      Begin btButtonEx.ButtonEx bttnSecessionAdd 
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "Add"
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
      Begin VB.ListBox ListSecessions 
         Height          =   2220
         Left            =   120
         TabIndex        =   16
         Top             =   3960
         Width           =   10215
      End
      Begin VB.TextBox txtSecessionName 
         Height          =   375
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Local"
         Height          =   255
         Left            =   7560
         TabIndex        =   41
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent"
         Height          =   255
         Left            =   9600
         TabIndex        =   40
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Foreginers"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Max."
         Height          =   255
         Left            =   3360
         TabIndex        =   38
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Secession "
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Fee"
         Height          =   255
         Left            =   5160
         TabIndex        =   35
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital Fee"
         Height          =   255
         Left            =   8160
         TabIndex        =   34
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Local Patients"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Foreginers"
         Height          =   255
         Left            =   8400
         TabIndex        =   32
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Bookings"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   255
         Left            =   6000
         TabIndex        =   29
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Usual Duration                                     minutes"
         Height          =   255
         Left            =   6000
         TabIndex        =   28
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent"
         Height          =   255
         Left            =   6360
         TabIndex        =   26
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label65 
         BackStyle       =   0  'Transparent
         Caption         =   "Foreginers"
         Height          =   255
         Left            =   5280
         TabIndex        =   25
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Local"
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital Fee"
         Height          =   255
         Left            =   4200
         TabIndex        =   23
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Fee"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Secession "
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum No."
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTabDates 
      Height          =   7695
      Left            =   4080
      TabIndex        =   17
      Top             =   360
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   8
      Tab             =   7
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Monday"
      TabPicture(0)   =   "frmChannellingScheduling.frx":0481
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tuesday"
      TabPicture(1)   =   "frmChannellingScheduling.frx":049D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Wednesday"
      TabPicture(2)   =   "frmChannellingScheduling.frx":04B9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Thursday"
      TabPicture(3)   =   "frmChannellingScheduling.frx":04D5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Friday"
      TabPicture(4)   =   "frmChannellingScheduling.frx":04F1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Saturday"
      TabPicture(5)   =   "frmChannellingScheduling.frx":050D
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Sunday"
      TabPicture(6)   =   "frmChannellingScheduling.frx":0529
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Other Leave"
      TabPicture(7)   =   "frmChannellingScheduling.frx":0545
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "Label13"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "ChkFullDayLeave"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "dtpAlteredDate"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).ControlCount=   3
      Begin MSComCtl2.DTPicker dtpAlteredDate 
         Height          =   375
         Left            =   2160
         TabIndex        =   53
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   56623107
         CurrentDate     =   39454
      End
      Begin VB.CheckBox ChkFullDayLeave 
         Caption         =   "FullDayLeave"
         Height          =   255
         Left            =   5280
         TabIndex        =   52
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Date"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Label LblDoctor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   27
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmChannelingScheduling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemDoctorID As Long
Private Sub Setcolours()
    bttnSecessionAdd.BackColor = BttnBackColour
    bttnSecessionAdd.ForeColor = BttnForeColour
    bttnSecessionDelete.BackColor = BttnBackColour
    bttnSecessionDelete.ForeColor = BttnForeColour
    frmChannelingScheduling.BackColor = FrmBackColour
    frmChannelingScheduling.ForeColor = FrmForeColour
End Sub


Private Sub bttnSecessionAdd_Click()
        
    ListSecessionIDs.Clear
    ListSecessions.Clear
    ListSecessionStart.Clear
    
    If SSTabDates.Tab = 7 Then
        Call AddLeaveDays
        Call ListLeaveDays
    Else
        If CanAdd = False Then Exit Sub
        Call SaveSecession
        Call ListFacilitySecessions
    End If
  
End Sub

Private Sub AddLeaveDays()
    If ChkFullDayLeave.Value <> 1 Then
        If CanAdd = False Then Exit Sub
    End If
    With DataEnvironment1.rssqlTem5
        If .State = 1 Then .Close
        .Source = "Select * from tblfacilitysecession"
        .Open
        .AddNew
        If ChkFullDayLeave.Value = 1 Then
            !fulldayleave = True
        Else
            !fulldayleave = False
        End If
        !altereddate = dtpAlteredDate.Value
        !Staff_ID = TemDoctorID
        !HospitalFacility_id = 10
        !secessionname = txtSecessionName.Text
        !startingtime = dtpStart.Value
        !usualduration = Val(txtUsualDuration.Text)
        !Maximum = Val(txtMaximum.Text)
        !localdoctorFee = Val(txtLocalDoctorFee.Text)
        !localhospitalFee = Val(txtLocalHospitalFee.Text)
        !foreigndoctorfee = Val(txtforeignerDoctorFee.Text)
        !ForeignHospitalFee = Val(txtForeginerDoctorFee.Text)
        !agentDoctorFee = Val(txtAgentDoctorFee.Text)
        !agenthospitalFee = Val(txtAgentHospitalFee.Text)
        If chkBypassOrder.Value = 1 Then
            !CanByPassOrder = True
        Else
            !CanByPassOrder = False
        End If
        If ChkCalculateTime.Value = 1 Then
            !calculateappointment = True
        Else
            !calculateappointment = False
        End If
        !Comments = txtComments.Text
        !SecessionWeekday = 8
        .Update
    End With
End Sub
    



Private Sub ListLeaveDays()
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = "Select * from  tblfacilitysecession where staff_id = " & TemDoctorID & " and SecessionWeekday = 8 order by altereddate"
        .Open
        If .RecordCount = 0 Then Exit Sub
        
        Dim TemText As String
        
        While .EOF = False
        
            If Not IsNull(!altereddate) Then
                TemText = Format(!altereddate, "dd-mm-yyyy")
            Else
                TemText = vbTab
            End If
            If Not IsNull(!fulldayleave) Then
                If !fulldayleave = True Then
                    TemText = TemText & vbTab & "Full Day Leave"
                ElseIf Not IsNull(!startingtime) Then
                    TemText = TemText & vbTab & !startingtime
                Else
                    TemText = TemText & vbTab & vbTab
                End If
            End If
            If Not IsNull(!Maximum) Then
                TemText = TemText & vbTab & !Maximum
            Else
                TemText = TemText & vbTab & vbTab
            End If
                    If Not IsNull(!localdoctorFee) Then
                        TemText = TemText & vbTab & Format(!localdoctorFee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    If Not IsNull(!foreigndoctorfee) Then
                        TemText = TemText & vbTab & Format(!foreigndoctorfee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    If Not IsNull(!agentDoctorFee) Then
                        TemText = TemText & vbTab & Format(!agentDoctorFee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    If Not IsNull(!localhospitalFee) Then
                        TemText = TemText & vbTab & Format(!localhospitalFee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    If Not IsNull(!ForeignHospitalFee) Then
                        TemText = TemText & vbTab & Format(!ForeignHospitalFee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    If Not IsNull(!agenthospitalFee) Then
                        TemText = TemText & vbTab & Format(!agenthospitalFee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    
            ListSecessions.AddItem TemText
            ListSecessionIDs.AddItem !facilitysecession_ID
            ListSecessionStart.AddItem !startingtime
            .MoveNext
        Wend
        .Close
    End With

End Sub

Private Sub ListFacilitySecessions()
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        Select Case SSTabDates.Tab
            Case 0:         .Source = "Select * from  tblfacilitysecession where staff_id = " & TemDoctorID & " and SecessionWeekday = " & vbMonday & " order by StartingTime"
            Case 1:         .Source = "Select * from  tblfacilitysecession where staff_id = " & TemDoctorID & " and SecessionWeekday = " & vbTuesday & " order by StartingTime"
            Case 2:         .Source = "Select * from  tblfacilitysecession where staff_id = " & TemDoctorID & " and SecessionWeekday = " & vbWednesday & " order by StartingTime"
            Case 3:         .Source = "Select * from  tblfacilitysecession where staff_id = " & TemDoctorID & " and SecessionWeekday = " & vbThursday & " order by StartingTime"
            Case 4:         .Source = "Select * from  tblfacilitysecession where staff_id = " & TemDoctorID & " and SecessionWeekday = " & vbFriday & " order by StartingTime"
            Case 5:         .Source = "Select * from  tblfacilitysecession where staff_id = " & TemDoctorID & " and SecessionWeekday = " & vbSaturday & " order by StartingTime"
            Case 6:         .Source = "Select * from  tblfacilitysecession where staff_id = " & TemDoctorID & " and SecessionWeekday = " & vbSunday & " order by StartingTime"
        End Select
        .Open
        If .RecordCount = 0 Then Exit Sub
        
        Dim TemText As String
        
        While .EOF = False
        
            If Not IsNull(!secessionname) Then
                TemText = !secessionname
            Else
                TemText = vbTab
            End If
            If Not IsNull(!startingtime) Then
                TemText = TemText & vbTab & !startingtime
            Else
                TemText = TemText & vbTab & vbTab
            End If
            If Not IsNull(!Maximum) Then
                TemText = TemText & vbTab & !Maximum
            Else
                TemText = TemText & vbTab & vbTab
            End If
                    If Not IsNull(!localdoctorFee) Then
                        TemText = TemText & vbTab & Format(!localdoctorFee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    If Not IsNull(!foreigndoctorfee) Then
                        TemText = TemText & vbTab & Format(!foreigndoctorfee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    If Not IsNull(!agentDoctorFee) Then
                        TemText = TemText & vbTab & Format(!agentDoctorFee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    If Not IsNull(!localhospitalFee) Then
                        TemText = TemText & vbTab & Format(!localhospitalFee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    If Not IsNull(!ForeignHospitalFee) Then
                        TemText = TemText & vbTab & Format(!ForeignHospitalFee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    If Not IsNull(!agenthospitalFee) Then
                        TemText = TemText & vbTab & Format(!agenthospitalFee, "#0.00")
                    Else
                        TemText = TemText & vbTab & vbTab
                    End If
                    
            ListSecessions.AddItem TemText
            ListSecessionIDs.AddItem !facilitysecession_ID
            ListSecessionStart.AddItem !startingtime
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub SaveSecession()
    With DataEnvironment1.rssqlTem3
        If .State = 1 Then .Close
        .Source = "Select * from tblfacilitysecession"
        .Open
        .AddNew
                        
                        !Staff_ID = TemDoctorID
                        !HospitalFacility_id = 10
                        
            Select Case SSTabDates.Tab
                Case 0: !SecessionWeekday = vbMonday
                Case 1: !SecessionWeekday = vbTuesday
                Case 2: !SecessionWeekday = vbWednesday
                Case 3: !SecessionWeekday = vbThursday
                Case 4: !SecessionWeekday = vbFriday
                Case 5: !SecessionWeekday = vbSaturday
                Case 6: !SecessionWeekday = vbSunday
                Case Else
            End Select
                    !secessionname = txtSecessionName.Text
                    !startingtime = dtpStart.Value
                    !usualduration = Val(txtUsualDuration.Text)
                    !Maximum = Val(txtMaximum.Text)
                    !localdoctorFee = Val(txtLocalDoctorFee.Text)
                    !localhospitalFee = Val(txtLocalHospitalFee.Text)
                    !foreigndoctorfee = Val(txtForeginerDoctorFee.Text)
                    !ForeignHospitalFee = Val(txtFogrignerHospitalFee.Text)
                    !agentDoctorFee = Val(txtAgentDoctorFee.Text)
                    !agenthospitalFee = Val(txtAgentHospitalFee.Text)
                    If chkBypassOrder.Value = 1 Then
                        !CanByPassOrder = True
                    Else
                        !CanByPassOrder = False
                    End If
                    If ChkCalculateTime.Value = 1 Then
                        !calculateappointment = True
                    Else
                        !calculateappointment = False
                    End If
                    !Comments = txtComments.Text
                    .Update
    End With
End Sub

Private Function CanAdd() As Boolean
    CanAdd = False
    Dim TemResponce As Integer
    If Not IsNumeric(DataComboDoctors.BoundText) Then
        TemResponce = MsgBox("You have not selected a doctor", vbCritical, "Doctor?")
        DataComboDoctors.SetFocus
        Exit Function
    End If
    If dtpStart.Value = TimeSerial(0, 0, 0) Then
        TemResponce = MsgBox("You have not enterd an starting time for the secession", vbCritical, "Starting time?")
        Exit Function
    End If
    If Trim(txtSecessionName.Text) = "" Then txtSecessionName.Text = Format(dtpStart.Value, "hh : mm ") & " Secession"
    If Not IsNumeric(txtLocalDoctorFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Doctor fee for local patients", vbCritical, "No doctor charge")
        Exit Function
    End If
    If Not IsNumeric(txtLocalHospitalFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Hospital fee for local patients", vbCritical, "No doctor charge")
        Exit Function
    End If
    If Not IsNumeric(txtForeginerDoctorFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Doctor fee for foreign patients", vbCritical, "No doctor charge")
        Exit Function
    End If
    If Not IsNumeric(txtFogrignerHospitalFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Hospital fee for foreign patients", vbCritical, "No doctor charge")
        Exit Function
    End If
    If Not IsNumeric(txtLocalDoctorFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Doctor fee for patients booking through agents", vbCritical, "No doctor charge")
        Exit Function
    End If
    If Not IsNumeric(txtLocalHospitalFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Hospital fee for patients booking through agents", vbCritical, "No doctor charge")
        Exit Function
    End If
    Dim TemNum As Long
    For TemNum = 0 To ListSecessionStart.ListCount - 1
        ListSecessionStart.ListIndex = TemNum
        If SSTabDates.Tab = 7 Then
            If IsDate(ListSecessionStart.Text) Then
                If ListSecessionStart.Text = dtpStart.Value Then
                    TemResponce = MsgBox("You have already entered details about this date earlier, Are you sure you want to add additional details?", vbCritical, "Same Date")
                    If TemResponce = vbNo Then Exit Function
                End If
            End If
        Else
            If IsDate(ListSecessionStart.Text) Then
                If Abs(ListSecessionStart.Text - dtpStart.Value) < 4 Then
                    TemResponce = MsgBox("You have entered a starting time which is very closer to the starting time of an existing secession. Please change the starting time", vbCritical, "Same Starting Time")
                    Exit Function
                End If
            End If
        End If
    Next
    CanAdd = True
End Function

Private Sub bttnSecessionDelete_Click()
Dim TemResponce As Integer
If ListSecessions.ListIndex < 0 Then
    TemResponce = MsgBox("You have not selected a secesion to delete", vbCritical, "No secession to delete")
    Exit Sub
End If
    ListSecessionIDs.ListIndex = ListSecessions.ListIndex
    ListSecessionStart.ListIndex = ListSecessions.ListIndex
    With DataEnvironment1.rssqlTem4
        If .State = 1 Then .Close
        .Source = "Select * from tblfacilitysecession where facilitysecession_ID = " & Val(ListSecessionIDs.Text)
        .Open
        If .RecordCount Then
            .Delete adAffectCurrent
        Else
            .Close
            Exit Sub
        End If
        .Close
    End With

ListSecessionIDs.RemoveItem (ListSecessions.ListIndex)
ListSecessionStart.RemoveItem (ListSecessions.ListIndex)
ListSecessions.RemoveItem (ListSecessions.ListIndex)

End Sub


Private Sub cmbSecession_Click()
txtSecessionName.Text = cmbSecession.Text
End Sub

Private Sub DataComboDoctors_Change()
    On Error GoTo ErrorHandler
    If DataComboDoctors.Text = "" Then LblDoctor.Caption = "Doctor not selected": Exit Sub
    If Not IsNumeric(DataComboDoctors.BoundText) Then LblDoctor.Caption = "Doctor not selected": Exit Sub
    TemDoctorID = DataComboDoctors.BoundText
    LblDoctor.Caption = FindDoctorFromID(TemDoctorID)
    SSTabDates.Tab = 0
    ListSecessionIDs.Clear
    ListSecessions.Clear
    ListSecessionStart.Clear
    ListFacilitySecessions
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub


Private Sub Form_Load()
    Call Setcolours
    
    dtpStart.Value = TimeSerial(0, 0, 0)
    dtpAlteredDate.Value = Date
    dtpAlteredDate.MinDate = Date
    
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT * FROM tblDoctor order by doctorlistedname"
        .Open
    End With
    DataComboDoctors.RowMember = "sqltem"
    DataComboDoctors.ListField = "DoctorListedName"
    DataComboDoctors.BoundText = "Doctor_ID"

Dim ingRet As Long

Dim TabDatesSecessions(8) As Long

TabDatesSecessions(0) = 70
TabDatesSecessions(1) = 125
TabDatesSecessions(2) = 165
TabDatesSecessions(3) = 205
TabDatesSecessions(4) = 240
TabDatesSecessions(5) = 288
TabDatesSecessions(6) = 315
TabDatesSecessions(7) = 355
TabDatesSecessions(8) = 400

ingRet = SendMessage(ListSecessions.hwnd, LB_SETTABSTOPS, 8, TabDatesSecessions(0))

End Sub



Private Sub DisplayDetails()
    With DataEnvironment1.rssqlTem2
        If .State = 1 Then .Close
        .Source = "select * from tblfacilitysecession where FacilitySecession_ID = " & Val(ListSecessionIDs.Text)
        .Open
        If .RecordCount Then
            If IsNull(!SecessionWeekday) Then Exit Sub
            Select Case !SecessionWeekday
                Case vbMonday: SSTabDates.Tab = 0
                
                Case vbTuesday: SSTabDates.Tab = 1
                
                Case vbWednesday: SSTabDates.Tab = 2
                
                Case vbThursday: SSTabDates.Tab = 3
                
                Case vbFriday: SSTabDates.Tab = 4
                
                Case vbSaturday: SSTabDates.Tab = 5
                
                Case vbSunday: SSTabDates.Tab = 6
                
                Case Else
            End Select
                    If Not IsNull(!secessionname) Then txtSecessionName.Text = !secessionname
                    If Not IsNull(!startingtime) Then dtpStart.Value = !startingtime
                    If Not IsNull(!usualduration) Then txtUsualDuration.Text = !usualduration
                    If Not IsNull(!Maximum) Then txtMaximum.Text = !Maximum
                    If Not IsNull(!localdoctorFee) Then txtLocalDoctorFee.Text = Format(!localdoctorFee, "#0.00")
                    If Not IsNull(!localhospitalFee) Then txtLocalHospitalFee.Text = Format(!localhospitalFee, "#0.00")
                    If Not IsNull(!foreigndoctorfee) Then txtForeginerDoctorFee.Text = Format(!foreigndoctorfee, "#0.00")
                    If Not IsNull(!ForeignHospitalFee) Then txtFogrignerHospitalFee.Text = Format(!ForeignHospitalFee, "#0.00")
                    If Not IsNull(!agentDoctorFee) Then txtAgentDoctorFee.Text = Format(!agentDoctorFee, "#0.00")
                    If Not IsNull(!agenthospitalFee) Then txtAgentHospitalFee.Text = Format(!agenthospitalFee, "#0.00")
                    If !CanByPassOrder = True Then chkBypassOrder.Value = 1
                    If !calculateappointment = True Then ChkCalculateTime.Value = 1
                    If Not IsNull(!Comments) Then txtComments.Text = !Comments
        End If
    End With
End Sub


Private Sub SSTabDates_Click(PreviousTab As Integer)
    ListSecessionIDs.Clear
    ListSecessions.Clear
    ListSecessionStart.Clear
    
    If SSTabDates.Tab = 7 Then
        Frame2.Visible = False
        ListLeaveDays
        Exit Sub
    Else
        Frame2.Visible = True
        ListFacilitySecessions
    End If
End Sub
