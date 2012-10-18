VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStaffSecessions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff & Facilities"
   ClientHeight    =   8625
   ClientLeft      =   3720
   ClientTop       =   1755
   ClientWidth     =   12105
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
   ScaleHeight     =   8625
   ScaleWidth      =   12105
   Begin MSDataListLib.DataCombo DataComboDoctors 
      Bindings        =   "frmStaffSecessions.frx":0000
      Height          =   7380
      Left            =   120
      TabIndex        =   36
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   13018
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   "DataCombo1"
   End
   Begin VB.Frame FrameSecessions 
      Height          =   6375
      Left            =   4320
      TabIndex        =   8
      Top             =   1320
      Width           =   7335
      Begin VB.ComboBox cmbSecessionName 
         Height          =   360
         ItemData        =   "frmStaffSecessions.frx":001F
         Left            =   1560
         List            =   "frmStaffSecessions.frx":0029
         TabIndex        =   37
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtComments 
         Height          =   360
         Left            =   2160
         TabIndex        =   34
         Top             =   3480
         Width           =   4935
      End
      Begin VB.TextBox txtUsualDuration 
         Height          =   360
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   32
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CheckBox CheckCalculateTime 
         Caption         =   "Calculate Appointment Time"
         Height          =   255
         Left            =   4440
         TabIndex        =   31
         Top             =   720
         Width           =   2775
      End
      Begin VB.ListBox ListSecessions 
         Height          =   1980
         Left            =   120
         TabIndex        =   18
         Top             =   4320
         Width           =   7095
      End
      Begin VB.CheckBox chkBypassOrder 
         Caption         =   "Can bypass order"
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtAgentHospitalFee 
         Height          =   360
         Left            =   3960
         MaxLength       =   250
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtAgentDoctorFee 
         Height          =   360
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   15
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtFogrignerHospitalFee 
         Height          =   360
         Left            =   3960
         MaxLength       =   250
         TabIndex        =   14
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtForeginerDoctorFee 
         Height          =   360
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtLocalHospitalFee 
         Height          =   360
         Left            =   3960
         MaxLength       =   250
         TabIndex        =   12
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtLocalDoctorFee 
         Height          =   360
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   11
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtMaximum 
         Height          =   360
         Left            =   1560
         MaxLength       =   250
         TabIndex        =   9
         Top             =   1200
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   360
         Left            =   1560
         TabIndex        =   19
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Format          =   19267586
         CurrentDate     =   39401
      End
      Begin btButtonEx.ButtonEx bttnSecessionDelete 
         Height          =   255
         Left            =   3960
         TabIndex        =   20
         Top             =   3960
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
         TabIndex        =   21
         Top             =   3960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "A&dd"
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
      Begin VB.TextBox txtSecessionName 
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Usual Duration                                     minutes"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   3120
         Width           =   5055
      End
      Begin VB.Label Label66 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Bookings"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label65 
         BackStyle       =   0  'Transparent
         Caption         =   "Foreginers"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Local Patients"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital Fee"
         Height          =   255
         Left            =   4200
         TabIndex        =   26
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Fee"
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Secession "
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum No."
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1455
      End
   End
   Begin btButtonEx.ButtonEx bttnChange 
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&hange"
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
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Ca&ncel"
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
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackStyle       =   1
      Caption         =   "&Save"
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
      Height          =   495
      Left            =   9960
      TabIndex        =   6
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
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
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Edit"
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
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin TabDlg.SSTab SSTabDates 
      Height          =   7215
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Monday"
      TabPicture(0)   =   "frmStaffSecessions.frx":003F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tuesday"
      TabPicture(1)   =   "frmStaffSecessions.frx":005B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Wednesday"
      TabPicture(2)   =   "frmStaffSecessions.frx":0077
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Thursday"
      TabPicture(3)   =   "frmStaffSecessions.frx":0093
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Friday"
      TabPicture(4)   =   "frmStaffSecessions.frx":00AF
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Saturday"
      TabPicture(5)   =   "frmStaffSecessions.frx":00CB
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Sunday"
      TabPicture(6)   =   "frmStaffSecessions.frx":00E7
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Other Leave"
      TabPicture(7)   =   "frmStaffSecessions.frx":0103
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
   End
   Begin VB.Label LblDoctorStaff 
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
      TabIndex        =   30
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmStaffSecessions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbSecessionName_Click()
txtSecessionName = cmbSecessionName
End Sub

Private Sub Form_Load()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "select * from tbldoctor order by DoctorListedName"
    .Open
End With
DataComboDoctors.RowSource = DataEnvironment1.rssqlTem
DataComboDoctors.RowMember = "DoctorListedName"
End Sub


