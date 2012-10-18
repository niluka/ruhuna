VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAbsantPatientList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Absent Patient Details "
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
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
   ScaleHeight     =   3810
   ScaleWidth      =   8370
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   6375
      Begin btButtonEx.ButtonEx bttnViewDetails 
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "View Details"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   5318
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Today"
      TabPicture(0)   =   "frmAbsantPatientDetails.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDate"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected Day"
      TabPicture(1)   =   "frmAbsantPatientDetails.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Period"
      TabPicture(2)   =   "frmAbsantPatientDetails.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DTPicker3"
      Tab(2).Control(1)=   "DTPicker2"
      Tab(2).Control(2)=   "Label1"
      Tab(2).Control(3)=   "Label2"
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   -70200
         TabIndex        =   2
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   39489
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -74040
         TabIndex        =   3
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   39489
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -72240
         TabIndex        =   4
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   39489
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   375
         Left            =   -74640
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   375
         Left            =   -70800
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblDate 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   600
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmAbsantPatientList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnViewDetails_Click()
With DataEnvironment1.rscAgents

If .State = 1 Then .Close

    Select Case SSTab1.Tab
    Case 0
        .Open "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionName FROM tblInstitutions RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID Where ((FullyPaid = 1) And (patientabsent = 1)and (AppointmentDate = '" & Date & "'))ORDER BY tblPatientMainDetails.FirstName"
    Case 1
        .Open "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionName FROM tblInstitutions RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID Where ((FullyPaid = 1) And (patientabsent = 1)and (AppointmentDate = '" & DTPicker1.Value & "'))ORDER BY tblPatientMainDetails.FirstName"
    Case 2
        .Open "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionName FROM tblInstitutions RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID Where ((FullyPaid = 1) And (patientabsent = 1)and (AppointmentDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "'))ORDER BY tblPatientMainDetails.FirstName"
    End Select
    
    Select Case SSTab1.Tab
    Case 0
        dtrAbsantPatientList.Sections("PageHeader").Controls("lblDate").Caption = "Date     :  " & Date
    Case 1
        dtrAbsantPatientList.Sections("PageHeader").Controls("lblDate").Caption = "Date     :  " & DTPicker1.Value
    Case 2
        dtrAbsantPatientList.Sections("PageHeader").Controls("lblDate").Caption = "Date From   :  " & DTPicker2.Value & "       To     " & DTPicker3.Value
    End Select
    
    Set dtrAbsantPatientList.DataSource = DataEnvironment1.rscAgents
    dtrAbsantPatientList.Show
    
End With

End Sub

Private Sub Form_Load()
    lblDate.Caption = Date
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker3.Value = Date
    If UserAuthority <> AuthorityUser Then
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
    End If
End Sub
