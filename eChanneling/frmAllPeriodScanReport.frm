VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAllPeriodScanReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan Detail Reports"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAllPeriodScanReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8145
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Process"
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
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   4320
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   3615
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   7335
      Begin VB.Label lblIncome 
         Caption         =   "0.00"
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label lblPatients 
         Caption         =   "0.00"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "Hospital Income"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "No Of Patients"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   72876033
      CurrentDate     =   39489
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   72876033
      CurrentDate     =   39489
   End
   Begin btButtonEx.ButtonEx bttnbttnDoctorsIncomeReport 
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   4320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Hospital Income By Scanning Report"
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
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmAllPeriodScanReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemTital As String
    Dim CSetPrinter As New cSetDfltPrinter
    Dim A

Private Sub FindDoctorIncomePatients()
    Dim TemDoctorFee   As Double
    With DataEnvironment1.rsSqlTem1111
        If .State = 1 Then .Close
        .Open "Select* From tblPatientFacility Where (FullyPaid = 1 and IsScan = 1 And cancelled = 0 and refund = 0 and AppointmentDate Between  '" & DTPicker2.Value & "'  and '" & DTPicker3.Value & "' )  Order By PatientFacility_ID "
        If .RecordCount = 0 Then Exit Sub
        Do While .EOF = False
            TemDoctorFee = TemDoctorFee + !institutiondue
            .MoveNext
        Loop
        lblIncome.Caption = Format(TemDoctorFee, "0.00")
        lblPatients.Caption = .RecordCount
        .Close
    End With
End Sub

Private Sub ClearValus()
    lblIncome.Caption = ""
    lblPatients.Caption = ""
End Sub

Private Sub btnProcess_Click()
    FindDoctorIncomePatients
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnbttnDoctorsIncomeReport_Click()
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    Call PrintThismonthDoctorIncome
End Sub

Private Sub PrintThismonthDoctorIncome()
    Dim firstDay As Date
    Dim LastDay As Date
    
    Const PreSHape = "SHAPE {"
    Const Sql = "SELECT tblPatientFacility.*,  tblDoctor.Doctor_ID,tblDoctor.DoctorName FROM tblDoctor INNER JOIN tblPatientFacility ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID "
    Const PostSHape = "ORDER BY tblDoctor.DoctorListedName }  AS cnmdDoctorIncome COMPUTE cnmdDoctorIncome, SUM(cnmdDoctorIncome.'PersonalDue') AS TotalDoctorDue, COUNT(cnmdDoctorIncome.'PatientFacility_ID') AS TotalPatients BY 'DoctorName'"
    
    firstDay = DateSerial(Year(Date), Month(Date), 1)
    If Month(Date) = 12 Then
        LastDay = DateSerial(Year(Date), Month(Date), 31)
    Else
        LastDay = DateSerial(Year(Date), Month(Date) + 1, 1) - 1
    End If
    
    With DataEnvironment1
        If .rscnmdDoctorIncome_Grouping.State = 1 Then .rscnmdDoctorIncome_Grouping.Close
        .Commands!cnmdDoctorIncome_Grouping.CommandText = PreSHape & Sql & " Where appointmentdate Between  '" & Format(DTPicker2.Value, DefaultLongDate) & "' and '" & Format(DTPicker3.Value, DefaultLongDate) & "' and FullyPaid = 1 and IsScan = 1 And refund = 0 and cancelled = 0 " & PostSHape
        .cnmdDoctorIncome_Grouping
        dtrDoctorsIncomeReport2.Sections("PageHeader").Controls.Item("lblDate").Caption = "Date  From   : " & Format(DTPicker2.Value, DefaultLongDate) & "  To  " & Format(DTPicker3.Value, DefaultLongDate)
        Set dtrDoctorsIncomeReport2.DataSource = DataEnvironment1
        dtrDoctorsIncomeReport2.Show
    End With

End Sub

Private Sub DTPicker2_Change()
    ClearValus
    FindDoctorIncomePatients
End Sub

Private Sub DTPicker3_Change()
    ClearValus
    FindDoctorIncomePatients
End Sub

Private Sub Form_Load()
    DTPicker2 = Date
    DTPicker3 = Date
    FindDoctorIncomePatients
End Sub
