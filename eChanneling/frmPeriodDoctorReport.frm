VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPeriodDoctorReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctors Detail Reports"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPeriodDoctorReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8520
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   360
      TabIndex        =   10
      Top             =   0
      Width           =   7695
      Begin MSDataListLib.DataCombo DtcDoctor 
         Height          =   360
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Doctor Name"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   5520
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
      Height          =   2775
      Left            =   720
      TabIndex        =   0
      Top             =   2400
      Width           =   6975
      Begin btButtonEx.ButtonEx bttnbttnDoctorsIncomeReport 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Doctors Income  Report"
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
      Begin btButtonEx.ButtonEx bttnbttnDoctorsIncomeCatagorReport 
         Height          =   375
         Left            =   3600
         TabIndex        =   17
         Top             =   2280
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Doctors Income Catagory Report"
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
      Begin VB.Label lblIncome 
         Caption         =   "0.00"
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label lblPatients 
         Caption         =   "0.00"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label5 
         Caption         =   "Doctor Income"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "No Of Patients"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmPeriodDoctorReport.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblDate"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected Day"
      TabPicture(1)   =   "frmPeriodDoctorReport.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Period"
      TabPicture(2)   =   "frmPeriodDoctorReport.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DTPicker3"
      Tab(2).Control(1)=   "DTPicker2"
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(3)=   "Label1"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "This Month"
      TabPicture(3)   =   "frmPeriodDoctorReport.frx":0496
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).ControlCount=   0
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   -70200
         TabIndex        =   4
         Top             =   900
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   77856769
         CurrentDate     =   39489
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -74040
         TabIndex        =   3
         Top             =   900
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   77856769
         CurrentDate     =   39489
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -72240
         TabIndex        =   2
         Top             =   900
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   77856769
         CurrentDate     =   39489
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
         Left            =   -72240
         TabIndex        =   7
         Top             =   900
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   375
         Left            =   -70800
         TabIndex        =   6
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   375
         Left            =   -74640
         TabIndex        =   5
         Top             =   900
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPeriodDoctorReport"
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
    Select Case SSTab1.Tab
    
    Case 0
    .Open "Select* From tblPatientFacility Where (Staff_ID = " & DtcDoctor.BoundText & " and FullyPaid = 1 and cancelled = 0 and refund = 0 and AppointmentDate = '" & Date & "' )  Order By PatientFacility_ID "
    
    Case 1
    .Open "Select* From tblPatientFacility Where (Staff_ID = " & DtcDoctor.BoundText & " and FullyPaid = 1 and cancelled = 0 and refund = 0 and AppointmentDate = '" & DTPicker1.Value & "' )  Order By PatientFacility_ID "
    
    Case 2
    .Open "Select* From tblPatientFacility Where (Staff_ID = " & DtcDoctor.BoundText & " and FullyPaid = 1 and cancelled = 0 and refund = 0 and AppointmentDate Between  '" & DTPicker2.Value & "'  and '" & DTPicker3.Value & "' )  Order By PatientFacility_ID "
    
    Case 3
    Call PrintThismonthDoctorIncome
    Exit Sub
    End Select
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    
    TemDoctorFee = TemDoctorFee + !personaldue
    .MoveNext
    Loop
    
    
    lblIncome.Caption = Format(TemDoctorFee, "0.00")
    lblPatients.Caption = .RecordCount
    
    End With
    

End Sub

Private Sub ClearValus()
lblIncome.Caption = ""
lblPatients.Caption = ""

End Sub

Private Sub FindTital()
If IsNumeric(DtcDoctor.BoundText) = False Then Exit Sub

With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    .Open "SELECT tblDoctor.*, tblTitle.Title FROM tblTitle RIGHT JOIN tblDoctor ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID Where (Doctor_Id = " & DtcDoctor.BoundText & ") "
    TemTital = !Title
    If .State = 1 Then .Close

End With
End Sub

Private Sub bttnbttnDoctorsIncomeCatagorReport_Click()
Const pershape = "SHAPE {"
Const Sql = "SELECT tblTitle.Title , tblDoctor.DoctorListedName, tblPatientFacility.* FROM ((tblDoctor LEFT OUTER JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) RIGHT OUTER JOIN tblPatientFacility ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID)"
Const PostSHape = "}  AS cmdDoctorIncomeC COMPUTE cmdDoctorIncomeC, SUM(cmdDoctorIncomeC.'PersonalFee') AS TotalPersonalFee, ANY(cmdDoctorIncomeC.'Title') AS TitalName, COUNT(cmdDoctorIncomeC.'PersonalFee') AS TotalPatients BY 'DoctorListedName','PersonalFee'"
' SHAPE {SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.* FROM ( ( tblDoctor RIGHT OUTER JOIN tblPatientFacility ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID ) LEFT OUTER JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID )}  AS cmdDoctorIncomeC COMPUTE cmdDoctorIncomeC, SUM(cmdDoctorIncomeC.'PersonalFee') AS TotalPersonalFee, ANY(cmdDoctorIncomeC.'Title') AS TitalName, COUNT(cmdDoctorIncomeC.'PersonalFee') AS TotalPatients BY 'DoctorName','PersonalFee'
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1
    If .rscmdDoctorIncomeC_Grouping.State = 1 Then .rscmdDoctorIncomeC_Grouping.Close
    If PayToDoctor = True Then
        .Commands!cmdDoctorIncomeC_Grouping.CommandText = pershape & Sql & " Where (FullyPaid = 1) and (cancelled = 0) and (refund = 0) Order By DoctorListedName" & PostSHape
    Else
        .Commands!cmdDoctorIncomeC_Grouping.CommandText = pershape & Sql & " Where (FullyPaid = 1) and (cancelled = 0) and (refund = 0) and (patientabsent = 0) Order By DoctorListedName" & PostSHape
    End If
    .cmdDoctorIncomeC_Grouping
    
    Set dtrDoctorIncomeC.DataSource = DataEnvironment1
    dtrDoctorIncomeC.Sections("PageFooter").Controls("lblAdd").Caption = LongAd
    
    dtrDoctorIncomeC.Show
End With


'' SHAPE {SELECT tblTitle.Title , tblDoctor.DoctorName, tblPatientFacility.* FROM ((tblDoctor LEFT OUTER JOIN tblTitle ON tblDoctor.Doct6orTitle_ID = tblTitle.Title_ID) RIGHT OUTER JOIN tblPatientFacility ON tbl6Doctor.Doctor_ID = tblPatientFacility.Staff_ID)}  AS cmdDoctorIncomeC COMPUTE cmdDoctorIncomeC, SUM(cmdDoctorIncomeC.'PersonalFee') AS TotalPersonalFee, ANY(cmdDoctorIncomeC.'Title') AS TitalName, COUNT(cmdDoctorIncomeC.'PersonalFee') AS TotalPatients BY 'DoctorName','PersonalFee'


End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnbttnDoctorsIncomeReport_Click()

CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

If SSTab1.Tab = 3 Then
Call PrintThismonthDoctorIncome
Exit Sub
End If

If IsNumeric(DtcDoctor.BoundText) = False Then Exit Sub

'SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblPatientFacility.PatientFacility_ID FROM tblPatientMainDetails INNER JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID ORDER BY tblPatientFacility.PatientFacility_ID
Dim TemPatient As Long

With DataEnvironment1.rssqlTem12
    If .State = 1 Then .Close
    Select Case SSTab1.Tab
    
    Case 0
    .Open " SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblPatientFacility.PatientFacility_ID FROM tblPatientMainDetails INNER JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID Where (Staff_ID = " & DtcDoctor.BoundText & " and FullyPaid = 1 and cancelled = 0 and refund = 0 and AppointmentDate = '" & Date & "' ) ORDER BY tblPatientFacility.PatientFacility_ID"

    
    Case 1
    .Open " SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblPatientFacility.PatientFacility_ID FROM tblPatientMainDetails INNER JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID Where (Staff_ID = " & DtcDoctor.BoundText & " and FullyPaid = 1 and cancelled = 0 and refund = 0 and AppointmentDate = '" & DTPicker1.Value & "' ) ORDER BY tblPatientFacility.PatientFacility_ID"

    Case 2
    
    .Open " SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblPatientFacility.PatientFacility_ID FROM tblPatientMainDetails INNER JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID Where (Staff_ID = " & DtcDoctor.BoundText & " and FullyPaid = 1 and cancelled = 0 and refund = 0 and AppointmentDate Between  '" & DTPicker2.Value & "'  and '" & DTPicker3.Value & "' ) ORDER BY tblPatientFacility.PatientFacility_ID"
    Case 3
    Call PrintThismonthDoctorIncome
    End Select
    
    TemPatient = .RecordCount
End With

        Select Case SSTab1.Tab

        Case 0
            dtrDoctorIncomePeriodRt.Sections("Section2").Controls.Item("lblDate").Caption = "Date    : " & Format(Date, DefaultLongDate)
        Case 1
            dtrDoctorIncomePeriodRt.Sections("Section2").Controls.Item("lbldate").Caption = "Date    : " & DTPicker1.Value
        Case 2
            dtrDoctorIncomePeriodRt.Sections("Section2").Controls.Item("lbldate").Caption = "Date From   : " & DTPicker2.Value & "    To   : " & DTPicker3.Value
        End Select

    With dtrDoctorIncomePeriodRt

        If HospitalDetails = True Then
            .Sections("ReportHeader10").Controls.Item("RptName").Caption = InstitutionName
            .Sections("ReportHeader10").Controls.Item("RptAddress").Caption = InstitutionAddress
            .Sections("Section5").Controls.Item("lblAd1").Caption = LongAd
            .Sections("Section3").Controls.Item("lblAds").Caption = ShortAd
        Else
            .Sections("ReportHeader10").Controls.Item("RptName").Caption = Empty
            .Sections("ReportHeader10").Controls.Item("RptAddress").Caption = Empty
            .Sections("Section5").Controls.Item("lblAd1").Caption = LongAd
            .Sections("Section3").Controls.Item("lblAds").Caption = ShortAd
            
        End If
      
    .Sections("Section2").Controls.Item("lblDoctorName").Caption = TemTital & "  " & DtcDoctor.Text
    .Sections("Section5").Controls.Item("lblTotalPation").Caption = TemPatient

    Set .DataSource = DataEnvironment1.rssqlTem12
    .Show
        
    End With

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
    .Commands!cnmdDoctorIncome_Grouping.CommandText = PreSHape & Sql & " Where appointmentdate Between  '" & firstDay & "' and '" & LastDay & "' and FullyPaid = 1 and refund = 0 and cancelled = 0 " & PostSHape
    .cnmdDoctorIncome_Grouping
    dtrDoctorsIncomeReport2.Sections("PageHeader").Controls.Item("lblDate").Caption = "Date  From   : " & Format(firstDay, DefaultLongDate) & "  To  " & Format(LastDay, DefaultLongDate)
    Set dtrDoctorsIncomeReport2.DataSource = DataEnvironment1
    dtrDoctorsIncomeReport2.Show
End With

End Sub

Private Sub DtcDoctor_Click(Area As Integer)
If IsNumeric(DtcDoctor.BoundText) = False Then Exit Sub
If IsNumeric(DtcDoctor.BoundText) = False Then A = MsgBox("Select Doctot", vbCritical + vbOKOnly, "Error"): Exit Sub
ClearValus
FindDoctorIncomePatients
FindTital
End Sub

Private Sub DTPicker1_Change()
If IsNumeric(DtcDoctor.BoundText) = False Then A = MsgBox("Select Doctot", vbCritical + vbOKOnly, "Error"): Exit Sub
ClearValus
FindDoctorIncomePatients
End Sub

Private Sub DTPicker2_Change()
If IsNumeric(DtcDoctor.BoundText) = False Then A = MsgBox("Select Doctot", vbCritical + vbOKOnly, "Error"): Exit Sub
ClearValus
FindDoctorIncomePatients

End Sub

Private Sub DTPicker3_Change()
If IsNumeric(DtcDoctor.BoundText) = False Then A = MsgBox("Select Doctot", vbCritical + vbOKOnly, "Error"): Exit Sub
ClearValus
FindDoctorIncomePatients

End Sub

Private Sub Form_Load()
lblDate = Date
DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
SSTab1.Tab = 0
Call FillDoctorName
    If UserAuthority <> AuthorityOwner Then
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
        SSTab1.TabVisible(3) = False
    End If

End Sub

Private Sub FillDoctorName()

With DataEnvironment1.rssqlTemDoctorView1
    If .State = 1 Then .Close
    
        If SurnameFirst = True Then
            .Open "Select* From tblDoctor Order By DoctorListedName"
            Set DtcDoctor.RowSource = DataEnvironment1.rssqlTemDoctorView1
            DtcDoctor.BoundColumn = "Doctor_Id"
            DtcDoctor.ListField = "DoctorListedName"
        Else
            .Open "Select* From tblDoctor Order By DoctorName"
            Set DtcDoctor.RowSource = DataEnvironment1.rssqlTemDoctorView1
            DtcDoctor.BoundColumn = "Doctor_Id"
            DtcDoctor.ListField = "DoctorName"
        End If


End With


End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 3 Then
    DtcDoctor.Enabled = False
Else
    DtcDoctor.Enabled = True
End If
End Sub

