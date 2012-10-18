VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPeriodReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Bookings"
   ClientHeight    =   5160
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
   Icon            =   "frmPeriodReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8520
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   4680
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
      Top             =   1320
      Width           =   6975
      Begin btButtonEx.ButtonEx bttnChannelingPatients 
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print Channeling Patients"
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
      Begin VB.Label lblTotalPatients 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Total Patients"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   7646
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
      TabPicture(0)   =   "frmPeriodReport.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDate"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected Day"
      TabPicture(1)   =   "frmPeriodReport.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Period"
      TabPicture(2)   =   "frmPeriodReport.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(2)=   "DTPicker2"
      Tab(2).Control(3)=   "DTPicker3"
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   -70200
         TabIndex        =   4
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61997057
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
         Format          =   61997057
         CurrentDate     =   39489
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -72240
         TabIndex        =   2
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61997057
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
         Left            =   2760
         TabIndex        =   7
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   375
         Left            =   -70800
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   375
         Left            =   -74640
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPeriodReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CSetPrinter As New cSetDfltPrinter


Private Sub bttnChannelingPatients_Click()
Const PreSHape = "SHAPE {"
Const Sql = "SELECT tblPatientFacility.*, tblDoctor.DoctorListedName, tblTitle.Title FROM tblTitle RIGHT JOIN (tblPatientFacility LEFT JOIN tblDoctor ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID Where "
Const PostSHape = " (((tblPatientFacility.HospitalFacility_ID) = 10)) } AS cmmdAllDoctorPatients COMPUTE cmmdAllDoctorPatients, COUNT(cmmdAllDoctorPatients.'PatientFacility_ID') AS TotalPatientCount, COUNT(cmmdAllDoctorPatients.'CancelledNull') AS TotalCancellations, SUM(cmmdAllDoctorPatients.'RefundNull') AS TotalRefunds, SUM(cmmdAllDoctorPatients.'FullyPaidNull') AS TotalFullyPaid, COUNT(cmmdAllDoctorPatients.'PatientAbsentNull') AS TotalAbsent, ANY(cmmdAllDoctorPatients.'Title') AS DoctorTitle BY 'DoctorListedName'"

CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
' SHAPE {SELECT tblPatientFacility.*, tblDoctor.DoctorName, tblTitle.TitleFROM tblTitle RIGHT JOIN (tblPatientFacility LEFT JOIN tblDoctor ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID Where (((tblPatientFacility.HospitalFacility_ID) = 10)) }  AS cmmdAllDoctorPatients COMPUTE cmmdAllDoctorPatients, COUNT(cmmdAllDoctorPatients.'PatientFacility_ID') AS TotalPatientCount, COUNT(cmmdAllDoctorPatients.'CancelledNull') AS TotalCancellations, SUM(cmmdAllDoctorPatients.'RefundNull') AS TotalRefunds, SUM(cmmdAllDoctorPatients.'FullyPaidNull') AS TotalFullyPaid, COUNT(cmmdAllDoctorPatients.'PatientAbsentNull') AS TotalAbsent, ANY(cmmdAllDoctorPatients.'Title') AS DoctorTitle BY 'DoctorName'
    
    With DataEnvironment1
    
        If .rscmmdAllDoctorPatients_Grouping.State = 1 Then .rscmmdAllDoctorPatients_Grouping.Close
        If UserAuthority <> AuthorityOwner Then
        Select Case SSTab1.Tab
        
        Case 0
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate = '" & Date & "') and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        Case 1
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate = '" & DTPicker1.Value & "') and (tblPatientFacility.DaySerial % " & IncomeDeflation & " = 0) and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        Case 2
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and (tblPatientFacility.DaySerial % " & IncomeDeflation & " = 0) and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        End Select
        Else
        Select Case SSTab1.Tab
        
        Case 0
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate = '" & Date & "') and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        Case 1
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate = '" & DTPicker1.Value & "') and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        Case 2
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        End Select
        End If
    
    End With
    
    
        Select Case SSTab1.Tab

        Case 0
            dtrAllPatientsbyDate.Sections("ReportHeader").Controls.Item("lbldate").Caption = "Date    : " & Format(Date, DefaultLongDate)
        Case 1
            dtrAllPatientsbyDate.Sections("ReportHeader").Controls.Item("lbldate").Caption = "Date    : " & DTPicker1.Value
        Case 2
            dtrAllPatientsbyDate.Sections("ReportHeader").Controls.Item("lbldate").Caption = "Date From   : " & DTPicker2.Value & "    To   : " & DTPicker3.Value
        End Select
    
    With dtrAllPatientsbyDate
    
        If HospitalDetails = True Then
            .Sections("ReportHeader").Controls.Item("InstitutionName").Caption = InstitutionName
            .Sections("ReportHeader").Controls.Item("InstitutionAddress").Caption = InstitutionAddress
            .Sections("ReportFooter").Controls.Item("ad1").Caption = LongAd
            
        Else
            .Sections("ReportHeader").Controls.Item("InstitutionName").Caption = Empty
            .Sections("ReportHeader").Controls.Item("InstitutionAddress").Caption = Empty
            .Sections("ReportFooter").Controls.Item("ad1").Caption = LongAd
        End If
        Set .DataSource = DataEnvironment1
        .Show
    End With

End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub FindPatients()
Dim TemPatients As Long

Const PreSHape = "SHAPE {"
Const Sql = "SELECT tblPatientFacility.*, tblDoctor.DoctorName, tblTitle.Title FROM tblTitle RIGHT JOIN (tblPatientFacility LEFT JOIN tblDoctor ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID Where "
Const PostSHape = " (((tblPatientFacility.HospitalFacility_ID) = 10)) } AS cmmdAllDoctorPatients COMPUTE cmmdAllDoctorPatients, COUNT(cmmdAllDoctorPatients.'PatientFacility_ID') AS TotalPatientCount, COUNT(cmmdAllDoctorPatients.'CancelledNull') AS TotalCancellations, SUM(cmmdAllDoctorPatients.'RefundNull') AS TotalRefunds, SUM(cmmdAllDoctorPatients.'FullyPaidNull') AS TotalFullyPaid, COUNT(cmmdAllDoctorPatients.'PatientAbsentNull') AS TotalAbsent, ANY(cmmdAllDoctorPatients.'Title') AS DoctorTitle BY 'DoctorName'"

    With DataEnvironment1
    
        If .rscmmdAllDoctorPatients_Grouping.State = 1 Then .rscmmdAllDoctorPatients_Grouping.Close
        If UserAuthority <> AuthorityOwner Then
        Select Case SSTab1.Tab
        
        Case 0
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate = '" & Date & "') and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        Case 1
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate = '" & DTPicker1.Value & "') and (tblPatientFacility.DaySerial % " & IncomeDeflation & " = 0) and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        Case 2
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and (tblPatientFacility.DaySerial % " & IncomeDeflation & " = 0) and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        End Select
        Else
        Select Case SSTab1.Tab
        
        Case 0
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate = '" & Date & "') and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        Case 1
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate = '" & DTPicker1.Value & "') and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        Case 2
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " (appointmentdate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and (fullypaid = 1) and (cancelled = 0) and (refund = 0) and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
        End Select
        End If
        Do While .rscmmdAllDoctorPatients_Grouping.EOF = False
        
            TemPatients = Val(TemPatients) + Val(.rscmmdAllDoctorPatients_Grouping!TotalFullyPaid)
        
        .rscmmdAllDoctorPatients_Grouping.MoveNext
        Loop
        
        If .rscmmdAllDoctorPatients_Grouping.State = 1 Then .rscmmdAllDoctorPatients_Grouping.Close
        
        lblTotalPatients.Caption = TemPatients
        
    End With
    
End Sub

Private Sub DTPicker1_Change()
Call FindPatients
End Sub

Private Sub DTPicker2_Change()
Call FindPatients
End Sub

Private Sub DTPicker3_Change()
Call FindPatients
End Sub

Private Sub Form_Load()
lblDate = Date
DTPicker1.Value = Date - 1
DTPicker1.MaxDate = Date - 1
DTPicker2.Value = Date - 1
DTPicker2.MaxDate = Date - 1
DTPicker3.Value = Date - 1
DTPicker3.MaxDate = Date - 1
Call FindPatients
    If UserAuthority <> AuthorityOwner And UserAuthority <> AuthorityOwnerCOvered Then
        SSTab1.TabVisible(1) = True
        SSTab1.TabVisible(2) = True
    End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then Call FindPatients
End Sub

