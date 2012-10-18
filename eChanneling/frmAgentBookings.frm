VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAgentBookings 
   Caption         =   "Agent Bookings"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAgentBookings.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin btButtonEx.ButtonEx bttnExit 
      Height          =   495
      Left            =   13920
      TabIndex        =   6
      Top             =   8640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Bindings        =   "frmAgentBookings.frx":0442
      Height          =   5535
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   17
      FixedCols       =   0
      MergeCells      =   1
      AllowUserResizing=   3
      DataMember      =   "cmmdAgentBookings_Grouping"
      _NumberOfBands  =   2
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   5
      _Band(0)._MapCol(0)._Name=   "InstitutionName"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "DoctorName"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "AgentCode"
      _Band(0)._MapCol(2)._RSIndex=   4
      _Band(0)._MapCol(3)._Name=   "AgentName"
      _Band(0)._MapCol(3)._RSIndex=   2
      _Band(0)._MapCol(3)._Hidden=   -1  'True
      _Band(0)._MapCol(4)._Name=   "AgentDoctorName"
      _Band(0)._MapCol(4)._RSIndex=   3
      _Band(0)._MapCol(4)._Hidden=   -1  'True
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   14
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   0
      _Band(1)._ParentBand=   0
      _Band(1)._NumMapCols=   14
      _Band(1)._MapCol(0)._Name=   "FirstName"
      _Band(1)._MapCol(0)._RSIndex=   3
      _Band(1)._MapCol(1)._Name=   "PatientFacility_ID"
      _Band(1)._MapCol(1)._RSIndex=   4
      _Band(1)._MapCol(1)._Alignment=   7
      _Band(1)._MapCol(2)._Name=   "FullyPaid"
      _Band(1)._MapCol(2)._RSIndex=   5
      _Band(1)._MapCol(3)._Name=   "Cancelled"
      _Band(1)._MapCol(3)._RSIndex=   6
      _Band(1)._MapCol(4)._Name=   "Refund"
      _Band(1)._MapCol(4)._RSIndex=   7
      _Band(1)._MapCol(5)._Name=   "PersonalDue"
      _Band(1)._MapCol(5)._RSIndex=   8
      _Band(1)._MapCol(5)._Alignment=   7
      _Band(1)._MapCol(6)._Name=   "InstitutionDue"
      _Band(1)._MapCol(6)._RSIndex=   9
      _Band(1)._MapCol(6)._Alignment=   7
      _Band(1)._MapCol(7)._Name=   "TotalDue"
      _Band(1)._MapCol(7)._RSIndex=   10
      _Band(1)._MapCol(7)._Alignment=   7
      _Band(1)._MapCol(8)._Name=   "BookingDate"
      _Band(1)._MapCol(8)._RSIndex=   11
      _Band(1)._MapCol(9)._Name=   "BookingTime"
      _Band(1)._MapCol(9)._RSIndex=   12
      _Band(1)._MapCol(10)._Name=   "AppointmentDate"
      _Band(1)._MapCol(10)._RSIndex=   13
      _Band(1)._MapCol(11)._Name=   "InstitutionName"
      _Band(1)._MapCol(11)._RSIndex=   0
      _Band(1)._MapCol(12)._Name=   "InstitutionCode"
      _Band(1)._MapCol(12)._RSIndex=   1
      _Band(1)._MapCol(13)._Name=   "DoctorName"
      _Band(1)._MapCol(13)._RSIndex=   2
   End
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2520
      Left            =   3600
      TabIndex        =   3
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   80412673
      CurrentDate     =   39471
   End
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   12600
      TabIndex        =   1
      Top             =   8640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Print"
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
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2520
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   80412673
      CurrentDate     =   39471
   End
   Begin btButtonEx.ButtonEx btnSame 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   8640
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Print All Bookings with same appointment date"
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
   Begin btButtonEx.ButtonEx btnDifferent 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   9120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Print All Bookings with different appointment date"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date To"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmAgentBookings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CSetPrinter As New cSetDfltPrinter

Const PreSHape = "SHAPE {"
Const Sql = "SELECT tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode, tblDoctor.DoctorName, tblPatientMainDetails.FirstName, tblPatientFacility.PatientFacility_ID, tblPatientFacility.FullyPaid, tblPatientFacility.Cancelled, tblPatientFacility.Refund, tblPatientFacility.PersonalDue, tblPatientFacility.InstitutionDue, tblPatientFacility.TotalDue, tblPatientFacility.BookingDate, tblPatientFacility.BookingTime, tblPatientFacility.AppointmentDate FROM tblInstitutions RIGHT JOIN (tblDoctor RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID WHERE "
Const PostSHape = "(((tblPatientFacility.HospitalFacility_ID)=10)) AND (((tblPatientFacility.PaymentMode)= 'Agent'))}  AS cmmdAgentBookings COMPUTE cmmdAgentBookings, ANY(cmmdAgentBookings.'InstitutionName') AS AgentName, ANY(cmmdAgentBookings.'DoctorName') AS AgentDoctorName, ANY(cmmdAgentBookings.'InstitutionCode') AS AgentCode BY 'InstitutionName','DoctorName'"


Private Sub ListResults()
    'On Error Resume Next
    
    Grid1.Visible = False
    
    With DataEnvironment1
        If .rscmmdAgentBookings_Grouping.State = 1 Then .rscmmdAgentBookings_Grouping.Close
        .Commands!cmmdAgentBookings_Grouping.CommandText = PreSHape & Sql & " (bookingdate Between  '" & MonthView1.Value & "' and '" & MonthView2.Value & "')  and (paymentmode = 'Agent') and " & PostSHape
        .cmmdAgentBookings_Grouping
        If .rscmmdAgentBookings_Grouping.State = 0 Then .rscmmdAgentBookings_Grouping.Open
    End With
    
    Grid1.Visible = True
    Set Grid1.DataSource = DataEnvironment1
    Grid1.Refresh
    
'    Grid1.AllowUserResizing = flexResizeColumns
'
'    Grid1.MergeCells = flexMergeFree
    Grid1.ExpandAll
    
End Sub

Private Sub FormatGrid()
With Grid1
    .ColWidth(0, 0) = 2200
    .ColWidth(1, 0) = 2200
    .ColWidth(2, 0) = 1200
    .ColWidth(0, 1) = 2000
    .ColWidth(1, 1) = 1200
    .ColWidth(2, 1) = 1
    .ColWidth(3, 1) = 1
    .ColWidth(4, 1) = 1
    .ColWidth(5, 1) = 1
    .ColWidth(6, 1) = 1
    .ColWidth(7, 1) = 1
    .ColWidth(8, 1) = 1500
    .ColWidth(9, 1) = 1500
    .ColWidth(10, 1) = 1700
    .ColWidth(11, 1) = 1
    .ColWidth(12, 1) = 1
    .ColWidth(13, 1) = 1
    .ColWidth(14, 1) = 1
    .ColWidth(15, 1) = 1
    .ColWidth(16, 1) = 1
    .ColWidth(17, 1) = 1
    .ColWidth(18, 1) = 1
    
End With
End Sub

Private Sub btnDifferent_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1
        If .rscmmdAgentBookings_Grouping.State = 1 Then .rscmmdAgentBookings_Grouping.Close
        .Commands!cmmdAgentBookings_Grouping.CommandText = PreSHape & Sql & " (bookingdate Between  '" & MonthView1.Value & "' and '" & MonthView2.Value & "') and ((appointmentdate > '" & MonthView2.Value & "') ) and (paymentmode = 'Agent') and " & PostSHape
        .cmmdAgentBookings_Grouping
        If .rscmmdAgentBookings_Grouping.State = 0 Then .rscmmdAgentBookings_Grouping.Open
    End With

If HospitalDetails = True Then
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptName").Caption = InstitutionName
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateFrom").Caption = Format(MonthView1.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateTo").Caption = Format(MonthView2.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageFooter").Controls.Item("Ad1").Caption = LongAd
    DataReportBookingThroughAgents.Refresh
    DataReportBookingThroughAgents.Show
Else
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptName").Caption = Empty
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptAddress").Caption = Empty
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateFrom").Caption = Format(MonthView1.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateTo").Caption = Format(MonthView2.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageFooter").Controls.Item("Ad1").Caption = LongAd
    DataReportBookingThroughAgents.Refresh
    DataReportBookingThroughAgents.Show
End If

End Sub

Private Sub btnSame_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1
        If .rscmmdAgentBookings_Grouping.State = 1 Then .rscmmdAgentBookings_Grouping.Close
        .Commands!cmmdAgentBookings_Grouping.CommandText = PreSHape & Sql & " (bookingdate Between  '" & MonthView1.Value & "' and '" & MonthView2.Value & "')  and ((appointmentdate Between  '" & MonthView1.Value & "' and '" & MonthView2.Value & "') ) and (paymentmode = 'Agent') and " & PostSHape
        .cmmdAgentBookings_Grouping
        If .rscmmdAgentBookings_Grouping.State = 0 Then .rscmmdAgentBookings_Grouping.Open
    End With

If HospitalDetails = True Then
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptName").Caption = InstitutionName
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateFrom").Caption = Format(MonthView1.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateTo").Caption = Format(MonthView2.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageFooter").Controls.Item("Ad1").Caption = LongAd
    DataReportBookingThroughAgents.Refresh
    DataReportBookingThroughAgents.Show
Else
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptName").Caption = Empty
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptAddress").Caption = Empty
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateFrom").Caption = Format(MonthView1.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateTo").Caption = Format(MonthView2.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageFooter").Controls.Item("Ad1").Caption = LongAd
    DataReportBookingThroughAgents.Refresh
    DataReportBookingThroughAgents.Show
End If

End Sub

Private Sub bttnExit_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1
        If .rscmmdAgentBookings_Grouping.State = 1 Then .rscmmdAgentBookings_Grouping.Close
        .Commands!cmmdAgentBookings_Grouping.CommandText = PreSHape & Sql & " (bookingdate Between  '" & MonthView1.Value & "' and '" & MonthView2.Value & "')  and (paymentmode = 'Agent') and " & PostSHape
        .cmmdAgentBookings_Grouping
        If .rscmmdAgentBookings_Grouping.State = 0 Then .rscmmdAgentBookings_Grouping.Open
    End With

If HospitalDetails = True Then
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptName").Caption = InstitutionName
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateFrom").Caption = Format(MonthView1.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateTo").Caption = Format(MonthView2.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageFooter").Controls.Item("Ad1").Caption = LongAd
    DataReportBookingThroughAgents.Refresh
    DataReportBookingThroughAgents.Show
Else
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptName").Caption = Empty
    DataReportBookingThroughAgents.Sections("ReportHeader").Controls.Item("RptAddress").Caption = Empty
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateFrom").Caption = Format(MonthView1.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageHeader").Controls.Item("RptDateTo").Caption = Format(MonthView2.Value, DefaultLongDate)
    DataReportBookingThroughAgents.Sections("PageFooter").Controls.Item("Ad1").Caption = LongAd
    DataReportBookingThroughAgents.Refresh
    DataReportBookingThroughAgents.Show
End If

End Sub

Private Sub Form_Load()
    MonthView1 = Date
    MonthView2 = Date
    ListResults
    FormatGrid
End Sub

Private Sub Grid1_Click()
'    Text1.Text = Grid1.Col
'    Text2.Text = Grid1.Row
FormatGrid
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Call ListResults
End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
    Call ListResults
End Sub
