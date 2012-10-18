VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSelectedDoctorPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pyments for Doctors for selected days appointments"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7215
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2370
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   65601537
      CurrentDate     =   39480
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   4080
      TabIndex        =   6
      Top             =   480
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   65601537
      CurrentDate     =   39480
   End
   Begin btButtonEx.ButtonEx bttnPaymentsDueSummery 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Doctor Payments to Complete - (Summery)"
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
   Begin btButtonEx.ButtonEx bttnPaymentsCompletedSummery 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Doctor payments completed - (Summery)"
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
   Begin btButtonEx.ButtonEx bttnTotalPaymentsSummery 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Doctor payments - (Summery)"
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
   Begin btButtonEx.ButtonEx bttnTotaPaymentsDetail 
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   3120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Doctor payments - (Details)"
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
   Begin btButtonEx.ButtonEx bttnPaymentsCompletedDetail 
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   3720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Doctor payments completed - (Details)"
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
   Begin btButtonEx.ButtonEx bttnPaymentsDueDetails 
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   4320
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Doctor payments to complete - (Details)"
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
      Caption         =   "To"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelectedDoctorPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CSetPrinter As New cSetDfltPrinter
Const PreSHape = "SHAPE {"
Const Sql = "SELECT tblPatientFacility.*, tblDoctor.DoctorName, tblPatientMainDetails.FirstName, tblStaff.StaffName, tblStaffPayment.* FROM tblStaff RIGHT JOIN (((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblDoctor ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) LEFT JOIN tblStaffPayment ON tblPatientFacility.StaffPayment_ID = tblStaffPayment.StaffPayment_ID) ON tblStaff.Staff_ID = tblStaffPayment.User_ID where "
Const PostSHape = "(((tblPatientFacility.HospitalFacility_ID)=10))}  AS cmmdDoctorPayments COMPUTE cmmdDoctorPayments, ANY(cmmdDoctorPayments.'DoctorName') AS PaidDoctorName, ANY(cmmdDoctorPayments.'StaffName') AS PaidStaffName, ANY(cmmdDoctorPayments.'PaidDate') AS DoctorPaidDate, ANY(cmmdDoctorPayments.'PaidToSTaff') AS PaidOrNot, SUM(cmmdDoctorPayments.'PersonalDue') AS ToPayDoctor, SUM(cmmdDoctorPayments.'StaffPayment') AS PaidAmountToDoctor, ANY(cmmdDoctorPayments.'PaidDate') AS DocPaidDate, ANY(cmmdDoctorPayments.'PaidTime') AS DocPaidTime, ANY(cmmdDoctorPayments.'tblPatientFacility.Staff_ID') AS PaidID, ANY(cmmdDoctorPayments.'AppointmentDate') AS ForAppointmentDate BY 'tblPatientFacility.Staff_ID','AppointmentDate','PaidToSTaff','tblPatientFacility.StaffPayment_ID'"


Private Sub bttnPaymentsCompletedSummery_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    With DataEnvironment1
        If .rscmmdDoctorPayments_Grouping.State = 1 Then .rscmmdDoctorPayments_Grouping.Close
        .Commands!cmmdDoctorPayments_Grouping.CommandText = PreSHape & Sql & "  appointmentdate between  '" & MonthView1.Value & "'  and  '" & MonthView2.Value & "' and tblpatientfacility.paidtostaff = 1 and " & PostSHape
        .cmmdDoctorPayments_Grouping
    End With
    With DataReportPatientViceDoctorPaymentsSummery
        If HospitalDetails = True Then
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = InstitutionName
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = InstitutionAddress
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments completed (Summery)"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        Else
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = Empty
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = Empty
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments completed (Summery)"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        End If
        .Sections("ReportHeader").Controls("lblappdate").Caption = "Appointment Date : " & Format(Date, DefaultLongDate)
        .Show
    End With

End Sub

Private Sub bttnPaymentsDueDetails_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1
        If .rscmmdDoctorPayments_Grouping.State = 1 Then .rscmmdDoctorPayments_Grouping.Close
        .Commands!cmmdDoctorPayments_Grouping.CommandText = PreSHape & Sql & "  appointmentdate between  '" & MonthView1.Value & "'  and  '" & MonthView2.Value & "'  and  tblpatientfacility.paidtostaff =0  and " & PostSHape
        .cmmdDoctorPayments_Grouping
    End With
    With DataReportPatientViceDoctorPayments
        If HospitalDetails = True Then
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = InstitutionName
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = InstitutionAddress
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments to complete"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        Else
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = Empty
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = Empty
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments to complete"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        End If
        .Show
    End With

End Sub

Private Sub bttnPaymentsDueSummery_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1
        If .rscmmdDoctorPayments_Grouping.State = 1 Then .rscmmdDoctorPayments_Grouping.Close
        .Commands!cmmdDoctorPayments_Grouping.CommandText = PreSHape & Sql & "  appointmentdate between  '" & MonthView1.Value & "'  and  '" & MonthView2.Value & "' and  tblpatientfacility.paidtostaff =0  and " & PostSHape
        .cmmdDoctorPayments_Grouping
    End With
    With DataReportPatientViceDoctorPaymentsSummery
        If HospitalDetails = True Then
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = InstitutionName
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = InstitutionAddress
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments to complete (Summery)"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        Else
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = Empty
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = Empty
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments to complete (Summery)"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        End If
        .Sections("ReportHeader").Controls("lblappdate").Caption = "Appointment Date : " & Format(Date, DefaultLongDate)
        .Show
    End With

End Sub

Private Sub bttnTotalPaymentsSummery_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1
        If .rscmmdDoctorPayments_Grouping.State = 1 Then .rscmmdDoctorPayments_Grouping.Close
        .Commands!cmmdDoctorPayments_Grouping.CommandText = PreSHape & Sql & "  appointmentdate between  '" & MonthView1.Value & "'  and  '" & MonthView2.Value & "'    and " & PostSHape
        .cmmdDoctorPayments_Grouping
    End With
    With DataReportPatientViceDoctorPaymentsSummery
        If HospitalDetails = True Then
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = InstitutionName
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = InstitutionAddress
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments (Summery)"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        Else
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = Empty
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = Empty
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments (Summery)"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        End If
        .Sections("ReportHeader").Controls("lblappdate").Caption = "Appointment Date : " & Format(Date, DefaultLongDate)
        .Show
    End With
End Sub

Private Sub bttnTotaPaymentsDetail_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1
        If .rscmmdDoctorPayments_Grouping.State = 1 Then .rscmmdDoctorPayments_Grouping.Close
        .Commands!cmmdDoctorPayments_Grouping.CommandText = PreSHape & Sql & "  appointmentdate between  '" & MonthView1.Value & "'  and  '" & MonthView2.Value & "'  and " & PostSHape
        .cmmdDoctorPayments_Grouping
    End With
    With DataReportPatientViceDoctorPayments
        If HospitalDetails = True Then
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = InstitutionName
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = InstitutionAddress
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        Else
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = Empty
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = Empty
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        End If
        .Show
    End With
End Sub

Private Sub bttnPaymentsCompletedDetail_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1
        If .rscmmdDoctorPayments_Grouping.State = 1 Then .rscmmdDoctorPayments_Grouping.Close
        .Commands!cmmdDoctorPayments_Grouping.CommandText = PreSHape & Sql & "  appointmentdate between  '" & MonthView1.Value & "'  and  '" & MonthView2.Value & "'  and  tblpatientfacility.paidtostaff = 1 and " & PostSHape
        .cmmdDoctorPayments_Grouping
    End With
    With DataReportPatientViceDoctorPayments
        If HospitalDetails = True Then
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = InstitutionName
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = InstitutionAddress
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments Completed (summery)"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        Else
            .Sections("ReportHeader").Controls("lblinstitutionname").Caption = Empty
            .Sections("ReportHeader").Controls("lblinstitutionaddress").Caption = Empty
            .Sections("ReportHeader").Controls("lblReportTitle").Caption = "Doctor Payments Completed (summery)"
            .Sections("pagefooter").Controls("ad1").Caption = LongAd
        End If
        .Show
    End With
End Sub

Private Sub Form_Load()
    MonthView1.Value = Date
    MonthView2.Value = Date
    If UserAuthority <> AuthorityOwner Then
        MonthView1.Enabled = False
        MonthView2.Enabled = False
    End If
    
End Sub
