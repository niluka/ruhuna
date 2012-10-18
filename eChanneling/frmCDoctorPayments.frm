VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmCDoctorPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Payments"
   ClientHeight    =   11145
   ClientLeft      =   375
   ClientTop       =   1755
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCDoctorPayments.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FramePatientList 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCDoctorPayments.frx":0442
         Height          =   3135
         Left            =   120
         TabIndex        =   28
         Top             =   3840
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataMember      =   "cmmdDoctorPaymentTem"
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "PatientFacility_ID"
            Caption         =   "Receipt"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "FirstName"
            Caption         =   "Patient"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "BookingDate"
            Caption         =   "Booked On"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "FullyPaid"
            Caption         =   "Payment"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Paid"
               FalseValue      =   "To Pay"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Cancelled"
            Caption         =   "Cancellations"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Cancelled"
               FalseValue      =   ""
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Refund"
            Caption         =   "Refunds"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Refunded"
               FalseValue      =   ""
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "PatientAbsent"
            Caption         =   "Presecence"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "Absent"
               FalseValue      =   ""
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   7
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "PersonalDue"
            Caption         =   "Doctor Payments"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """Rs. ""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Alignment       =   2
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   2505.26
            EndProperty
         EndProperty
      End
      Begin VB.ListBox ListDatesAndSecessions 
         Height          =   2820
         IntegralHeight  =   0   'False
         ItemData        =   "frmCDoctorPayments.frx":0461
         Left            =   11280
         List            =   "frmCDoctorPayments.frx":0463
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
      Begin VB.ListBox ListConsultants 
         Height          =   2820
         IntegralHeight  =   0   'False
         ItemData        =   "frmCDoctorPayments.frx":0465
         Left            =   3600
         List            =   "frmCDoctorPayments.frx":0467
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox ListSpecialities 
         Height          =   2820
         IntegralHeight  =   0   'False
         ItemData        =   "frmCDoctorPayments.frx":0469
         Left            =   120
         List            =   "frmCDoctorPayments.frx":046B
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.ListBox List1 
         Height          =   1980
         Left            =   14160
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox ListConsultantIDs 
         Height          =   1980
         Left            =   5880
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox ListSpecialityIDs 
         Height          =   1980
         Left            =   2760
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   14655
         Begin VB.OptionButton OptionPayable 
            Caption         =   "All Payable"
            Height          =   255
            Left            =   2400
            TabIndex        =   6
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton OptionPaid 
            Caption         =   "Paid"
            Height          =   255
            Left            =   8520
            TabIndex        =   8
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton OptionToPay 
            Caption         =   "To Pay"
            Height          =   255
            Left            =   5760
            TabIndex        =   7
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton OptionAll 
            Caption         =   "All"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.ListBox ListSecessionIDs 
         Height          =   540
         Left            =   14160
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid GridList1 
         Height          =   3135
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5530
         _Version        =   393216
      End
      Begin btButtonEx.ButtonEx bttnCloseList 
         Height          =   375
         Left            =   13200
         TabIndex        =   11
         Top             =   8760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin btButtonEx.ButtonEx bttnView 
         Height          =   375
         Left            =   11520
         TabIndex        =   10
         Top             =   8760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "View - To Pay"
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
      Begin btButtonEx.ButtonEx bttnPayDoctor 
         Height          =   375
         Left            =   9840
         TabIndex        =   9
         Top             =   8760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Pay Doctor"
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
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2820
         Left            =   8160
         TabIndex        =   3
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   56426497
         CurrentDate     =   39472
      End
      Begin btButtonEx.ButtonEx bttnViewAll 
         Height          =   375
         Left            =   11520
         TabIndex        =   27
         Top             =   8160
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "View - All"
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
      Begin VB.Label lblAbsentFee 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   31
         Top             =   7800
         Width           =   1935
      End
      Begin VB.Label lblAbsentDescreption 
         BackStyle       =   0  'Transparent
         Caption         =   "Absent Patients"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   7800
         Width           =   5535
      End
      Begin VB.Label lblDuePayments1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   29
         Top             =   8760
         Width           =   1935
      End
      Begin VB.Label lblPaidToDoctor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   23
         Top             =   8160
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Payments"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   8160
         Width           =   4455
      End
      Begin VB.Label lblCRA 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancellations / Refunds"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   7440
         Width           =   4455
      End
      Begin VB.Label lblTotalDoctorFee 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   21
         Top             =   7080
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Doctor Fee"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   7080
         Width           =   4455
      End
      Begin VB.Label lblDuePayments 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   17
         Top             =   8760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Due"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   8760
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "List Criteria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblTotalRepayment 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   7440
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmCDoctorPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemPatientID As Long
    Dim TemHospitalFacilityID As Long
    Dim TemstaffID As Long
    Dim TemPatientFacilityID As Long
    Dim TemBillId As Long
    Dim TemSecessionID  As Integer
    Dim IsACancellation As Boolean
    Dim IsARefund As Boolean
    Dim ChoosenOption As OptionButton
    
    Dim TemPreviousDate As Date
    Dim TemPreviousSecession As Long
    Dim TemPreviousDoctorID As Long
    Dim TemPreviousOptionChanged As Boolean
    
    
    
Private Sub bttnViewAll_Click()
Dim TemResponce As Integer
If Not IsNumeric(ListConsultantIDs.Text) Then
    TemResponce = MsgBox("You have not selected a doctor", vbCritical, "Doctor")
    ListConsultants.SetFocus
    Exit Sub
End If
With DataEnvironment1.rssqlTem14
    If .State = 1 Then .Close
    If ListSecessionIDs.Text = "All" Then
        If PayToDoctor = True Then
            .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.paidtostaff = false and  tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and   appointmentdate = #" & MonthView1.Value & "#  ORDER BY tblPatientFacility.DaySerial")
        Else
            .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.paidtostaff = false and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and   appointmentdate = #" & MonthView1.Value & "#  and patientabsent = false ORDER BY tblPatientFacility.DaySerial")
        End If
    Else
        If Not IsNumeric(ListSecessionIDs.Text) Or Val(ListSecessionIDs.Text) < 0 Then
            TemResponce = MsgBox("You have not selected a secession", vbCritical, "Secession")
            ListDatesAndSecessions.SetFocus
            Exit Sub
        End If
        If PayToDoctor = True Then
            .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.paidtostaff = false and  tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and    appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.secession = " & ListSecessionIDs.Text & " ORDER BY tblPatientFacility.DaySerial")
        Else
            .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.paidtostaff = false and  tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and    appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.secession = " & ListSecessionIDs.Text & " and  patientabsent = false ORDER BY tblPatientFacility.DaySerial")
        End If

    
    End If
    Set DataReportDoctorPayments.DataSource = DataEnvironment1.rssqlTem14
    DataReportDoctorPayments.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportDoctorPayments.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportDoctorPayments.Sections("Section4").Controls.Item("Rptlbltopic").Caption = "Doctor Payment"
    DataReportDoctorPayments.Sections("Section2").Controls.Item("rptdoctorname").Caption = FindDoctorFromID(Val(ListConsultantIDs.Text))
    DataReportDoctorPayments.Sections("Section2").Controls.Item("RptDate").Caption = Format(MonthView1.Value, DefaultLongDate)
    If ListSecessionIDs.Text = "All" Then
    Else
        DataReportDoctorPayments.Sections("Section2").Controls.Item("rptSecession").Caption = ListDatesAndSecessions.Text
    End If
    DataReportDoctorPayments.Sections("Section5").Controls.Item("lblpaidamount").Caption = Format(lblDuePayments.Caption, "#0.00")
'    DataReportDoctorPayments.Sections("Section5").Controls.Item("lblusername").Caption = FindStaffFromID(UserID)
    DataReportDoctorPayments.Sections("Section5").Controls.Item("lbldoctorname").Caption = "Nurse Signature"
DataReportDoctorPayments.Show
End With
End Sub

Private Sub Form_Load()
    Call Setcolours
    FillSpeciality
    MonthView1.Value = Date
    If UserAuthority <> AuthorityOwner Then
        MonthView1.Enabled = False
    End If
    
End Sub

Private Sub FillSpeciality()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblspeciality order by speciality "
    .Open
    ListSpecialities.AddItem "All"
    ListSpecialityIDs.AddItem "All"
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
    
    

Private Sub ListConsultants_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then MonthView1.SetFocus
MonthView1.BackColor = &H80C0FF
End Sub

Private Sub ListDatesAndSecessions_Click()
    ListSecessionIDs.ListIndex = ListDatesAndSecessions.ListIndex
    Call FillDataGrid
End Sub

Private Sub ListSpecialities_Click()
    ListSpecialityIDs.ListIndex = ListSpecialities.ListIndex
    ListConsultantIDs.Clear
    ListConsultants.Clear
    Call FormatSecessionsList
    If ListSpecialities.Text = "All" Then
        ListAllConsultants
    ElseIf ListSpecialities.Text <> "All" And IsNumeric(ListSpecialityIDs.Text) = True Then
        ListSelectedConsultants
    Else
        FormatGridConsultants
    End If
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

Private Sub ListConsultants_Click()
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    Call FillDates
End Sub

Private Sub FillDates()
    ListDatesAndSecessions.Visible = False
    Me.MousePointer = vbHourglass
    Call FormatSecessionsList
    
    Dim TemBookingDate As Date
    Dim NowROw As Long
    
    With DataEnvironment1.rssqlTem5
        If .State = 1 Then .Close
        .Source = "SELECT tblfacilitysecession.* from tblfacilitysecession where hospitalfacility_ID = 10 and staff_ID = " & Val(ListConsultantIDs.Text)
        If .State = 0 Then .Open
        If .RecordCount = 0 Then .Close: ListDatesAndSecessions.Visible = True:     Me.MousePointer = vbDefault: Exit Sub
        .Close
    End With
        
        TemBookingDate = MonthView1.Value
        
        With DataEnvironment1.rssqlTem4
            If .State = 1 Then .Close
            .Source = "Select * from tblfacilitysecession where hospitalfacility_ID =  10  and staff_ID = " & Val(ListConsultantIDs.Text) & " and AlteredDate = #" & TemBookingDate & "# order by StartingTime"
            .Open
            If .RecordCount <> 0 Then
                If !fulldayleave = False Then
                    While .EOF = False
                        ListSecessionIDs.AddItem !facilitysecession_ID
                        ListDatesAndSecessions.AddItem FindSecessionFromID(!facilitysecession_ID)
                        .MoveNext
                    Wend
                End If
                .Close
            Else
                If .State = 1 Then .Close
                .Source = "Select * from tblfacilitysecession where hospitalfacility_ID = 10 and staff_ID = " & Val(ListConsultantIDs.Text) & " and SecessionWeekday = " & Weekday(TemBookingDate) & " order by StartingTime"
                .Open
                If .RecordCount <> 0 Then
                    While .EOF = False
                        ListSecessionIDs.AddItem !facilitysecession_ID
                        ListDatesAndSecessions.AddItem FindSecessionFromID(!facilitysecession_ID)
                        .MoveNext
                    Wend
                End If
            End If
        End With
    
    ListDatesAndSecessions.Visible = True
    Me.MousePointer = vbDefault
End Sub


Private Sub ListSpecialities_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ListConsultants.SetFocus
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    Call FormatSecessionsList
    Call FillDates
End Sub

Private Sub FormatGridSpeciality()
    ListSpecialities.Clear
    ListSpecialityIDs.Clear
End Sub

Private Sub FormatGridConsultants()
    ListConsultants.Clear
    ListConsultantIDs.Clear
End Sub

  
Private Sub bttnView_Click()
Dim TemResponce As Integer

If Val(lblDuePayments.Caption) <= 0 Then TemResponce = MsgBox("You have already Paid to doctor", vbCritical, "No Due Payments"): Exit Sub

If Not IsNumeric(ListConsultantIDs.Text) Then
    TemResponce = MsgBox("You have not selected a doctor", vbCritical, "Doctor")
    ListConsultants.SetFocus
    Exit Sub
End If

With DataEnvironment1.rssqlTem14
    If .State = 1 Then .Close
    If ListSecessionIDs.Text = "All" Then
        If PayToDoctor = True Then
            .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and   appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.paidtostaff = false ORDER BY tblPatientFacility.DaySerial")
        Else
            .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and   appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.paidtostaff = false and patientabsent = false ORDER BY tblPatientFacility.DaySerial")
        End If
    Else
        If Not IsNumeric(ListSecessionIDs.Text) Or Val(ListSecessionIDs.Text) < 0 Then
            TemResponce = MsgBox("You have not selected a secession", vbCritical, "Secession")
            ListDatesAndSecessions.SetFocus
            Exit Sub
        End If
        If PayToDoctor = True Then
            .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and    appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.secession = " & ListSecessionIDs.Text & " and tblPatientFacility.paidtostaff = false ORDER BY tblPatientFacility.DaySerial")
        Else
            .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and    appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.secession = " & ListSecessionIDs.Text & " and tblPatientFacility.paidtostaff = false  and patientabsent = false ORDER BY tblPatientFacility.DaySerial")
        End If
    End If
   
    Set DataReportDoctorPaymentsView.DataSource = DataEnvironment1.rssqlTem14
    DataReportDoctorPaymentsView.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportDoctorPaymentsView.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportDoctorPaymentsView.Sections("Section4").Controls.Item("Rptlbltopic").Caption = "Doctor Payment"
    DataReportDoctorPaymentsView.Sections("Section2").Controls.Item("rptdoctorname").Caption = FindDoctorFromID(Val(ListConsultantIDs.Text))
    
    DataReportDoctorPaymentsView.Sections("Section2").Controls.Item("RptDate").Caption = Format(MonthView1.Value, DefaultLongDate)
    
    If ListSecessionIDs.Text = "All" Then
        
    Else
        DataReportDoctorPaymentsView.Sections("Section2").Controls.Item("rptSecession").Caption = ListDatesAndSecessions.Text
    End If
    
    DataReportDoctorPaymentsView.Sections("Section5").Controls.Item("lblpaidamount").Caption = Format(lblDuePayments.Caption, "#0.00")
'    DataReportDoctorPayments.Sections("Section5").Controls.Item("lblusername").Caption = FindStaffFromID(UserID)
    DataReportDoctorPaymentsView.Sections("Section5").Controls.Item("lbldoctorname").Caption = "Nurse Signature"
    DataReportDoctorPaymentsView.Show
End With
End Sub


Private Sub FormatSecessionsList()
    ListSecessionIDs.Clear
    ListDatesAndSecessions.Clear
    ListSecessionIDs.AddItem "All"
    ListDatesAndSecessions.AddItem "All"
End Sub

Private Sub FillSecessionsList()
If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    With DataEnvironment1.rssqlTem4
        If .State = 1 Then .Close
        .Source = "Select * from tblfacilitysecession where hospitalfacility_ID =  10  and staff_ID = " & Val(ListConsultantIDs.Text) & " and AlteredDate = #" & MonthView1.Value & "# order by StartingTime"
        .Open
            If .RecordCount <> 0 Then
                While .EOF = False
                    ListSecessionIDs.AddItem !facilitysecession_ID
                    ListDatesAndSecessions.AddItem FindSecessionFromID(!facilitysecession_ID)
                    .MoveNext
                Wend
                .Close
            Else
                If .State = 1 Then .Close
                .Source = "Select * from tblfacilitysecession where hospitalfacility_ID = 10 and staff_ID = " & Val(ListConsultantIDs.Text) & " and SecessionWeekday = " & Weekday(MonthView1.Value) & " order by StartingTime"
                .Open
                If .RecordCount <> 0 Then
                    While .EOF = False
                        ListSecessionIDs.AddItem !facilitysecession_ID
                        ListDatesAndSecessions.AddItem FindSecessionFromID(!facilitysecession_ID)
                        .MoveNext
                    Wend
                End If
                .Close
            End If
        End With
End Sub


Private Sub FillDataGrid()
    
    Set DataGrid1.DataSource = Nothing
    
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    If ListDatesAndSecessions.ListIndex < 0 Then Exit Sub
    If Not IsNumeric(ListSecessionIDs.Text) And ListSecessionIDs.Text <> "All" Then Exit Sub
    
    Dim SqlSelect As String
    Dim SqlWhere As String
    Dim SqlOrderby As String
    
    SqlSelect = "SELECT tblPatientFacility.PatientFacility_ID, tblPatientMainDetails.FirstName, tblPatientFacility.BookingDate, tblPatientFacility.FullyPaid, tblPatientFacility.Cancelled, tblPatientFacility.Refund, tblPatientFacility.PatientAbsent, tblPatientFacility.PersonalDue "
    SqlSelect = SqlSelect & "FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID "
    
    SqlOrderby = "ORDER BY tblPatientFacility.Secession, tblPatientFacility.DaySerial"
    With DataEnvironment1.rssqlTem3
        If .State = 1 Then .Close
        If ListSecessionIDs.Text = "All" Then
            If OptionAll.Value = True Then
                SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#)) "
            ElseIf OptionPayable.Value = True Then
                If PayToDoctor = True Then
                    SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#)) "
                Else
                    SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PatientAbsent)=False)) "
                End If
            ElseIf OptionPaid.Value = True Then
                    SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) and ((tblPatientFacility.paidtostaff) = true)) "
            ElseIf OptionToPay.Value = True Then
                If PayToDoctor = True Then
                    SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PaidToSTaff)=False)) "
                Else
                    SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PaidToSTaff)=False) AND ((tblPatientFacility.PatientAbsent)=False)) "
                End If
            End If
        Else
            If OptionAll.Value = True Then
                SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#)) "
            ElseIf OptionPayable.Value = True Then
                If PayToDoctor = True Then
                    SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#)) "
                Else
                    SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PatientAbsent)=False)) "
                End If
            ElseIf OptionPaid.Value = True Then
                    SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) and ((tblPatientFacility.paidtostaff) = true)) "
            ElseIf OptionToPay.Value = True Then
                If PayToDoctor = True Then
                    SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PaidToSTaff)=False)) "
                Else
                    SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PaidToSTaff)=False) AND ((tblPatientFacility.PatientAbsent)=False)) "
                End If
            End If
        End If
        
        .Source = SqlSelect & SqlWhere & SqlOrderby
        If .State = 0 Then .Open
        Set DataGrid1.DataSource = DataEnvironment1
        DataGrid1.DataMember = "sqltem3"
    End With
    
    With DataEnvironment1.rssqlTem2
        If .State = 1 Then .Close
        If ListSecessionIDs.Text = "All" Then
                If PayToDoctor = True Then
                    SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PaidToSTaff)=False)) "
                Else
                    SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PaidToSTaff)=False) AND ((tblPatientFacility.PatientAbsent)=False)) "
                End If
        Else
                If PayToDoctor = True Then
                    SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND (tblPatientFacility.PaidToSTaff=False)) "
                Else
                    SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PaidToSTaff)=False) AND ((tblPatientFacility.PatientAbsent)=False)) "
                End If
        End If
        .Source = "SELECT DISTINCTROW Sum(tblPatientFacility.PersonalDue) AS SumOfPersonal From tblPatientFacility " & SqlWhere
        .Open
        If Not IsNull(!sumofpersonal) Then
            lblDuePayments.Caption = !sumofpersonal
            lblDuePayments1.Caption = "Rs. " & Format(!sumofpersonal, "#0.00")
        Else
            lblDuePayments.Caption = 0
            lblDuePayments1.Caption = "Rs. " & Format(0, "#0.00")
        End If
        
        If .State = 1 Then .Close
        If ListSecessionIDs.Text = "All" Then
            SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#)) "
        Else
            SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#)) "
        End If
        .Source = "SELECT DISTINCTROW Sum(tblPatientFacility.Personalfee) AS SumOfPersonal From tblPatientFacility " & SqlWhere
        .Open
        If Not IsNull(!sumofpersonal) Then
            lblTotalDoctorFee.Caption = "Rs. " & Format(!sumofpersonal, "#0.00")
        Else
            lblTotalDoctorFee.Caption = "Rs. " & Format(0, "#0.00")
        End If
        
        If PayToDoctor = True Then
            lblAbsentDescreption.Caption = "Absent Patients' Fee(Included in doctor payments)"
        Else
            lblAbsentDescreption.Caption = "Absent Patients' Fee(Not included in doctor payments)"
        End If
        
        If .State = 1 Then .Close
        If ListSecessionIDs.Text = "All" Then
            SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PatientAbsent)=true)) "
        Else
            SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND ((tblPatientFacility.PatientAbsent)=true)) "
        End If
        
        .Source = "SELECT DISTINCTROW Sum(tblPatientFacility.Personaldue) AS SumOfPersonal From tblPatientFacility " & SqlWhere
        .Open
        If Not IsNull(!sumofpersonal) Then
            lblAbsentFee.Caption = "Rs. " & Format(!sumofpersonal, "#0.00")
        Else
            lblAbsentFee.Caption = "Rs. " & Format(0, "#0.00")
        End If
        
        If .State = 1 Then .Close
        If ListSecessionIDs.Text = "All" Then
                If PayToDoctor = True Then
                    SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND  (((tblPatientFacility.Cancelled) = true) or ((tblPatientFacility.refund)= True) )) "
                Else
                    SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND  (((tblPatientFacility.Cancelled) = true) or ((tblPatientFacility.patientabsent) = true) or ((tblPatientFacility.refund)= True) )) "
                End If
        Else
                If PayToDoctor = True Then
                    SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND (((tblPatientFacility.Cancelled) = true) or ((tblPatientFacility.refund)= True) )) "
                Else
                    SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) AND (((tblPatientFacility.Cancelled) = true) or ((tblPatientFacility.patientabsent) = true) or ((tblPatientFacility.refund)= True) )) "
                End If
        End If
        .Source = "SELECT DISTINCTROW Sum(tblPatientFacility.Personalrefund) AS SumOfPersonal From tblPatientFacility " & SqlWhere
        .Open
        If Not IsNull(!sumofpersonal) Then
            lblTotalRepayment.Caption = "Rs. " & Format(!sumofpersonal, "#0.00")
        Else
            lblTotalRepayment.Caption = "Rs. " & Format(0, "#0.00")
        End If
        
        If .State = 1 Then .Close
        If ListSecessionIDs.Text = "All" Then
            SqlWhere = " WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) and ((tblPatientFacility.paidtostaff) = true)) "
        Else
            SqlWhere = " WHERE (((tblPatientFacility.Secession)= " & ListSecessionIDs.Text & " ) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.Staff_ID)= " & ListConsultantIDs.Text & " ) AND ((tblPatientFacility.AppointmentDate)=#" & MonthView1.Value & "#) and ((tblPatientFacility.paidtostaff) = true)) "
        End If
        .Source = "SELECT DISTINCTROW Sum(tblPatientFacility.Personaldue) AS SumOfPersonal From tblPatientFacility " & SqlWhere
        .Open
        If Not IsNull(!sumofpersonal) Then
            lblPaidToDoctor.Caption = "Rs. " & Format(!sumofpersonal, "#0.00")
        Else
            lblPaidToDoctor.Caption = "Rs. 0.00"
        End If
        If .State = 1 Then .Close
    End With
    
End Sub



Private Sub bttnCloseList_Click()
    Unload Me
End Sub

Private Sub bttnPayDoctor_Click()
Dim TemResponce As Integer
LoginSucceeded = False
If Not IsNumeric(ListConsultantIDs.Text) Then
    TemResponce = MsgBox("You have not selected a doctor", vbCritical, "Doctor")
    ListConsultants.SetFocus
    Exit Sub
End If

If Val(lblDuePayments.Caption) <= 0 Then TemResponce = MsgBox("You have already Paid to doctor", vbCritical, "No Due Payments"): Exit Sub

'If LoginSucceeded = False Then
TemResponce = MsgBox("Are You Sure You want to pay " & lblDuePayments.Caption & " to " & ListConsultants.Text & "?", vbCritical + vbYesNo, "Payments")
If TemResponce = vbNo Then Exit Sub


If DoctorPaymentDetailedReport = True Then
        With DataEnvironment1.rssqlTem14
            If .State = 1 Then .Close
            If ListSecessionIDs.Text = "All" Then
                If PayToDoctor = True Then
                    .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and   appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.paidtostaff = false ORDER BY tblPatientFacility.DaySerial")
                Else
                    .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and   appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.paidtostaff = false and patientabsent = false ORDER BY tblPatientFacility.DaySerial")
                End If
            Else
                If Not IsNumeric(ListSecessionIDs.Text) Or Val(ListSecessionIDs.Text) < 0 Then
                    TemResponce = MsgBox("You have not selected a secession", vbCritical, "Secession")
                    ListDatesAndSecessions.SetFocus
                    Exit Sub
                End If
                If PayToDoctor = True Then
                    .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and    appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.secession = " & ListSecessionIDs.Text & " and tblPatientFacility.paidtostaff = false ORDER BY tblPatientFacility.DaySerial")
                Else
                    .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and    appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.secession = " & ListSecessionIDs.Text & " and tblPatientFacility.paidtostaff = false  and patientabsent = false ORDER BY tblPatientFacility.DaySerial")
                End If
            End If
           
            Set DataReportDoctorPaymentsView.DataSource = DataEnvironment1.rssqlTem14
            DataReportDoctorPaymentsView.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
            DataReportDoctorPaymentsView.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
            DataReportDoctorPaymentsView.Sections("Section4").Controls.Item("Rptlbltopic").Caption = "Doctor Payment"
            DataReportDoctorPaymentsView.Sections("Section2").Controls.Item("rptdoctorname").Caption = FindDoctorFromID(Val(ListConsultantIDs.Text))
            
            DataReportDoctorPaymentsView.Sections("Section2").Controls.Item("RptDate").Caption = Format(MonthView1.Value, DefaultLongDate)
            
            If ListSecessionIDs.Text = "All" Then
                
            Else
                DataReportDoctorPaymentsView.Sections("Section2").Controls.Item("rptSecession").Caption = ListDatesAndSecessions.Text
            End If
            
            DataReportDoctorPaymentsView.Sections("Section5").Controls.Item("lblpaidamount").Caption = Format(lblDuePayments.Caption, "#0.00")
        '    DataReportDoctorPayments.Sections("Section5").Controls.Item("lblusername").Caption = FindStaffFromID(UserID)
            DataReportDoctorPaymentsView.Sections("Section5").Controls.Item("lbldoctorname").Caption = "Nurse Signature"
            DataReportDoctorPaymentsView.Show
        End With


Else



    
    With DataEnvironment1.rssqlTem14
        If .State = 1 Then .Close
        If ListSecessionIDs.Text = "All" Then
            If PayToDoctor = True Then
                .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.Personaldue > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and tblPatientFacility.paidtostaff = false and  appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.paidtostaff = false ORDER BY tblPatientFacility.DaySerial")
            Else
                .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.Personaldue > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and tblPatientFacility.paidtostaff = false and  appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.paidtostaff = false and patientabsent = false ORDER BY tblPatientFacility.DaySerial")
            End If
        Else
            If Not IsNumeric(ListSecessionIDs.Text) Or Val(ListSecessionIDs.Text) < 0 Then
                TemResponce = MsgBox("You have not selected a secession", vbCritical, "Secession")
                ListDatesAndSecessions.SetFocus
                Exit Sub
            End If
            If PayToDoctor = True Then
                .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.Personaldue > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and tblPatientFacility.paidtostaff = false and  appointmentdate = #" & MonthView1.Value & "# and tblPatientFacility.secession = " & ListSecessionIDs.Text & " ORDER BY tblPatientFacility.DaySerial")
      '          .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.Personaldue > 0) and tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & " and tblPatientFacility.paidtostaff = false and  appointmentdate = #" & MonthView1.Value & "# ORDER BY tblPatientFacility.DaySerial")
            Else
                .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.Personaldue > 0) and (tblPatientFacility.staff_ID = " & ListConsultantIDs.Text & ") and (tblPatientFacility.paidtostaff = false) and  (appointmentdate = #" & MonthView1.Value & "#) and (tblPatientFacility.secession = " & ListSecessionIDs.Text & ")and (patientabsent = false) ORDER BY tblPatientFacility.DaySerial")
            End If
        End If
        If .RecordCount = 0 Then Exit Sub
        
        Set DataReportDoctorPayments.DataSource = DataEnvironment1.rssqlTem14
        
        If HospitalDetails = True Then
            DataReportDoctorPayments.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
            DataReportDoctorPayments.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        Else
            DataReportDoctorPayments.Sections("Section4").Controls.Item("RptName").Caption = Empty
            DataReportDoctorPayments.Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        End If
        
        DataReportDoctorPayments.Sections("Section4").Controls.Item("Rptlbltopic").Caption = "Doctor Payment"
        DataReportDoctorPayments.Sections("Section2").Controls.Item("rptdoctorname").Caption = FindDoctorFromID(Val(ListConsultantIDs.Text))
        DataReportDoctorPayments.Sections("Section2").Controls.Item("RptDate").Caption = Format(MonthView1.Value, DefaultLongDate)
        DataReportDoctorPayments.Sections("Section2").Controls.Item("rptSecession").Caption = ListDatesAndSecessions.Text
        DataReportDoctorPayments.Sections("Section5").Controls.Item("lblpaidamount").Caption = Format(lblDuePayments.Caption, "#0.00")
    '    DataReportDoctorPayments.Sections("Section5").Controls.Item("lblusername").Caption = FindStaffFromID(UserID)
        DataReportDoctorPayments.Sections("Section5").Controls.Item("lbldoctorname").Caption = "Nurse Signature"
        DataReportDoctorPayments.Show
    
    End With

End If

Dim TemStaffPaymentID As Long

If Val(lblDuePayments.Caption) = 0 Then Exit Sub

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblStaffPayment"
    If .State = 0 Then .Open
    .AddNew
    !HospitalFacility_ID = 10
    !Staff_ID = (Val(ListConsultantIDs.Text))
    !PaidAmount = Val(lblDuePayments.Caption)
    !paiddate = Date
    !paidtime = Time
    !user_ID = UserID
    !isadoctor = True
    .Update
    TemStaffPaymentID = !staffpayment_ID
    .Close
End With

Dim i As Integer

        With DataEnvironment1.rssqlTem
            If .State = 1 Then .Close
            
        If ListSecessionIDs.Text = "All" Then
            If PayToDoctor = True Then
                .Open ("SELECT tblPatientFacility.* FROM tblPatientFacility WHERE HospitalFacility_ID = 10 AND FullyPaid = true AND Personaldue > 0 and staff_ID = " & ListConsultantIDs.Text & " and  appointmentdate = #" & MonthView1.Value & "# and paidtostaff = false")
            Else
                .Open ("SELECT tblPatientFacility.* FROM tblPatientFacility WHERE HospitalFacility_ID = 10 AND FullyPaid = true AND Personaldue > 0 and staff_ID = " & ListConsultantIDs.Text & " and  appointmentdate = #" & MonthView1.Value & "# and paidtostaff = false and patientabsent = false")
            End If
        Else
            If Not IsNumeric(ListSecessionIDs.Text) Or Val(ListSecessionIDs.Text) < 0 Then
                TemResponce = MsgBox("You have not selected a secession", vbCritical, "Secession")
                ListDatesAndSecessions.SetFocus
                Exit Sub
            End If
            If PayToDoctor = True Then
                .Open ("SELECT tblPatientFacility.* FROM tblPatientFacility WHERE HospitalFacility_ID = 10 AND FullyPaid = true AND Personaldue > 0 and staff_ID = " & ListConsultantIDs.Text & " and  appointmentdate = #" & MonthView1.Value & "# and secession = " & ListSecessionIDs.Text & " and paidtostaff = false")
            Else
                .Open ("SELECT tblPatientFacility.* FROM tblPatientFacility WHERE HospitalFacility_ID = 10 AND FullyPaid = true AND Personaldue > 0 and staff_ID = " & ListConsultantIDs.Text & " and  appointmentdate = #" & MonthView1.Value & "# and secession = " & ListSecessionIDs.Text & " and paidtostaff = false  and patientabsent = false")
            End If
        End If
            
            If .State = 0 Then .Open
            If .RecordCount = 0 Then Exit Sub
            While .EOF = False
                !paidtostaff = True
                !paidtostaffon = Date
                !paidtostaffuser = UserID
                !staffpayment = !personaldue
                !staffpayment_ID = TemStaffPaymentID
                .Update
                .MoveNext
            Wend
            .Close
        End With

FillDataGrid

End Sub



Private Sub MonthView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ListDatesAndSecessions.SetFocus
End Sub

Private Sub MonthView1_LostFocus()
MonthView1.BackColor = &H8000000F

End Sub

Private Sub OptionAll_Click()
    If OptionAll.Value = True Then
        FillDataGrid
    End If
End Sub





Private Sub Setcolours()
    
    bttnCloseList.BackColor = BttnBackColour
    bttnCloseList.ForeColor = BttnForeColour
    bttnViewAll.BackColor = BttnBackColour
    bttnViewAll.ForeColor = BttnForeColour
    bttnPayDoctor.BackColor = BttnBackColour
    bttnPayDoctor.ForeColor = BttnForeColour
    bttnView.BackColor = BttnBackColour
    bttnView.ForeColor = BttnForeColour
    FramePatientList.BackColor = FrmBackColour
    FramePatientList.ForeColor = FrmForeColour
    Me.BackColor = FrameBackColour
    Me.ForeColor = FrameForeColour
    OptionPayable.BackColor = FrmBackColour
    OptionPayable.ForeColor = FrmForeColour
    OptionAll.BackColor = FrmBackColour
    OptionAll.ForeColor = FrmForeColour
    OptionToPay.BackColor = FrmBackColour
    OptionToPay.ForeColor = FrmForeColour
    OptionPaid.BackColor = FrmBackColour
    OptionPaid.ForeColor = FrmForeColour
    Frame1.BackColor = FrmBackColour
    Frame1.ForeColor = FrmForeColour
End Sub


Private Sub OptionPaid_Click()
    If OptionPaid.Value = True Then
        FillDataGrid
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionPayable_Click()
    If OptionPayable.Value = True Then
        FillDataGrid
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionToPay_Click()
    If OptionToPay.Value = True Then
        FillDataGrid
    End If
    TemPreviousOptionChanged = True
End Sub

