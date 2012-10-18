VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmBulkRefund 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Refund"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "frmBulkRefund.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   15240
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2820
      Left            =   6720
      TabIndex        =   17
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   67960833
      CurrentDate     =   39476
   End
   Begin VB.ListBox ListDatesAndSecessions 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      ItemData        =   "frmBulkRefund.frx":0442
      Left            =   6720
      List            =   "frmBulkRefund.frx":0444
      TabIndex        =   2
      ToolTipText     =   "List of Date, Secession, Maximum number per secession, Starting Time and already given numbers of the selected consultant"
      Top             =   3600
      Width           =   3015
   End
   Begin VB.ListBox ListPatientFacilities 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      ItemData        =   "frmBulkRefund.frx":0446
      Left            =   9960
      List            =   "frmBulkRefund.frx":0448
      Style           =   1  'Checkbox
      TabIndex        =   3
      ToolTipText     =   "List of number, patient, paid or not, cancelled or refunded, agent code and present or absent"
      Top             =   360
      Width           =   5055
   End
   Begin VB.ListBox ListConsultants 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      ItemData        =   "frmBulkRefund.frx":044A
      Left            =   3000
      List            =   "frmBulkRefund.frx":044C
      TabIndex        =   1
      ToolTipText     =   "List of Consultants of selected speciality"
      Top             =   360
      Width           =   3495
   End
   Begin VB.ListBox ListSpecialities 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      ItemData        =   "frmBulkRefund.frx":044E
      Left            =   240
      List            =   "frmBulkRefund.frx":0450
      TabIndex        =   0
      ToolTipText     =   "List of Specialities"
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox ListSecessionStartingTime 
      Height          =   4380
      Left            =   13920
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListPatientFacilityIDs 
      Height          =   4380
      Left            =   14280
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox ListConsultantIDs 
      Height          =   4380
      Left            =   14280
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox ListSecessionIDs 
      Height          =   4380
      Left            =   13920
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListDates 
      Height          =   4380
      Left            =   13920
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListSpecialityIDs 
      Height          =   4380
      Left            =   14040
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   13440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.ListBox ListSecessionMax 
      Height          =   4380
      Left            =   13920
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListRoomNo 
      Height          =   2220
      Left            =   4680
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin btButtonEx.ButtonEx bttnRefund 
      Height          =   375
      Left            =   11640
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Refund"
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
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "No.    Name             Paid Can/Ref  Agent  P/Ab"
      Height          =   255
      Left            =   9960
      TabIndex        =   16
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Secession"
      Height          =   255
      Left            =   6720
      TabIndex        =   15
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Consultant"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Speciality"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   2535
   End
   Begin VB.Shape BoxPatients 
      BackStyle       =   1  'Opaque
      Height          =   4770
      Left            =   9840
      Top             =   120
      Width           =   5295
   End
   Begin VB.Shape BoxDates 
      BackStyle       =   1  'Opaque
      Height          =   4770
      Left            =   6600
      Top             =   120
      Width           =   3255
   End
   Begin VB.Shape BoxConsultant 
      BackStyle       =   1  'Opaque
      Height          =   4770
      Left            =   2880
      Top             =   120
      Width           =   3735
   End
   Begin VB.Shape BoxSpeciality 
      BackStyle       =   1  'Opaque
      Height          =   4770
      Left            =   120
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmBulkRefund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
    Dim TemRoomNo As String
    Dim TemDoctorFee As Double
    Dim TemFDoctorFee As Double
    Dim TemADoctorFee As Double
    Dim TemInstitutionFee As Double
    Dim TemFInstitutionFee As Double
    Dim TemAInstitutionFee As Double
    Dim TemOtherFee As Double
    Dim TemFOtherFee As Double
    Dim TemAOtherFee As Double
    Dim TemSecession As Long
    Dim CSetPrinter As New cSetDfltPrinter
    Dim SecessionMax As Long
    Dim TemCanByPassOrder As Boolean
    Dim TemCalculateAppointment As Boolean
    Dim TemAgentRefNo As String
'    Dim TemSecession  As Integer
    Dim TemAgentCredit As Double
    Dim TemPatientID As Long
    Dim TemAgentMaxCredit As Double
    Dim TemPatientFacilityID As Long
    Dim TemAppointmentDate As Date
    Dim TemAppointmentTime As Date
    Dim TemDaySerial As Long
    Dim TemAgentBookingID As Long
    Dim TemSecessionStartingTime As Date
    Dim TemUsualDuration As Long
    Dim TemPatient As String
    Dim TemConsultant As String
    Dim TemNonCancelledVisits As Long
    Dim TemBillId As Long
    Dim TemPreviousDate As Date
    Dim TemTextForList As String


Private Sub bttnAllSecessionPatients_Click()

    Const PreSHape = "SHAPE {"
    Const Sql = "SELECT tblPatientFacility.*, tblDoctor.DoctorListedName, tblFacilitySecession.SecessionName, tblTitle.Title , tblPatientMainDetails.FirstName FROM tblTitle RIGHT JOIN (tblDoctor RIGHT JOIN (tblFacilitySecession RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblFacilitySecession.FacilitySecession_ID = tblPatientFacility.Secession) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID Where "
    Const PostSHape = "(((tblPatientFacility.HospitalFacility_ID)=10))}  AS AllSecessionPatients COMPUTE AllSecessionPatients, ANY(AllSecessionPatients.'DoctorListedName') AS SecessionDoctorName, ANY(AllSecessionPatients.'SecessionName') AS ThisSecessionName, SUM(AllSecessionPatients.'CancelledNull') AS AllCancelled, SUM(AllSecessionPatients.'RefundNull') AS AllRefunds, SUM(AllSecessionPatients.'PatientAbsentNull') AS AllAbsent, SUM(AllSecessionPatients.'FullyPaidNull') AS AllFullyPaid, COUNT(AllSecessionPatients.'PatientFacility_ID') AS AllPatients, ANY(AllSecessionPatients.'Title') AS DoctorTitle BY 'DoctorListedName','SecessionName' "
    
    With DataEnvironment1
    
    
        If .rsAllSecessionPatients_Grouping.State = 1 Then .rsAllSecessionPatients_Grouping.Close
        
        If DetailedCount = False Then
            If PayToDoctor = True Then
                .Commands!AllSecessionPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & MonthView2.Value & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and " & PostSHape
            Else
                .Commands!AllSecessionPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & MonthView2.Value & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and patientabsent = 0 and " & PostSHape
            End If
            .AllSecessionPatients_Grouping
            dtrAllSecessionPatients.Sections("Section1").Controls.Item("txt1").Visible = True
            dtrAllSecessionPatients.Sections("Section1").Controls.Item("txt2").Visible = False
            dtrAllSecessionPatients.Sections("Section1").Controls.Item("txt3").Visible = False
            dtrAllSecessionPatients.Sections("Section1").Controls.Item("txt4").Visible = False
            dtrAllSecessionPatients.Sections("Section1").Controls.Item("txt5").Visible = False
            dtrAllSecessionPatients.Sections("PageHeader").Controls.Item("lbl1").Visible = True
            dtrAllSecessionPatients.Sections("PageHeader").Controls.Item("lbl2").Visible = False
            dtrAllSecessionPatients.Sections("PageHeader").Controls.Item("lbl3").Visible = False
            dtrAllSecessionPatients.Sections("PageHeader").Controls.Item("lbl4").Visible = False
            dtrAllSecessionPatients.Sections("PageHeader").Controls.Item("lbl5").Visible = False
            dtrAllSecessionPatients.Sections("ReportFooter").Controls.Item("Function1").Visible = True
            dtrAllSecessionPatients.Sections("ReportFooter").Controls.Item("Function2").Visible = False
            dtrAllSecessionPatients.Sections("ReportFooter").Controls.Item("Function3").Visible = False
            dtrAllSecessionPatients.Sections("ReportFooter").Controls.Item("Function4").Visible = False
            dtrAllSecessionPatients.Sections("ReportFooter").Controls.Item("Function5").Visible = False
        Else
            If PayToDoctor = True Then
                .Commands!AllSecessionPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & MonthView2.Value & "' and  " & PostSHape
            Else
                .Commands!AllSecessionPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & MonthView2.Value & "' and patientabsent = 0 and " & PostSHape
            End If
            .AllSecessionPatients_Grouping
            dtrAllSecessionPatients.Sections("Section1").Controls.Item("txt1").Visible = True
            dtrAllSecessionPatients.Sections("Section1").Controls.Item("txt2").Visible = True
            dtrAllSecessionPatients.Sections("Section1").Controls.Item("txt3").Visible = True
            dtrAllSecessionPatients.Sections("Section1").Controls.Item("txt4").Visible = True
            dtrAllSecessionPatients.Sections("Section1").Controls.Item("txt5").Visible = True
            dtrAllSecessionPatients.Sections("PageHeader").Controls.Item("lbl1").Visible = True
            dtrAllSecessionPatients.Sections("PageHeader").Controls.Item("lbl2").Visible = True
            dtrAllSecessionPatients.Sections("PageHeader").Controls.Item("lbl3").Visible = True
            dtrAllSecessionPatients.Sections("PageHeader").Controls.Item("lbl4").Visible = True
            dtrAllSecessionPatients.Sections("PageHeader").Controls.Item("lbl5").Visible = True
            dtrAllSecessionPatients.Sections("ReportFooter").Controls.Item("Function1").Visible = True
            dtrAllSecessionPatients.Sections("ReportFooter").Controls.Item("Function2").Visible = True
            dtrAllSecessionPatients.Sections("ReportFooter").Controls.Item("Function3").Visible = True
            dtrAllSecessionPatients.Sections("ReportFooter").Controls.Item("Function4").Visible = True
            dtrAllSecessionPatients.Sections("ReportFooter").Controls.Item("Function5").Visible = True
        End If
    End With
    With dtrAllSecessionPatients
        If HospitalDetails = True Then
            .Sections("ReportHeader").Controls.Item("InstitutionName").Caption = InstitutionName
            .Sections("ReportHeader").Controls.Item("InstitutionAddress").Caption = InstitutionAddress
            .Sections("ReportHeader").Controls.Item("lbldate").Caption = Format(MonthView2.Value, DefaultLongDate)
            .Sections("ReportFooter").Controls.Item("ad1").Caption = LongAd
        Else
            .Sections("ReportHeader").Controls.Item("InstitutionName").Caption = Empty
            .Sections("ReportHeader").Controls.Item("InstitutionAddress").Caption = Empty
            .Sections("ReportHeader").Controls.Item("lbldate").Caption = Format(MonthView2.Value, DefaultLongDate)
            .Sections("ReportFooter").Controls.Item("ad1").Caption = LongAd
        End If
        Set .DataSource = DataEnvironment1
        .Show
    End With
End Sub

Private Sub Form_Load()
    MonthView2.Value = Date
    Call FormatGridSpeciality
    Call FormatGridConsultants
    Call FormatGridDates
    Call FormatGridPatients
    Call FillSpeciality
    Dim ingRet As Long
    Dim TabDates(1) As Long
    Dim TabPatientFacilities(6) As Long
    TabDates(0) = 48
    TabDates(1) = 166
    TabPatientFacilities(0) = 3 * 4
    TabPatientFacilities(1) = 15 * 4
    TabPatientFacilities(2) = 20 * 4
    TabPatientFacilities(3) = 28 * 4
    TabPatientFacilities(4) = 29 * 4
    ingRet = SendMessage(ListDates.hwnd, LB_SETTABSTOPS, 2, TabDates(0))
    ingRet = SendMessage(ListPatientFacilities.hwnd, LB_SETTABSTOPS, 7, TabPatientFacilities(0))

End Sub

Private Sub FormatGridSpeciality()
    ListSpecialities.Clear
    ListSpecialityIDs.Clear
End Sub

Private Sub FormatGridConsultants()
    ListConsultants.Clear
    ListConsultantIDs.Clear
End Sub

Private Sub FormatGridDates()
    ListDatesAndSecessions.Clear
    ListSecessionIDs.Clear
End Sub

Private Sub FormatGridPatients()
    ListPatientFacilities.Clear
    ListPatientFacilityIDs.Clear
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


Private Sub ListAllConsultants()
Call FormatGridConsultants
With DataEnvironment1.rssqlTem1
    If .State = 1 Then .Close
    If SurnameFirst = True Then
        .Source = "SELECT  tbldoctor.*  FROM  tbldoctor  order by DoctorListedName"
    Else
        .Source = "SELECT  tbldoctor.*  FROM  tbldoctor  order by DoctorName"
    End If
    .Open
    If .RecordCount = 0 Then Exit Sub
    While Not .EOF
            If SurnameFirst = True Then
                ListConsultants.AddItem !doctorlistedname
            Else
                ListConsultants.AddItem !doctorname
            End If
        ListConsultantIDs.AddItem !Doctor_ID
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
            .Source = "SELECT tbldoctor.* FROM tbldoctor where  doctorspeciality_ID = " & Val(ListSpecialityIDs.Text) & " order by DoctorName"
        End If
        .Open
        If .RecordCount = 0 Then Exit Sub
        While Not .EOF
            
            If SurnameFirst = True Then
                ListConsultants.AddItem !doctorlistedname
            Else
                ListConsultants.AddItem !doctorname
            End If
            
            ListConsultantIDs.AddItem !Doctor_ID
            .MoveNext
        Wend
        .Close
    End With
End Sub



Private Sub ListSpecialities_Click()
    ListSpecialityIDs.ListIndex = ListSpecialities.ListIndex
    ListConsultantIDs.Clear
    ListConsultants.Clear
    ListSecessionIDs.Clear
    ListSecessionMax.Clear
    ListSecessionStartingTime.Clear
    ListDates.Clear
    ListDatesAndSecessions.Clear
    ListRoomNo.Clear
    ListPatientFacilities.Clear
    ListPatientFacilityIDs.Clear
    If ListSpecialities.Text = "All" Then
        ListAllConsultants
    ElseIf ListSpecialities.Text <> "All" And IsNumeric(ListSpecialityIDs.Text) = True Then
        ListSelectedConsultants
    Else
        FormatGridConsultants
    End If
End Sub

Private Sub ListSpecialities_GotFocus()
    BoxSpeciality.BackColor = BttnBackColour ' vbRed
End Sub

Private Sub ListSpecialities_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    ListConsultants.SetFocus
    KeyCode = Empty
Else

End If
End Sub

Private Sub ListSpecialities_LostFocus()
     BoxSpeciality.BackColor = FrameBackColour ' - 2147483633
End Sub

Private Sub ListConsultants_Click()
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    Call FormatGridDates
    Call FormatGridPatients
    TemPatientFacilityID = 0
    TemDoctorFee = 0
    TemFDoctorFee = 0
    TemInstitutionFee = 0
    TemFInstitutionFee = 0
    TemOtherFee = 0
    TemAppointmentDate = Empty
    TemAppointmentTime = Empty
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    Call FillDates
End Sub


Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    FillDates
End Sub


Private Sub FillDates()
    ListDatesAndSecessions.Visible = False:     Me.MousePointer = vbHourglass:
    Call FormatGridDates
    ListSecessionIDs.AddItem "All"
    ListDatesAndSecessions.AddItem "All"
    Dim TemBookingDate As Date
    With DataEnvironment1.rssqlTem5
        If .State = 1 Then .Close
        .Source = "SELECT tblfacilitysecession.* from tblfacilitysecession where hospitalfacility_ID = 10 and staff_ID = " & Val(ListConsultantIDs.Text)
        If .State = 0 Then .Open
        If .RecordCount = 0 Then .Close: ListDatesAndSecessions.Visible = True:     Me.MousePointer = vbDefault: Exit Sub
        .Close
    End With
    TemBookingDate = MonthView2.Value
    With DataEnvironment1.rssqlTem4
        If .State = 1 Then .Close
        .Source = "Select * from tblfacilitysecession where hospitalfacility_ID =  10  and staff_ID = " & Val(ListConsultantIDs.Text) & " and AlteredDate = '" & TemBookingDate & "' order by StartingTime"
        .Open
            
        If .RecordCount <> 0 Then
            If !fulldayleave = False Then
                While .EOF = False
                    TemTextForList = !SecessionName
                    ListSecessionIDs.AddItem !facilitysecession_ID
                    ListDatesAndSecessions.AddItem TemTextForList
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
                    TemTextForList = !SecessionName
                    ListSecessionIDs.AddItem !facilitysecession_ID
                    ListDatesAndSecessions.AddItem TemTextForList
                    .MoveNext
                Wend
            End If
        End If
    End With
    ListDatesAndSecessions.Visible = True
    Me.MousePointer = vbDefault
End Sub

Private Sub ListDatesAndSecessions_Click()
    ListSecessionIDs.ListIndex = ListDatesAndSecessions.ListIndex
    TemAppointmentDate = MonthView2.Value
    Call FormatGridPatients
    If Not IsNumeric(ListSecessionIDs.Text) And ListSecessionIDs.Text <> "All" Then Exit Sub
    Call FillGridPatients
End Sub

Private Sub ListPatientFacilities_Click()
    ListPatientFacilityIDs.ListIndex = ListPatientFacilities.ListIndex
End Sub


Private Sub FillGridPatients()
    Dim TemTextForList As String

    Call FormatGridPatients
        With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        If ListSecessionIDs.Text = "All" Then
            .Source = "SELECT * from tblpatientfacility where hospitalfacility_ID = 10 and Staff_ID = " & Val(ListConsultantIDs.Text) & " and AppointmentDate = '" & MonthView2.Value & "' order by secession , DaySerial"
        Else
            .Source = "SELECT * from tblpatientfacility where hospitalfacility_ID = 10 and Staff_ID = " & Val(ListConsultantIDs.Text) & " and AppointmentDate = '" & MonthView2.Value & "' and Secession = " & Val(ListSecessionIDs.Text) & " order by DaySerial"
        End If
        .Open
        If .RecordCount = 0 Then Exit Sub
        While Not .EOF
            TemTextForList = !DaySerial & vbTab & Left(FindPatientByID(!patientid), 11)
            If !FullyPaid = 1 Then
                TemTextForList = TemTextForList & vbTab & "Paid"
            Else
                TemTextForList = TemTextForList & vbTab & Space(4)
            End If
            If !cancelled = True Then
                TemTextForList = TemTextForList & vbTab & "Cancel"
            ElseIf !Refund = True Then
                TemTextForList = TemTextForList & vbTab & "Refund"
            Else
                TemTextForList = TemTextForList & vbTab & Space(6)
            End If
            If Not IsNull(!Agent_ID) Then
                If !Agent_ID <> 0 Then
                    TemTextForList = TemTextForList & vbTab & Left(FindAgentCodeFromID(!Agent_ID), 3)
                Else
                    TemTextForList = TemTextForList & vbTab & Space(3)
                End If
            End If
            If !patientabsent = True Then
                TemTextForList = TemTextForList & vbTab & "A"
            Else
                TemTextForList = TemTextForList & vbTab & " "
            End If
            ListPatientFacilities.AddItem TemTextForList
            ListPatientFacilityIDs.AddItem !patientfacility_ID
            .MoveNext
        Wend
    End With

End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnRefund_Click()
    Dim i As Integer
    Dim TemResponce As Integer
    Dim ThisPatientID As Long
    Dim ThisStaffRepayR As Double
    Dim ThisInstitutionRepayR As Double
    Dim ThisOtherRepayR As Double
    Dim ThisRepayTotalR As Double
    
    For i = 0 To ListPatientFacilities.ListCount - 1
        If ListPatientFacilities.Selected(i) = True Then
            With DataEnvironment1.rssqlTem
                If .State = 1 Then .Close
                .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(ListPatientFacilityIDs.List(i))
                .Open
                If .RecordCount > 0 Then
'                    If !paidtostaff = 0 And !cancelled = 0 And !Refund = 0 And !FullyPaid = 1 Then
                    If !paidtostaff = False And !Refund = False And !cancelled = False And !Refund = False And !FullyPaid = True Then
                        ThisPatientID = !patientid
                        ThisStaffRepayR = Val(Format(!personaldue, "0.00"))
                        ThisInstitutionRepayR = Val(Format(!institutiondue, "0.00"))
                        ThisOtherRepayR = Val(Format(!otherdue, "0.00"))
                        ThisRepayTotalR = Val(Format(!TotalDue, "0.00"))
                        If IsNull(!Personalrefund) Then
                            !personaldue = !personalfee - ThisStaffRepayR
                            !Personalrefund = ThisStaffRepayR
                        Else
                            !personaldue = 0
                            !Personalrefund = ThisStaffRepayR
                        End If
                        If IsNull(!totalrefund) Then
                            !TotalDue = !totalfee - ThisStaffRepayR
                            !totalrefund = ThisStaffRepayR
                        Else
                            !TotalDue = !totalfee - ThisStaffRepayR
                            !totalrefund = ThisStaffRepayR
                        End If
                        !RepayComments = "refund"
                        !repaydate = !AppointmentDate
                        !repaytime = !appointmenttime
                        !cancelled = False
                        !Refund = True
                        !refundnull = 1
                        !repayUser_ID = UserID
                        !RefundToPatient = 1
                        .Update
                        .Close
                        .Source = "select * from tblpatientrepay"
                        .Open
                        .AddNew
                        !patient_ID = ThisPatientID
                        !HospitalFacility_ID = 10
                        !repayUser_ID = UserID
                        !repaydate = Date
                        !repaytime = Time
                        !StaffRepay = ThisStaffRepayR
                        !InstitutionRepay = 0
                        !OtherRepay = 0
                        !TotalRepay = ThisStaffRepayR
                        !Staff_ID = Val(ListConsultantIDs.Text)
                        !RepayComments = "refund"
                        !patientfacility_ID = TemPatientFacilityID
                        !RefundToPatient = 1
                        .Update
                        .Close
                    End If
                End If
            End With
        End If
    Next
    
    frmDoctorPayments.ListDatesAndSecessions_Click
    Unload Me
    
'    Call FormatGridPatients
'    Call ListDatesAndSecessions_Click

End Sub






Private Sub ListConsultants_GotFocus()
    BoxConsultant.BackColor = BttnBackColour ' vbRed
End Sub

Private Sub ListConsultants_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    
    Call FormatGridDates
    Call FormatGridPatients
    TemDoctorFee = 0
    TemFDoctorFee = 0
    TemInstitutionFee = 0
    TemFInstitutionFee = 0
    TemOtherFee = 0
'    TemDoctorID = 0
    TemAppointmentDate = Empty
    TemAppointmentTime = Empty
'    TwoSecessions = True
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
'    TemDoctorID = Val(ListConsultantIDs.Text)
    Call FillDates
    ListDatesAndSecessions.SetFocus
    KeyCode = Empty
ElseIf KeyCode = vbKeyLeft Then
    ListSpecialities.SetFocus
    KeyCode = Empty
ElseIf KeyCode = vbKeyUp Or vbKeyDown Then
    FormatGridDates
End If

End Sub


Private Sub ListConsultants_LostFocus()
    BoxConsultant.BackColor = FrameBackColour ' vbRed
End Sub



Private Sub FindAppointmentTime()
'    If TemUsualDuration = 0 Then Exit Sub
    If TemSecessionStartingTime = TimeSerial(0, 0, 0) Then Exit Sub
    TemAppointmentTime = TimeSerial(Hour(TemSecessionStartingTime), Minute(TemSecessionStartingTime) + (TemUsualDuration * TemNonCancelledVisits), 0)
End Sub


Private Sub ListDatesAndSecessions_GotFocus()
    BoxDates.BackColor = BttnBackColour ' vbRed
End Sub

Private Sub ListDatesAndSecessions_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    KeyCode = Empty
    ListPatientFacilities.SetFocus
ElseIf KeyCode = vbKeyRight Then
    ListPatientFacilities.SetFocus
    KeyCode = Empty
ElseIf KeyCode = vbKeyLeft Then
    ListConsultants.SetFocus
    KeyCode = Empty
End If
End Sub

Private Sub ListDatesAndSecessions_LostFocus()
    BoxDates.BackColor = FrameBackColour ' vbRed
End Sub



Private Sub ListPatientFacilities_GotFocus()
    BoxPatients.BackColor = BttnBackColour ' vbRed
End Sub


Private Sub ListPatientFacilities_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    bttnRefund.SetFocus
    KeyCode = Empty
ElseIf KeyCode = vbKeyLeft Then
    ListDatesAndSecessions.SetFocus
    KeyCode = Empty
Else

End If
End Sub

Private Sub ListPatientFacilities_LostFocus()
    BoxPatients.BackColor = FrameBackColour ' vbRed
End Sub

Private Sub MonthView2_GotFocus()
    BoxDates.BackColor = BttnBackColour ' vbRed
End Sub

Private Sub MonthView2_LostFocus()
    BoxDates.BackColor = FrameBackColour ' vbRed
End Sub

