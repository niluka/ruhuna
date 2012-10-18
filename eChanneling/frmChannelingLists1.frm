VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmChannelingLists 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Channeling Lists"
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
   Icon            =   "frmChannelingLists1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin btButtonEx.ButtonEx bttnCloseList 
      Height          =   375
      Left            =   13680
      TabIndex        =   0
      Top             =   9000
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
   Begin VB.Frame FramePatientList 
      Height          =   8895
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   15015
      Begin VB.ListBox ListSpecialityIDs 
         Height          =   1980
         Left            =   2760
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox ListConsultantIDs 
         Height          =   1980
         Left            =   5880
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox ListSecessionIDs 
         Height          =   1980
         Left            =   14160
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox ListSpecialities 
         Height          =   2820
         IntegralHeight  =   0   'False
         ItemData        =   "frmChannelingLists1.frx":0442
         Left            =   120
         List            =   "frmChannelingLists1.frx":0444
         TabIndex        =   18
         Top             =   240
         Width           =   3375
      End
      Begin VB.ListBox ListConsultants 
         Height          =   2820
         IntegralHeight  =   0   'False
         ItemData        =   "frmChannelingLists1.frx":0446
         Left            =   3600
         List            =   "frmChannelingLists1.frx":0448
         TabIndex        =   17
         Top             =   240
         Width           =   4455
      End
      Begin VB.ListBox ListDatesAndSecessions 
         Height          =   2820
         IntegralHeight  =   0   'False
         ItemData        =   "frmChannelingLists1.frx":044A
         Left            =   11280
         List            =   "frmChannelingLists1.frx":044C
         TabIndex        =   16
         Top             =   240
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         Caption         =   "List Criteria"
         Height          =   1095
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   14895
         Begin VB.OptionButton OptionAgentBookings 
            Caption         =   "Agent Bookings"
            Height          =   255
            Left            =   7560
            TabIndex        =   8
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton OptionCancellationsAndRefunds 
            Caption         =   "Cancellations And Refunds"
            Height          =   255
            Left            =   4440
            TabIndex        =   13
            Top             =   600
            Width           =   2775
         End
         Begin VB.OptionButton OptionRefunds 
            Caption         =   "Refunds"
            Height          =   255
            Left            =   2280
            TabIndex        =   12
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton OptionCancellations 
            Caption         =   "Cancellations"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton OptionCashBookings 
            Caption         =   "Cash Bookings"
            Height          =   255
            Left            =   10680
            TabIndex        =   9
            Top             =   600
            Width           =   2175
         End
         Begin VB.OptionButton OptionWithoutCancellationsAndRefunds 
            Caption         =   "Without Cancellations And Refunds"
            Height          =   255
            Left            =   10680
            TabIndex        =   7
            Top             =   240
            Width           =   3375
         End
         Begin VB.OptionButton OptionWithoutCancellation 
            Caption         =   "Without Cancellations"
            Height          =   255
            Left            =   4440
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton OptionFullyPaid 
            Caption         =   "Fully Paid Only"
            Height          =   255
            Left            =   2280
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton OptionWithoutRefunds 
            Caption         =   "Without Refunds"
            Height          =   255
            Left            =   7560
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton OptionAll 
            Caption         =   "List All"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridList1 
         Height          =   4455
         Left            =   120
         TabIndex        =   15
         Top             =   4320
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7858
         _Version        =   393216
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2820
         Left            =   8160
         TabIndex        =   20
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   65798145
         CurrentDate     =   39472
      End
   End
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   375
      Left            =   12000
      TabIndex        =   10
      Top             =   9000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Print"
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
Attribute VB_Name = "frmChannelingLists"
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
    

Private Sub Form_Load()
    Call Setcolours
    FillSpeciality
    FormatPatientFacilityList
    MonthView1.Value = Date
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
    
    
Private Sub ListDatesAndSecessions_Click()
    ListSecessionIDs.ListIndex = ListDatesAndSecessions.ListIndex
    Call FillPatientFacilityList
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
            ListConsultantIDs.AddItem !Doctor_ID
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub ListConsultants_Click()
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    Call FormatPatientFacilityList
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
            .Source = "Select * from tblfacilitysecession where hospitalfacility_ID =  10  and staff_ID = " & Val(ListConsultantIDs.Text) & " and AlteredDate = '" & TemBookingDate & "' order by StartingTime"
            .Open
            If .RecordCount <> 0 Then
                If !fulldayleave = 0 Then
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


Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    
    Call FormatPatientFacilityList
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

  
Private Sub bttnPrint_Click()
Dim TemResponce As Long

If ListConsultants.ListIndex < 0 Or IsNumeric(ListConsultantIDs.Text) = False Then
    TemResponce = MsgBox("You have not selected a consultant", vbCritical, "No COnsultant")
    ListConsultants.SetFocus
    Exit Sub
End If

If ListDatesAndSecessions.ListIndex < 0 Or (IsNumeric(ListSecessionIDs.Text) = False And ListSecessionIDs.Text <> "All") Then
    TemResponce = MsgBox("You have not selected a Date and secession", vbCritical, "No Date & Secession")
    ListDatesAndSecessions.SetFocus
    Exit Sub
End If

    With DataEnvironment1.rssqlDoctorView
        If ListSecessionIDs.Text = "All" Then
            If .State = 1 Then .Close
            If OptionAll.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where (hospitalFacility_ID = 10) and  (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "')  order by dayserial"
            ElseIf OptionFullyPaid.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "')   and fullypaid = 1 order by dayserial"
            ElseIf OptionWithoutCancellation.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "')  and cancelled = 0 order by dayserial"
            ElseIf OptionWithoutRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and refund = 0 order by dayserial"
            ElseIf OptionWithoutCancellationsAndRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "')  and (cancelled = 0 and refund = 0 ) order by dayserial"
            ElseIf OptionAgentBookings.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and  (PaymentMode = 'Agent')  order by dayserial"
            ElseIf OptionCashBookings.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and  (PaymentMode = 'Cash') order by dayserial"
            ElseIf OptionCancellations.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and  (cancelled = 1) order by dayserial"
            ElseIf OptionRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and  (refund = 1) order by dayserial"
            ElseIf OptionCancellationsAndRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (secession = " & Val(ListSecessionIDs.Text) & ") and (cancelled = 1 or refund = 1 ) order by dayserial"
            End If
        Else
            If .State = 1 Then .Close
                '.Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where staff_ID = " & Val(ListConsultantIDs.Text) & " and appointmentdate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " order by dayserial"
            If OptionAll.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where (hospitalFacility_ID = 10) and  (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") order by dayserial"
            ElseIf OptionFullyPaid.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ")  and fullypaid = 1 order by dayserial"
            ElseIf OptionWithoutCancellation.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and cancelled = 0 order by dayserial"
            ElseIf OptionWithoutRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and refund = 0 order by dayserial"
            ElseIf OptionWithoutCancellationsAndRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and (cancelled = 0 and refund = 0 ) order by dayserial"
            ElseIf OptionAgentBookings.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and (PaymentMode = 'Agent')  order by dayserial"
            ElseIf OptionCashBookings.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and (PaymentMode = 'Cash') order by dayserial"
            ElseIf OptionCancellations.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and (cancelled = 1) order by dayserial"
            ElseIf OptionRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and (refund = 1) order by dayserial"
            ElseIf OptionCancellationsAndRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (secession = " & Val(ListSecessionIDs.Text) & ") and (cancelled = 1 or refund = 1 ) order by dayserial"
            End If
        End If
        
        .Open
        
    End With
    
    With DataReportDoctorView
    Set .DataSource = DataEnvironment1.rssqlDoctorView
      If HospitalDetails = True Then
            .Sections.Item("ReportHeader10").Controls.Item("RptName").Caption = InstitutionName
            .Sections.Item("ReportHeader10").Controls.Item("RptAddress").Caption = InstitutionAddress
'            .Sections.Item("ReportHeader10").Controls.Item("lblinstitutiontelephone").Caption = InstitutionTelephone
            .Sections.Item("Section2").Controls.Item("lbldoctorname").Caption = "Consultant : " & FindLDoctorFromID(Val(ListConsultantIDs.Text))
            .Sections.Item("section3").Controls.Item("lblAds").Caption = LongAd
        Else
            .Sections.Item("ReportHeader10").Controls.Item("RptName").Caption = Empty
            .Sections.Item("ReportHeader10").Controls.Item("RptAddress").Caption = Empty
'            .Sections.Item("ReportHeader10").Controls.Item("lblinstitutiontelephone").Caption = InstitutionTelephone
            .Sections.Item("Section2").Controls.Item("lbldoctorname").Caption = "Consultant : " & FindLDoctorFromID(Val(ListConsultantIDs.Text))
            .Sections.Item("section3").Controls.Item("lblAds").Caption = LongAd
        End If

        If ListSecessionIDs.Text = "All" Then
            .Sections.Item("Section2").Controls.Item("lbldatesecession").Caption = "Date : " & Format(MonthView1.Value, DefaultLongDate)
        Else
            .Sections.Item("Section2").Controls.Item("lbldatesecession").Caption = "Date : " & Format(MonthView1.Value, DefaultLongDate) & "   Secession : " & FindSecessionFromID(Val(ListSecessionIDs.Text))
        End If
        .Sections.Item("Section5").Controls.Item("lblad1").Caption = LongAd
        .Show
    End With

End Sub

Private Sub Form_Activate()
If SetPrinter = False Then Unload Me: Exit Sub
If UserAuthority = AuthorityOwner Then
    MonthView1.Enabled = True
ElseIf UserAuthority = AuthorityAnalyzer Then
    MonthView1.Enabled = True
Else
    MonthView1.Enabled = False
End If

End Sub


Private Function SetPrinter() As Boolean
SetPrinter = False
Dim MyPrinter As Printer

For Each MyPrinter In Printers
    If MyPrinter.DeviceName = ReportPrinterName Then
        Set Printer = MyPrinter
        SetPrinter = True
    End If
Next

If SetPrinter = False Then
        Dim TemResponce  As Integer
        TemResponce = MsgBox("You have not selected a valied printer for bill printing, Please select a printer", vbCritical, "No printer")
        frmPrintingPreferances.Show
        frmPrintingPreferances.ZOrder 0
        frmPrintingPreferances.SSTab1.Tab = 1
        frmPrintingPreferances.ComboBillPrinter.SetFocus
End If


End Function


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
        .Source = "Select * from tblfacilitysecession where hospitalfacility_ID =  10  and staff_ID = " & ListConsultantIDs.Text & " and AlteredDate = '" & MonthView1.Value & "' order by StartingTime"
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
                .Source = "Select * from tblfacilitysecession where hospitalfacility_ID = 10 and staff_ID = " & ListConsultantIDs.Text & " and SecessionWeekday = " & Weekday(MonthView1.Value) & " order by StartingTime"
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

Public Sub FormatPatientFacilityList()
    Dim BorderMargin As Long
    BorderMargin = 150
    With GridList1
        .Clear
        .Rows = 1
        .Row = 0
        .Cols = 16
        
        .ColWidth(0) = 600
        .Col = 0
        .CellAlignment = 4
        .Text = "Serial"
        
        .ColWidth(1) = 3000
        .Col = 1
        .CellAlignment = 4
        .Text = "Patient Name"
    
        .ColWidth(2) = 1100
        .Col = 2
        .CellAlignment = 4
        .Text = "Fully Paid"
    
        .ColWidth(3) = 1100
        .Col = 3
        .CellAlignment = 4
        .Text = "Refunds"
        
        .ColWidth(4) = 1100
        .Col = 4
        .CellAlignment = 4
        .Text = "Doc. Fee"
        
        .ColWidth(5) = 1100
        .Col = 5
        .CellAlignment = 4
        .Text = "Hos. Fee"

        .ColWidth(6) = 1100
        .Col = 6
        .CellAlignment = 4
        .Text = "Other Fee"

        .ColWidth(7) = 1100
        .Col = 7
        .CellAlignment = 4
        .Text = "Doc. Refund"
        
        .ColWidth(8) = 1100
        .Col = 8
        .CellAlignment = 4
        .Text = "Hos. Refund"

        .ColWidth(9) = 1100
        .Col = 9
        .CellAlignment = 4
        .Text = "Other Refund"

        .ColWidth(10) = 1100
        .Col = 10
        .CellAlignment = 4
        .Text = "Paid to Doctor"

        .ColWidth(11) = 1100
        .Col = 11
        .CellAlignment = 4
        .Text = "Booking ID"
    
        .ColWidth(12) = 1
        .Col = 12
        .Text = "Patient_ID"
        
        .ColWidth(13) = 1
        .Col = 13
        .Text = "HospitalFacility_ID"
        
        .ColWidth(14) = 1
        .Col = 14
        .Text = "Staff_ID"
        
        .ColWidth(15) = 1
        .Col = 15
        .Text = "PatientBill_ID"
    End With
End Sub

Public Sub FillPatientFacilityList()
    Dim TemResponce As Integer
    
    If ListConsultants.ListIndex < 0 Or IsNumeric(ListConsultantIDs.Text) = False Then
        TemResponce = MsgBox("You have not selected a consultant", vbCritical, "No COnsultant")
        ListConsultants.SetFocus
        Exit Sub
    End If
    
    If ListDatesAndSecessions.ListIndex < 0 Or (IsNumeric(ListSecessionIDs.Text) = False And ListSecessionIDs.Text <> "All") Then
        TemResponce = MsgBox("You have not selected a Date and secession", vbCritical, "No Date & Secession")
        ListDatesAndSecessions.SetFocus
        Exit Sub
    End If
    
    FormatPatientFacilityList
    
    GridList1.Visible = False
    Me.MousePointer = vbHourglass
    
    Dim NowROw As Long
    Dim TemNum As Long
    With DataEnvironment1.rssqlDoctorView
        If ListSecessionIDs.Text = "All" Then
            If .State = 1 Then .Close
            If OptionAll.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where (hospitalFacility_ID = 10) and  (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "')  order by dayserial"
            ElseIf OptionFullyPaid.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "')   and fullypaid = 1 order by dayserial"
            ElseIf OptionWithoutCancellation.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "')  and cancelled = 0 order by dayserial"
            ElseIf OptionWithoutRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "')  and refund = 0 order by dayserial"
            ElseIf OptionWithoutCancellationsAndRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "')  and (cancelled = 1 or refund = 1 ) order by dayserial"
            ElseIf OptionAgentBookings.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and  (PaymentMode = 'Agent')  order by dayserial"
            ElseIf OptionCashBookings.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and  (PaymentMode = 'Cash') order by dayserial"
            ElseIf OptionCancellations.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and  (cancelled = 1) order by dayserial"
            ElseIf OptionRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and  (refund = 1) order by dayserial"
            ElseIf OptionCancellationsAndRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (secession = " & Val(ListSecessionIDs.Text) & ") and (cancelled = 1 or refund = 1 ) order by dayserial"
            End If
        Else
            If .State = 1 Then .Close
                '.Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where staff_ID = " & Val(ListConsultantIDs.Text) & " and appointmentdate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " order by dayserial"
            If OptionAll.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where (hospitalFacility_ID = 10) and  (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") order by dayserial"
            ElseIf OptionFullyPaid.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ")  and fullypaid = 1 order by dayserial"
            ElseIf OptionWithoutCancellation.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and cancelled = 0 order by dayserial"
            ElseIf OptionWithoutRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and refund = 0 order by dayserial"
            ElseIf OptionWithoutCancellationsAndRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and (cancelled = 0 or refund = 0 ) order by dayserial"
            ElseIf OptionAgentBookings.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and (PaymentMode = 'Agent')  order by dayserial"
            ElseIf OptionCashBookings.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and (PaymentMode = 'Cash') order by dayserial"
            ElseIf OptionCancellations.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and (cancelled = 1) order by dayserial"
            ElseIf OptionRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (hospitalFacility_ID = 10) and (staff_ID = " & ListConsultantIDs.Text & ") and (appointmentdate = '" & MonthView1.Value & "') and (secession = " & Val(ListSecessionIDs.Text) & ") and (refund = 1) order by dayserial"
            ElseIf OptionCancellationsAndRefunds.Value = True Then
                .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where  (secession = " & Val(ListSecessionIDs.Text) & ") and (cancelled = 1 or refund = 1 ) order by dayserial"
            End If
        End If
        
        .Open
     
        
        
        If .State = 0 Then .Open
        
        If .RecordCount = 0 Then
            GridList1.Visible = True
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        .MoveFirst
        GridList1.Rows = .RecordCount + 1
        
        NowROw = 0
        
        While Not .EOF
            NowROw = NowROw + 1
            
            GridList1.Row = NowROw
            
            GridList1.Col = 0
            GridList1.CellAlignment = 7
            GridList1.Text = !DaySerial
            
            GridList1.Col = 1
            GridList1.CellAlignment = 1
            GridList1.Text = FindPatientByID(Val(!patientid))
        
            GridList1.Col = 2
            GridList1.CellAlignment = 7
            If !FullyPaid = 1 Then
                GridList1.Text = "Yes"
            Else
                GridList1.Text = "No"
            End If
        
            GridList1.Col = 3
            GridList1.CellAlignment = 7
            If !cancelled = True Then
                GridList1.Text = "Cancelled"
            ElseIf !Refund = True Then
                GridList1.Text = "Repaied"
            End If
            
            
            GridList1.Col = 4
            GridList1.CellAlignment = 7
            GridList1.Text = Format(!personalfee, "0.00")
            
            GridList1.Col = 5
            GridList1.CellAlignment = 7
            GridList1.Text = Format(!institutionfee, "0.00")
    
            GridList1.Col = 6
            GridList1.CellAlignment = 7
            GridList1.Text = Format(!otherfee, "0.00")
    
            GridList1.Col = 7
            GridList1.CellAlignment = 7
            If Not IsNull(!Personalrefund) Then
                GridList1.Text = Format(!Personalrefund, "0.00")
            Else
                GridList1.Text = Empty
            End If
            
            GridList1.Col = 8
            GridList1.CellAlignment = 7
            If Not IsNull(!institutionrefund) Then
                GridList1.Text = Format(!institutionrefund, "0.00")
            Else
                GridList1.Text = Empty
            End If
    
            GridList1.Col = 9
            GridList1.CellAlignment = 7
            
            If Not IsNull(!otherrefund) Then
                GridList1.Text = Format(!otherrefund, "0.00")
            Else
                GridList1.Text = Empty
            End If
            GridList1.Col = 10
            GridList1.CellAlignment = 7
            If !paidtostaff = True Then
                GridList1.Text = "Yes"
            Else
                GridList1.Text = "No"
            End If
    
            GridList1.Col = 11
            GridList1.CellAlignment = 7
            GridList1.Text = !patientfacility_ID
            
            GridList1.Col = 12
            GridList1.Text = !patientid
            GridList1.Col = 13
            GridList1.Text = !HospitalFacility_ID
            GridList1.Col = 14
            GridList1.Text = !FacilityStaff_ID
            GridList1.Col = 15
            GridList1.Text = !PatientBill_ID
            
            
            .MoveNext
        Wend
        .Close
    End With
    
    With GridList1
    
    Dim TemDoctorTotalFee As Double
    Dim TemHospitalTotalFee As Double
    Dim TemOtherTotalFee As Double
    Dim TemDoctorTotalRepayment As Double
    Dim TemHospitalTotalRepayment As Double
    Dim TemOtherTotalRepayment As Double
    
    TemDoctorTotalFee = 0
    TemHospitalTotalFee = 0
    TemOtherTotalFee = 0
    TemDoctorTotalRepayment = 0
    TemHospitalTotalRepayment = 0
    TemOtherTotalRepayment = 0
    
    
    For TemNum = 1 To GridList1.Rows - 1
        .Col = 0
        .Row = TemNum
        .Text = TemNum
        
        .Col = 4
        TemDoctorTotalFee = TemDoctorTotalFee + Val(.Text)
        
        .Col = 5
        TemHospitalTotalFee = TemHospitalTotalFee + Val(.Text)

        .Col = 6
        TemOtherTotalFee = TemOtherTotalFee + Val(.Text)

        .Col = 7
        TemDoctorTotalRepayment = TemDoctorTotalRepayment + Val(.Text)
        
        .Col = 8
        TemHospitalTotalRepayment = TemHospitalTotalRepayment + Val(.Text)

        .Col = 9
        TemOtherTotalRepayment = TemOtherTotalRepayment + Val(.Text)

    Next
    
    .Rows = .Rows + 1
    .Row = .Rows - 1
    
        .Col = 4
        .CellAlignment = 7
        .Text = Format(TemDoctorTotalFee, "0.00")
        
        .Col = 5
        .CellAlignment = 7
        .Text = Format(TemHospitalTotalFee, "0.00")

        .Col = 6
        .CellAlignment = 7
        .Text = Format(TemOtherTotalFee, "0.00")

        .Col = 7
        .CellAlignment = 7
        .Text = Format(TemDoctorTotalRepayment, "0.00")
        
        .Col = 8
        .CellAlignment = 7
        .Text = Format(TemHospitalTotalRepayment, "0.00")

        .Col = 9
        .CellAlignment = 7
        .Text = Format(TemOtherTotalRepayment, "0.00")

    
    
    
    .Row = 0
    .Col = 0
    
    
    End With
    
    TemPreviousDoctorID = Val(ListConsultantIDs.Text)
    TemPreviousSecession = Val(ListSecessionIDs.Text)
    TemPreviousDate = MonthView1.Value
    TemPreviousOptionChanged = False
    
    GridList1.Visible = True
    Me.MousePointer = vbDefault
    
End Sub

Private Sub bttnCloseList_Click()
    Unload Me
End Sub

Private Sub OptionAgentBookings_Click()
    If OptionAgentBookings.Value = True Then
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionAll_Click()
    If OptionAll.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionCancellations_Click()
    If OptionCancellations.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionCancellationsAndRefunds_Click()
    If OptionCancellationsAndRefunds.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionCashBookings_Click()
    If OptionCashBookings.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionFullyPaid_Click()
    If OptionFullyPaid.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub


Private Sub Setcolours()
    bttnCloseList.BackColor = BttnBackColour
    bttnCloseList.ForeColor = BttnForeColour
    bttnPrint.BackColor = BttnBackColour
    bttnPrint.ForeColor = BttnForeColour
    FramePatientList.BackColor = FrmBackColour
    FramePatientList.ForeColor = FrmForeColour
    Me.BackColor = FrameBackColour
    Me.ForeColor = FrameForeColour
'    DataComboDoctorStaff.BackColor = FrmBackColour
'    DataComboDoctorStaff.ForeColor = FrmForeColour
    OptionAgentBookings.BackColor = FrmBackColour
    OptionAgentBookings.ForeColor = FrmForeColour
    OptionAll.BackColor = FrmBackColour
    OptionAll.ForeColor = FrmForeColour
    OptionCancellations.BackColor = FrmBackColour
    OptionCancellations.ForeColor = FrmForeColour
    OptionCancellationsAndRefunds.BackColor = FrmBackColour
    OptionCancellationsAndRefunds.ForeColor = FrmForeColour
    OptionCashBookings.BackColor = FrmBackColour
    OptionCashBookings.ForeColor = FrmForeColour
    OptionFullyPaid.BackColor = FrmBackColour
    OptionFullyPaid.ForeColor = FrmForeColour
    OptionWithoutCancellation.BackColor = FrmBackColour
    OptionWithoutCancellation.ForeColor = FrmForeColour
    OptionWithoutCancellationsAndRefunds.BackColor = FrmBackColour
    OptionWithoutCancellationsAndRefunds.ForeColor = FrmForeColour
    OptionWithoutRefunds.BackColor = FrmBackColour
    OptionWithoutRefunds.ForeColor = FrmForeColour
    OptionRefunds.BackColor = FrmBackColour
    OptionRefunds.ForeColor = FrmForeColour
    Frame1.BackColor = FrmBackColour
    Frame1.ForeColor = FrmForeColour
End Sub

Private Sub OptionRefunds_Click()
    If OptionRefunds.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionWithoutCancellation_Click()
    If OptionWithoutCancellation.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionWithoutCancellationsAndRefunds_Click()
    If OptionWithoutCancellationsAndRefunds.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionWithoutRefunds_Click()
    If OptionWithoutRefunds.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub
