VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmCreditAppointments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Appointments"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14100
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
   ScaleHeight     =   8340
   ScaleWidth      =   14100
   Begin btButtonEx.ButtonEx btnFill 
      Height          =   375
      Left            =   10560
      TabIndex        =   12
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Fill"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   12720
      TabIndex        =   9
      Top             =   7800
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
   Begin MSFlexGridLib.MSFlexGrid gridBookings 
      Height          =   5775
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   10186
      _Version        =   393216
   End
   Begin VB.OptionButton optStaffSettled 
      Caption         =   "Staff Booking Settled"
      Height          =   240
      Left            =   4200
      TabIndex        =   7
      Top             =   1560
      Width           =   3375
   End
   Begin VB.OptionButton optTelephoneSettled 
      Caption         =   "Telephone Booking Settled"
      Height          =   240
      Left            =   4200
      TabIndex        =   6
      Top             =   1200
      Width           =   3375
   End
   Begin VB.OptionButton optStaffBookingToSettle 
      Caption         =   "Staff Bookings To Settle"
      Height          =   240
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Width           =   3375
   End
   Begin VB.OptionButton optTelephoneBookingsToSettle 
      Caption         =   "Telephone Bookings To Settle"
      Height          =   240
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Value           =   -1  'True
      Width           =   3375
   End
   Begin MSDataListLib.DataCombo cmbStaff 
      Height          =   360
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   64225283
      CurrentDate     =   40445
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
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
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Excel"
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   6120
      TabIndex        =   14
      Top             =   240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   64225283
      CurrentDate     =   40445
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Doctor"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCreditAppointments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temTopic As String
    Dim temSubTopic As String
    Dim rsView As New ADODB.Recordset
    Dim temSQL  As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridBookings, temTopic, temSubTopic
End Sub

Private Sub btnFill_Click()
    Dim TemSelect As String
    Dim TemWhere As String
    Dim temFrom As String
    Dim temOrderBy As String
    Dim temSQL As String
    
    Dim P(0) As Integer
    Dim D(3) As Integer
    
    D(0) = 1
    D(1) = 2
    D(2) = 3
    
    'CONVERT(DATETIME, '" & MonthView1.Value & "', 102)
    
    If optStaffBookingToSettle.Value = True Then
        TemSelect = "SELECT tblPatientFacility.PatientFacility_ID as [Serial No],  tblPatientFacility.personalfee as [Doctor Fee], tblPatientFacility.institutionfee as [Hospital Fee], (tblPatientFacility.institutionfee +  tblPatientFacility.personalfee )as [Total Fee],  "
        If dtpFrom.Value <> dtpTo.Value Then
            TemSelect = TemSelect & " convert(datetime , tblPatientFacility.AppointmentDate , 102) as [Appointment Date] , "
        End If
        TemSelect = TemSelect & "    tblTitle.Title + ' ' + tblDoctor.DoctorName as Doctor, tblPatientFacility.BookingDate as [Booked Date],  "
        If IsNumeric(cmbStaff.BoundText) = False Then
            TemSelect = TemSelect & " tblBookedUser.StaffName as [Booked By],  "
        End If
        TemSelect = TemSelect & "  tblPatientMainDetails.FirstName as [Patient] , tblBookedForStaff.StaffName as [Booked For],  [FullyPaid] AS [Settled], tblCreditSettledUser.StaffName as [Settled User], tblPatientFacility.SettleCashDate as [Settled Date],  tblPatientFacility.SettleCashTime as [Settled Time], (tblPatientFacility.Cancelled)  as [Booking Cancelled] , (tblPatientFacility.Refund) as [Booking Refunded], tblRepayUser.StaffName as [Repaid User], (tblPatientFacility.RepayDate) as [Repaid Date]  , ( tblPatientFacility.RepayTime) as [Repaid Time] "
        temFrom = "FROM dbo.tblPatientMainDetails RIGHT OUTER JOIN dbo.tblPatientFacility ON dbo.tblPatientMainDetails.Patient_ID = dbo.tblPatientFacility.PatientID LEFT OUTER JOIN dbo.tblDoctor LEFT OUTER JOIN dbo.tblTitle ON dbo.tblDoctor.DoctorTitle_ID = dbo.tblTitle.Title_ID ON dbo.tblPatientFacility.Staff_ID = dbo.tblDoctor.Doctor_ID LEFT OUTER JOIN dbo.tblStaff tblBookedUser ON dbo.tblPatientFacility.User_ID = tblBookedUser.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblBookedForStaff ON dbo.tblPatientFacility.CreditStaff_ID = tblBookedForStaff.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblCreditSettledUser ON dbo.tblPatientFacility.CreditSettleUser_ID = tblCreditSettledUser.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblRepayUser ON dbo.tblPatientFacility.RepayUser_ID = tblRepayUser.Staff_ID "
        TemWhere = " Where tblPatientFacility.fullypaid = 0 and tblPatientFacility.AppointmentDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbStaff.BoundText) = True Then
            TemWhere = TemWhere & " AND tblPatientFacility.User_ID = " & Val(cmbStaff.BoundText)
        End If
        TemWhere = TemWhere & " AND tblPatientFacility.CreditStaff_ID <> 0 AND tblPatientFacility.CreditStaff_ID IS NOT NULL"
        TemWhere = TemWhere & " AND tblPatientFacility.PaymentMethod_ID = 4 "
        
        temOrderBy = "Order By tblPatientFacility.PatientFacility_ID"
    ElseIf optStaffSettled.Value = True Then
        TemSelect = "SELECT tblPatientFacility.PatientFacility_ID as [Serial No],  tblPatientFacility.personalfee as [Doctor Fee], tblPatientFacility.institutionfee as [Hospital Fee], (tblPatientFacility.institutionfee +  tblPatientFacility.personalfee )as [Total Fee],  "
        If dtpFrom.Value <> dtpTo.Value Then
            TemSelect = TemSelect & " (tblPatientFacility.AppointmentDate) as [Appointment Date] , "
        End If
        TemSelect = TemSelect & "   (tblPatientFacility.BookingDate) as [Booked Date], tblTitle.Title + ' ' + tblDoctor.DoctorName as Doctor,  "
        TemSelect = TemSelect & " tblBookedUser.StaffName as [Booked By],  "
        TemSelect = TemSelect & "  tblPatientMainDetails.FirstName as [Patient] , tblBookedForStaff.StaffName as [Booked For],  ([FullyPaid]) AS [Settled], "
        If IsNumeric(cmbStaff.BoundText) = False Then
            TemSelect = TemSelect & " tblCreditSettledUser.StaffName as [Settled User], "
        End If
        TemSelect = TemSelect & " (tblPatientFacility.SettleCashDate) as [Settled Date], ( tblPatientFacility.SettleCashTime ) as [Settled Time], (tblPatientFacility.Cancelled ) as [Booking Cancelled] , (tblPatientFacility.Refund) as [Booking Refunded], tblRepayUser.StaffName as [Repaid User], (tblPatientFacility.RepayDate) as [Repaid Date]  , ( tblPatientFacility.RepayTime) as [Repaid Time] "
        temFrom = "FROM dbo.tblPatientMainDetails RIGHT OUTER JOIN dbo.tblPatientFacility ON dbo.tblPatientMainDetails.Patient_ID = dbo.tblPatientFacility.PatientID LEFT OUTER JOIN dbo.tblDoctor LEFT OUTER JOIN dbo.tblTitle ON dbo.tblDoctor.DoctorTitle_ID = dbo.tblTitle.Title_ID ON dbo.tblPatientFacility.Staff_ID = dbo.tblDoctor.Doctor_ID LEFT OUTER JOIN dbo.tblStaff tblBookedUser ON dbo.tblPatientFacility.User_ID = tblBookedUser.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblBookedForStaff ON dbo.tblPatientFacility.CreditStaff_ID = tblBookedForStaff.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblCreditSettledUser ON dbo.tblPatientFacility.CreditSettleUser_ID = tblCreditSettledUser.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblRepayUser ON dbo.tblPatientFacility.RepayUser_ID = tblRepayUser.Staff_ID "
        TemWhere = " Where tblPatientFacility.fullypaid = 1 and tblPatientFacility.AppointmentDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbStaff.BoundText) = True Then
            TemWhere = TemWhere & " AND tblPatientFacility.CreditSettleUser_ID = " & Val(cmbStaff.BoundText)
        End If
        TemWhere = TemWhere & " AND tblPatientFacility.CreditStaff_ID <> 0 AND tblPatientFacility.CreditStaff_ID IS NOT NULL"
        TemWhere = TemWhere & " AND tblPatientFacility.PaymentMethod_ID = 4 "
        
        temOrderBy = "Order By tblPatientFacility.PatientFacility_ID"
    
    ElseIf optTelephoneBookingsToSettle.Value = True Then
        TemSelect = "SELECT tblPatientFacility.PatientFacility_ID as [Serial No],  tblPatientFacility.personalfee as [Doctor Fee], tblPatientFacility.institutionfee as [Hospital Fee], (tblPatientFacility.institutionfee +  tblPatientFacility.personalfee )as [Total Fee],  "
        If dtpFrom.Value <> dtpTo.Value Then
            TemSelect = TemSelect & " (tblPatientFacility.AppointmentDate) as [Appointment Date] , "
        End If
        TemSelect = TemSelect & "   (tblPatientFacility.BookingDate) as [Booked Date], tblTitle.Title + ' ' + tblDoctor.DoctorName as Doctor,  "
        If IsNumeric(cmbStaff.BoundText) = False Then
            TemSelect = TemSelect & " tblBookedUser.StaffName as [Booked By],  "
        End If
        TemSelect = TemSelect & " tblPatientMainDetails.FirstName as [Patient] ,  ([FullyPaid]) AS [Settled], tblCreditSettledUser.StaffName as [Settled User], (tblPatientFacility.SettleCashDate) as [Settled Date], ( tblPatientFacility.SettleCashTime ) as [Settled Time], (tblPatientFacility.Cancelled ) as [Booking Cancelled] , (tblPatientFacility.Refund) as [Booking Refunded], tblRepayUser.StaffName as [Repaid User], (tblPatientFacility.RepayDate) as [Repaid Date]  , ( tblPatientFacility.RepayTime) as [Repaid Time] "
        temFrom = "FROM dbo.tblPatientMainDetails RIGHT OUTER JOIN dbo.tblPatientFacility ON dbo.tblPatientMainDetails.Patient_ID = dbo.tblPatientFacility.PatientID LEFT OUTER JOIN dbo.tblDoctor LEFT OUTER JOIN dbo.tblTitle ON dbo.tblDoctor.DoctorTitle_ID = dbo.tblTitle.Title_ID ON dbo.tblPatientFacility.Staff_ID = dbo.tblDoctor.Doctor_ID LEFT OUTER JOIN dbo.tblStaff tblBookedUser ON dbo.tblPatientFacility.User_ID = tblBookedUser.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblBookedForStaff ON dbo.tblPatientFacility.CreditStaff_ID = tblBookedForStaff.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblCreditSettledUser ON dbo.tblPatientFacility.CreditSettleUser_ID = tblCreditSettledUser.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblRepayUser ON dbo.tblPatientFacility.RepayUser_ID = tblRepayUser.Staff_ID "
        TemWhere = " Where tblPatientFacility.fullypaid = 0 and tblPatientFacility.AppointmentDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbStaff.BoundText) = True Then
            TemWhere = TemWhere & " AND tblPatientFacility.User_ID = " & Val(cmbStaff.BoundText)
        End If
        TemWhere = TemWhere & " AND (tblPatientFacility.CreditStaff_ID = 0 OR tblPatientFacility.CreditStaff_ID IS NULL) "
        TemWhere = TemWhere & " AND tblPatientFacility.PaymentMethod_ID = 4 "
        
        temOrderBy = "Order By tblPatientFacility.PatientFacility_ID"
    
    ElseIf optTelephoneSettled.Value = True Then
    
    
    
        TemSelect = "SELECT tblPatientFacility.PatientFacility_ID as [Serial No],  tblPatientFacility.personalfee as [Doctor Fee], tblPatientFacility.institutionfee as [Hospital Fee], (tblPatientFacility.institutionfee +  tblPatientFacility.personalfee )as [Total Fee],  "
        If dtpFrom.Value <> dtpTo.Value Then
            TemSelect = TemSelect & " (tblPatientFacility.AppointmentDate) as [Appointment Date] , "
        End If
        TemSelect = TemSelect & "   (tblPatientFacility.BookingDate) as [Booked Date], tblTitle.Title + ' ' + tblDoctor.DoctorName as Doctor,  "
        TemSelect = TemSelect & " tblBookedUser.StaffName as [Booked By],  "
        TemSelect = TemSelect & "  tblPatientMainDetails.FirstName as [Patient],  ([FullyPaid]) AS [Settled], "
        If IsNumeric(cmbStaff.BoundText) = False Then
            TemSelect = TemSelect & " tblCreditSettledUser.StaffName as [Settled User], "
        End If
        TemSelect = TemSelect & " (tblPatientFacility.SettleCashDate) as [Settled Date], ( tblPatientFacility.SettleCashTime ) as [Settled Time], (tblPatientFacility.Cancelled ) as [Booking Cancelled] , (tblPatientFacility.Refund) as [Booking Refunded], tblRepayUser.StaffName as [Repaid User], (tblPatientFacility.RepayDate) as [Repaid Date]  , ( tblPatientFacility.RepayTime) as [Repaid Time] "
        temFrom = "FROM dbo.tblPatientMainDetails RIGHT OUTER JOIN dbo.tblPatientFacility ON dbo.tblPatientMainDetails.Patient_ID = dbo.tblPatientFacility.PatientID LEFT OUTER JOIN dbo.tblDoctor LEFT OUTER JOIN dbo.tblTitle ON dbo.tblDoctor.DoctorTitle_ID = dbo.tblTitle.Title_ID ON dbo.tblPatientFacility.Staff_ID = dbo.tblDoctor.Doctor_ID LEFT OUTER JOIN dbo.tblStaff tblBookedUser ON dbo.tblPatientFacility.User_ID = tblBookedUser.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblBookedForStaff ON dbo.tblPatientFacility.CreditStaff_ID = tblBookedForStaff.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblCreditSettledUser ON dbo.tblPatientFacility.CreditSettleUser_ID = tblCreditSettledUser.Staff_ID LEFT OUTER JOIN dbo.tblStaff tblRepayUser ON dbo.tblPatientFacility.RepayUser_ID = tblRepayUser.Staff_ID "
        TemWhere = " Where tblPatientFacility.fullypaid = 1 and tblPatientFacility.AppointmentDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbStaff.BoundText) = True Then
            TemWhere = TemWhere & " AND tblPatientFacility.CreditSettleUser_ID = " & Val(cmbStaff.BoundText)
        End If
        TemWhere = TemWhere & " AND (tblPatientFacility.CreditStaff_ID = 0 OR tblPatientFacility.CreditStaff_ID IS NULL) "
        TemWhere = TemWhere & " AND tblPatientFacility.PaymentMethod_ID = 4 "
        
        temOrderBy = "Order By tblPatientFacility.PatientFacility_ID"
    
    
    End If
    
    temSQL = TemSelect & " " & temFrom & " " & TemWhere & " " & temOrderBy
    
    
    
    FillTotalGrid temSQL, gridBookings, 0, D, P
    
    If IsNumeric(cmbStaff.BoundText) = True Then
        temSubTopic = " User - " & cmbStaff.Text
    End If
    If dtpFrom.Value <> dtpTo.Value Then
        temSubTopic = temSubTopic & " From - " & Format(dtpFrom.Value, "dd MMMM yyyy") & " To - " & Format(dtpTo.Value, "dd MMMM yyyy")
    Else
        temSubTopic = temSubTopic & " On - " & Format(dtpFrom.Value, "dd MMMM yyyy")
    End If
    
    If optStaffBookingToSettle.Value = True Then
        temTopic = "Staff Bookings To Settle"
    ElseIf optStaffSettled.Value = True Then
        temTopic = "Settling Staff Bookings Settled"
    ElseIf optTelephoneBookingsToSettle.Value = True Then
        temTopic = "Telephone Bookings To Settle"
    ElseIf optTelephoneSettled.Value = True Then
        temTopic = "Settling Telephone Bookings Settled"
    End If
    
        gridBookings.TextMatrix(gridBookings.Rows - 1, 0) = gridBookings.Rows - 2

    
End Sub

Private Sub btnPrint_Click()
    Dim myPR As PrintReport
    GetPrintDefaults myPR
    GridPrint gridBookings, myPR, temTopic, temSubTopic
End Sub

Private Sub Form_Load()
    Call getSettings
    Call FillCOmbos
End Sub

Private Sub cmbStaff_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbStaff.Text = Empty
    End If
End Sub


Private Sub FillCOmbos()
    With rsView
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblDoctor order by DoctorName"
        .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
    End With
    With cmbStaff
        Set .RowSource = rsView
        .ListField = "DoctorName"
        .BoundColumn = "Doctor_ID"
    End With
End Sub

Private Sub saveSettings()
    SaveCommonSettings Me
End Sub

Private Sub getSettings()
    GetCommonSettings Me
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call saveSettings
End Sub
