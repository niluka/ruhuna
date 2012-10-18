VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmCreditBookings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Bookings"
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
   Begin VB.OptionButton optStaffSettling 
      Caption         =   "Staff Booking Settling"
      Height          =   240
      Left            =   4200
      TabIndex        =   7
      Top             =   1560
      Width           =   3375
   End
   Begin VB.OptionButton optTelephoneSettling 
      Caption         =   "Telephone Booking Settling"
      Height          =   240
      Left            =   4200
      TabIndex        =   6
      Top             =   1200
      Width           =   3375
   End
   Begin VB.OptionButton optStaffBooking 
      Caption         =   "Staff Bookings"
      Height          =   240
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Width           =   3375
   End
   Begin VB.OptionButton optTelephoneBookings 
      Caption         =   "Telephone Bookings"
      Height          =   240
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
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
      Format          =   151781379
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
      Format          =   151781379
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
      Caption         =   "User"
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
Attribute VB_Name = "frmCreditBookings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temTopic As String
    Dim temSubTopic As String
    

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridBookings, temTopic, temSubTopic
End Sub

Private Sub btnFill_Click()
    Dim temSelect As String
    Dim temWhere As String
    Dim temFrom As String
    Dim temOrderBy As String
    Dim temSQL As String
    
    If optStaffBooking.Value = True Then
        temSelect = "SELECT tblPatientFacility.PatientFacility_ID as [Serial No], "
        If dtpFrom.Value <> dtpTo.Value Then
            temSelect = temSelect & " format(tblPatientFacility.BookingDate, 'dd MMM yyyy') as [Booked Date] , "
        End If
        temSelect = temSelect & "   format(tblPatientFacility.BookingTime, 'hh:mm AMPM') as [Booked Time], tblTitle.Title + ' ' + tblDoctor.DoctorName as Doctor, tblPatientFacility.AppointmentDate as [Appointment Date],  "
        If IsNumeric(cmbStaff.BoundText) = True Then
            temSelect = temSelect & " tblBookedUser.StaffName as [Booked By],  "
        End If
        temsqlect = temSelect & "  tblPatientMainDetails.FirstName as [Patient] , tblBookedForStaff.StaffName as [Booked For],  Format([FullyPaid],'Yes/No') AS [Settled], tblCreditSettledUser.StaffName as [Settled User], format(tblPatientFacility.SettleCashDate, 'dd MMM yyyy') as [Settled Date], format( tblPatientFacility.SettleCashTime , 'mm:hh AMPM') as [Settled Time], format(tblPatientFacility.Cancelled , 'Yes/No') as [Booking Cancelled] , format(tblPatientFacility.Refund, 'Yes/No') as [Booking Refunded], tblRepayUser.StaffName as [Repaid User], format(tblPatientFacility.RepayDate, 'dd MMM yyyy') as [Repaid Date]  , format( tblPatientFacility.RepayTime, 'mm:hh AMPM') as [Repaid Time] "
        temFrom = " FROM (((((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) LEFT JOIN tblStaff AS tblBookedUser ON tblPatientFacility.User_ID = tblBookedUser.Staff_ID) LEFT JOIN tblStaff AS tblBookedForStaff ON tblPatientFacility.CreditStaff_ID = tblBookedForStaff.Staff_ID) INNER JOIN tblStaff AS tblCreditSettledUser ON tblPatientFacility.CreditSettleUser_ID = tblCreditSettledUser.Staff_ID) LEFT JOIN tblStaff AS tblRepayUser ON tblPatientFacility.RepayUser_ID = tblRepayUser.Staff_ID "
        temWhere = " Where tblPatientFacility.BookingDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbStaff.BoundText) = True Then
            temWhere = temWhere & " AND tblPatientFacility.User_ID = " & Val(cmbStaff.BoundText)
        End If
        temWhere = temWhere & " AND tblPatientFacility.CreditStaff_ID <> 0 AND tblPatientFacility.CreditStaff_ID NOT NULL"
        temWhere = temWhere & " AND tblPatientFacility.PaymentMethod_ID = 4 "
        
        temOrderBy = "Order By tblPatientFacility.PatientFacility_ID"
    ElseIf optStaffSettling.Value = True Then
        temSelect = "SELECT tblPatientFacility.PatientFacility_ID as [Serial No], "
        temSelect = temSelect & " format(tblPatientFacility.BookingDate, 'dd MMM yyyy') as [Booked Date] , "
        temSelect = temSelect & "   format(tblPatientFacility.BookingTime, 'hh:mm AMPM') as [Booked Time], tblTitle.Title + ' ' + tblDoctor.DoctorName as Doctor, tblPatientFacility.AppointmentDate as [Appointment Date],  "
        temSelect = temSelect & " tblBookedUser.StaffName as [Booked By],  "
        temsqlect = temSelect & "  tblPatientMainDetails.FirstName as [Patient] , tblBookedForStaff.StaffName as [Booked For],  Format([FullyPaid],'Yes/No') AS [Settled], "
        If IsNumeric(cmbStaff.BoundText) = True Then
            temSelect = temSelect & " tblCreditSettledUser.StaffName as [Settled User], "
        End If
        temSelect = temSelect & " format(tblPatientFacility.SettleCashDate, 'dd MMM yyyy') as [Settled Date], format( tblPatientFacility.SettleCashTime , 'mm:hh AMPM') as [Settled Time], format(tblPatientFacility.Cancelled , 'Yes/No') as [Booking Cancelled] , format(tblPatientFacility.Refund, 'Yes/No') as [Booking Refunded], tblRepayUser.StaffName as [Repaid User], format(tblPatientFacility.RepayDate, 'dd MMM yyyy') as [Repaid Date]  , format( tblPatientFacility.RepayTime, 'mm:hh AMPM') as [Repaid Time] "
        temFrom = " FROM (((((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) LEFT JOIN tblStaff AS tblBookedUser ON tblPatientFacility.User_ID = tblBookedUser.Staff_ID) LEFT JOIN tblStaff AS tblBookedForStaff ON tblPatientFacility.CreditStaff_ID = tblBookedForStaff.Staff_ID) INNER JOIN tblStaff AS tblCreditSettledUser ON tblPatientFacility.CreditSettleUser_ID = tblCreditSettledUser.Staff_ID) LEFT JOIN tblStaff AS tblRepayUser ON tblPatientFacility.RepayUser_ID = tblRepayUser.Staff_ID "
        temWhere = " Where tblPatientFacility.SettleCashDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbStaff.BoundText) = True Then
            temWhere = temWhere & " AND tblPatientFacility.CreditSettleUser_ID = " & Val(cmbStaff.BoundText)
        End If
        temWhere = temWhere & " AND tblPatientFacility.CreditStaff_ID <> 0 AND tblPatientFacility.CreditStaff_ID NOT NULL"
        temWhere = temWhere & " AND tblPatientFacility.PaymentMethod_ID = 4 "
        
        temOrderBy = "Order By tblPatientFacility.PatientFacility_ID"
    
    ElseIf optTelephoneBookings.Value = True Then
        temSelect = "SELECT tblPatientFacility.PatientFacility_ID as [Serial No], "
        If dtpFrom.Value <> dtpTo.Value Then
            temSelect = temSelect & " format(tblPatientFacility.BookingDate, 'dd MMM yyyy') as [Booked Date] , "
        End If
        temSelect = temSelect & "   format(tblPatientFacility.BookingTime, 'hh:mm AMPM') as [Booked Time], tblTitle.Title + ' ' + tblDoctor.DoctorName as Doctor, tblPatientFacility.AppointmentDate as [Appointment Date],  "
        If IsNumeric(cmbStaff.BoundText) = True Then
            temSelect = temSelect & " tblBookedUser.StaffName as [Booked By],  "
        End If
        temsqlect = temSelect & " tblPatientMainDetails.FirstName as [Patient] ,  Format([FullyPaid],'Yes/No') AS [Settled], tblCreditSettledUser.StaffName as [Settled User], format(tblPatientFacility.SettleCashDate, 'dd MMM yyyy') as [Settled Date], format( tblPatientFacility.SettleCashTime , 'mm:hh AMPM') as [Settled Time], format(tblPatientFacility.Cancelled , 'Yes/No') as [Booking Cancelled] , format(tblPatientFacility.Refund, 'Yes/No') as [Booking Refunded], tblRepayUser.StaffName as [Repaid User], format(tblPatientFacility.RepayDate, 'dd MMM yyyy') as [Repaid Date]  , format( tblPatientFacility.RepayTime, 'mm:hh AMPM') as [Repaid Time] "
        temFrom = " FROM (((((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) LEFT JOIN tblStaff AS tblBookedUser ON tblPatientFacility.User_ID = tblBookedUser.Staff_ID) LEFT JOIN tblStaff AS tblBookedForStaff ON tblPatientFacility.CreditStaff_ID = tblBookedForStaff.Staff_ID) INNER JOIN tblStaff AS tblCreditSettledUser ON tblPatientFacility.CreditSettleUser_ID = tblCreditSettledUser.Staff_ID) LEFT JOIN tblStaff AS tblRepayUser ON tblPatientFacility.RepayUser_ID = tblRepayUser.Staff_ID "
        temWhere = " Where tblPatientFacility.BookingDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbStaff.BoundText) = True Then
            temWhere = temWhere & " AND tblPatientFacility.User_ID = " & Val(cmbStaff.BoundText)
        End If
        temWhere = temWhere & " AND (tblPatientFacility.CreditStaff_ID = 0 OR tblPatientFacility.CreditStaff_ID IS NULL) "
        temWhere = temWhere & " AND tblPatientFacility.PaymentMethod_ID = 4 "
        
        temOrderBy = "Order By tblPatientFacility.PatientFacility_ID"
    
    ElseIf optTelephoneSettling.Value = True Then
    
    
    
        temSelect = "SELECT tblPatientFacility.PatientFacility_ID as [Serial No], "
        temSelect = temSelect & " format(tblPatientFacility.BookingDate, 'dd MMM yyyy') as [Booked Date] , "
        temSelect = temSelect & "   format(tblPatientFacility.BookingTime, 'hh:mm AMPM') as [Booked Time], tblTitle.Title + ' ' + tblDoctor.DoctorName as Doctor, tblPatientFacility.AppointmentDate as [Appointment Date],  "
        temSelect = temSelect & " tblBookedUser.StaffName as [Booked By],  "
        temsqlect = temSelect & "  tblPatientMainDetails.FirstName as [Patient],  Format([FullyPaid],'Yes/No') AS [Settled], "
        If IsNumeric(cmbStaff.BoundText) = True Then
            temSelect = temSelect & " tblCreditSettledUser.StaffName as [Settled User], "
        End If
        temSelect = temSelect & " format(tblPatientFacility.SettleCashDate, 'dd MMM yyyy') as [Settled Date], format( tblPatientFacility.SettleCashTime , 'mm:hh AMPM') as [Settled Time], format(tblPatientFacility.Cancelled , 'Yes/No') as [Booking Cancelled] , format(tblPatientFacility.Refund, 'Yes/No') as [Booking Refunded], tblRepayUser.StaffName as [Repaid User], format(tblPatientFacility.RepayDate, 'dd MMM yyyy') as [Repaid Date]  , format( tblPatientFacility.RepayTime, 'mm:hh AMPM') as [Repaid Time] "
        temFrom = " FROM (((((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) LEFT JOIN tblStaff AS tblBookedUser ON tblPatientFacility.User_ID = tblBookedUser.Staff_ID) LEFT JOIN tblStaff AS tblBookedForStaff ON tblPatientFacility.CreditStaff_ID = tblBookedForStaff.Staff_ID) INNER JOIN tblStaff AS tblCreditSettledUser ON tblPatientFacility.CreditSettleUser_ID = tblCreditSettledUser.Staff_ID) LEFT JOIN tblStaff AS tblRepayUser ON tblPatientFacility.RepayUser_ID = tblRepayUser.Staff_ID "
        temWhere = " Where tblPatientFacility.SettleCashDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbStaff.BoundText) = True Then
            temWhere = temWhere & " AND tblPatientFacility.CreditSettleUser_ID = " & Val(cmbStaff.BoundText)
        End If
        temWhere = temWhere & " AND (tblPatientFacility.CreditStaff_ID = 0 OR tblPatientFacility.CreditStaff_ID IS NULL) "
        temWhere = temWhere & " AND tblPatientFacility.PaymentMethod_ID = 4 "
        
        temOrderBy = "Order By tblPatientFacility.PatientFacility_ID"
    
    
    End If
    
    temSQL = temSelect & " " & temFrom & " " & temWhere & " " & temOrderBy
    
    
    
    FillAnyGrid temSQL, gridBookings
    
    If IsNumeric(cmbStaff.BoundText) = True Then
        temSubTopic = " User - " & cmbStaff.Text
    End If
    If dtpFrom.Value <> dtpTo.Value Then
        temSubTopic = temSubTopic & " From - " & Format(dtpFrom.Value, "dd MMMM yyyy") & " To - " & Format(dtpTo.Value, "dd MMMM yyyy")
    Else
        temSubTopic = temSubTopic & " On - " & Format(dtpFrom.Value, "dd MMMM yyyy")
    End If
    
    If optStaffBooking.Value = True Then
        temTopic = "Staff Bookings"
    ElseIf optStaffSettling.Value = True Then
        temTopic = "Settling Staff Bookings"
    ElseIf optTelephoneBookings.Value = True Then
        temTopic = "Telephone Bookings"
    ElseIf optTelephoneSettling.Value = True Then
        temTopic = "Settling Telephone Bookings"
    End If
    
End Sub

Private Sub btnPrint_Click()
    Dim myPR As PrintReport
    GridPrint gridBookings, myPR, temTopic, temSubTopic
End Sub

Private Sub Form_Load()
    Call GetSettings
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
End Sub

Private Sub GetSettings()
    GetCommonSettings
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub
