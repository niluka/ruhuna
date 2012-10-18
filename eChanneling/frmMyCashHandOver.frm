VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMyCashHandOver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Cash Hand Over"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMyCashHandOver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   8985
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   0
      Width           =   8295
      Begin MSDataListLib.DataCombo dtcUserName 
         Height          =   360
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label4 
         Caption         =   "User Name"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   7695
      Begin btButtonEx.ButtonEx bttnParintSummary 
         Height          =   375
         Left            =   5280
         TabIndex        =   28
         Top             =   4680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print Summary"
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
      Begin VB.Line Line4 
         X1              =   2880
         X2              =   5160
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line3 
         X1              =   2880
         X2              =   5160
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line2 
         X1              =   2880
         X2              =   5160
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label lblPatientsTotalCreditsettle 
         Height          =   375
         Left            =   5280
         TabIndex        =   30
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblPatientsTotalCash 
         Height          =   375
         Left            =   5280
         TabIndex        =   29
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label20 
         Caption         =   "To Pay Today Doctor Fees"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label lblTodayBalanceDoctorPayment 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   26
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label lblDoctorPayment 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   25
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label17 
         Caption         =   "Less : Paid Today Doctor Fees"
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label lblTodayDoctorFees 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   23
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label15 
         Caption         =   "Total Today Doctor Fees"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   7560
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblHospitalCharges 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   21
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Hospital Charges"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblToPayDoctorPayment 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   19
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "All Doctors Paybal Fees"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label lblTotalIncome 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label lblCreditSettle 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblCashBooking 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Total Cash Income"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Total Credit Settle"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Total Cash Booking"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2175
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11456
      _Version        =   393216
      Tab             =   1
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
      TabPicture(0)   =   "frmMyCashHandOver.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblDate"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected day"
      TabPicture(1)   =   "frmMyCashHandOver.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Period"
      TabPicture(2)   =   "frmMyCashHandOver.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(2)=   "DTPicker2"
      Tab(2).Control(3)=   "DTPicker3"
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   -69360
         TabIndex        =   7
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   160366593
         CurrentDate     =   39489
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -73560
         TabIndex        =   6
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   160366593
         CurrentDate     =   39489
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   160366593
         CurrentDate     =   39489
      End
      Begin VB.Label Label3 
         Caption         =   "To"
         Height          =   375
         Left            =   -70320
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Date From"
         Height          =   375
         Left            =   -74640
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblDate 
         Caption         =   "Label1"
         Height          =   375
         Left            =   -72000
         TabIndex        =   5
         Top             =   600
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmMyCashHandOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemTotalCash As Double
Dim TemTotalCredit As Double
Dim TemDoctorPayment As Double
Dim TemTotalIncome As Double
Dim TemToPayConsultantfee As Double
Dim TemTodayConsultantfee As Double
Dim TemhospitalFee As Double

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnParintSummary_Click()
If IsNumeric(dtcUserName.BoundText) = False Then
    Dim TemResponce As Integer
    TemResponce = MsgBox("Please select a user", vbCritical, "User?")
    dtcUserName.SetFocus
    Exit Sub
End If
With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    .Open "Select * From tblTem"
    Set dtrShiftHandOver.DataSource = DataEnvironment1.rssqlTemSu1
End With
With dtrShiftHandOver
    If HospitalDetails = True Then
        .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    Else
        .Sections("Section4").Controls.Item("RptName").Caption = Empty
        .Sections("Section4").Controls.Item("RptAddress").Caption = Empty
    End If
    .Sections("Section4").Controls.Item("RptCashierName").Caption = "Cashier Name   :    " & dtcUserName.Text
    
    .Sections("Section2").Controls.Item("rptLCashChannaling").Caption = Format(lblCashBooking, "0.00")
    .Sections("Section2").Controls.Item("rptlCashFromCreditChaneling").Caption = Format(lblCreditSettle, "0.00")
    .Sections("Section2").Controls.Item("lblCashPatients").Caption = lblPatientsTotalCash.Caption
    .Sections("Section2").Controls.Item("lblCreditPatients").Caption = lblPatientsTotalCreditsettle.Caption
    .Sections("Section2").Controls.Item("rptTotalCashReceive").Caption = Format(lblTotalIncome, "0.00")
    .Sections("Section2").Controls.Item("rptHospitalCharges").Caption = Format(lblHospitalCharges, "0.00")
    .Sections("Section2").Controls.Item("rptTotalDoctorDue").Caption = Format(lblToPayDoctorPayment, "0.00")
    .Sections("Section2").Controls.Item("rptTodaydueAmount").Caption = Format(lblTodayDoctorFees, "0.00")
    .Sections("Section2").Controls.Item("lblPaidDoctorfeesToday").Caption = Format(lblDoctorPayment, "0.00")
    .Sections("Section2").Controls.Item("lblBalanceDoctotfee").Caption = Format(lblTodayBalanceDoctorPayment, "0.00")
        
        If SSTab1.Tab = 0 Then
        .Sections("Section4").Controls.Item("rptdate").Caption = "Date     : " & Format(Date, DefaultLongDate)
        ElseIf SSTab1.Tab = 1 Then
        .Sections("Section4").Controls.Item("rptdate").Caption = "Date     : " & DTPicker1.Value
        ElseIf SSTab1.Tab = 2 Then
        .Sections("Section4").Controls.Item("rptdate").Caption = "Date From  : " & DTPicker2.Value & "  To   " & DTPicker3.Value
        End If
        
    .Show
End With

End Sub

Private Sub dtcUserName_Click(Area As Integer)
If IsNumeric(dtcUserName.BoundText) = False Then Exit Sub
Call CalculateVales

End Sub

Private Sub DTPicker1_Change()
Call CalculateVales

End Sub

Private Sub DTPicker2_Change()
Call CalculateVales

End Sub


Private Sub DTPicker3_Change()
Call CalculateVales

End Sub

Private Sub Form_Load()
Call FillStaffNames
lblDate.Caption = Date
dtcUserName.BoundText = UserID
SSTab1.Tab = 0
SSTab1.TabEnabled(2) = False
Call CalculateVales
DTPicker1.Format = dtpCustom
DTPicker1.CustomFormat = DefaultLongDate
DTPicker1.Value = Date
End Sub

Private Sub ClearValues()
lblCashBooking.Caption = "0.00"
lblCreditSettle.Caption = "0.00"
lblTotalIncome.Caption = "0.00"
lblToPayDoctorPayment.Caption = "0.00"
lblHospitalCharges.Caption = "0.00"
lblTodayDoctorFees.Caption = "0.00"
lblDoctorPayment.Caption = "0.00"
lblTodayBalanceDoctorPayment.Caption = "0.00"
lblPatientsTotalCash.Caption = Empty
lblPatientsTotalCash.Caption = Empty

End Sub

Private Sub CalculateVales()
Call ClearValues
Call CashBooking
Call CreditSettle
Call TotalIncome
Call ToPayDoctorPayment
Call CaculateHospitalCharges
Call FindTodayDoctorPayment
Call DoctorPaidAmount
Call FindBalanceConsultonfee

End Sub

Private Sub FillStaffNames()

With DataEnvironment1.rssqlTemHandOver1
    If .State = 1 Then .Close
    .Open "Select* From tblStaff Order By StaffName"
    Set dtcUserName.RowSource = DataEnvironment1.rssqlTemHandOver1
    dtcUserName.ListField = "StaffName"
    dtcUserName.BoundColumn = "Staff_ID"


End With

End Sub

Private Sub CashBooking()
TemTotalCash = 0
TemhospitalFee = 0
With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
 
    Case 0
        .Open "Select* From tblPatientFacility Where (User_ID = " & dtcUserName.BoundText & ") and (BookingDate ='" & Date & "') and (PaymentMode = 'Cash')and (fullypaid = 1) and (((cancelled = 0) and (refund = 0))or(cancelled = 1 and refund = 0 and repayuser_ID <> " & dtcUserName.BoundText & ") or (refund = 1 and cancelled = 0 and repayuser_ID <> " & dtcUserName.BoundText & ")) "
    Case 1
        .Open "Select* From tblPatientFacility Where (User_ID = " & dtcUserName.BoundText & ") and (BookingDate ='" & DTPicker1.Value & "') and (PaymentMode = 'Cash')and (fullypaid = 1) and (((cancelled = 0) and (refund = 0))or(cancelled = 1 and refund = 0 and repayuser_ID <> " & dtcUserName.BoundText & ") or (refund = 1 and cancelled = 0 and repayuser_ID <> " & dtcUserName.BoundText & ")) "
    Case 2
'        .Open "Select* From tblPatientFacility Where (User_ID = " & dtcUserName.BoundText & ") and (BookingDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and (PaymentMode = 'Cash')and (fullypaid = 1) "
    
    End Select
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemTotalCash = Val(TemTotalCash) + Val(!totalfee)
    TemhospitalFee = Val(TemhospitalFee) + Val(!InstitutionFee)

    .MoveNext
    Loop
    
    lblPatientsTotalCash.Caption = "No Of Patients  :  " & DataEnvironment1.rssqlTem10.RecordCount
    
    If .State = 1 Then .Close
    
End With
lblCashBooking.Caption = Format(TemTotalCash, "0.00")
End Sub

Private Sub CreditSettle()
TemTotalCredit = 0
With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
 
    Case 0
        .Open "Select* From tblPatientFacility Where (CreditSettleUser_ID = " & dtcUserName.BoundText & ") and (SettleCashDate ='" & Date & "') and (PaymentMode = 'Credit')and (((cancelled = 0) and (refund = 0)) or (cancelled = 1 and refund = 0 and repayuser_ID <> " & dtcUserName.BoundText & ") or (refund = 1 and cancelled = 0 and repayuser_ID <> " & dtcUserName.BoundText & ")) "
    Case 1
        .Open "Select* From tblPatientFacility Where (CreditSettleUser_ID = " & dtcUserName.BoundText & ") and (SettleCashDate ='" & DTPicker1.Value & "') and (PaymentMode = 'Credit')and (((cancelled = 0) and (refund = 0)) or (cancelled = 1 and refund = 0 and repayuser_ID <> " & dtcUserName.BoundText & ") or (refund = 1 and cancelled = 0 and repayuser_ID <> " & dtcUserName.BoundText & ")) "
    Case 2
'        .Open "Select* From tblPatientFacility Where (CreditSettleUser_ID = " & dtcUserName.BoundText & ") and (SettleCashDate ='" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and (PaymentMode = 'Credit') and (fullypaid = 1) "
    
    End Select

    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemTotalCredit = Val(TemTotalCredit) + Val(!totalfee)
    TemhospitalFee = Val(TemhospitalFee) + Val(!InstitutionFee)
    .MoveNext
    Loop
    
    lblPatientsTotalCreditsettle.Caption = "No Of Patients  :  " & DataEnvironment1.rssqlTem10.RecordCount

    If .State = 1 Then .Close
End With
lblCreditSettle.Caption = Format(TemTotalCredit, "0.00")

End Sub

Private Sub TotalIncome()
TemTotalIncome = Val(TemTotalCash) + Val(TemTotalCredit)
lblTotalIncome.Caption = Format(TemTotalIncome, "0.00")

End Sub

Private Sub ToPayDoctorPayment()
TemToPayConsultantfee = 0
With DataEnvironment1.rssqlTem11
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
 
    Case 0
        .Open "Select* From tblPatientFacility Where (User_ID = " & dtcUserName.BoundText & ") and (BookingDate ='" & Date & "')and (cancelled = 0)and (refund = 0) "
'         .Open "Select* From tblPatientFacility Where (BookingDate ='" & Date & "')and (cancelled = 0)and (refund = 0) "
   
    Case 1
        .Open "Select* From tblPatientFacility Where (User_ID = " & dtcUserName.BoundText & ") and (BookingDate ='" & DTPicker1.Value & "') and (cancelled = 0)and (refund = 0) "
'        .Open "Select* From tblPatientFacility Where (BookingDate ='" & DTPicker1.Value & "') and (cancelled = 0)and (refund = 0) "
 
    Case 2
'        .Open "Select* From tblPatientFacility Where (User_ID = " & dtcUserName.BoundText & ") and (BookingDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and (fullypaid = 1) and (cancelled = 0)and (refund = 0) "
    
    End Select

    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemToPayConsultantfee = Val(TemToPayConsultantfee) + Val(!personaldue)
    .MoveNext
    Loop
    
    If .State = 1 Then .Close
End With
lblToPayDoctorPayment.Caption = Format(TemToPayConsultantfee, "0.00")
End Sub

Private Sub CaculateHospitalCharges()

lblHospitalCharges.Caption = Format(TemhospitalFee, "0.00")
End Sub

Private Sub FindTodayDoctorPayment()
TemTodayConsultantfee = 0
With DataEnvironment1.rssqlTem11
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
 
    Case 0
        .Open "Select* From tblPatientFacility Where (User_ID = " & dtcUserName.BoundText & ") and  (BookingDate = '" & Date & "') and  (AppointmentDate = '" & Date & "')and (fullypaid = 1) and (cancelled = 0)and (refund = 0) "
    Case 1
        .Open "Select* From tblPatientFacility Where (User_ID = " & dtcUserName.BoundText & ") and (BookingDate = '" & DTPicker1.Value & "') and (AppointmentDate = '" & DTPicker1.Value & "')and (fullypaid = 1) and (cancelled = 0)and (refund = 0) "
    Case 2
        '.Open "Select* From tblPatientFacility Where (User_ID = " & dtcUserName.BoundText & ") and (BookingDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and (AppointmentDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "')and (fullypaid = 1) and (cancelled = 0)and (refund = 0) "
    
    End Select

    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemTodayConsultantfee = Val(TemTodayConsultantfee) + Val(!personaldue)
    .MoveNext
    Loop
    
    If .State = 1 Then .Close
End With

lblTodayDoctorFees.Caption = Format(TemTodayConsultantfee, "0.00")
End Sub

Private Sub DoctorPaidAmount()
TemDoctorPayment = 0
With DataEnvironment1.rssqlTem1
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
 
    Case 0
        .Open "Select* From tblStaffPayment Where (User_ID = " & dtcUserName.BoundText & ") and (PaidDate = '" & Date & "') "
    Case 1
        .Open "Select* From tblStaffPayment Where (User_ID = " & dtcUserName.BoundText & ") and (PaidDate = '" & DTPicker1.Value & "') "
    Case 2
        .Open "Select* From tblStaffPayment Where (User_ID = " & dtcUserName.BoundText & ") and (PaidDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "' ) "
    
    End Select

    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemDoctorPayment = Val(TemDoctorPayment) + Val(!PaidAmount)
    .MoveNext
    Loop
    
    If .State = 1 Then .Close
End With

lblDoctorPayment.Caption = Format(TemDoctorPayment, "0.00")

End Sub

Private Sub FindBalanceConsultonfee()

lblTodayBalanceDoctorPayment.Caption = Format(Val(TemTodayConsultantfee) - Val(TemDoctorPayment), "0.00")
End Sub
