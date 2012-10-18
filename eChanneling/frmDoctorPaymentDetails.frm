VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDoctorPaymentDetails 
   Caption         =   "Doctor Payment Details"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   12840
      TabIndex        =   27
      Top             =   7680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   11400
      TabIndex        =   26
      Top             =   7680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
   Begin VB.Frame FrameDoctorSelection 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      Begin VB.Frame FrameSelectDoctor 
         Caption         =   "Select Doctor"
         Height          =   3135
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   6495
         Begin VB.ListBox ListConsultantIDs 
            Height          =   1020
            Left            =   5760
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ListBox ListSpecialityIDs 
            Height          =   1020
            Left            =   2040
            TabIndex        =   24
            Top             =   600
            Visible         =   0   'False
            Width           =   615
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
            Height          =   2400
            ItemData        =   "frmDoctorPaymentDetails.frx":0000
            Left            =   120
            List            =   "frmDoctorPaymentDetails.frx":0002
            TabIndex        =   7
            ToolTipText     =   "List of Specialities"
            Top             =   600
            Width           =   2535
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
            Height          =   2400
            ItemData        =   "frmDoctorPaymentDetails.frx":0004
            Left            =   2880
            List            =   "frmDoctorPaymentDetails.frx":0006
            TabIndex        =   6
            ToolTipText     =   "List of Consultants of selected speciality"
            Top             =   600
            Width           =   3495
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Speciality"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultant"
            Height          =   255
            Left            =   2880
            TabIndex        =   8
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.OptionButton OptionSelectedDoctors 
         Caption         =   "Selected Doctor"
         Height          =   240
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton OptionAllDoctors 
         Caption         =   "All Doctors"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Bindings        =   "frmDoctorPaymentDetails.frx":0008
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   100
      FillStyle       =   1
      SelectionMode   =   2
      AllowUserResizing=   1
      DataMember      =   "cmmdDoctorPayments_Grouping"
      _NumberOfBands  =   2
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).BandExpandable=   0   'False
      _Band(0)._NumMapCols=   14
      _Band(0)._MapCol(0)._Name=   "tblPatientFacility.Staff_ID"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Alignment=   7
      _Band(0)._MapCol(1)._Name=   "AppointmentDate"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "PaidToSTaff"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "tblPatientFacility.StaffPayment_ID"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(3)._Alignment=   7
      _Band(0)._MapCol(4)._Name=   "PaidDoctorName"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "PaidStaffName"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "DoctorPaidDate"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "PaidOrNot"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(8)._Name=   "ToPayDoctor"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(9)._Name=   "PaidAmountToDoctor"
      _Band(0)._MapCol(9)._RSIndex=   9
      _Band(0)._MapCol(10)._Name=   "DocPaidDate"
      _Band(0)._MapCol(10)._RSIndex=   10
      _Band(0)._MapCol(11)._Name=   "DocPaidTime"
      _Band(0)._MapCol(11)._RSIndex=   11
      _Band(0)._MapCol(12)._Name=   "PaidID"
      _Band(0)._MapCol(12)._RSIndex=   12
      _Band(0)._MapCol(13)._Name=   "ForAppointmentDate"
      _Band(0)._MapCol(13)._RSIndex=   13
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   85
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   0
      _Band(1)._ParentBand=   0
      _Band(1)._NumMapCols=   85
      _Band(1)._MapCol(0)._Name=   "PatientFacility_ID"
      _Band(1)._MapCol(0)._RSIndex=   0
      _Band(1)._MapCol(0)._Alignment=   7
      _Band(1)._MapCol(1)._Name=   "tblPatientFacility.User_ID"
      _Band(1)._MapCol(1)._RSIndex=   1
      _Band(1)._MapCol(1)._Alignment=   7
      _Band(1)._MapCol(2)._Name=   "RepayUser_ID"
      _Band(1)._MapCol(2)._RSIndex=   2
      _Band(1)._MapCol(2)._Alignment=   7
      _Band(1)._MapCol(3)._Name=   "CreditSettleUser_ID"
      _Band(1)._MapCol(3)._RSIndex=   3
      _Band(1)._MapCol(3)._Alignment=   7
      _Band(1)._MapCol(4)._Name=   "PatientID"
      _Band(1)._MapCol(4)._RSIndex=   4
      _Band(1)._MapCol(4)._Alignment=   7
      _Band(1)._MapCol(5)._Name=   "tblPatientFacility.HospitalFacility_ID"
      _Band(1)._MapCol(5)._RSIndex=   5
      _Band(1)._MapCol(5)._Alignment=   7
      _Band(1)._MapCol(6)._Name=   "tblPatientFacility.FacilityStaff_ID"
      _Band(1)._MapCol(6)._RSIndex=   6
      _Band(1)._MapCol(6)._Alignment=   7
      _Band(1)._MapCol(7)._Name=   "FacilityCatogery"
      _Band(1)._MapCol(7)._RSIndex=   7
      _Band(1)._MapCol(7)._Alignment=   7
      _Band(1)._MapCol(8)._Name=   "tblPatientFacility.Staff_ID"
      _Band(1)._MapCol(8)._RSIndex=   8
      _Band(1)._MapCol(8)._Alignment=   7
      _Band(1)._MapCol(9)._Name=   "PatientBill_ID"
      _Band(1)._MapCol(9)._RSIndex=   9
      _Band(1)._MapCol(9)._Alignment=   7
      _Band(1)._MapCol(10)._Name=   "BookingDate"
      _Band(1)._MapCol(10)._RSIndex=   10
      _Band(1)._MapCol(11)._Name=   "BookingTime"
      _Band(1)._MapCol(11)._RSIndex=   11
      _Band(1)._MapCol(12)._Name=   "AppointmentDate"
      _Band(1)._MapCol(12)._RSIndex=   12
      _Band(1)._MapCol(13)._Name=   "Secession"
      _Band(1)._MapCol(13)._RSIndex=   13
      _Band(1)._MapCol(13)._Alignment=   7
      _Band(1)._MapCol(14)._Name=   "AppointmentTime"
      _Band(1)._MapCol(14)._RSIndex=   14
      _Band(1)._MapCol(15)._Name=   "DaySerial"
      _Band(1)._MapCol(15)._RSIndex=   15
      _Band(1)._MapCol(15)._Alignment=   7
      _Band(1)._MapCol(16)._Name=   "PersonalFee"
      _Band(1)._MapCol(16)._RSIndex=   16
      _Band(1)._MapCol(16)._Alignment=   7
      _Band(1)._MapCol(17)._Name=   "PersonalFeeToPay"
      _Band(1)._MapCol(17)._RSIndex=   17
      _Band(1)._MapCol(17)._Alignment=   7
      _Band(1)._MapCol(18)._Name=   "InstitutionFee"
      _Band(1)._MapCol(18)._RSIndex=   18
      _Band(1)._MapCol(18)._Alignment=   7
      _Band(1)._MapCol(19)._Name=   "InstitutionFeeToPay"
      _Band(1)._MapCol(19)._RSIndex=   19
      _Band(1)._MapCol(19)._Alignment=   7
      _Band(1)._MapCol(20)._Name=   "OtherFee"
      _Band(1)._MapCol(20)._RSIndex=   20
      _Band(1)._MapCol(20)._Alignment=   7
      _Band(1)._MapCol(21)._Name=   "OtherFeeToPay"
      _Band(1)._MapCol(21)._RSIndex=   21
      _Band(1)._MapCol(21)._Alignment=   7
      _Band(1)._MapCol(22)._Name=   "TotalFee"
      _Band(1)._MapCol(22)._RSIndex=   22
      _Band(1)._MapCol(22)._Alignment=   7
      _Band(1)._MapCol(23)._Name=   "TotalFeeToPay"
      _Band(1)._MapCol(23)._RSIndex=   23
      _Band(1)._MapCol(23)._Alignment=   7
      _Band(1)._MapCol(24)._Name=   "FullyPaid"
      _Band(1)._MapCol(24)._RSIndex=   24
      _Band(1)._MapCol(25)._Name=   "FullyPaidNull"
      _Band(1)._MapCol(25)._RSIndex=   25
      _Band(1)._MapCol(25)._Alignment=   7
      _Band(1)._MapCol(26)._Name=   "ResultSuccess"
      _Band(1)._MapCol(26)._RSIndex=   26
      _Band(1)._MapCol(27)._Name=   "PersonalRefund"
      _Band(1)._MapCol(27)._RSIndex=   27
      _Band(1)._MapCol(27)._Alignment=   7
      _Band(1)._MapCol(28)._Name=   "PersonalRefundComment"
      _Band(1)._MapCol(28)._RSIndex=   28
      _Band(1)._MapCol(29)._Name=   "InstitutionRefund"
      _Band(1)._MapCol(29)._RSIndex=   29
      _Band(1)._MapCol(29)._Alignment=   7
      _Band(1)._MapCol(30)._Name=   "InstitutionRefundComment"
      _Band(1)._MapCol(30)._RSIndex=   30
      _Band(1)._MapCol(31)._Name=   "OtherRefund"
      _Band(1)._MapCol(31)._RSIndex=   31
      _Band(1)._MapCol(31)._Alignment=   7
      _Band(1)._MapCol(32)._Name=   "OtherRefundComment"
      _Band(1)._MapCol(32)._RSIndex=   32
      _Band(1)._MapCol(33)._Name=   "TotalRefund"
      _Band(1)._MapCol(33)._RSIndex=   33
      _Band(1)._MapCol(33)._Alignment=   7
      _Band(1)._MapCol(34)._Name=   "RefundToPatient"
      _Band(1)._MapCol(34)._RSIndex=   34
      _Band(1)._MapCol(35)._Name=   "RefundToAgent"
      _Band(1)._MapCol(35)._RSIndex=   35
      _Band(1)._MapCol(36)._Name=   "RepayComments"
      _Band(1)._MapCol(36)._RSIndex=   36
      _Band(1)._MapCol(37)._Name=   "NewBooking"
      _Band(1)._MapCol(37)._RSIndex=   37
      _Band(1)._MapCol(38)._Name=   "CarriedForwardID"
      _Band(1)._MapCol(38)._RSIndex=   38
      _Band(1)._MapCol(38)._Alignment=   7
      _Band(1)._MapCol(39)._Name=   "BoughtForwardID"
      _Band(1)._MapCol(39)._RSIndex=   39
      _Band(1)._MapCol(39)._Alignment=   7
      _Band(1)._MapCol(40)._Name=   "PaidToSTaff"
      _Band(1)._MapCol(40)._RSIndex=   40
      _Band(1)._MapCol(41)._Name=   "StaffPayment"
      _Band(1)._MapCol(41)._RSIndex=   41
      _Band(1)._MapCol(41)._Alignment=   7
      _Band(1)._MapCol(42)._Name=   "PaidToStaffOn"
      _Band(1)._MapCol(42)._RSIndex=   42
      _Band(1)._MapCol(43)._Name=   "PaidToStaffUser"
      _Band(1)._MapCol(43)._RSIndex=   43
      _Band(1)._MapCol(43)._Alignment=   7
      _Band(1)._MapCol(44)._Name=   "tblPatientFacility.StaffPayment_ID"
      _Band(1)._MapCol(44)._RSIndex=   44
      _Band(1)._MapCol(44)._Alignment=   7
      _Band(1)._MapCol(45)._Name=   "Cancelled"
      _Band(1)._MapCol(45)._RSIndex=   45
      _Band(1)._MapCol(46)._Name=   "CancelledNull"
      _Band(1)._MapCol(46)._RSIndex=   46
      _Band(1)._MapCol(46)._Alignment=   7
      _Band(1)._MapCol(47)._Name=   "CancelRemark"
      _Band(1)._MapCol(47)._RSIndex=   47
      _Band(1)._MapCol(48)._Name=   "Refund"
      _Band(1)._MapCol(48)._RSIndex=   48
      _Band(1)._MapCol(49)._Name=   "RefundNull"
      _Band(1)._MapCol(49)._RSIndex=   49
      _Band(1)._MapCol(49)._Alignment=   7
      _Band(1)._MapCol(50)._Name=   "RefundRemark"
      _Band(1)._MapCol(50)._RSIndex=   50
      _Band(1)._MapCol(51)._Name=   "PaymentMode"
      _Band(1)._MapCol(51)._RSIndex=   51
      _Band(1)._MapCol(52)._Name=   "Agent_ID"
      _Band(1)._MapCol(52)._RSIndex=   52
      _Band(1)._MapCol(52)._Alignment=   7
      _Band(1)._MapCol(53)._Name=   "CreditAgent_ID"
      _Band(1)._MapCol(53)._RSIndex=   53
      _Band(1)._MapCol(53)._Alignment=   7
      _Band(1)._MapCol(54)._Name=   "PaymentMethod_Id"
      _Band(1)._MapCol(54)._RSIndex=   54
      _Band(1)._MapCol(54)._Alignment=   7
      _Band(1)._MapCol(55)._Name=   "BillPrinted"
      _Band(1)._MapCol(55)._RSIndex=   55
      _Band(1)._MapCol(56)._Name=   "RepayDate"
      _Band(1)._MapCol(56)._RSIndex=   56
      _Band(1)._MapCol(57)._Name=   "RepayTime"
      _Band(1)._MapCol(57)._RSIndex=   57
      _Band(1)._MapCol(58)._Name=   "RepayComment"
      _Band(1)._MapCol(58)._RSIndex=   58
      _Band(1)._MapCol(59)._Name=   "SettleCashDate"
      _Band(1)._MapCol(59)._RSIndex=   59
      _Band(1)._MapCol(60)._Name=   "SettleCashTime"
      _Band(1)._MapCol(60)._RSIndex=   60
      _Band(1)._MapCol(61)._Name=   "tblPatientFacility.C"
      _Band(1)._MapCol(61)._RSIndex=   61
      _Band(1)._MapCol(62)._Name=   "tblPatientFacility.IsADoctor"
      _Band(1)._MapCol(62)._RSIndex=   62
      _Band(1)._MapCol(63)._Name=   "IsAStaffMember"
      _Band(1)._MapCol(63)._RSIndex=   63
      _Band(1)._MapCol(64)._Name=   "IsAnInvestigation"
      _Band(1)._MapCol(64)._RSIndex=   64
      _Band(1)._MapCol(65)._Name=   "AgentRefNo"
      _Band(1)._MapCol(65)._RSIndex=   65
      _Band(1)._MapCol(66)._Name=   "PatientAbsent"
      _Band(1)._MapCol(66)._RSIndex=   66
      _Band(1)._MapCol(67)._Name=   "PersonalDue"
      _Band(1)._MapCol(67)._RSIndex=   67
      _Band(1)._MapCol(67)._Alignment=   7
      _Band(1)._MapCol(68)._Name=   "InstitutionDue"
      _Band(1)._MapCol(68)._RSIndex=   68
      _Band(1)._MapCol(68)._Alignment=   7
      _Band(1)._MapCol(69)._Name=   "OtherDue"
      _Band(1)._MapCol(69)._RSIndex=   69
      _Band(1)._MapCol(69)._Alignment=   7
      _Band(1)._MapCol(70)._Name=   "TotalDue"
      _Band(1)._MapCol(70)._RSIndex=   70
      _Band(1)._MapCol(70)._Alignment=   7
      _Band(1)._MapCol(71)._Name=   "DoctorName"
      _Band(1)._MapCol(71)._RSIndex=   71
      _Band(1)._MapCol(72)._Name=   "FirstName"
      _Band(1)._MapCol(72)._RSIndex=   72
      _Band(1)._MapCol(73)._Name=   "StaffName"
      _Band(1)._MapCol(73)._RSIndex=   73
      _Band(1)._MapCol(74)._Name=   "tblStaffPayment.StaffPayment_ID"
      _Band(1)._MapCol(74)._RSIndex=   74
      _Band(1)._MapCol(74)._Alignment=   7
      _Band(1)._MapCol(75)._Name=   "tblStaffPayment.HospitalFacility_ID"
      _Band(1)._MapCol(75)._RSIndex=   75
      _Band(1)._MapCol(75)._Alignment=   7
      _Band(1)._MapCol(76)._Name=   "tblStaffPayment.FacilityStaff_ID"
      _Band(1)._MapCol(76)._RSIndex=   76
      _Band(1)._MapCol(76)._Alignment=   7
      _Band(1)._MapCol(77)._Name=   "tblStaffPayment.IsADoctor"
      _Band(1)._MapCol(77)._RSIndex=   77
      _Band(1)._MapCol(78)._Name=   "tblStaffPayment.Staff_ID"
      _Band(1)._MapCol(78)._RSIndex=   78
      _Band(1)._MapCol(78)._Alignment=   7
      _Band(1)._MapCol(79)._Name=   "PaidAmount"
      _Band(1)._MapCol(79)._RSIndex=   79
      _Band(1)._MapCol(79)._Alignment=   7
      _Band(1)._MapCol(80)._Name=   "PaidDate"
      _Band(1)._MapCol(80)._RSIndex=   80
      _Band(1)._MapCol(81)._Name=   "PaidTime"
      _Band(1)._MapCol(81)._RSIndex=   81
      _Band(1)._MapCol(82)._Name=   "tblStaffPayment.User_ID"
      _Band(1)._MapCol(82)._RSIndex=   82
      _Band(1)._MapCol(82)._Alignment=   7
      _Band(1)._MapCol(83)._Name=   "Comments"
      _Band(1)._MapCol(83)._RSIndex=   83
      _Band(1)._MapCol(84)._Name=   "tblStaffPayment.C"
      _Band(1)._MapCol(84)._RSIndex=   84
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   7080
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Selected Day"
      TabPicture(0)   =   "frmDoctorPaymentDetails.frx":0027
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MonthViewDay"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "OptionToday"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "OptionYesterday"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "OptionDayBeforeYesterday"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OptionTomorrow"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OptionDayAfterTomorrow"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Selected Period"
      TabPicture(1)   =   "frmDoctorPaymentDetails.frx":0043
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MonthViewFrom"
      Tab(1).Control(1)=   "OptionThisYear"
      Tab(1).Control(2)=   "OptionLastMonth"
      Tab(1).Control(3)=   "OptionThisMonth"
      Tab(1).Control(4)=   "OptionLastWeek"
      Tab(1).Control(5)=   "OptionThisweek"
      Tab(1).Control(6)=   "MonthViewTo"
      Tab(1).Control(7)=   "OptionLastYear"
      Tab(1).ControlCount=   8
      Begin MSComCtl2.MonthView MonthViewFrom 
         Height          =   2820
         Left            =   -73200
         TabIndex        =   10
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         ShowToday       =   0   'False
         StartOfWeek     =   61931522
         CurrentDate     =   39477
      End
      Begin VB.OptionButton OptionThisYear 
         Caption         =   "This year"
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   2400
         Width           =   2295
      End
      Begin VB.OptionButton OptionLastMonth 
         Caption         =   "Last month"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   1920
         Width           =   2295
      End
      Begin VB.OptionButton OptionThisMonth 
         Caption         =   "This month"
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton OptionLastWeek 
         Caption         =   "Last week"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton OptionThisweek 
         Caption         =   "This week"
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton OptionDayAfterTomorrow 
         Caption         =   "Day after tomorrow"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Width           =   2295
      End
      Begin VB.OptionButton OptionTomorrow 
         Caption         =   "Tomorrow"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   2295
      End
      Begin VB.OptionButton OptionDayBeforeYesterday 
         Caption         =   "Daybefore yesterday"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   2295
      End
      Begin VB.OptionButton OptionYesterday 
         Caption         =   "Yesterday"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   2295
      End
      Begin VB.OptionButton OptionToday 
         Caption         =   "Today"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Value           =   -1  'True
         Width           =   2295
      End
      Begin MSComCtl2.MonthView MonthViewTo 
         Height          =   2820
         Left            =   -70080
         TabIndex        =   11
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         ShowToday       =   0   'False
         StartOfWeek     =   61931522
         CurrentDate     =   39477
      End
      Begin MSComCtl2.MonthView MonthViewDay 
         Height          =   2820
         Left            =   3600
         TabIndex        =   12
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   61931522
         CurrentDate     =   39477
      End
      Begin VB.OptionButton OptionLastYear 
         Caption         =   "Last year"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   2880
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmDoctorPaymentDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CSetPrinter As New cSetDfltPrinter
Const PreSHape = "SHAPE {"
Const Sql = "SELECT tblPatientFacility.*, tblDoctor.DoctorName, tblPatientMainDetails.FirstName, tblStaff.StaffName, tblStaffPayment.* FROM tblStaff RIGHT JOIN (((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblDoctor ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) LEFT JOIN tblStaffPayment ON tblPatientFacility.StaffPayment_ID = tblStaffPayment.StaffPayment_ID) ON tblStaff.Staff_ID = tblStaffPayment.User_ID where "
Const PostSHape = "(((tblPatientFacility.HospitalFacility_ID)=10))}  AS cmmdDoctorPayments COMPUTE cmmdDoctorPayments, ANY(cmmdDoctorPayments.'DoctorName') AS PaidDoctorName, ANY(cmmdDoctorPayments.'StaffName') AS PaidStaffName, ANY(cmmdDoctorPayments.'PaidDate') AS DoctorPaidDate, ANY(cmmdDoctorPayments.'PaidToSTaff') AS PaidOrNot, SUM(cmmdDoctorPayments.'PersonalDue') AS ToPayDoctor, SUM(cmmdDoctorPayments.'StaffPayment') AS PaidAmountToDoctor, ANY(cmmdDoctorPayments.'PaidDate') AS DocPaidDate, ANY(cmmdDoctorPayments.'PaidTime') AS DocPaidTime, ANY(cmmdDoctorPayments.'tblPatientFacility.Staff_ID') AS PaidID, ANY(cmmdDoctorPayments.'AppointmentDate') AS ForAppointmentDate BY 'tblPatientFacility.Staff_ID','AppointmentDate','PaidToSTaff','tblPatientFacility.StaffPayment_ID'"

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnPrint_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    DataReportPatientViceDoctorPayments.Sections("ReportHeader").Controls.Item("lblInstitutionName").Caption = InstitutionName
    DataReportPatientViceDoctorPayments.Sections("ReportHeader").Controls.Item("lblInstitutionAddress").Caption = InstitutionAddress
    DataReportPatientViceDoctorPayments.Show
End Sub


Private Sub Form_Load()
    Call FillSpeciality
    SSTab1.Tab = 0
    OptionToday.Value = True
    OptionAllDoctors.Value = False
    Call CalculateDates
    Call FillGrid
    If UserAuthority <> AuthorityOwner Then
        SSTab1.Enabled = False
    End If
        If UserAuthority <> AuthorityOwner Then
        SSTab1.TabVisible(1) = False
    End If

End Sub

Private Sub FormatGridSpeciality()
    ListSpecialities.Clear
    ListSpecialityIDs.Clear
End Sub

Private Sub FormatGridConsultants()
    ListConsultants.Clear
    ListConsultantIDs.Clear
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


Private Sub CalculateDates()
If OptionToday.Value = True Then
    MonthViewDay.Value = Date
ElseIf OptionYesterday.Value = True Then
    MonthViewDay.Value = Date - 1
ElseIf OptionDayBeforeYesterday.Value = True Then
    MonthViewDay.Value = Date - 2
ElseIf OptionTomorrow.Value = True Then
    MonthViewDay.Value = Date + 1
ElseIf OptionDayAfterTomorrow.Value = True Then
    MonthViewDay.Value = Date + 2
End If

If OptionThisweek.Value = True Then
    Select Case Weekday(Date)
    Case vbMonday:
        MonthViewFrom.Value = Date
        MonthViewTo.Value = Date
    Case vbTuesday:
        MonthViewFrom.Value = Date - 1
        MonthViewTo.Value = Date
    Case vbWednesday:
        MonthViewFrom.Value = Date - 2
        MonthViewTo.Value = Date
    Case vbThursday:
        MonthViewFrom.Value = Date - 3
        MonthViewTo.Value = Date
    Case vbFriday:
        MonthViewFrom.Value = Date - 4
        MonthViewTo.Value = Date
    Case vbSaturday:
        MonthViewFrom.Value = Date - 5
        MonthViewTo.Value = Date
    Case vbSunday:
        MonthViewFrom.Value = Date - 6
        MonthViewTo.Value = Date
    End Select
ElseIf OptionLastWeek.Value = True Then
    Select Case Weekday(Date)
    Case vbMonday:
        MonthViewFrom.Value = Date - 7
        MonthViewTo.Value = Date - 1
    Case vbTuesday:
        MonthViewFrom.Value = Date - 1 - 7
        MonthViewTo.Value = Date - 2
    Case vbWednesday:
        MonthViewFrom.Value = Date - 2 - 7
        MonthViewTo.Value = Date - 3
    Case vbThursday:
        MonthViewFrom.Value = Date - 3 - 7
        MonthViewTo.Value = Date - 4
    Case vbFriday:
        MonthViewFrom.Value = Date - 4 - 7
        MonthViewTo.Value = Date - 5
    Case vbSaturday:
        MonthViewFrom.Value = Date - 5 - 7
        MonthViewTo.Value = Date - 6
    Case vbSunday:
        MonthViewFrom.Value = Date - 6 - 7
        MonthViewTo.Value = Date - 7
    End Select
ElseIf OptionThisMonth.Value = True Then
        MonthViewTo.Value = Date
        MonthViewFrom.Value = DateSerial(Year(Date), Month(Date), 1)
ElseIf OptionLastMonth.Value = True Then
        If Month(Date) = 1 Then
            MonthViewTo.Value = DateSerial(Year(Date) - 1, 12, 31)
            MonthViewFrom.Value = DateSerial(Year(Date) - 1, 12, 1)
        Else
            MonthViewFrom.Value = DateSerial(Year(Date), Month(Date) - 1, 1)
            MonthViewTo.Value = DateSerial(Year(Date), Month(Date), 1)
            MonthViewTo.Value = MonthViewTo.Value - 1
        End If
ElseIf OptionThisYear.Value = True Then
        MonthViewFrom.Value = DateSerial(Year(Date), 1, 1)
        MonthViewTo.Value = Date
ElseIf OptionLastYear.Value = True Then
        MonthViewFrom.Value = DateSerial(Year(Date) - 1, 1, 1)
        MonthViewTo.Value = DateSerial(Year(Date) - 1, 12, 31)
End If

End Sub

Private Sub ListConsultants_Click()
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    Call FillGrid
End Sub

Private Sub ListSpecialities_Click()
    ListSpecialityIDs.ListIndex = ListSpecialities.ListIndex
    ListConsultantIDs.Clear
    ListConsultants.Clear
    If ListSpecialities.Text = "All" Then
        ListAllConsultants
    ElseIf ListSpecialities.Text <> "All" And IsNumeric(ListSpecialityIDs.Text) = True Then
        ListSelectedConsultants
    Else
        FormatGridConsultants
    End If
End Sub

Private Sub MonthViewDay_DateClick(ByVal DateClicked As Date)
    Call FillGrid
End Sub

Private Sub MonthViewFrom_DateClick(ByVal DateClicked As Date)
    Call FillGrid
End Sub

Private Sub MonthViewTo_DateClick(ByVal DateClicked As Date)
    Call FillGrid
End Sub

Private Sub OptionAllDoctors_Click()
If OptionAllDoctors.Value = True Then
    FrameSelectDoctor.Enabled = False
    ListConsultants.Visible = False
    ListSpecialities.Visible = False
Else
    FrameSelectDoctor.Enabled = True
    ListConsultants.Visible = True
    ListSpecialities.Visible = True
End If
    Call FillGrid
End Sub

Private Sub OptionDayAfterTomorrow_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub OptionDayBeforeYesterday_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub OptionLastMonth_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub OptionLastWeek_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub OptionLastYear_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub OptionSelectedDoctors_Click()
If OptionSelectedDoctors.Value = True Then
    FrameSelectDoctor.Enabled = True
    ListConsultants.Visible = True
    ListSpecialities.Visible = True
Else
    FrameSelectDoctor.Enabled = False
    ListConsultants.Visible = False
    ListSpecialities.Visible = False
End If
    Call FillGrid
End Sub

Private Sub OptionThisMonth_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub OptionThisweek_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub OptionThisYear_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub OptionToday_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub OptionTomorrow_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub OptionYesterday_Click()
    Call CalculateDates
    Call FillGrid
End Sub

Private Sub FillGrid()
    If OptionAllDoctors.Value = True Then
    
    Else
        If ListSpecialities.ListIndex < 0 Then Exit Sub
        If ListSpecialityIDs.ListIndex < 0 Then Exit Sub
        If ListConsultants.ListIndex < 0 Then Exit Sub
        If ListConsultantIDs.ListIndex < 0 Then Exit Sub
        If IsNumeric(ListConsultantIDs.Text) = False Then Exit Sub
    End If
    Grid1.Visible = False
    Dim TemStartDate As Date
    Dim TemEndDate As Date
    
    If SSTab1.Tab = 0 Then
        TemStartDate = MonthViewDay.Value
        TemEndDate = MonthViewDay.Value
    Else
        TemStartDate = MonthViewFrom.Value
        TemEndDate = MonthViewTo.Value
    End If

    Set Grid1.DataSource = Nothing
    
    With DataEnvironment1
        If .rscmmdDoctorPayments_Grouping.State = 1 Then .rscmmdDoctorPayments_Grouping.Close
        If OptionAllDoctors.Value = True Then
            .Commands!cmmdDoctorPayments_Grouping.CommandText = PreSHape & Sql & "  (appointmentdate Between  '" & TemStartDate & "' and '" & TemEndDate & "')  and  tblpatientfacility.paidtostaff = 1 and " & PostSHape
        Else
            .Commands!cmmdDoctorPayments_Grouping.CommandText = PreSHape & Sql & "  (appointmentdate Between  '" & TemStartDate & "' and '" & TemEndDate & "')  and  tblpatientfacility.staff_ID = " & ListConsultantIDs.Text & " And tblpatientfacility.paidtostaff = 1 and " & PostSHape
        End If
        .cmmdDoctorPayments_Grouping
    End With
    Grid1.Visible = True
    Set Grid1.DataSource = DataEnvironment1
    
    With Grid1
        .ExpandAll
        .CollapseAll
        '.Refresh
        '.CollapseAll
        .Row = 0
        
        .ColWidth(1, 0) = 2600
        .Col = 1
        .CellAlignment = 4
        .Text = "Doctor"
        
        .ColWidth(2, 0) = 1600
        .Col = 2
        .CellAlignment = 4
        .Text = "Appointment Date"
        
        .ColWidth(3, 0) = 800
        .Col = 3
        .CellAlignment = 4
        .Text = "Paid or not"
        
        .ColWidth(4, 0) = 1600
        .Col = 4
        .CellAlignment = 4
        .Text = "Paid User"
        
        
    End With
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call FillGrid
End Sub

