VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNewAnyDayEndSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Day-End Summeries"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewAnyDayEndSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   12195
   Begin MSComCtl2.MonthView DTPicker1 
      Height          =   2820
      Left            =   120
      TabIndex        =   142
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   159711233
      CurrentDate     =   39534
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Summery"
      TabPicture(0)   =   "frmNewAnyDayEndSummary.frx":038A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "bttnPrintSummery"
      Tab(0).Control(1)=   "lblExpence"
      Tab(0).Control(2)=   "lblIncome"
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(6)=   "lblCashBookings"
      Tab(0).Control(7)=   "lblSettlingCredit"
      Tab(0).Control(8)=   "Label1(2)"
      Tab(0).Control(9)=   "Label1(3)"
      Tab(0).Control(10)=   "Label1(4)"
      Tab(0).Control(11)=   "Label1(5)"
      Tab(0).Control(12)=   "Label1(6)"
      Tab(0).Control(13)=   "Label1(7)"
      Tab(0).Control(14)=   "Label1(8)"
      Tab(0).Control(15)=   "Label1(9)"
      Tab(0).Control(16)=   "lblCashRepayments"
      Tab(0).Control(17)=   "lblDoctorPayments"
      Tab(0).Control(18)=   "lblAgentCashPayments"
      Tab(0).Control(19)=   "Label1(15)"
      Tab(0).Control(20)=   "lblAgentBoolings"
      Tab(0).Control(21)=   "lblAgentRepayments"
      Tab(0).Control(22)=   "lblNetCash"
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Cash Bookings"
      TabPicture(1)   =   "frmNewAnyDayEndSummary.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "bttnCashSummery"
      Tab(1).Control(1)=   "bttnPrintCashRepayments"
      Tab(1).Control(2)=   "bttnPrintCreditSettling"
      Tab(1).Control(3)=   "bttnPrintCashBookings"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Agent Bookings"
      TabPicture(2)   =   "frmNewAnyDayEndSummary.frx":03C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "ButtonEx2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "ButtonEx1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "bttnPrintAgentPayments"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "bttnPrintAgentRepayments"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame6"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame7"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Payments"
      TabPicture(3)   =   "frmNewAnyDayEndSummary.frx":03DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(1)=   "Frame3"
      Tab(3).Control(2)=   "bttnDoctorPaymentsDoneToday"
      Tab(3).Control(3)=   "bttnDoctorPaymentsDoneForTodayAppointments"
      Tab(3).Control(4)=   "bttnDoctorPaymentsToDoForTodayAppointments"
      Tab(3).Control(5)=   "bttnDoctorPaymentsForTodayAppointments"
      Tab(3).Control(6)=   "Frame5"
      Tab(3).ControlCount=   7
      Begin VB.Frame Frame7 
         Caption         =   "Doctor Fee from Agent Bookings"
         Height          =   4455
         Left            =   120
         TabIndex        =   123
         Top             =   3120
         Width           =   8535
         Begin VB.ListBox ListDoctorAgent 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1710
            Left            =   120
            TabIndex        =   124
            Top             =   2520
            Width           =   8055
         End
         Begin VB.Label lblDocAgentFeeAgentRepaymentsO 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   140
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label lblDocAgentFeeCashRepaymentsO 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   139
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label lblDocFeeBookingsAgentO 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   138
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblDocFeeAgentO 
            Alignment       =   1  'Right Justify
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
            Left            =   5640
            TabIndex        =   137
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label51 
            Caption         =   "Less - Repayments to agent"
            Height          =   255
            Left            =   840
            TabIndex        =   136
            Top             =   2160
            Width           =   3255
         End
         Begin VB.Label Label45 
            Caption         =   "Less - Cash Repayments"
            Height          =   255
            Left            =   840
            TabIndex        =   135
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label17 
            Caption         =   "Bookings"
            Height          =   255
            Left            =   840
            TabIndex        =   134
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label Label60 
            Caption         =   "Doctor Fee for todays' appointments"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label Label59 
            Caption         =   "Doctor Fee For appointments of other days"
            Height          =   255
            Left            =   360
            TabIndex        =   132
            Top             =   1440
            Width           =   5175
         End
         Begin VB.Label lblDocFeeAgentT 
            Alignment       =   1  'Right Justify
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
            Left            =   5640
            TabIndex        =   131
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label56 
            Caption         =   "Bookings"
            Height          =   255
            Left            =   720
            TabIndex        =   130
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblDocFeeBookingsAgentT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   129
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label54 
            Caption         =   "Less - Cash Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   128
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblDocAgentFeeCashRepaymentsT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   127
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblDocAgentFeeAgentRepaymentsT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   126
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label37 
            Caption         =   "Less - Repayments to agent"
            Height          =   255
            Left            =   720
            TabIndex        =   125
            Top             =   960
            Width           =   3255
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Doctor Payments"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   89
         Top             =   5400
         Width           =   8175
         Begin VB.Label Label33 
            Caption         =   "Payments done for today appointments"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   720
            Width           =   4095
         End
         Begin VB.Label lblDocPaymentsForTodaysApp 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   96
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblTotalDocPayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   95
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label21 
            Caption         =   "Total Doctor Payments for todays appointments"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1440
            Width           =   5655
         End
         Begin VB.Label Label20 
            Caption         =   "Payments Done today for Doctors"
            Height          =   255
            Left            =   240
            TabIndex        =   93
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label Label19 
            Caption         =   "Payments to pay for today appointments"
            Height          =   255
            Left            =   240
            TabIndex        =   92
            Top             =   960
            Width           =   4095
         End
         Begin VB.Label lblTodayDoctorPaymentsMade 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   91
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblPaymentsToDo 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   90
            Top             =   960
            Width           =   2175
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Todays' Total Agent Bookings"
         Height          =   2655
         Left            =   120
         TabIndex        =   64
         Top             =   480
         Width           =   8535
         Begin VB.Label Label32 
            Caption         =   "Net agent bookings Value"
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   2280
            Width           =   4335
         End
         Begin VB.Label lblNetagentBooking 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            TabIndex        =   81
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label31 
            Caption         =   "Less -Agent Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   67
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label44 
            Caption         =   "Doctor Fee From Agent Bookings"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label43 
            Caption         =   "Hospital Fee From Agent Bookings"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   1200
            Width           =   4335
         End
         Begin VB.Label lblAgentDoctorFee 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            TabIndex        =   78
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblAgentHospitalFee 
            Alignment       =   1  'Right Justify
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
            Left            =   5040
            TabIndex        =   77
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label40 
            Caption         =   "Agent Bookings"
            Height          =   255
            Left            =   720
            TabIndex        =   76
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblDoctorAgentBookings 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   75
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label38 
            Caption         =   "Less -Cash Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   74
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblDoctorAgentCashRepayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   73
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label36 
            Caption         =   "Agent Bookings"
            Height          =   255
            Left            =   720
            TabIndex        =   72
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblAgentHospitalBookings 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   71
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label34 
            Caption         =   "Less - Cash Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   70
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblHospitalAgentCashRepayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   69
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblDoctorAgentAgentRepayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3360
            TabIndex        =   68
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "Less - Agent Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   65
            Top             =   1920
            Width           =   2775
         End
         Begin VB.Label lblHospitalAgentAgentRepayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   3840
            TabIndex        =   66
            Top             =   1920
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Total Repayments"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   57
         Top             =   360
         Width           =   8175
         Begin VB.Label lblRepaidToAgent 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   63
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblRepaidToPatient 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   62
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label14 
            Caption         =   "Repaied to agent"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label15 
            Caption         =   "Repayed to Patient"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Total Repayments"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lblTotalRepayments 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            TabIndex        =   58
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Doctor Fee from Cash"
         Height          =   4455
         Left            =   -74880
         TabIndex        =   25
         Top             =   3360
         Width           =   8535
         Begin VB.ListBox ListDOctorCash 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1710
            Left            =   120
            TabIndex        =   28
            Top             =   2520
            Width           =   8055
         End
         Begin VB.Label Label30 
            Caption         =   "Less - Cash Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   56
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label lblOtherDaysDoctorCashRepay 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   55
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label Label27 
            Caption         =   "Less - Cash Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   54
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lblTodaysDoctorCashRepay 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   53
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label lblOtherDaysDoctorCashSC 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   48
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label28 
            Caption         =   "Settling Credit Cash"
            Height          =   255
            Left            =   720
            TabIndex        =   47
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label lblOtherdaysDoctorCashDC 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   46
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label Label26 
            Caption         =   "Direct Cash"
            Height          =   255
            Left            =   720
            TabIndex        =   45
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblTodaysDoctorCashSC 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   44
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label24 
            Caption         =   "Settling Credit Cash"
            Height          =   255
            Left            =   720
            TabIndex        =   43
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblTOdaysDoctorCashDC 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   42
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label22 
            Caption         =   "Direct Cash"
            Height          =   255
            Left            =   720
            TabIndex        =   41
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblOtherDaysDoctorCash 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            TabIndex        =   32
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblTodaysDoctorCash 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            TabIndex        =   31
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label10 
            Caption         =   "Doctor Fee For appointments of other days"
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   1440
            Width           =   5175
         End
         Begin VB.Label Label8 
            Caption         =   "Doctor Fee for todays' appointments"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   5535
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Todays' Total Cash Collection"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   8535
         Begin VB.Label Label23 
            Caption         =   "Less - Cash Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   52
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label lblHospitalCashRepayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   51
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label13 
            Caption         =   "Less - Cash Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   50
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblDoctorCashRepayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   49
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblHospitalCashSC 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   40
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label Label16 
            Caption         =   "Settling Credit Cash"
            Height          =   255
            Left            =   720
            TabIndex        =   39
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label lblHospitalCashDC 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   38
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label12 
            Caption         =   "Direct Cash"
            Height          =   255
            Left            =   720
            TabIndex        =   37
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblDoctorCashSC 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   36
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label7 
            Caption         =   "Settling Credit Cash"
            Height          =   255
            Left            =   720
            TabIndex        =   35
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lblDoctorCashDC 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   34
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label5 
            Caption         =   "Direct Cash"
            Height          =   255
            Left            =   720
            TabIndex        =   33
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label lblHospitalCash 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            TabIndex        =   30
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblDoctorCash 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            TabIndex        =   29
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label9 
            Caption         =   "Hospital Fee From Cash"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Doctor Fee From Cash"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   2175
         End
      End
      Begin btButtonEx.ButtonEx bttnPrintSummery 
         Height          =   495
         Left            =   -68160
         TabIndex        =   83
         Top             =   8040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Summery"
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
      Begin btButtonEx.ButtonEx bttnPrintCashBookings 
         Height          =   495
         Left            =   -72720
         TabIndex        =   84
         Top             =   7920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Cash Bookings"
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
      Begin btButtonEx.ButtonEx bttnPrintCreditSettling 
         Height          =   495
         Left            =   -70560
         TabIndex        =   85
         Top             =   7920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Credit Settling"
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
      Begin btButtonEx.ButtonEx bttnPrintCashRepayments 
         Height          =   495
         Left            =   -68520
         TabIndex        =   86
         Top             =   7920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Cash Repayments"
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
      Begin btButtonEx.ButtonEx bttnPrintAgentRepayments 
         Height          =   495
         Left            =   6600
         TabIndex        =   87
         Top             =   7920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Agent Repayments"
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
      Begin btButtonEx.ButtonEx bttnPrintAgentPayments 
         Height          =   495
         Left            =   4440
         TabIndex        =   88
         Top             =   7920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Agent Payments"
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
      Begin btButtonEx.ButtonEx bttnDoctorPaymentsDoneToday 
         Height          =   375
         Left            =   -70560
         TabIndex        =   98
         Top             =   7320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Doctor Payments done today"
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
      Begin btButtonEx.ButtonEx bttnDoctorPaymentsDoneForTodayAppointments 
         Height          =   495
         Left            =   -70560
         TabIndex        =   99
         Top             =   7800
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Doctor payments done for todays appointments"
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
      Begin btButtonEx.ButtonEx bttnDoctorPaymentsToDoForTodayAppointments 
         Height          =   495
         Left            =   -74760
         TabIndex        =   100
         Top             =   7800
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Doctor Payments to do for todays appointments"
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
      Begin btButtonEx.ButtonEx bttnDoctorPaymentsForTodayAppointments 
         Height          =   375
         Left            =   -74760
         TabIndex        =   101
         Top             =   7320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Payments for todays appointments"
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
      Begin VB.Frame Frame5 
         Caption         =   "Todays Appointments"
         Height          =   3855
         Left            =   -74760
         TabIndex        =   104
         Top             =   1560
         Width           =   8175
         Begin VB.Line Line4 
            X1              =   7440
            X2              =   240
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line3 
            X1              =   7440
            X2              =   240
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Line Line2 
            X1              =   7440
            X2              =   240
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line1 
            X1              =   7440
            X2              =   240
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label lblToday 
            Alignment       =   1  'Right Justify
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
            Left            =   5160
            TabIndex        =   122
            Top             =   3480
            Width           =   2175
         End
         Begin VB.Label Label52 
            Caption         =   "Total Fee"
            Height          =   255
            Left            =   240
            TabIndex        =   121
            Top             =   3480
            Width           =   3495
         End
         Begin VB.Label lblTOdayHos 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   120
            Top             =   3000
            Width           =   2175
         End
         Begin VB.Label Label50 
            Caption         =   "Total Hospital Fee"
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   3000
            Width           =   3495
         End
         Begin VB.Label lblTodayHosAgent 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   118
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label Label48 
            Caption         =   "Hospital Fee For Agent Bookings"
            Height          =   255
            Left            =   240
            TabIndex        =   117
            Top             =   2640
            Width           =   3495
         End
         Begin VB.Label Label47 
            Caption         =   "Hospital Fee For Cash Bookings"
            Height          =   255
            Left            =   240
            TabIndex        =   116
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Label46 
            Caption         =   "Hospital Fee For Credit Bookings"
            Height          =   255
            Left            =   240
            TabIndex        =   115
            Top             =   2280
            Width           =   3975
         End
         Begin VB.Label lblTodayHosCash 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   114
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label lblTodayHosCredit 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   113
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label lblTodayDoc 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   112
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label39 
            Caption         =   "Total Doctor Fee"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   1440
            Width           =   3495
         End
         Begin VB.Label lblTodayDocAgent 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   110
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label35 
            Caption         =   "Doctor Fee For Agent Bookings"
            Height          =   255
            Left            =   240
            TabIndex        =   109
            Top             =   1080
            Width           =   3495
         End
         Begin VB.Label Label29 
            Caption         =   "Doctor Fee For Cash Bookings"
            Height          =   255
            Left            =   240
            TabIndex        =   108
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label25 
            Caption         =   "Doctor Fee For Credit Bookings"
            Height          =   255
            Left            =   240
            TabIndex        =   107
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label lblTodayDocCash 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   106
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblTodayDocCredit 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   5160
            TabIndex        =   105
            Top             =   720
            Width           =   2175
         End
      End
      Begin btButtonEx.ButtonEx bttnCashSummery 
         Height          =   495
         Left            =   -74760
         TabIndex        =   141
         Top             =   7920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Cash Summery"
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
      Begin btButtonEx.ButtonEx ButtonEx1 
         Height          =   495
         Left            =   2280
         TabIndex        =   145
         Top             =   7920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Agent Bookings"
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
      Begin btButtonEx.ButtonEx ButtonEx2 
         Height          =   495
         Left            =   120
         TabIndex        =   146
         Top             =   7920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Agent Summery"
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
      Begin VB.Label lblExpence 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -67800
         TabIndex        =   103
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblIncome 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -67800
         TabIndex        =   102
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cash Income"
         Height          =   255
         Index           =   0
         Left            =   -74160
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Non-Cash Income"
         Height          =   255
         Left            =   -74280
         TabIndex        =   20
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Agent Repayments"
         Height          =   255
         Left            =   -73800
         TabIndex        =   19
         Top             =   6480
         Width           =   2055
      End
      Begin VB.Label lblCashBookings 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -69600
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblSettlingCredit 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -69600
         TabIndex        =   17
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Non cash Expences"
         Height          =   255
         Index           =   2
         Left            =   -74280
         TabIndex        =   16
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Cash Expences"
         Height          =   255
         Index           =   3
         Left            =   -74160
         TabIndex        =   15
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Settling Credit Bookings"
         Height          =   255
         Index           =   4
         Left            =   -73680
         TabIndex        =   14
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Agent Cash Payments"
         Height          =   255
         Index           =   5
         Left            =   -73680
         TabIndex        =   13
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Cash Bookings"
         Height          =   255
         Index           =   6
         Left            =   -73680
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Doctor Payments"
         Height          =   255
         Index           =   7
         Left            =   -73680
         TabIndex        =   11
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Agent Bookings"
         Height          =   255
         Index           =   8
         Left            =   -73800
         TabIndex        =   10
         Top             =   5400
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Cash Repayments"
         Height          =   255
         Index           =   9
         Left            =   -73680
         TabIndex        =   9
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblCashRepayments 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -69600
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblDoctorPayments 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -69600
         TabIndex        =   7
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblAgentCashPayments 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -69600
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Net Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   -74280
         TabIndex        =   5
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lblAgentBoolings 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -69600
         TabIndex        =   4
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label lblAgentRepayments 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   -69600
         TabIndex        =   3
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label lblNetCash 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69360
         TabIndex        =   2
         Top             =   4320
         Width           =   2775
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   10320
      TabIndex        =   0
      Top             =   8880
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
   Begin VB.Label Label41 
      Caption         =   "Select Date"
      Height          =   255
      Left            =   120
      TabIndex        =   144
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   495
      Left            =   6480
      TabIndex        =   143
      Top             =   4800
      Width           =   1215
   End
End
Attribute VB_Name = "frmNewAnyDayEndSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CSetPrinter As New cSetDfltPrinter

Private Sub bttnCashSummery_Click()
With dtrTemCashBookings
    If HospitalDetails = True Then
        .Sections("Section4").Controls.Item("lblName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("lblAddress").Caption = InstitutionAddress
    End If
        .Sections("Section4").Controls.Item("lblTopic").Caption = "Cash Bookings"
        .Sections("Section4").Controls.Item("lblSubTopic").Caption = Format(DTPicker1.Value, DefaultLongDate)
        
        .Sections("Section4").Controls.Item("lbl1").Caption = Frame1.Caption
        .Sections("Section4").Controls.Item("lbl2").Caption = Label4.Caption
        .Sections("Section4").Controls.Item("lbl3").Caption = Label5.Caption
        .Sections("Section4").Controls.Item("lbl4").Caption = Label7.Caption
        .Sections("Section4").Controls.Item("lbl5").Caption = Label13.Caption
        .Sections("Section4").Controls.Item("lbl6").Caption = Label9.Caption
        .Sections("Section4").Controls.Item("lbl7").Caption = Label12.Caption
        .Sections("Section4").Controls.Item("lbl8").Caption = Label16.Caption
        .Sections("Section4").Controls.Item("lbl9").Caption = Label23.Caption
        .Sections("Section4").Controls.Item("lbl10").Caption = Frame2.Caption
        .Sections("Section4").Controls.Item("lbl11").Caption = Label8.Caption
        .Sections("Section4").Controls.Item("lbl12").Caption = Label22.Caption
        .Sections("Section4").Controls.Item("lbl13").Caption = Label24.Caption
        .Sections("Section4").Controls.Item("lbl14").Caption = Label27.Caption
        .Sections("Section4").Controls.Item("lbl15").Caption = Label10.Caption
        .Sections("Section4").Controls.Item("lbl16").Caption = Label26.Caption
        .Sections("Section4").Controls.Item("lbl17").Caption = Label28.Caption
        .Sections("Section4").Controls.Item("lbl18").Caption = Label30.Caption
        
         .Sections("Section4").Controls.Item("val1").Caption = lblDoctorCash.Caption
        .Sections("Section4").Controls.Item("val2").Caption = lblDoctorCashDC.Caption
        .Sections("Section4").Controls.Item("val3").Caption = lblDoctorCashSC.Caption
        .Sections("Section4").Controls.Item("val4").Caption = lblDoctorCashRepayments.Caption
        .Sections("Section4").Controls.Item("val5").Caption = lblHospitalCash.Caption
        .Sections("Section4").Controls.Item("val6").Caption = lblHospitalCashDC.Caption
        .Sections("Section4").Controls.Item("val7").Caption = lblHospitalCashSC.Caption
        .Sections("Section4").Controls.Item("val8").Caption = lblHospitalCashRepayments.Caption
        .Sections("Section4").Controls.Item("val9").Caption = lblTodaysDoctorCash.Caption
        .Sections("Section4").Controls.Item("val10").Caption = lblTOdaysDoctorCashDC.Caption
        .Sections("Section4").Controls.Item("val11").Caption = lblTodaysDoctorCashSC.Caption
        .Sections("Section4").Controls.Item("val12").Caption = lblTodaysDoctorCashRepay.Caption
        .Sections("Section4").Controls.Item("val13").Caption = lblOtherDaysDoctorCash.Caption
        .Sections("Section4").Controls.Item("val14").Caption = lblOtherdaysDoctorCashDC.Caption
        .Sections("Section4").Controls.Item("val15").Caption = lblOtherDaysDoctorCashSC.Caption
        .Sections("Section4").Controls.Item("val16").Caption = lblOtherDaysDoctorCashRepay.Caption
       
        
        .Show
End With
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub


Private Sub bttnDoctorPaymentsDoneForTodayAppointments_Click()
    Const PreSHape = "SHAPE {"
    Const Sql = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.* FROM tblPatientFacility LEFT JOIN (tblTitle RIGHT JOIN tblDoctor ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID "
    Dim SqlWhere  As String
    SqlWhere = " WHERE (((tblPatientFacility.FullyPaid)=1) AND  ((tblPatientFacility.PaidToSTaff)=1) AND ((tblPatientFacility.AppointmentDate)='" & DTPicker1.Value & "') "
    If PayToDoctor = True Then
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

        SqlWhere = SqlWhere & " AND ((tblPatientFacility.PatientAbsent)=0)"
    End If
    SqlWhere = SqlWhere & ")"
    Const PostSHape = "}  AS DocPayments COMPUTE DocPayments, SUM(DocPayments.'PersonalDue') AS DocFee, ANY(DocPayments.'Title') AS DocTitle BY 'PaidToSTaff','DoctorName'"
    With DataEnvironment1
        If .rsDocPayments_Grouping.State = 1 Then .rsDocPayments_Grouping.Close
        .Commands!DocPayments_Grouping.CommandText = PreSHape & Sql & SqlWhere & PostSHape
        .DocPayments_Grouping
    End With
    With dtrDocPayments
        If HospitalDetails = True Then
            .Sections.Item("ReportHeader").Controls.Item("lblInstitutionName").Caption = InstitutionName
            .Sections.Item("ReportHeader").Controls.Item("lblInstitutionAddress").Caption = InstitutionAddress
        End If
        .Sections.Item("ReportHeader").Controls.Item("lblreport").Caption = "Doctor Payments For All Bookings"
        .Sections.Item("ReportHeader").Controls.Item("lblreport").Caption = Format(DTPicker1.Value, DefaultLongDate)
        .Sections.Item("PageFooter").Controls.Item("lblAd").Caption = LongAd
        .Show
    End With
End Sub

Private Sub bttnDoctorPaymentsDoneToday_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    With dtrTodayDocPayments
        .DataMember = Empty
    End With
    With DataEnvironment1
        If .rsTodayDoctorPayments.State = 1 Then .rsTodayDoctorPayments.Close
        .rsTodayDoctorPayments.Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblStaff.StaffName, tblStaffPayment.* FROM tblStaff RIGHT JOIN (tblTitle RIGHT JOIN (tblDoctor RIGHT JOIN tblStaffPayment ON tblDoctor.Doctor_ID = tblStaffPayment.Staff_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID) ON tblStaff.Staff_ID = tblStaffPayment.User_ID where paiddate = '" & DTPicker1.Value & "' order by StaffPayment_ID"
        .rsTodayDoctorPayments.Open
    End With
    With dtrTodayDocPayments
        .DataMember = "TodayDoctorPayments"
        If HospitalDetails = True Then
            .Sections("section4").Controls("lblinstitutionname").Caption = InstitutionName
            .Sections("section4").Controls("lblinstitutionaddress").Caption = InstitutionAddress
            .Sections("section4").Controls("lblReport").Caption = "Doctor Payments done Today"
            .Sections("section5").Controls("lblad").Caption = LongAd
            .Sections("section4").Controls("lblReportsub").Caption = Format(DTPicker1.Value, DefaultLongDate)
        Else
            .Sections("section4").Controls("lblinstitutionname").Caption = Empty
            .Sections("section4").Controls("lblinstitutionaddress").Caption = Empty
            .Sections("section4").Controls("lblReport").Caption = "Doctor Payments done Today"
            .Sections("section4").Controls("lblReportsub").Caption = Format(DTPicker1.Value, DefaultLongDate)
            .Sections("section5").Controls("lblad").Caption = LongAd
        End If
        .Show
    End With


End Sub

Private Sub bttnDoctorPaymentsForTodayAppointments_Click()
    Const PreSHape = "SHAPE {"
    Const Sql = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.* FROM tblPatientFacility LEFT JOIN (tblTitle RIGHT JOIN tblDoctor ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID "
    Dim SqlWhere  As String
    SqlWhere = " WHERE (((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.AppointmentDate)='" & DTPicker1.Value & "') "
    ' ((tblPatientFacility.PaidToSTaff)=1) AND
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    If PayToDoctor = True Then
        SqlWhere = SqlWhere & " AND ((tblPatientFacility.PatientAbsent)=0)"
    End If
    SqlWhere = SqlWhere & ")"
    Const PostSHape = "}  AS DocPayments COMPUTE DocPayments, SUM(DocPayments.'PersonalDue') AS DocFee, ANY(DocPayments.'Title') AS DocTitle BY 'PaidToSTaff','DoctorName'"
    With DataEnvironment1
        If .rsDocPayments_Grouping.State = 1 Then .rsDocPayments_Grouping.Close
        .Commands!DocPayments_Grouping.CommandText = PreSHape & Sql & SqlWhere & PostSHape
        .DocPayments_Grouping
    End With
    With dtrDocPayments
        If HospitalDetails = True Then
            .Sections.Item("ReportHeader").Controls.Item("lblInstitutionName").Caption = InstitutionName
            .Sections.Item("ReportHeader").Controls.Item("lblInstitutionAddress").Caption = InstitutionAddress
        End If
        .Sections.Item("ReportHeader").Controls.Item("lblreport").Caption = "Doctor Payments For All Bookings"
        .Sections.Item("ReportHeader").Controls.Item("lblreport").Caption = Format(DTPicker1.Value, DefaultLongDate)
        .Sections.Item("PageFooter").Controls.Item("lblAd").Caption = LongAd
        .Show
    End With
End Sub

Private Sub bttnDoctorPaymentsToDoForTodayAppointments_Click()
    Const PreSHape = "SHAPE {"
    Const Sql = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.* FROM tblPatientFacility LEFT JOIN (tblTitle RIGHT JOIN tblDoctor ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID "
    Dim SqlWhere  As String
    SqlWhere = " WHERE (((tblPatientFacility.FullyPaid)=1) AND  ((tblPatientFacility.PaidToSTaff)=0)AND ((tblPatientFacility.AppointmentDate)='" & DTPicker1.Value & "') "
    If PayToDoctor = True Then
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

        SqlWhere = SqlWhere & " AND ((tblPatientFacility.PatientAbsent)=0)"
    End If
    SqlWhere = SqlWhere & ")"
    Const PostSHape = "}  AS DocPayments COMPUTE DocPayments, SUM(DocPayments.'PersonalDue') AS DocFee, ANY(DocPayments.'Title') AS DocTitle BY 'PaidToSTaff','DoctorName'"
    With DataEnvironment1
        If .rsDocPayments_Grouping.State = 1 Then .rsDocPayments_Grouping.Close
        .Commands!DocPayments_Grouping.CommandText = PreSHape & Sql & SqlWhere & PostSHape
        .DocPayments_Grouping
    End With
    With dtrDocPayments
        If HospitalDetails = True Then
            .Sections.Item("ReportHeader").Controls.Item("lblInstitutionName").Caption = InstitutionName
            .Sections.Item("ReportHeader").Controls.Item("lblInstitutionAddress").Caption = InstitutionAddress
        End If
        .Sections.Item("ReportHeader").Controls.Item("lblreport").Caption = "Doctor Payments For All Bookings"
        .Sections.Item("ReportHeader").Controls.Item("lblreport").Caption = Format(DTPicker1.Value, DefaultLongDate)
        .Sections.Item("PageFooter").Controls.Item("lblAd").Caption = LongAd
        .Show
    End With
End Sub

Private Sub bttnPrintAgentPayments_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1.rssqlTem3
    If .State = 1 Then .Close
'    Select Case SSTab1.Tab
'    Case 0
     .Source = "Select tblAgentCashSettle.*, tblInstitutions.* fROM tblAgentCashSettle Left Join tblInstitutions On tblAgentCashSettle.Institution_Id = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.SettledDate = '" & DTPicker1.Value & "')"
    .Open
'    Case 1
'    .Source = "Select tblAgentCashSettle.*, tblInstitutions.* fROM tblAgentCashSettle Left Join tblInstitutions On tblAgentCashSettle.Institution_Id = tblInstitutions.Institution_ID Where (tblAgentCashSettle.SettledDate = '" & DTPicker1 & "')"
'    .Open
'    Case 2
'    .Source = "SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID Where (tblAgentCashSettle.SettledDate between '" & DTPicker2 & "' and '" & DTPicker3 & "')"
'    .Open
'    End Select
'    If .RecordCount = 0 Then A = MsgBox("No Cash receive to view", vbCritical + vbOKOnly, "No Data"): Exit Sub
    End With
    With dtrAgentCashReceive
        If HospitalDetails = True Then
        .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        Else
        .Sections("Section4").Controls.Item("RptName").Caption = Empty
        .Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        End If
'         Select Case SSTab1.Tab
'         Case 0
         .Sections("Section2").Controls.Item("rptFromdate").Caption = Format(DTPicker1.Value, DefaultLongDate)
         .Sections("Section2").Controls.Item("rptTodate").Caption = Format(DTPicker1.Value, DefaultLongDate)
'         Case 1
'         .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
'         .Sections("Section2").Controls.Item("rptTodate").Caption = DTPicker1.Value
'         Case 2
'         .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
'         .Sections("Section2").Controls.Item("rptTodate").Caption = DTPicker3.Value
'         End Select
        .Sections("Section2").Controls.Item("rptlHeding1").Caption = ""
        .Sections("Section2").Controls.Item("RptCashierName").Caption = ""
        Set .DataSource = DataEnvironment1.rssqlTem3
        .Show
    End With
End Sub

Private Sub bttnPrintAgentRepayments_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

dtrCashRefunds.DataMember = Empty
With DataEnvironment1.rsRefunds
    If .State = 1 Then .Close
    .Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.*, tblPatientMainDetails.FirstName, tblPatientFacility.SettleCashDate, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode FROM tblInstitutions RIGHT JOIN (tblPatientMainDetails RIGHT JOIN (tblPatientFacility LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID Where (((tblPatientFacility.HospitalFacility_ID) = 10) And ((tblPatientFacility.RefundToAgent) = 1) And ((tblPatientFacility.Repaydate) = '" & DTPicker1.Value & "') ) ORDER BY tblPatientFacility.PatientFacility_ID"
    .Open
End With
With dtrCashRefunds
    .DataMember = "refunds"
    If HospitalDetails = True Then
        .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = InstitutionName
        .Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
    End If
    .Sections.Item("Section4").Controls.Item("lblreport").Caption = "All Agent Repayments"
    .Sections.Item("Section4").Controls.Item("lblreportsub").Caption = Format(DTPicker1.Value, DefaultLongDate)
    .Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
    .Show
End With
End Sub

Private Sub bttnPrintCashBookings_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

dtrCashBookings.DataMember = Empty
With DataEnvironment1.rsCashBookings
    If .State = 1 Then .Close
    .Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.*, tblPatientMainDetails.FirstName FROM tblPatientMainDetails RIGHT JOIN (tblPatientFacility LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash') AND ((tblPatientFacility.bookingdate)='" & DTPicker1.Value & "') ) ORDER BY tblPatientFacility.PatientFacility_ID "
    .Open
End With
With dtrCashBookings
    .DataMember = "CashBookings"
    If HospitalDetails = True Then
        .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = InstitutionName
        .Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
    End If
    .Sections.Item("Section4").Controls.Item("lblreport").Caption = "All Cash Bookings"
    .Sections.Item("Section4").Controls.Item("lblreportsub").Caption = Format(DTPicker1.Value, DefaultLongDate)
    .Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
    .Show
End With
End Sub

Private Sub bttnPrintCashRepayments_Click()
dtrCashRefunds.DataMember = Empty
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1.rsRefunds
    If .State = 1 Then .Close
    .Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.*, tblPatientMainDetails.FirstName, tblPatientFacility.SettleCashDate, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode FROM tblInstitutions RIGHT JOIN (tblPatientMainDetails RIGHT JOIN (tblPatientFacility LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID Where (((tblPatientFacility.HospitalFacility_ID) = 10) And ((tblPatientFacility.RefundToPatient) = 1) And ((tblPatientFacility.Repaydate) = '" & DTPicker1.Value & "') ) ORDER BY tblPatientFacility.PatientFacility_ID"
    .Open
End With
With dtrCashRefunds
    .DataMember = "refunds"
    If HospitalDetails = True Then
        .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = InstitutionName
        .Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
    End If
    .Sections.Item("Section4").Controls.Item("lblreport").Caption = "All Cash Repayments"
    .Sections.Item("Section4").Controls.Item("lblreportsub").Caption = Format(DTPicker1.Value, DefaultLongDate)
    .Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
    .Show
End With

End Sub

Private Sub bttnPrintCreditSettling_Click()
dtrCreditBookings.DataMember = Empty
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1.rsCashBookings
    If .State = 1 Then .Close
    .Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.*, tblPatientMainDetails.FirstName FROM tblPatientMainDetails RIGHT JOIN (tblPatientFacility LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Credit') AND ((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.SettleCashDate)='" & DTPicker1.Value & "') ) ORDER BY tblPatientFacility.PatientFacility_ID "
    .Open
End With
With dtrCreditBookings
    .DataMember = "CashBookings"
    If HospitalDetails = True Then
        .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = InstitutionName
        .Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
    End If
    .Sections.Item("Section4").Controls.Item("lblreport").Caption = "All Cash For Credit Bookings"
    .Sections.Item("Section4").Controls.Item("lblreportsub").Caption = Format(DTPicker1.Value, DefaultLongDate)
    .Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
    .Show
End With
End Sub

Private Sub bttnPrintSummery_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1.rssqlTemSu1
        If .State = 1 Then .Close
        .Open "Select * From tblTem"
        Set dtrSummeryReport.DataSource = DataEnvironment1.rssqlTemSu1
    End With
    With dtrSummeryReport
        If HospitalDetails = True Then
            .Sections("Section4").Controls.Item("lblinstitutionname").Caption = InstitutionName
            .Sections("Section4").Controls.Item("lblinstitutionaddress").Caption = InstitutionAddress
        End If
        .Sections("Section4").Controls.Item("lblreport").Caption = "Day End Summery"
        .Sections("Section4").Controls.Item("lblreportsub").Caption = Format(DTPicker1.Value, DefaultLongDate)
        .Sections("Section2").Controls.Item("lblcash").Caption = lblCashBookings.Caption
        .Sections("Section2").Controls.Item("lblcredit").Caption = lblSettlingCredit.Caption
        .Sections("Section2").Controls.Item("lblagent").Caption = lblAgentCashPayments.Caption
        .Sections("Section2").Controls.Item("lbltotalcash").Caption = lblIncome.Caption
        .Sections("Section2").Controls.Item("lblrepayments").Caption = lblCashRepayments.Caption
        .Sections("Section2").Controls.Item("lbldoctorpayments").Caption = lblDoctorPayments.Caption
        .Sections("Section2").Controls.Item("lbltotalpayments").Caption = lblExpence.Caption
        .Sections("Section2").Controls.Item("lblnetcash").Caption = lblNetCash.Caption
        .Show
    End With
End Sub


Private Sub ButtonEx1_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

dtrAgentBookings.DataMember = Empty
With DataEnvironment1.rsCashBookings
    If .State = 1 Then .Close
    .Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.*, tblPatientMainDetails.FirstName FROM tblPatientMainDetails RIGHT JOIN (tblPatientFacility LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent') AND ((tblPatientFacility.bookingdate)='" & DTPicker1.Value & "') ) ORDER BY tblPatientFacility.PatientFacility_ID "
    .Open
End With
With dtrAgentBookings
    .DataMember = "CashBookings"
    If HospitalDetails = True Then
        .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = InstitutionName
        .Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
    End If
    .Sections.Item("Section4").Controls.Item("lblreport").Caption = "All Agent Bookings"
    .Sections.Item("Section4").Controls.Item("lblreportsub").Caption = Format(DTPicker1.Value, DefaultLongDate)
    .Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
    .Show
End With

End Sub

Private Sub ButtonEx2_Click()
With dtrTemCashBookings
    If HospitalDetails = True Then
        .Sections("Section4").Controls.Item("lblName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("lblAddress").Caption = InstitutionAddress
    End If
        .Sections("Section4").Controls.Item("lblTopic").Caption = "Agent Bookings"
        .Sections("Section4").Controls.Item("lblSubTopic").Caption = Format(DTPicker1.Value, DefaultLongDate)
        
        .Sections("Section4").Controls.Item("lbl1").Caption = Frame6.Caption
        .Sections("Section4").Controls.Item("lbl2").Caption = Label44.Caption
        .Sections("Section4").Controls.Item("lbl3").Caption = Label40.Caption
        .Sections("Section4").Controls.Item("lbl4").Caption = Label38.Caption
        .Sections("Section4").Controls.Item("lbl5").Caption = Label31.Caption
        .Sections("Section4").Controls.Item("lbl6").Caption = Label43.Caption
        .Sections("Section4").Controls.Item("lbl7").Caption = Label36.Caption
        .Sections("Section4").Controls.Item("lbl8").Caption = Label34.Caption
        .Sections("Section4").Controls.Item("lbl9").Caption = Label11.Caption
        .Sections("Section4").Controls.Item("lbl10").Caption = Frame7.Caption
        .Sections("Section4").Controls.Item("lbl11").Caption = Label60.Caption
        .Sections("Section4").Controls.Item("lbl12").Caption = Label56.Caption
        .Sections("Section4").Controls.Item("lbl13").Caption = Label54.Caption
        .Sections("Section4").Controls.Item("lbl14").Caption = Label37.Caption
        .Sections("Section4").Controls.Item("lbl15").Caption = Label59.Caption
        .Sections("Section4").Controls.Item("lbl16").Caption = Label17.Caption
        .Sections("Section4").Controls.Item("lbl17").Caption = Label45.Caption
        .Sections("Section4").Controls.Item("lbl18").Caption = Label51.Caption
        
         .Sections("Section4").Controls.Item("val1").Caption = lblAgentDoctorFee.Caption
        .Sections("Section4").Controls.Item("val2").Caption = lblDoctorAgentBookings.Caption
        .Sections("Section4").Controls.Item("val3").Caption = lblDoctorAgentCashRepayments.Caption
        .Sections("Section4").Controls.Item("val4").Caption = lblDoctorAgentAgentRepayments.Caption
        .Sections("Section4").Controls.Item("val5").Caption = lblAgentHospitalFee.Caption
        .Sections("Section4").Controls.Item("val6").Caption = lblAgentHospitalBookings.Caption
        .Sections("Section4").Controls.Item("val7").Caption = lblHospitalAgentCashRepayments.Caption
        .Sections("Section4").Controls.Item("val8").Caption = lblHospitalAgentAgentRepayments.Caption
        .Sections("Section4").Controls.Item("val9").Caption = lblDocFeeAgentT.Caption
        .Sections("Section4").Controls.Item("val10").Caption = lblDocFeeBookingsAgentT.Caption
        .Sections("Section4").Controls.Item("val11").Caption = lblDocAgentFeeCashRepaymentsT.Caption
        .Sections("Section4").Controls.Item("val12").Caption = lblDocAgentFeeAgentRepaymentsT.Caption
        .Sections("Section4").Controls.Item("val13").Caption = lblDocFeeAgentO.Caption
        .Sections("Section4").Controls.Item("val14").Caption = lblDocFeeBookingsAgentO.Caption
        .Sections("Section4").Controls.Item("val15").Caption = lblDocAgentFeeCashRepaymentsO.Caption
        .Sections("Section4").Controls.Item("val16").Caption = lblDocAgentFeeAgentRepaymentsO.Caption
        
        .Show
End With

End Sub

Private Sub DTPicker1_DateClick(ByVal DateClicked As Date)
    Me.MousePointer = vbHourglass
    DoEvents
    Call CalculateIncome
    Me.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    DTPicker1.Value = Date
    Call CalculateIncome
    If UserAuthority <> AuthorityOwner Then
        DTPicker1.Enabled = False
    End If

End Sub



Private Sub CalculateIncome()
    Dim TemCash As Double
    Dim TemCash1 As Double
    Dim temSQL As String
    Dim TemWhere As String
    Dim TemNum As Long
    Dim temText As String
    Dim TemCashDC As Double
    Dim TemCashSC As Double
    Dim TemCashRepay As Double
    Dim TemCash2 As Double
    Dim TemCash3 As Double
    Dim TemCash4 As Double
    Dim TemMaxDate As Date
    Dim TemMinDate As Date
    Dim TemDate As Date
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (TotalFee) as TotalGrand, sum(personalfee) as TotalDoctorFee , sum(institutionfee) as TotalHospitalFee  "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
        End If
    End With
    lblCashBookings.Caption = Format(TemCash, "0.00")
    lblDoctorCashDC.Caption = Format(TemCash1, "0.00")
    lblHospitalCashDC.Caption = Format(TemCash2, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (TotalFee) as TotalGrand, sum(personalfee) as TotalDoctorFee , sum(institutionfee) as TotalHospitalFee  "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.PaymentMode)='Credit')) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
        End If
    End With
    lblSettlingCredit.Caption = Format(TemCash, "0.00")
    lblDoctorCashSC.Caption = Format(TemCash1, "0.00")
    lblHospitalCashSC.Caption = Format(TemCash2, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (Totalrefund) as TotalGrand, sum(personalrefund) as TotalDoctorFee , sum(institutionrefund) as TotalHospitalFee  "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
        End If
    End With
    lblCashRepayments.Caption = Format(TemCash, "0.00")
    lblRepaidToPatient.Caption = Format(TemCash, "0.00")
    lblDoctorCashRepayments.Caption = Format(TemCash1, "0.00")
    lblHospitalCashRepayments.Caption = Format(TemCash2, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (Totalrefund) as TotalGrand, sum(personalrefund) as TotalDoctorFee , sum(institutionrefund) as TotalHospitalFee  "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToAgent)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1))"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
        End If
    End With
    lblAgentRepayments.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (Totalrefund) as TotalGrand, sum(personalrefund) as TotalDoctorFee , sum(institutionrefund) as TotalHospitalFee  "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToAgent)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and paymentmode = 'Agent'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
        End If
    End With
    lblDoctorAgentAgentRepayments.Caption = Format(TemCash1, "0.00")
    lblHospitalAgentAgentRepayments.Caption = Format(TemCash2, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (Totalrefund) as TotalGrand, sum(personalrefund) as TotalDoctorFee , sum(institutionrefund) as TotalHospitalFee  "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and paymentmode = 'Agent' "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
        End If
    End With
    lblDoctorAgentCashRepayments.Caption = Format(TemCash1, "0.00")
    lblHospitalAgentCashRepayments.Caption = Format(TemCash2, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (Totalfee) as TotalGrand, sum(personalfee) as TotalDoctorFee , sum(institutionfee) as TotalHospitalFee  "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = "WHERE (((tblPatientFacility.BookingDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent'))"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
        End If
    End With
    lblAgentBoolings.Caption = Format(TemCash, "0.00")
    lblDoctorAgentBookings.Caption = Format(TemCash1, "0.00")
    lblAgentHospitalBookings.Caption = Format(TemCash2, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(cash) as TotalGrand "
    temSQL = temSQL & " FROM tblagentcashsettle "
    TemWhere = " where SettledDate = '" & DTPicker1.Value & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblAgentCashPayments.Caption = Format(TemCash, "0.00")
' ***********************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(paidamount) as TotalGrand "
    temSQL = temSQL & " FROM tblstaffpayment "
    TemWhere = " where PaidDate = '" & DTPicker1.Value & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblDoctorPayments.Caption = Format(TemCash, "0.00")
    lblTodayDoctorPaymentsMade.Caption = Format(TemCash, "0.00")
' ***********************************************
    TemCash = 0
    TemCash1 = 0
    TemCash = Val(lblCashBookings.Caption)
    TemCash = TemCash + Val(lblSettlingCredit.Caption)
    TemCash = TemCash + Val(lblAgentCashPayments.Caption)
    TemCash = TemCash - Val(lblCashRepayments.Caption)
    TemCash = TemCash - Val(lblDoctorPayments.Caption)
    lblNetCash.Caption = Format(TemCash, "0.00")
    lblDoctorCash.Caption = Format(Val(lblDoctorCashDC.Caption) + Val(lblDoctorCashSC.Caption) - Val(lblDoctorCashRepayments.Caption), "0.00")
    lblHospitalCash.Caption = Format(Val(lblHospitalCashDC.Caption) + Val(lblHospitalCashSC.Caption) - Val(lblHospitalCashRepayments.Caption), "0.00")
    lblIncome.Caption = Format(Val(lblCashBookings.Caption) + Val(lblSettlingCredit.Caption) + Val(lblAgentCashPayments.Caption), "0.00")
    lblExpence.Caption = "(" & Format(Val(lblCashRepayments.Caption) + Val(lblDoctorPayments.Caption), "0.00") & ")"
' Doctor Cash By AppointmentDate
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (personalFee) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) and appointmentdate = '" & DTPicker1.Value & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblTOdaysDoctorCashDC.Caption = Format(TemCash, "0.00")
    temSQL = "SELECT sum (personalFee) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Credit')) and appointmentdate = '" & DTPicker1.Value & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash1 = !TotalGrand
        End If
    End With
    lblTodaysDoctorCashSC.Caption = Format(TemCash1, "0.00")
    temSQL = "SELECT sum(personalrefund) as TotalDoctorFee "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and appointmentdate = '" & DTPicker1.Value & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!totaldoctorfee) Then TemCash2 = !totaldoctorfee
        End If
    End With
    lblTodaysDoctorCashRepay.Caption = Format(TemCash2, "0.00")
    lblTodaysDoctorCash.Caption = Format((Val(lblTOdaysDoctorCashDC.Caption) + Val(lblTodaysDoctorCashSC.Caption) - Val(lblTodaysDoctorCashRepay.Caption)), "0.00")
    ListDOctorCash.Clear
    ListDOctorCash.AddItem "Date     " & vbTab & "Direct Cash" & vbTab & "Credit settling" & vbTab & "Cash Repayments" & vbTab & vbTab & "Total"
    With DataEnvironment1.rssqlTem
        temSQL = "Select max(appointmentDate) as MaxBookingDate , min(appointmentdate) as MinBookingDate from tblpatientfacility "
        TemWhere = " where (paymentmode = 'Cash' and bookingdate ='" & DTPicker1.Value & "') or (paymentmode = 'Credit' and settlecashdate = '" & DTPicker1.Value & "') or ( repaydate = '" & DTPicker1.Value & "'  ) "
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If Not IsNull(!MaxBookingDate) Then
            TemMaxDate = !MaxBookingDate
        Else
            TemMaxDate = DTPicker1.Value
        End If
        If Not IsNull(!minbookingdate) Then
            TemMinDate = !minbookingdate
        Else
            TemMinDate = DTPicker1.Value
        End If
        If .RecordCount > 0 Then
            TemNum = 0
            TemCash = 0
            TemCashDC = 0
            TemCashSC = 0
            TemCashRepay = 0
            TemDate = TemMinDate
            While TemMinDate + TemNum <= TemMaxDate
                TemCash1 = 0
                TemCash2 = 0
                TemCash3 = 0
                TemCash4 = 0
                temSQL = "SELECT sum (personalFee) as TotalGrand "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) and appointmentdate = '" & TemDate & "'"
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!TotalGrand) Then TemCash1 = !TotalGrand: TemCashDC = TemCashDC + TemCash1
                    End If
                temSQL = "SELECT sum (personalFee) as TotalGrand "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Credit')) and appointmentdate = '" & TemDate & "'"
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!TotalGrand) Then TemCash2 = !TotalGrand: TemCashSC = TemCashSC + TemCash2
                    End If
                temSQL = "SELECT sum(personalrefund) as TotalDoctorFee "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and appointmentdate = '" & TemDate & "'"
                With DataEnvironment1.rssqlTem1
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!totaldoctorfee) Then TemCash4 = !totaldoctorfee: TemCashRepay = TemCashRepay + !totaldoctorfee
                    End If
                End With
                TemCash3 = TemCash1 + TemCash2 - TemCash4
                If TemCash1 + TemCash2 + TemCash4 > 0 Then
                   temText = Format(TemDate, DefaultShortDate)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash1, "0.00"), 10)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash2, "0.00"), 10)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash4, "0.00"), 10)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash3, "0.00"), 10)
                   ListDOctorCash.AddItem temText
                   TemCash = TemCash + TemCash3
                End If
                TemNum = TemNum + 1
                TemDate = TemMinDate + TemNum
            Wend
        End If
        If .State = 1 Then .Close
    End With
    lblOtherdaysDoctorCashDC.Caption = Format(TemCashDC - Val(lblTOdaysDoctorCashDC.Caption), "0.00")
    lblOtherDaysDoctorCashSC.Caption = Format(TemCashSC - Val(lblTodaysDoctorCashSC.Caption), "0.00")
    lblOtherDaysDoctorCashRepay.Caption = Format(TemCashRepay - Val(lblTodaysDoctorCashRepay.Caption), "0.00")
    lblOtherDaysDoctorCash.Caption = Format(TemCash - Val(lblTodaysDoctorCash.Caption), "0.00")
' ******************************************
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (TotalFee) as TotalGrand, sum(personalfee) as TotalDoctorFee , sum(institutionfee) as TotalHospitalFee  "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent')) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
        End If
    End With
    lblDoctorAgentBookings.Caption = Format(TemCash1, "0.00")
    lblAgentHospitalBookings.Caption = Format(TemCash2, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (Totalrefund) as TotalGrand, sum(personalrefund) as TotalDoctorFee , sum(institutionrefund) as TotalHospitalFee  "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and paymentmode = 'Agent' "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
        End If
    End With
    lblDoctorAgentCashRepayments.Caption = Format(TemCash1, "0.00")
    lblHospitalAgentCashRepayments.Caption = Format(TemCash2, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    
    temSQL = "SELECT sum (Totalrefund) as TotalGrand, sum(personalrefund) as TotalDoctorFee , sum(institutionrefund) as TotalHospitalFee  "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToAgent)=1))   "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
        End If
    End With
    lblDoctorAgentAgentRepayments.Caption = Format(TemCash1, "0.00")
    lblHospitalAgentAgentRepayments.Caption = Format(TemCash2, "0.00")
' *******************************************************
' ************* TotalRepayments ****************
            
        
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum (Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblRepaidToPatient.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToAgent)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1))"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblRepaidToAgent.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE tblPatientFacility.RepayDate='" & DTPicker1.Value & "' AND tblPatientFacility.HospitalFacility_ID=10  AND (tblPatientFacility.Cancelled=1 or tblPatientFacility.Refund=1)"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblTotalRepayments.Caption = Format(TemCash, "0.00")

' *******************************************************
' Total Doctor Payments made for today
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(personaldue) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE tblPatientFacility.AppointmentDate='" & DTPicker1.Value & "' AND tblPatientFacility.HospitalFacility_ID=10  AND tblPatientFacility.paidtostaff = 1"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblDocPaymentsForTodaysApp.Caption = Format(TemCash, "0.00")
' *******************************************************
' Total Doctor Payments to make for today
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(personaldue) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE tblPatientFacility.appointmentdate='" & DTPicker1.Value & "' AND tblPatientFacility.HospitalFacility_ID=10  AND tblPatientFacility.paidtostaff =0  and tblPatientFacility.fullypaid = 1 "
    If PayToDoctor = False Then
        TemWhere = TemWhere & " and patientabsent = 0 "
    End If
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblPaymentsToDo.Caption = Format(TemCash, "0.00")
' ***********************************************
' Total Doctor Payments for today
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(personaldue) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE tblPatientFacility.appointmentdate='" & DTPicker1.Value & "' AND tblPatientFacility.HospitalFacility_ID=10 and tblPatientFacility.fullypaid = 1 "
    If PayToDoctor = False Then
        TemWhere = TemWhere & " and patientabsent = 0 "
    End If
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblTotalDocPayments.Caption = Format(TemCash, "0.00")
' ***********************************************
' Total Doctor Payments for cash bookings
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    TemCash3 = 0
    temSQL = "SELECT sum(personaldue) as TotalGrand , sum(personaldue) as DocFee , sum(institutiondue) as HosFee "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE tblPatientFacility.appointmentdate='" & DTPicker1.Value & "' AND tblPatientFacility.HospitalFacility_ID=10 and tblPatientFacility.paymentmode = 'Cash' "
    If PayToDoctor = False Then
        TemWhere = TemWhere & " and patientabsent = 0 "
    End If
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!docfee) Then TemCash1 = !docfee
            If Not IsNull(!hosfee) Then TemCash2 = !hosfee
        End If
    End With
    lblTodayDocCash.Caption = Format(TemCash1, "0.00")
    lblTodayHosCash.Caption = Format(TemCash2, "0.00")
' *******************************************************
' Total Doctor Payments for credit bookings
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    TemCash3 = 0
    temSQL = "SELECT sum(personaldue) as TotalGrand , sum(personaldue) as DocFee , sum(institutiondue) as HosFee "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE tblPatientFacility.appointmentdate='" & DTPicker1.Value & "' AND tblPatientFacility.HospitalFacility_ID=10 and tblPatientFacility.paymentmode = 'Credit' and tblPatientFacility.fullypaid = 1 "
    If PayToDoctor = False Then
        TemWhere = TemWhere & " and patientabsent = 0 "
    End If
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!docfee) Then TemCash1 = !docfee
            If Not IsNull(!hosfee) Then TemCash2 = !hosfee
        End If
    End With
    lblTodayDocCredit.Caption = Format(TemCash1, "0.00")
    lblTodayHosCredit.Caption = Format(TemCash2, "0.00")
' *******************************************************
' Total Doctor Payments for Agent bookings
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    TemCash3 = 0
    temSQL = "SELECT sum(personaldue) as TotalGrand , sum(personaldue) as DocFee , sum(institutiondue) as HosFee "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE tblPatientFacility.appointmentdate='" & DTPicker1.Value & "' AND tblPatientFacility.HospitalFacility_ID=10 and tblPatientFacility.paymentmode = 'Agent' "
    If PayToDoctor = False Then
        TemWhere = TemWhere & " and patientabsent = 0 "
    End If
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!docfee) Then TemCash1 = !docfee
            If Not IsNull(!hosfee) Then TemCash2 = !hosfee
        End If
    End With
    lblTodayDocAgent.Caption = Format(TemCash1, "0.00")
    lblTodayHosAgent.Caption = Format(TemCash2, "0.00")

' *******************************************************

lblTOdayHos.Caption = Format((Val(lblTodayHosCash.Caption) + Val(lblTodayHosCredit.Caption) + Val(lblTodayHosAgent.Caption)), "0.00")
lblTodayDoc.Caption = Format((Val(lblTodayDocCash.Caption) + Val(lblTodayDocCredit.Caption) + Val(lblTodayDocAgent.Caption)), "0.00")
lblToday.Caption = Format((Val(lblTOdayHos.Caption) + Val(lblTodayDoc.Caption)), "0.00")

' *******************************************************

lblAgentDoctorFee.Caption = Format(Val(lblDoctorAgentBookings.Caption) - Val(lblDoctorAgentCashRepayments.Caption) - Val(lblDoctorAgentAgentRepayments.Caption), "0.00")
lblAgentHospitalFee.Caption = Format(Val(lblAgentHospitalBookings.Caption) - Val(lblHospitalAgentCashRepayments.Caption) - Val(lblHospitalAgentAgentRepayments.Caption), "0.00")
lblNetagentBooking.Caption = Format(Val(lblAgentDoctorFee.Caption) + Val(lblAgentHospitalFee.Caption), "0.00")







' Doctor Agent By AppointmentDate
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (personalFee) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent')) and appointmentdate = '" & DTPicker1.Value & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblDocFeeBookingsAgentT.Caption = Format(TemCash, "0.00")
    
    temSQL = "SELECT sum(personalrefund) as TotalDoctorFee "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE  ((tblPatientFacility.PaymentMode)='Agent') and (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1)  AND appointmentdate = '" & DTPicker1.Value & "' )"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!totaldoctorfee) Then TemCash2 = !totaldoctorfee
        End If
    End With
    lblDocAgentFeeCashRepaymentsT.Caption = Format(TemCash2, "0.00")
    
    temSQL = "SELECT sum(personalrefund) as TotalDoctorFee "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE  ((tblPatientFacility.PaymentMode)='Agent') and (((tblPatientFacility.RepayDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToAgent)=1)  AND appointmentdate = '" & DTPicker1.Value & "')"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!totaldoctorfee) Then TemCash2 = !totaldoctorfee
        End If
    End With
    lblDocAgentFeeAgentRepaymentsT.Caption = Format(TemCash2, "0.00")
    
    lblDocFeeAgentT.Caption = Format((Val(lblDocFeeBookingsAgentT.Caption) - Val(lblDocAgentFeeCashRepaymentsT) - Val(lblDocAgentFeeAgentRepaymentsT.Caption)), "0.00")
    
    ListDoctorAgent.Clear
    ListDoctorAgent.AddItem "Date     " & vbTab & "Agent Bookings" & vbTab & "Cash Repay" & vbTab & "Paid to Agent" & vbTab & vbTab & "Total"
    
    With DataEnvironment1.rssqlTem
        temSQL = "Select max(appointmentDate) as MaxBookingDate , min(appointmentdate) as MinBookingDate from tblpatientfacility "
        TemWhere = " where (paymentmode = 'Agent' and bookingdate ='" & DTPicker1.Value & "') or ( repaydate = '" & DTPicker1.Value & "'  ) "
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If Not IsNull(!MaxBookingDate) Then
            TemMaxDate = !MaxBookingDate
        Else
            TemMaxDate = DTPicker1.Value
        End If
        If Not IsNull(!minbookingdate) Then
            TemMinDate = !minbookingdate
        Else
            TemMinDate = DTPicker1.Value
        End If
        
        Dim TemCashB As Double
        Dim TemCashCR As Double
        Dim TemCashAR As Double
        
        If .RecordCount > 0 Then
            TemNum = 0
            TemCashB = 0
            TemCashCR = 0
            TemCashAR = 0
            TemCashRepay = 0
            TemDate = TemMinDate
            While TemMinDate + TemNum <= TemMaxDate
                TemCash1 = 0
                TemCash2 = 0
                TemCash3 = 0
                TemCash4 = 0
                temSQL = "SELECT sum (personalFee) as TotalGrand "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & DTPicker1.Value & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent')) and appointmentdate = '" & TemDate & "'"
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!TotalGrand) Then TemCash1 = !TotalGrand: TemCashB = TemCashB + TemCash1
                    End If
                temSQL = "SELECT sum(personalrefund) as TotalDoctorFee "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE tblPatientFacility.RepayDate='" & DTPicker1.Value & "' AND tblPatientFacility.HospitalFacility_ID = 10 AND tblPatientFacility.PaymentMode='Agent' AND tblPatientFacility.RefundToPatient= 1  and appointmentdate = '" & TemDate & "'"
                With DataEnvironment1.rssqlTem1
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!totaldoctorfee) Then TemCash2 = !totaldoctorfee: TemCashCR = TemCashCR + !totaldoctorfee
                    End If
                End With
                temSQL = "SELECT sum(personalrefund) as TotalDoctorFee "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE tblPatientFacility.RepayDate='" & DTPicker1.Value & "' AND tblPatientFacility.HospitalFacility_ID=10 AND tblPatientFacility.PaymentMode='Agent' AND tblPatientFacility.RefundToAgent=1  and appointmentdate = '" & TemDate & "'"
                With DataEnvironment1.rssqlTem1
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!totaldoctorfee) Then TemCash3 = !totaldoctorfee: TemCashAR = TemCashAR + !totaldoctorfee
                    End If
                End With
                
                TemCash4 = TemCash1 - TemCash2 - TemCash3
                
                If Abs(TemCash1) + Abs(TemCash2) + Abs(TemCash3) <> 0 Then
                
                   temText = Format(TemDate, DefaultShortDate)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash1, "0.00"), 10)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash2, "0.00"), 10)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash3, "0.00"), 10)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash4, "0.00"), 10)
                   ListDoctorAgent.AddItem temText
                   TemCash = TemCash + TemCash4
                End If
                TemNum = TemNum + 1
                TemDate = TemMinDate + TemNum
            Wend
        End If
        If .State = 1 Then .Close
    End With
    
    lblDocFeeBookingsAgentO.Caption = Format(TemCashB - Val(lblDocFeeBookingsAgentT.Caption), "0.00")
    lblDocAgentFeeCashRepaymentsO.Caption = Format(TemCashCR - Val(lblDocAgentFeeCashRepaymentsT.Caption), "0.00")
    lblDocAgentFeeAgentRepaymentsO.Caption = Format(TemCashAR - Val(lblDocAgentFeeAgentRepaymentsT.Caption), "0.00")
    lblDocFeeAgentO.Caption = Format(TemCash - Val(lblDocFeeAgentT.Caption), "0.00")
' ******************************************


























End Sub

