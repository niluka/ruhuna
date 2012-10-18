VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNewMyShiftEndSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift End Summery"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewMyShiftEndSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   8925
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Summery"
      TabPicture(0)   =   "frmNewMyShiftEndSummary.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNetCash"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAgentRepayments"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblAgentBoolings"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(15)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblAgentCashPayments"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblDoctorPayments"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCashRepayments"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(9)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(6)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblSettlingCredit"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblCashBookings"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblIncome"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblExpence"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "bttnPrintSummery"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Cash Bookings"
      TabPicture(1)   =   "frmNewMyShiftEndSummary.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "bttnPrintCashRepayments"
      Tab(1).Control(1)=   "bttnPrintCreditSettling"
      Tab(1).Control(2)=   "bttnPrintCashBookings"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(4)=   "Frame2"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Agent Bookings"
      TabPicture(2)   =   "frmNewMyShiftEndSummary.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "bttnPrintAgentRepayments"
      Tab(2).Control(2)=   "bttnPrintAgentPayments"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Payments"
      TabPicture(3)   =   "frmNewMyShiftEndSummary.frx":03DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "bttnDoctorPaymentsDoneToday"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "bttnDoctorPaymentsDoneForTodayAppointments"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "bttnDoctorPaymentsToDoForTodayAppointments"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "bttnDoctorPaymentsForTodayAppointments"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame5"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
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
         Height          =   3135
         Left            =   -74760
         TabIndex        =   64
         Top             =   480
         Width           =   8175
         Begin VB.Label Label32 
            Caption         =   "Net agent bookings Value"
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   2760
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
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label31 
            Caption         =   "Less -Agent Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   67
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label44 
            Caption         =   "Doctor Fee From Agent Bookings"
            Height          =   255
            Left            =   240
            TabIndex        =   80
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label43 
            Caption         =   "Hospital Fee From Agent Bookings"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   1560
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
            Top             =   360
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
            Left            =   5160
            TabIndex        =   77
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label Label40 
            Caption         =   "Agent Bookings"
            Height          =   255
            Left            =   720
            TabIndex        =   76
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label lblDoctorAgentBookings 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   75
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label38 
            Caption         =   "Less -Cash Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   74
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label lblDoctorAgentCashRepayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   73
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label36 
            Caption         =   "Agent Bookings"
            Height          =   255
            Left            =   720
            TabIndex        =   72
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblAgentHospitalBookings 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   71
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label34 
            Caption         =   "Less - Cash Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   70
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label lblHospitalAgentCashRepayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   69
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label lblDoctorAgentAgentRepayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2400
            TabIndex        =   68
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "Less - Agent Repayments"
            Height          =   255
            Left            =   720
            TabIndex        =   65
            Top             =   2280
            Width           =   2775
         End
         Begin VB.Label lblHospitalAgentAgentRepayments 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   255
            Left            =   2880
            TabIndex        =   66
            Top             =   2280
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
         Left            =   -74760
         TabIndex        =   25
         Top             =   3360
         Width           =   8295
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
         Left            =   -74760
         TabIndex        =   22
         Top             =   480
         Width           =   8295
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
         Left            =   6600
         TabIndex        =   83
         Top             =   7920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Print Summery"
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
         Left            =   -74640
         TabIndex        =   84
         Top             =   7920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Print Cash Bookings"
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
         Left            =   -71880
         TabIndex        =   85
         Top             =   7920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Print Credit Settling"
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
         Left            =   -69120
         TabIndex        =   86
         Top             =   7920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Print Cash Repayments"
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
         Left            =   -69240
         TabIndex        =   87
         Top             =   4440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Print Agent Repayments"
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
         Left            =   -69240
         TabIndex        =   88
         Top             =   3840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Print Agent Payments"
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
      Begin VB.Label lblExpence 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   7200
         TabIndex        =   103
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label lblIncome 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   7200
         TabIndex        =   102
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cash Income"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Non-Cash Income"
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Agent Repayments"
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   6960
         Width           =   2055
      End
      Begin VB.Label lblCashBookings 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5400
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblSettlingCredit 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5400
         TabIndex        =   17
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Non cash Expences"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   16
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Cash Expences"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   15
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Settling Credit Bookings"
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   14
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Agent Cash Payments"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   13
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Cash Bookings"
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Doctor Payments"
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   11
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Agent Bookings"
         Height          =   255
         Index           =   8
         Left            =   1200
         TabIndex        =   10
         Top             =   5280
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Cash Repayments"
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   9
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblCashRepayments 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblDoctorPayments 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5400
         TabIndex        =   7
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblAgentCashPayments 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5400
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
         Left            =   720
         TabIndex        =   5
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lblAgentBoolings 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5400
         TabIndex        =   4
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label lblAgentRepayments 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5520
         TabIndex        =   3
         Top             =   6960
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
         Left            =   5640
         TabIndex        =   2
         Top             =   4320
         Width           =   2775
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      Top             =   8880
      Width           =   1695
      _ExtentX        =   2990
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
End
Attribute VB_Name = "frmNewMyShiftEndSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CSetPrinter As New cSetDfltPrinter

Private Sub bttnClose_Click()
    Unload Me
End Sub


Private Sub bttnDoctorPaymentsDoneForTodayAppointments_Click()
    Const PreSHape = "SHAPE {"
    Const Sql = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.* FROM tblPatientFacility LEFT JOIN (tblTitle RIGHT JOIN tblDoctor ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID "
    Dim SqlWhere  As String
    SqlWhere = " WHERE (((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.PaidToStaffUser)=" & UserID & ") AND   ((tblPatientFacility.PaidToSTaff)=1) AND ((tblPatientFacility.AppointmentDate)='" & Date & "') "
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
        .Sections.Item("ReportHeader").Controls.Item("lblreport").Caption = "Doctor Payments For All Bookings by " & FindStaffFromID(UserID)
        .Sections.Item("ReportHeader").Controls.Item("lblreport").Caption = Format(Date, DefaultLongDate)
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
        .rsTodayDoctorPayments.Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblStaff.StaffName, tblStaffPayment.* FROM tblStaff RIGHT JOIN (tblTitle RIGHT JOIN (tblDoctor RIGHT JOIN tblStaffPayment ON tblDoctor.Doctor_ID = tblStaffPayment.Staff_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID) ON tblStaff.Staff_ID = tblStaffPayment.User_ID where paiddate = '" & Date & "' order by StaffPayment_ID"
        .rsTodayDoctorPayments.Open
    End With
    With dtrTodayDocPayments
        .DataMember = "TodayDoctorPayments"
        If HospitalDetails = True Then
            .Sections("section4").Controls("lblinstitutionname").Caption = InstitutionName
            .Sections("section4").Controls("lblinstitutionaddress").Caption = InstitutionAddress
            .Sections("section4").Controls("lblReport").Caption = "Doctor Payments done Today"
            .Sections("section5").Controls("lblad").Caption = LongAd
            .Sections("section4").Controls("lblReportsub").Caption = Format(Date, DefaultLongDate)
        Else
            .Sections("section4").Controls("lblinstitutionname").Caption = Empty
            .Sections("section4").Controls("lblinstitutionaddress").Caption = Empty
            .Sections("section4").Controls("lblReport").Caption = "Doctor Payments done Today"
            .Sections("section4").Controls("lblReportsub").Caption = Format(Date, DefaultLongDate)
            .Sections("section5").Controls("lblad").Caption = LongAd
        End If
        .Show
    End With
End Sub

Private Sub bttnDoctorPaymentsForTodayAppointments_Click()
    Const PreSHape = "SHAPE {"
    Const Sql = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.* FROM tblPatientFacility LEFT JOIN (tblTitle RIGHT JOIN tblDoctor ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID "
    Dim SqlWhere  As String
    SqlWhere = " WHERE (((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.AppointmentDate)='" & Date & "') "
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
        .Sections.Item("ReportHeader").Controls.Item("lblreport").Caption = Format(Date, DefaultLongDate)
        .Sections.Item("PageFooter").Controls.Item("lblAd").Caption = LongAd
        .Show
    End With
End Sub

Private Sub bttnDoctorPaymentsToDoForTodayAppointments_Click()
    Const PreSHape = "SHAPE {"
    Const Sql = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.* FROM tblPatientFacility LEFT JOIN (tblTitle RIGHT JOIN tblDoctor ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID "
    Dim SqlWhere  As String
    SqlWhere = " WHERE (((tblPatientFacility.FullyPaid)=1) AND  ((tblPatientFacility.PaidToSTaff)=0)AND ((tblPatientFacility.AppointmentDate)='" & Date & "') "
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
        .Sections.Item("ReportHeader").Controls.Item("lblreport").Caption = Format(Date, DefaultLongDate)
        .Sections.Item("PageFooter").Controls.Item("lblAd").Caption = LongAd
        .Show
    End With
End Sub

Private Sub bttnPrintAgentPayments_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
With DataEnvironment1.rssqlTem3
    If .State = 1 Then .Close
     .Source = "Select tblAgentCashSettle.*, tblInstitutions.* fROM tblAgentCashSettle Left Join tblInstitutions On tblAgentCashSettle.Institution_Id = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.SettledDate = '" & Date & "') and User_ID =" & UserID
    .Open
    End With
    With dtrAgentCashReceive
        If HospitalDetails = True Then
        .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        Else
        .Sections("Section4").Controls.Item("RptName").Caption = Empty
        .Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        End If
         .Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
         .Sections("Section2").Controls.Item("rptTodate").Caption = Format(Date, DefaultLongDate)
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
    .Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.*, tblPatientMainDetails.FirstName, tblPatientFacility.SettleCashDate, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode FROM tblInstitutions RIGHT JOIN (tblPatientMainDetails RIGHT JOIN (tblPatientFacility LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID Where (((tblPatientFacility.HospitalFacility_ID) = 10) And ((tblPatientFacility.RepayUser_ID) = " & UserID & ") And ((tblPatientFacility.RefundToAgent) = 1) And ((tblPatientFacility.Repaydate) = '" & Date & "') ) ORDER BY tblPatientFacility.PatientFacility_ID"
    .Open
End With
With dtrCashRefunds
    .DataMember = "refunds"
    If HospitalDetails = True Then
        .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = InstitutionName
        .Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
    End If
    .Sections.Item("Section4").Controls.Item("lblreport").Caption = "All Agent Repayments"
    .Sections.Item("Section4").Controls.Item("lblreportsub").Caption = Format(Date, DefaultLongDate)
    .Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
    .Show
End With
End Sub

Private Sub bttnPrintCashBookings_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
dtrCashBookings.DataMember = Empty
With DataEnvironment1.rsCashBookings
    If .State = 1 Then .Close
    .Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.*, tblPatientMainDetails.FirstName FROM tblPatientMainDetails RIGHT JOIN (tblPatientFacility LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.User_ID)=" & UserID & ") AND ((tblPatientFacility.PaymentMode)='Cash') AND ((tblPatientFacility.bookingdate)='" & Date & "') ) ORDER BY tblPatientFacility.PatientFacility_ID "
    .Open
End With
With dtrCashBookings
    .DataMember = "CashBookings"
    If HospitalDetails = True Then
        .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = InstitutionName
        .Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
    End If
    .Sections.Item("Section4").Controls.Item("lblreport").Caption = "All Cash Bookings"
    .Sections.Item("Section4").Controls.Item("lblreportsub").Caption = Format(Date, DefaultLongDate)
    .Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
    .Show
End With
End Sub

Private Sub bttnPrintCashRepayments_Click()
dtrCashRefunds.DataMember = Empty
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
With DataEnvironment1.rsRefunds
    If .State = 1 Then .Close
    .Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.*, tblPatientMainDetails.FirstName, tblPatientFacility.SettleCashDate, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode FROM tblInstitutions RIGHT JOIN (tblPatientMainDetails RIGHT JOIN (tblPatientFacility LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID Where (((tblPatientFacility.HospitalFacility_ID) = 10) And ((tblPatientFacility.RepayUser_ID) = " & UserID & ") And ((tblPatientFacility.RefundToPatient) = 1) And ((tblPatientFacility.Repaydate) = '" & Date & "') ) ORDER BY tblPatientFacility.PatientFacility_ID"
    .Open
End With
With dtrCashRefunds
    .DataMember = "refunds"
    If HospitalDetails = True Then
        .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = InstitutionName
        .Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
    End If
    .Sections.Item("Section4").Controls.Item("lblreport").Caption = "All Cash Repayments"
    .Sections.Item("Section4").Controls.Item("lblreportsub").Caption = Format(Date, DefaultLongDate)
    .Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
    .Show
End With

End Sub

Private Sub bttnPrintCreditSettling_Click()
dtrCreditBookings.DataMember = Empty
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
With DataEnvironment1.rsCashBookings
    If .State = 1 Then .Close
    .Source = "SELECT tblTitle.Title, tblDoctor.DoctorName, tblPatientFacility.*, tblPatientMainDetails.FirstName FROM tblPatientMainDetails RIGHT JOIN (tblPatientFacility LEFT JOIN (tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.CreditSettleUser_ID)=" & UserID & ") AND ((tblPatientFacility.PaymentMode)='Credit') AND ((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.SettleCashDate)='" & Date & "') ) ORDER BY tblPatientFacility.PatientFacility_ID "
    .Open
End With
With dtrCreditBookings
    .DataMember = "CashBookings"
    If HospitalDetails = True Then
        .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = InstitutionName
        .Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
    End If
    .Sections.Item("Section4").Controls.Item("lblreport").Caption = "All Cash For Credit Bookings"
    .Sections.Item("Section4").Controls.Item("lblreportsub").Caption = Format(Date, DefaultLongDate)
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
        .Sections("Section4").Controls.Item("lblreportsub").Caption = Format(Date, DefaultLongDate)
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

Private Sub Form_Load()
    SSTab1.Tab = 0
    Call CalculateIncome
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
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.User_ID)=" & UserID & ") AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) "
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
    TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.CreditSettleUser_ID)=" & UserID & ")  AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.PaymentMode)='Credit')) "
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
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ")  AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
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
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ") AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToagent)=True))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1))"
    With DataEnvironment1.rssqlTem1     'RepayUser_ID
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
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ") AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToAgent)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and paymentmode = 'Agent'"
    With DataEnvironment1.rssqlTem1                                  ' RepayUser_ID
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
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ") AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and paymentmode = 'Agent' "
    With DataEnvironment1.rssqlTem1                                                       ' RepayUser_ID
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
    TemWhere = "WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.user_ID)=" & UserID & ") AND ((tblPatientFacility.PaymentMode)='Agent'))"
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
'    TemCash = 0
'    TemCash1 = 0
'    TemCash2 = 0
'    temSql = "SELECT sum (Totalfee) as TotalGrand, sum(personalfee) as TotalDoctorFee , sum(institutionfee) as TotalHospitalFee  "
'    temSql = temSql & " FROM tblPatientFacility "
'    TemWhere = "WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.user_ID)=" & UserID & ") AND ((tblPatientFacility.PaymentMode)='Credit')  AND ( (tblPatientFacility.CreditStaff_ID)=0 OR (tblPatientFacility.CreditStaff_ID is Null)  ))"
'    With DataEnvironment1.rssqlTem1
'        If .State = 1 Then .Close
'        .Source = temSql & TemWhere
'        .Open
'        If .RecordCount <> 0 Then
'            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
'            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
'            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
'        End If
'    End With
'    lblTelephoneBookings.Caption = Format(TemCash, "0.00")
''    lblDoctorAgentBookings.Caption = Format(TemCash1, "0.00")
''    lblAgentHospitalBookings.Caption = Format(TemCash2, "0.00")

' *******************************************************
'    TemCash = 0
'    TemCash1 = 0
'    TemCash2 = 0
'    temSql = "SELECT sum (Totalfee) as TotalGrand, sum(personalfee) as TotalDoctorFee , sum(institutionfee) as TotalHospitalFee  "
'    temSql = temSql & " FROM tblPatientFacility "
'    TemWhere = "WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.user_ID)=" & UserID & ") AND ((tblPatientFacility.PaymentMode)='Credit')  AND ( (tblPatientFacility.CreditStaff_ID) <> 0 AND (tblPatientFacility.CreditStaff_ID IS NOT Null)  ))"
'    With DataEnvironment1.rssqlTem1
'        If .State = 1 Then .Close
'        .Source = temSql & TemWhere
'        .Open
'        If .RecordCount <> 0 Then
'            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
'            If Not IsNull(!totaldoctorfee) Then TemCash1 = !totaldoctorfee
'            If Not IsNull(!totalhospitalfee) Then TemCash2 = !totalhospitalfee
'        End If
'    End With
'    lblStaffBookings.Caption = Format(TemCash, "0.00")
'    lblDoctorAgentBookings.Caption = Format(TemCash1, "0.00")
'    lblAgentHospitalBookings.Caption = Format(TemCash2, "0.00")



' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(cash) as TotalGrand "
    temSQL = temSQL & " FROM tblagentcashsettle "
    TemWhere = " where SettledDate = '" & Date & "'"
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
    TemWhere = " where PaidDate = '" & Date & "' and User_id =" & UserID
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
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.user_ID)=" & UserID & ") AND ((tblPatientFacility.PaymentMode)='Cash')) and appointmentdate = '" & Date & "'"
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
    TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.CreditSettleUser_ID)=" & UserID & ")  AND ((tblPatientFacility.PaymentMode)='Credit')) and appointmentdate = '" & Date & "'"
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
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ")  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and appointmentdate = '" & Date & "'"
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
        TemWhere = " where (paymentmode = 'Cash' and bookingdate ='" & Date & "') or (paymentmode = 'Credit' and settlecashdate = '" & Date & "') or ( repaydate = '" & Date & "'  ) "
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If Not IsNull(!MaxBookingDate) Then
            TemMaxDate = !MaxBookingDate
        Else
            TemMaxDate = Date
        End If
        If Not IsNull(!minbookingdate) Then
            TemMinDate = !minbookingdate
        Else
            TemMinDate = Date
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
                TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.user_ID)=" & UserID & ") AND  ((tblPatientFacility.PaymentMode)='Cash')) and appointmentdate = '" & TemDate & "'"
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!TotalGrand) Then TemCash1 = !TotalGrand: TemCashDC = TemCashDC + TemCash1
                    End If
                temSQL = "SELECT sum (personalFee) as TotalGrand "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.CreditSettleUser_ID)=" & UserID & ") AND  ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Credit')) and appointmentdate = '" & TemDate & "'"
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!TotalGrand) Then TemCash2 = !TotalGrand: TemCashSC = TemCashSC + TemCash2
                    End If
                temSQL = "SELECT sum(personalrefund) as TotalDoctorFee "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND  ((tblPatientFacility.RepayUser_ID)=" & UserID & ")  AND  ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and appointmentdate = '" & TemDate & "'"
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
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.user_ID)=" & UserID & ") AND ((tblPatientFacility.PaymentMode)='Agent')) "
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
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ") AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and paymentmode = 'Agent' "
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
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ")  AND ((tblPatientFacility.RefundToAgent)=1))   "
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
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ")   AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
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
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ")   AND ((tblPatientFacility.RefundToAgent)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1))"
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
    TemWhere = " WHERE tblPatientFacility.RepayDate='" & Date & "' AND  ((tblPatientFacility.RepayUser_ID)=" & UserID & ") and  tblPatientFacility.HospitalFacility_ID=10  AND (tblPatientFacility.Cancelled=1 or tblPatientFacility.Refund=1)"
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
    TemWhere = " WHERE tblPatientFacility.AppointmentDate='" & Date & "' AND tblPatientFacility.HospitalFacility_ID=10  AND tblPatientFacility.paidtostaff = 1"
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
    TemWhere = " WHERE tblPatientFacility.appointmentdate='" & Date & "' AND tblPatientFacility.HospitalFacility_ID=10  AND tblPatientFacility.paidtostaff =0  and tblPatientFacility.fullypaid = 1 "
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
    TemWhere = " WHERE tblPatientFacility.appointmentdate='" & Date & "' AND tblPatientFacility.HospitalFacility_ID=10 and tblPatientFacility.fullypaid = 1 "
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
    TemWhere = " WHERE tblPatientFacility.appointmentdate='" & Date & "' AND tblPatientFacility.HospitalFacility_ID=10 and tblPatientFacility.paymentmode = 'Cash' "
    If PayToDoctor = False Then
        TemWhere = TemWhere & " and patientabsent = 0 and user_ID = " & UserID & " "
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
    TemWhere = " WHERE tblPatientFacility.appointmentdate='" & Date & "' AND tblPatientFacility.HospitalFacility_ID=10 and tblPatientFacility.paymentmode = 'Credit' and tblPatientFacility.fullypaid = 1 "
    If PayToDoctor = False Then
        TemWhere = TemWhere & " and patientabsent = 0 and CreditSettleUser_ID = " & UserID & " "
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
    TemWhere = " WHERE tblPatientFacility.appointmentdate='" & Date & "' AND tblPatientFacility.HospitalFacility_ID=10 and tblPatientFacility.paymentmode = 'Agent' "
    If PayToDoctor = False Then
        TemWhere = TemWhere & " and patientabsent = 0 and user_ID = " & UserID & " "
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

End Sub

