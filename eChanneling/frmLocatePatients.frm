VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmLocatePatients 
   Caption         =   "Locating Bookings"
   ClientHeight    =   10230
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
   Icon            =   "frmLocatePatients.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   15240
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2820
      Left            =   6360
      TabIndex        =   133
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   73334785
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
      ItemData        =   "frmLocatePatients.frx":0442
      Left            =   6360
      List            =   "frmLocatePatients.frx":0444
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
      Height          =   4350
      ItemData        =   "frmLocatePatients.frx":0446
      Left            =   9480
      List            =   "frmLocatePatients.frx":0448
      TabIndex        =   3
      ToolTipText     =   "List of number, patient, paid or not, cancelled or refunded, agent code and present or absent"
      Top             =   360
      Width           =   5775
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
      ItemData        =   "frmLocatePatients.frx":044A
      Left            =   3000
      List            =   "frmLocatePatients.frx":044C
      TabIndex        =   1
      ToolTipText     =   "List of Consultants of selected speciality"
      Top             =   360
      Width           =   3135
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
      ItemData        =   "frmLocatePatients.frx":044E
      Left            =   240
      List            =   "frmLocatePatients.frx":0450
      TabIndex        =   0
      ToolTipText     =   "List of Specialities"
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox ListSecessionStartingTime 
      Height          =   4380
      Left            =   13920
      TabIndex        =   55
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListPatientFacilityIDs 
      Height          =   4380
      Left            =   14280
      TabIndex        =   53
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox ListConsultantIDs 
      Height          =   4380
      Left            =   14280
      TabIndex        =   52
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox ListSecessionIDs 
      Height          =   4380
      Left            =   13920
      TabIndex        =   51
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListDates 
      Height          =   4380
      Left            =   13920
      TabIndex        =   50
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListSpecialityIDs 
      Height          =   4380
      Left            =   14040
      TabIndex        =   49
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   8
      Tab             =   4
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Booking"
      TabPicture(0)   =   "frmLocatePatients.frx":0452
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FramePatientDetails"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Reprint"
      TabPicture(1)   =   "frmLocatePatients.frx":046E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameReprints"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cancel"
      TabPicture(2)   =   "frmLocatePatients.frx":048A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameCancellations"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Refund"
      TabPicture(3)   =   "frmLocatePatients.frx":04A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameRefunds"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Settle Credit"
      TabPicture(4)   =   "frmLocatePatients.frx":04C2
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "FrameSettleCredit"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Absent"
      TabPicture(5)   =   "frmLocatePatients.frx":04DE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "bttnMarkPresent"
      Tab(5).Control(1)=   "bttnMarkAbsent"
      Tab(5).Control(2)=   "bttnChangeName"
      Tab(5).Control(3)=   "txtNameChange"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Search"
      TabPicture(6)   =   "frmLocatePatients.frx":04FA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "txtSearchAgentRefNo"
      Tab(6).Control(1)=   "txtSearchBookingID"
      Tab(6).Control(2)=   "ComboPatientName"
      Tab(6).Control(3)=   "DTPickerFindPatientDate"
      Tab(6).Control(4)=   "gridPatient"
      Tab(6).Control(5)=   "bttnSearch"
      Tab(6).Control(6)=   "bttnAgentRefSearch"
      Tab(6).Control(7)=   "Label31"
      Tab(6).Control(8)=   "Label26"
      Tab(6).Control(9)=   "Label20"
      Tab(6).Control(10)=   "lblAgentName"
      Tab(6).Control(11)=   "Label27"
      Tab(6).ControlCount=   12
      TabCaption(7)   =   "Views"
      TabPicture(7)   =   "frmLocatePatients.frx":0516
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "MonthView1"
      Tab(7).Control(1)=   "bttnNurseView"
      Tab(7).Control(2)=   "bttnDoctorView"
      Tab(7).Control(3)=   "bttnAllPatients"
      Tab(7).Control(4)=   "bttnAllDoctors"
      Tab(7).Control(5)=   "bttnAllSecessionPatients"
      Tab(7).ControlCount=   6
      Begin VB.TextBox txtSearchAgentRefNo 
         Height          =   375
         Left            =   -62640
         TabIndex        =   146
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtNameChange 
         Height          =   360
         Left            =   -74760
         TabIndex        =   36
         Top             =   3000
         Width           =   3135
      End
      Begin VB.TextBox txtSearchBookingID 
         Height          =   375
         Left            =   -62640
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox ComboPatientName 
         Height          =   360
         Left            =   -72000
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   480
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker DTPickerFindPatientDate 
         Height          =   375
         Left            =   -74160
         TabIndex        =   38
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MM yyyy"
         Format          =   73334787
         CurrentDate     =   39470
      End
      Begin VB.Frame FramePatientDetails 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   115
         Top             =   360
         Width           =   14775
         Begin VB.TextBox txtPhone 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   150
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox txtConsultant 
            Height          =   360
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   144
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtAppAt 
            Height          =   360
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   142
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox txtAppOn 
            Height          =   360
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   140
            Top             =   720
            Width           =   3495
         End
         Begin VB.TextBox txtBookingOn 
            Height          =   360
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   138
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txtAgentRefNo 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   136
            Top             =   2640
            Width           =   3135
         End
         Begin VB.TextBox txtAgentCode 
            Height          =   375
            Left            =   3840
            TabIndex        =   135
            Top             =   1920
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtBookedPatientName 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtBookedPatientID 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   116
            Top             =   240
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.TextBox txtBookingID 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox txtPaymentMethod 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1680
            Width           =   3135
         End
         Begin VB.TextBox txtBookingUser 
            Height          =   360
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   2400
            Width           =   3135
         End
         Begin VB.TextBox txtCancelRefund 
            Height          =   1440
            Left            =   11520
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtCreditSettle 
            Height          =   1200
            Left            =   11520
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1800
            Width           =   3135
         End
         Begin VB.TextBox txtAgentAndCode 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   2160
            Width           =   3135
         End
         Begin VB.Label Label30 
            Caption         =   "Phone"
            Height          =   255
            Left            =   120
            TabIndex        =   151
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label29 
            Caption         =   "Consultant"
            Height          =   255
            Left            =   5040
            TabIndex        =   145
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label28 
            Caption         =   "App. at"
            Height          =   255
            Left            =   5040
            TabIndex        =   143
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label19 
            Caption         =   "App. on"
            Height          =   255
            Left            =   5040
            TabIndex        =   141
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label18 
            Caption         =   "Booking On"
            Height          =   255
            Left            =   4920
            TabIndex        =   139
            Top             =   1920
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Agent Ref. No."
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   2640
            Width           =   2415
         End
         Begin VB.Label Label32 
            Caption         =   "Agent"
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Patient Name"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label46 
            Caption         =   "Booking ID"
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label47 
            Caption         =   "Payment Method"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label48 
            Caption         =   "Booking User"
            Height          =   255
            Left            =   4920
            TabIndex        =   119
            Top             =   2400
            Width           =   2415
         End
         Begin VB.Label Label49 
            Caption         =   "Cancel / Refund"
            Height          =   255
            Left            =   10080
            TabIndex        =   118
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label50 
            Caption         =   "Credit Settling"
            Height          =   255
            Left            =   10080
            TabIndex        =   117
            Top             =   1800
            Width           =   2415
         End
      End
      Begin MSFlexGridLib.MSFlexGrid gridPatient 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   40
         Top             =   960
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4260
         _Version        =   393216
      End
      Begin VB.Frame FrameCancellations 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   81
         Top             =   360
         Width           =   9615
         Begin VB.Frame Frame2 
            Height          =   975
            Left            =   7320
            TabIndex        =   125
            Top             =   1440
            Width           =   2175
            Begin VB.OptionButton OptionDoNotPrintCancel 
               Caption         =   "Do not print"
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   600
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton OptionPrintCancel 
               Caption         =   "Print"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.OptionButton OptionRepayAgent 
            Caption         =   "Repay Agent"
            Height          =   255
            Left            =   7560
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.OptionButton OptionRepayPatient 
            Caption         =   "Repay Patient"
            Height          =   255
            Left            =   7560
            TabIndex        =   19
            Top             =   840
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtStaffRepayC 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   13
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtInstitutionRepayC 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   14
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtOtherRepayC 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   15
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtRepayTotalC 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtCancellationComments 
            Height          =   375
            Left            =   2040
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   2520
            Width           =   5175
         End
         Begin btButtonEx.ButtonEx bttnCancellation 
            Height          =   375
            Left            =   7800
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Cancel Booking"
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
         Begin VB.Label lblStaffFeePaidC 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2040
            TabIndex        =   92
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label51 
            Caption         =   "Doctor Fee :"
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
            Left            =   240
            TabIndex        =   97
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label11 
            Caption         =   "Institution Fee:"
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
            Left            =   240
            TabIndex        =   96
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Paid Amount"
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
            Left            =   2040
            TabIndex        =   94
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Re-Payment"
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
            Left            =   5880
            TabIndex        =   93
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblInstitutionFeePaidC 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2040
            TabIndex        =   91
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblOtherFeePaidC 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2040
            TabIndex        =   90
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblTotalPaidC 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2040
            TabIndex        =   89
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Total"
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
            Left            =   240
            TabIndex        =   88
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label lblPreviousStaffRepayC 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   3840
            TabIndex        =   87
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblPreviousInstitutionRepayC 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   3840
            TabIndex        =   86
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblPreviousOtherRepayC 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   3840
            TabIndex        =   85
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblPreviousTotalRepayC 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   3840
            TabIndex        =   84
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Previous Repays"
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
            Left            =   3480
            TabIndex        =   83
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Comments"
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
            Left            =   360
            TabIndex        =   82
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Other Fee:"
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
            Left            =   240
            TabIndex        =   95
            Top             =   1440
            Width           =   2055
         End
      End
      Begin VB.Frame FrameSettleCredit 
         Height          =   3015
         Left            =   120
         TabIndex        =   71
         Top             =   360
         Width           =   9615
         Begin VB.Frame Frame4 
            Height          =   975
            Left            =   7320
            TabIndex        =   127
            Top             =   1200
            Width           =   2175
            Begin VB.OptionButton OptionSettleCreditPrint 
               Caption         =   "Print"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   1935
            End
            Begin VB.OptionButton OptionSettleCreditDoNotPrint 
               Caption         =   "Do not print"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   600
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin btButtonEx.ButtonEx bttnCashSettle 
            Height          =   375
            Left            =   7320
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Settle Credit"
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
         Begin VB.Label Label45 
            Caption         =   "Doctor Fee :"
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
            Left            =   360
            TabIndex        =   80
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label44 
            Caption         =   "Institution Fee:"
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
            Left            =   360
            TabIndex        =   79
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label43 
            Caption         =   "Other Fee:"
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
            Left            =   360
            TabIndex        =   78
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lblDoctorFeeToPay 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2520
            TabIndex        =   77
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblHospitalFeeToPay 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2520
            TabIndex        =   76
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblOtherFeeToPay 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2520
            TabIndex        =   75
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblTotalFeeToPay 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2520
            TabIndex        =   74
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label37 
            Caption         =   "Total"
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
            Left            =   360
            TabIndex        =   73
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            Caption         =   "To Pay"
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
            Left            =   2520
            TabIndex        =   72
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame FrameReprints 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   59
         Top             =   360
         Width           =   9615
         Begin btButtonEx.ButtonEx bttnReprint 
            Height          =   375
            Left            =   5160
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Reprint"
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
         Begin VB.Label lblTotalFeePaid 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2400
            TabIndex        =   111
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblOtherFeePaid 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2400
            TabIndex        =   110
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblHospitalFeePaid 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2400
            TabIndex        =   109
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblDoctorFeePaid 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2400
            TabIndex        =   108
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblPaymentMethod 
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
            Left            =   2640
            TabIndex        =   70
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label Label36 
            Caption         =   "Payment Method"
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
            Left            =   360
            TabIndex        =   69
            Top             =   2400
            Width           =   1935
         End
         Begin VB.Label Label34 
            Caption         =   "Total"
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
            Left            =   360
            TabIndex        =   68
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            Caption         =   "Paid"
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
            Left            =   2520
            TabIndex        =   63
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label24 
            Caption         =   "Other Fee:"
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
            Left            =   360
            TabIndex        =   62
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label23 
            Caption         =   "Institution Fee:"
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
            Left            =   360
            TabIndex        =   61
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label22 
            Caption         =   "Doctor Fee :"
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
            Left            =   360
            TabIndex        =   60
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame FrameRefunds 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   56
         Top             =   360
         Width           =   9615
         Begin VB.Frame Frame3 
            Height          =   975
            Left            =   7320
            TabIndex        =   126
            Top             =   1440
            Width           =   2175
            Begin VB.OptionButton OptionRefundPrint 
               Caption         =   "Print"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   1935
            End
            Begin VB.OptionButton OptionRefundDoNotPrint 
               Caption         =   "Do not print"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   600
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin VB.TextBox txtRefundComments 
            Height          =   375
            Left            =   2040
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   2520
            Width           =   5175
         End
         Begin VB.TextBox txtStaffRepayR 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   23
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtInstitutionRepayR 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   24
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtOtherRepayR 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   25
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtRepayTotalR 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   1920
            Width           =   1335
         End
         Begin btButtonEx.ButtonEx bttnRefund 
            Height          =   375
            Left            =   7800
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
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
         Begin VB.Label Label17 
            Caption         =   "Doctor Fee :"
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
            Left            =   240
            TabIndex        =   57
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label16 
            Caption         =   "Institution Fee:"
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
            Left            =   240
            TabIndex        =   64
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Other Fee:"
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
            Left            =   240
            TabIndex        =   65
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "Re-Payment"
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
            Left            =   5880
            TabIndex        =   66
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblStaffFeePaidR 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2040
            TabIndex        =   107
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblInstitutionFeePaidR 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2040
            TabIndex        =   106
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblOtherFeePaidR 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2040
            TabIndex        =   105
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblTotalPaidR 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   2040
            TabIndex        =   104
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Total"
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
            Left            =   240
            TabIndex        =   103
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblPreviousStaffRepayR 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   3840
            TabIndex        =   102
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblPreviousInstitutionRepayR 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   3840
            TabIndex        =   101
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblPreviousOtherRepayR 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   3840
            TabIndex        =   100
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblPreviousTotalRepayR 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   3840
            TabIndex        =   99
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Previous Repays"
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
            Left            =   3360
            TabIndex        =   98
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label21 
            Caption         =   "Comments"
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
            Left            =   360
            TabIndex        =   58
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label Label14 
            Caption         =   "Paid Amount"
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
            Left            =   2160
            TabIndex        =   67
            Top             =   240
            Width           =   1575
         End
      End
      Begin btButtonEx.ButtonEx bttnSearch 
         Height          =   375
         Left            =   -61320
         TabIndex        =   42
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Search"
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
         Left            =   -68280
         TabIndex        =   47
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   73334785
         CurrentDate     =   39446
      End
      Begin btButtonEx.ButtonEx bttnNurseView 
         Height          =   375
         Left            =   -74640
         TabIndex        =   43
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Nurse View"
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
      Begin btButtonEx.ButtonEx bttnDoctorView 
         Height          =   375
         Left            =   -74640
         TabIndex        =   44
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Doctor View"
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
      Begin btButtonEx.ButtonEx bttnChangeName 
         Height          =   375
         Left            =   -71520
         TabIndex        =   37
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Change Name"
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
      Begin btButtonEx.ButtonEx bttnMarkAbsent 
         Height          =   375
         Left            =   -74760
         TabIndex        =   34
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Mark as absent"
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
      Begin btButtonEx.ButtonEx bttnMarkPresent 
         Height          =   375
         Left            =   -74760
         TabIndex        =   35
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Mark as present"
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
      Begin btButtonEx.ButtonEx bttnAllPatients 
         Height          =   375
         Left            =   -74640
         TabIndex        =   45
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "All Patients"
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
      Begin btButtonEx.ButtonEx bttnAllDoctors 
         Height          =   375
         Left            =   -74640
         TabIndex        =   46
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "All Doctors"
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
      Begin btButtonEx.ButtonEx bttnAgentRefSearch 
         Height          =   375
         Left            =   -61320
         TabIndex        =   147
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Search"
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
      Begin btButtonEx.ButtonEx bttnAllSecessionPatients 
         Height          =   375
         Left            =   -74640
         TabIndex        =   149
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "All Secessions"
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
      Begin VB.Label Label31 
         Caption         =   "Agent Ref. No."
         Height          =   375
         Left            =   -64080
         TabIndex        =   148
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "Date"
         Height          =   255
         Left            =   -74880
         TabIndex        =   128
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Booking ID"
         Height          =   375
         Left            =   -64080
         TabIndex        =   124
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblAgentName 
         Height          =   375
         Left            =   -74760
         TabIndex        =   114
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label27 
         Caption         =   "Name"
         Height          =   255
         Left            =   -72600
         TabIndex        =   112
         Top             =   480
         Width           =   975
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   13080
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   8640
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.ListBox ListSecessionMax 
      Height          =   4380
      Left            =   13920
      TabIndex        =   54
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListRoomNo 
      Height          =   2220
      Left            =   4680
      TabIndex        =   113
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Left            =   6360
      TabIndex        =   134
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "No.    Name             Paid Can/Ref  Agent  P/Ab"
      Height          =   255
      Left            =   9960
      TabIndex        =   132
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Secession"
      Height          =   255
      Left            =   6360
      TabIndex        =   131
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Consultant"
      Height          =   255
      Left            =   3000
      TabIndex        =   130
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Speciality"
      Height          =   255
      Left            =   240
      TabIndex        =   129
      Top             =   120
      Width           =   2535
   End
   Begin VB.Shape BoxPatients 
      BackStyle       =   1  'Opaque
      Height          =   4770
      Left            =   9480
      Top             =   120
      Width           =   5775
   End
   Begin VB.Shape BoxDates 
      BackStyle       =   1  'Opaque
      Height          =   4770
      Left            =   6240
      Top             =   120
      Width           =   3255
   End
   Begin VB.Shape BoxConsultant 
      BackStyle       =   1  'Opaque
      Height          =   4770
      Left            =   2880
      Top             =   120
      Width           =   3375
   End
   Begin VB.Shape BoxSpeciality 
      BackStyle       =   1  'Opaque
      Height          =   4770
      Left            =   120
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmLocatePatients"
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
    Dim TemPhone As String
    Dim TemConsultant As String
    Dim TemNonCancelledVisits As Long
    Dim TemBillId As Long
    Dim TemPreviousDate As Date
    Dim TemTextForList As String

Private Sub bttnAgentRefSearch_Click()
    Call SearchAgentRefNo
    SSTab2.Tab = 0
    txtSearchBookingID.Text = Empty
    txtSearchAgentRefNo.Text = Empty
End Sub

Private Sub SearchAgentRefNo()
Dim TemResponce As Integer

With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    .Source = "Select * from tblpatientfacility where agentrefno = '" & txtSearchAgentRefNo.Text & "'"
    .Open
    
    If .RecordCount = 0 Then
        TemResponce = MsgBox("There is no such agent referance. Please re-check", vbCritical, "Wrong referance No.")
        .Close
        Exit Sub
    Else
        txtSearchBookingID.Text = !patientfacility_ID
    End If
    
    Call ListAllConsultants
   
    Dim TemNum As Long
    
    If ListConsultants.ListCount = 0 Then
        TemResponce = MsgBox("The consultant is deleted", vbCritical, "Consultant Deleted")
        Exit Sub
    End If
    
    MonthView2.Value = !AppointmentDate
        
    Dim ConsultantFound As Boolean
    ConsultantFound = False
    For TemNum = 0 To ListConsultantIDs.ListCount - 1
        ListConsultantIDs.ListIndex = TemNum
        If Val(ListConsultantIDs.Text) = !Staff_ID Then
            ListConsultants.ListIndex = TemNum
            ListConsultants_Click
            TemNum = ListConsultantIDs.ListCount
            ConsultantFound = True
        End If
    Next
    If ConsultantFound = False Then
        TemResponce = MsgBox("The consultant the patient booked is deleted", vbCritical, "Deleted")
        Exit Sub
    End If
    
    If ListDatesAndSecessions.ListCount = 0 Then
        TemResponce = MsgBox("The booking date for the patient is deleted", vbCritical, "Deleted")
        Exit Sub
    End If
    
    
    Dim DateFound As Boolean
    Dim FoundSecession As Long
    DateFound = False
    
    For TemNum = 0 To ListSecessionIDs.ListCount - 1
        ListSecessionIDs.ListIndex = TemNum
            ListSecessionIDs.ListIndex = TemNum
            If ListSecessionIDs.Text = !Secession Then
                FoundSecession = TemNum
                DateFound = True
                TemNum = ListSecessionIDs.ListCount - 1
            End If

    Next
    
    If DateFound = True Then
        ListDatesAndSecessions.ListIndex = FoundSecession
        ListDatesAndSecessions_Click
    Else
        ListDatesAndSecessions.ListIndex = 0
        ListDatesAndSecessions_Click
    End If
    
    If ListPatientFacilities.ListCount = 0 Then Exit Sub
    
    For TemNum = 0 To ListPatientFacilities.ListCount - 1
        ListPatientFacilityIDs.ListIndex = TemNum
        If Val(ListPatientFacilityIDs.Text) = Val(txtSearchBookingID.Text) Then
            ListPatientFacilities.ListIndex = TemNum
            ListPatientFacilities_Click
            TemNum = ListPatientFacilities.ListCount
        End If
    Next

End With

End Sub


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
    DTPickerFindPatientDate.Value = Date
    MonthView1.Value = Date
    MonthView2.Value = Date
    Call FillPatientName
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
    FrameCancellations.Enabled = False
    FrameRefunds.Enabled = False
    FrameReprints.Enabled = False
    FrameSettleCredit.Enabled = False
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
    ClearPatientDetails
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
    MonthView1.Value = MonthView2.Value
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

Public Sub ListDatesAndSecessions_Click()
    ListSecessionIDs.ListIndex = ListDatesAndSecessions.ListIndex
    TemAppointmentDate = MonthView2.Value
    MonthView1.Value = TemAppointmentDate
    Call ClearPatientDetails
    Call FormatGridPatients
    If Not IsNumeric(ListSecessionIDs.Text) And ListSecessionIDs.Text <> "All" Then Exit Sub
    Call FillGridPatients
    DTPickerFindPatientDate.Value = MonthView2.Value
End Sub

Private Sub ListPatientFacilities_Click()
    ListPatientFacilityIDs.ListIndex = ListPatientFacilities.ListIndex
    If IsNumeric(ListPatientFacilityIDs.Text) Then
        FrameCancellations.Enabled = True
        FrameRefunds.Enabled = True
        FrameReprints.Enabled = True
        FrameSettleCredit.Enabled = True
        TemPatientFacilityID = Val(ListPatientFacilityIDs.Text)
        Call ClearPatientDetails
        Call GetPatientDetails
        SSTab2.Tab = 0
    Else
        FrameCancellations.Enabled = True
        FrameRefunds.Enabled = True
        FrameReprints.Enabled = True
        FrameSettleCredit.Enabled = True
        TemPatientFacilityID = Empty
        Call ClearPatientDetails
        SSTab2.Tab = 3
    End If
End Sub

Private Sub GetPatientDetails()
    With DataEnvironment1.rssqlTem8
        If .State = 1 Then .Close
        .Source = "select * from tblpatientfacility where patientfacility_ID = " & TemPatientFacilityID
        .Open
        If .RecordCount = 0 Then Exit Sub
    TemPatientID = !patientid
    TemPatient = FindPatientByID(!patientid)
    TemPhone = FindPhoneByID(!patientid)
    TemPatientFacilityID = !patientfacility_ID
    TemAppointmentDate = Format(!AppointmentDate, DefaultLongDate)
    TemAppointmentTime = !appointmenttime
    txtBookingOn.Text = Format(!BookingDate, "dd MMM yyyy") & " - " & Format(!BookingTime, "hh:mm AMPM")
    txtAppOn.Text = Format(!AppointmentDate, DefaultLongDate)
    txtAppAt.Text = !appointmenttime
    TemDaySerial = !DaySerial
    txtBookedPatientName.Text = TemPatient
    txtPhone.Text = TemPhone
    txtNameChange.Text = TemPatient
    txtBookedPatientID.Text = TemPatientID
    txtBookingID.Text = TemPatientFacilityID
    txtPaymentMethod.Text = !PaymentMode
    txtBookingUser.Text = FindStaffFromID(!user_ID)
    txtConsultant.Text = ListConsultants.Text
    If Not IsNull(!PersonalFee) Then
        TemDoctorFee = !PersonalFee
        lblDoctorFeePaid.Caption = Format(!PersonalFee, "0.00")
        lblStaffFeePaidC.Caption = Format(!PersonalFee, "0.00")
        lblStaffFeePaidR.Caption = Format(!PersonalFee, "0.00")
        txtStaffRepayC.Text = Format(!PersonalFee, "0.00")
        txtStaffRepayR.Text = Format(!PersonalFee, "0.00")
    Else
        lblDoctorFeePaid.Caption = Format(0, "0.00")
        lblStaffFeePaidC.Caption = Format(0, "0.00")
        lblStaffFeePaidR.Caption = Format(0, "0.00")
    End If
    If Not IsNull(!InstitutionFee) Then
        TemInstitutionFee = !InstitutionFee
        lblHospitalFeePaid.Caption = Format(!InstitutionFee, "0.00")
        lblInstitutionFeePaidC.Caption = Format(!InstitutionFee, "0.00")
        lblInstitutionFeePaidR.Caption = Format(!InstitutionFee, "0.00")
        txtInstitutionRepayC.Text = Format(!InstitutionFee, "0.00")
    Else
        lblHospitalFeePaid.Caption = "0.00"
        lblInstitutionFeePaidC.Caption = "0.00"
        lblInstitutionFeePaidR.Caption = Format(0, "0.00")
    End If
    If Not IsNull(!otherfee) Then
        TemOtherFee = !otherfee
        lblOtherFeePaid.Caption = Format(!otherfee, "0.00")
        lblOtherFeePaidR.Caption = Format(!otherfee, "0.00")
        lblOtherFeePaidC.Caption = Format(!otherfee, "0.00")
    Else
        lblOtherFeePaid.Caption = "0.00"
        lblOtherFeePaidR.Caption = "0.00"
        lblOtherFeePaidC.Caption = Format(0, "0.00")
    End If
    If Not IsNull(!totalfee) Then
        TemTotalPayment = !otherfee
        lblTotalFeePaid.Caption = Format(!totalfee, "0.00")
        lblTotalPaidC.Caption = Format(!totalfee, "0.00")
        lblTotalPaidR.Caption = Format(!totalfee, "0.00")
    Else
        lblTotalFeePaid.Caption = "0.00"
        lblTotalPaidC.Caption = "0.00"
        lblTotalPaidR.Caption = "0.00"
    End If
    
    If Not IsNull(!Personalrefund) Then
        
        lblPreviousStaffRepayC.Caption = Format(!Personalrefund, "0.00")
        lblPreviousStaffRepayR.Caption = Format(!Personalrefund, "0.00")
    Else
        lblPreviousStaffRepayC.Caption = "0.00"
        lblPreviousStaffRepayR.Caption = Format(0, "0.00")
    End If
        
    If Not IsNull(!institutionrefund) Then
        lblPreviousInstitutionRepayC.Caption = Format(!institutionrefund, "0.00")
        lblPreviousInstitutionRepayR.Caption = Format(!institutionrefund, "0.00")
    Else
        lblPreviousInstitutionRepayC.Caption = "0.00"
        lblPreviousInstitutionRepayR.Caption = "0.00"
    End If
        
    If Not IsNull(!otherrefund) Then
        lblPreviousOtherRepayC.Caption = Format(!otherrefund, "0.00")
        lblPreviousOtherRepayR.Caption = Format(!otherrefund, "0.00")
    Else
        lblPreviousOtherRepayC.Caption = "0.00"
        lblPreviousOtherRepayR.Caption = "0.00"
    End If
    
    If Not IsNull(!totalrefund) Then
        lblPreviousTotalRepayC.Caption = Format(!totalrefund, "0.00")
        lblPreviousTotalRepayR.Caption = Format(!totalrefund, "0.00")
    Else
        lblPreviousTotalRepayC.Caption = "0.00"
        lblPreviousTotalRepayR.Caption = "0.00"
    End If
    
    If Not IsNull(!PersonalFeeToPay) Then
        lblDoctorFeeToPay.Caption = Format(!PersonalFeeToPay, "0.00")
    Else
        lblDoctorFeeToPay.Caption = "0.00"
    End If
    
    If Not IsNull(!InstitutionFeeToPay) Then
        lblHospitalFeeToPay.Caption = Format(!InstitutionFeeToPay, "0.00")
    Else
        lblHospitalFeeToPay.Caption = "0.00"
    End If
    
    If Not IsNull(!otherfeetopay) Then
        lblOtherFeeToPay.Caption = Format(!otherfeetopay, "0.00")
    Else
        lblOtherFeeToPay.Caption = "0.00"
    End If
    
    If Not IsNull(!totalfeetopay) Then
        lblTotalFeeToPay.Caption = Format(!totalfeetopay, "0.00")
    Else
        lblTotalFeeToPay.Caption = "0.00"
    End If
    
    If Not IsNull(!Personalrefund) Then
        lblPreviousStaffRepayR.Caption = Format(!Personalrefund, "0.00")
    Else
        lblPreviousStaffRepayR.Caption = "0.00"
    End If
    If Not IsNull(!institutionrefund) Then
        lblPreviousInstitutionRepayR.Caption = Format(!institutionrefund, "0.00")
    Else
        lblPreviousInstitutionRepayR.Caption = "0.00"
    End If
    If Not IsNull(!otherrefund) Then
        lblPreviousOtherRepayR.Caption = Format(!otherrefund, "0.00")
    Else
        lblPreviousOtherRepayR.Caption = "0.00"
    End If
    
    If Not IsNull(!Personalrefund) Then
        lblPreviousStaffRepayC.Caption = Format(!Personalrefund, "0.00")
    Else
        lblPreviousStaffRepayC.Caption = "0.00"
    End If
    If Not IsNull(!institutionrefund) Then
        lblPreviousInstitutionRepayC.Caption = Format(!institutionrefund, "0.00")
    Else
        lblPreviousInstitutionRepayC.Caption = "0.00"
    End If
    If Not IsNull(!otherrefund) Then
        lblPreviousOtherRepayC.Caption = Format(!otherrefund, "0.00")
    Else
        lblPreviousOtherRepayC.Caption = "0.00"
    End If
    
    If !PaymentMode = "Credit" Then
        If IsNull(!CreditSettleUser_ID) Or !CreditSettleUser_ID = 0 Then
            txtCreditSettle.Text = "The booking done for credit. The patient has to pay Rs." & Format(!totalfeetopay, "0.00")
            FrameCancellations.Enabled = False
            FrameRefunds.Enabled = False
            FrameSettleCredit.Enabled = True
        Else
            txtCreditSettle.Text = "The booking done for credit. The patient had settled it by paying Rs." & Format(!totalfee, "0.00") & " to " & FindStaffFromID(!CreditSettleUser_ID)
            txtCreditSettle.Text = "No credit issues"
            FrameCancellations.Enabled = True
            FrameRefunds.Enabled = True
            FrameSettleCredit.Enabled = False
        End If
        If IsNull(!CreditStaff_ID) = False Then
            If !CreditStaff_ID <> 0 Then
                txtCreditSettle.Text = txtCreditSettle.Text + " (Staff Booking for " & FindStaffFromID(!CreditStaff_ID) & ")"
            End If
        End If
    ElseIf !PaymentMode = "Agent" Then
        txtAgentAndCode.Text = FindAgentFromID(!Agent_ID)
        txtAgentCode.Text = FindAgentCodeFromID(!Agent_ID)
        txtAgentRefNo.Text = !AgentRefNo
        txtCreditSettle.Text = "No credit issues"
        FrameCancellations.Enabled = True
        FrameRefunds.Enabled = True
        FrameSettleCredit.Enabled = False
    Else
        txtCreditSettle.Text = "No credit issues"
        FrameCancellations.Enabled = True
        FrameRefunds.Enabled = True
        FrameSettleCredit.Enabled = False
    End If
    
    If !Cancelled = True Then
        txtCancelRefund.Text = "Cancelled on " & Format(!repaydate, DefaultLongDate) & " at " & Format(!RepayTime, "hh:mm AMPM") & " by " & FindStaffFromID(!repayUser_ID) & ". Rs. " & Format(!totalrefund, "0.00") & " was repaied."
        FrameCancellations.Enabled = False
        FrameRefunds.Enabled = False
        txtRepayTotalC.Text = Empty
        txtRepayTotalR.Text = Empty
        txtStaffRepayC.Text = Empty
        txtStaffRepayR.Text = Empty
        txtInstitutionRepayC.Text = Empty
        txtInstitutionRepayR.Text = Empty
        txtOtherRepayC.Text = Empty
    ElseIf !Refund = True Then
        txtCancelRefund.Text = "Refunded on " & Format(!repaydate, DefaultLongDate) & " by  at " & Format(!RepayTime, "hh:mm AMPM") & "  " & FindStaffFromID(!repayUser_ID) & ". Rs. " & Format(!totalrefund, "0.00") & " was repaied."
        FrameCancellations.Enabled = False
        FrameRefunds.Enabled = False
        txtRepayTotalC.Text = Empty
        txtRepayTotalR.Text = Empty
        txtStaffRepayC.Text = Empty
        txtStaffRepayR.Text = Empty
        txtInstitutionRepayC.Text = Empty
        txtInstitutionRepayR.Text = Empty
        txtOtherRepayC.Text = Empty
    Else
        txtCancelRefund.Text = "No cencellations or refunds"
        FrameCancellations.Enabled = True
        FrameRefunds.Enabled = True
    End If
    If !PaymentMode = "Agent" Then
        OptionRepayAgent.Visible = True
        OptionRepayPatient.Visible = True
        OptionRepayAgent.Value = False
        OptionRepayPatient.Value = False
    Else
        OptionRepayAgent.Visible = False
        OptionRepayPatient.Visible = False
    End If

' **************************

        .Close
    End With
End Sub


Private Sub ClearPatientDetails()
    txtBookingOn.Text = Empty
    txtAppOn.Text = Empty
    txtAppAt.Text = Empty
    txtBookedPatientID.Text = Empty
    txtBookedPatientName.Text = Empty
    txtPhone.Text = Empty
    txtNameChange.Text = Empty
    txtBookingID.Text = Empty
    txtBookingUser.Text = Empty
    txtCancellationComments.Text = Empty
    txtCancelRefund.Text = Empty
    txtCancelRefund.Text = Empty
    txtCreditSettle.Text = Empty
    txtInstitutionRepayC.Text = Empty
    txtInstitutionRepayR.Text = Empty
    txtOtherRepayC.Text = Empty
    txtOtherRepayR.Text = Empty
    txtPaymentMethod.Text = Empty
    txtRepayTotalC.Text = Empty
    txtRepayTotalR.Text = Empty
    txtStaffRepayC.Text = Empty
    txtStaffRepayR.Text = Empty
    txtAgentAndCode.Text = Empty
    txtAgentCode.Text = Empty
    txtAgentRefNo.Text = Empty
    txtConsultant.Text = Empty
'    lblAgentAmount.Caption = Empty
'    lblCashDue.Caption = Empty
'    lblCredit.Caption = Empty
    lblDoctorFeePaid.Caption = Empty
    lblDoctorFeeToPay.Caption = Empty
    lblHospitalFeePaid.Caption = Empty
    lblHospitalFeeToPay.Caption = Empty
    lblInstitutionFeePaidC.Caption = Empty
    lblInstitutionFeePaidR.Caption = Empty
    lblOtherFeePaid.Caption = Empty
    lblOtherFeePaidC.Caption = Empty
    lblOtherFeePaidR.Caption = Empty
    lblOtherFeeToPay.Caption = Empty
    lblPaymentMethod.Caption = Empty
    lblPreviousInstitutionRepayC.Caption = Empty
    lblPreviousInstitutionRepayR.Caption = Empty
    lblPreviousOtherRepayC.Caption = Empty
    lblPreviousOtherRepayR.Caption = Empty
    lblPreviousStaffRepayC.Caption = Empty
    lblPreviousTotalRepayC.Caption = Empty
    lblPreviousTotalRepayR.Caption = Empty
    lblStaffFeePaidC.Caption = Empty
    lblStaffFeePaidR.Caption = Empty
    lblTotalFeePaid.Caption = Empty
    lblTotalFeeToPay.Caption = Empty
    lblTotalPaidC.Caption = Empty
    lblTotalPaidR.Caption = Empty

End Sub


Private Sub FillGridPatients()
    Dim TemTextForList As String
    Call ClearPatientDetails

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
            
' *************************************
            
'            If !FullyPaid = 1 Then
'                TemTextForList = TemTextForList & vbTab & "Paid"
'            Else
'                TemTextForList = TemTextForList & vbTab & Space(4)
'            End If
            
' *************************************
            
            
            
            
            If !PaymentMethod_ID = 1 Then
                TemTextForList = TemTextForList & Space(2) & "Ch" & Space(5)
            
            ElseIf !PaymentMethod_ID = 2 Then
                If Not IsNull(!Agent_ID) Then
                    If !Agent_ID <> 0 Then
                        TemTextForList = TemTextForList & Space(2) & "Ag" & Space(1) & Left(FindAgentCodeFromID(!Agent_ID) & Space(4), 4)
                    Else
                        TemTextForList = TemTextForList & Space(2) & "Ag" & Space(1) & Space(4)
                    End If
                Else
                    TemTextForList = TemTextForList & Space(2) & "Ag" & Space(5)
                End If
            
            ElseIf !PaymentMethod_ID = 4 Then
                If Not IsNull(!CreditStaff_ID) Then
                    If !CreditStaff_ID <> 0 Then
                        TemTextForList = TemTextForList & Space(2) & "St" & Space(1) & Left(FindStaffCodeFromID(!CreditStaff_ID) & Space(4), 4)
                    Else
                        TemTextForList = TemTextForList & Space(2) & "Tp" & Space(1) & Space(4)
                    End If
                Else
                    TemTextForList = TemTextForList & Space(2) & "Tp" & Space(5)
                End If
            End If
            
            
            
            If !FullyPaid = True Then
                TemTextForList = TemTextForList & Space(2) & "Pd"
            Else
                TemTextForList = TemTextForList & Space(2) & "Np"
            End If
            
            
' *************************************
            
            If !Cancelled = True Then
                TemTextForList = TemTextForList & Space(2) & "Cancel"
            ElseIf !Refund = True Then
                TemTextForList = TemTextForList & Space(2) & "Refund"
            Else
                TemTextForList = TemTextForList & Space(2) & Space(6)
            End If
            If Not IsNull(!Agent_ID) Then
                If !Agent_ID <> 0 Then
                    TemTextForList = TemTextForList & Space(2) & Left(FindAgentCodeFromID(!Agent_ID), 3)
                Else
                    TemTextForList = TemTextForList & Space(2) & Space(3)
                End If
            End If
            If !patientabsent = True Then
                TemTextForList = TemTextForList & Space(2) & "Ab"
            Else
                TemTextForList = TemTextForList & Space(2) & " "
            End If
            ListPatientFacilities.AddItem TemTextForList
            ListPatientFacilityIDs.AddItem !patientfacility_ID
            .MoveNext
        Wend
    End With

End Sub














Private Sub SetBillPrinter()
    CSetPrinter.SetPrinterAsDefault (BillPrinterName)
End Sub

Private Sub SetBillPaper()
Dim TemResponce As Long
Dim RetVal As Integer
RetVal = SelectForm(BillPaperName, Me.hwnd)
Select Case RetVal
    Case FORM_NOT_SELECTED   ' 0
        TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
    Case FORM_SELECTED   ' 1
        Call SelectPrint
    Case FORM_ADDED   ' 2
        TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
End Select
End Sub

Private Sub SelectPrint()
        If PrintingOnBlankPaper = True Then
            BillPrint2
        ElseIf PrintingOnPrintedPaper = True Then
            BillPrint3
        End If
End Sub

Private Sub BillPrint3()
    Dim TemRows As Long

    With DataEnvironment1.rssqlTem15
        If .State = 1 Then .Close
        .Source = "SELECT * from tblchannellingPrintingPreferances"
        .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        
        Dim TemBoolean As Boolean
        Printer.Font.Name = "Arial"
        Printer.Font.Size = 10
        Printer.Font.Bold = False
        'Printer.Line (100, 100)-(Printer.ScaleWidth - 100, Printer.ScaleHeight - 100)
        TemBoolean = PrintingPlainText(!date1x, !date1y, Date)
        'TemBoolean = PrintingPlainText(!refno1x, !refno1y, TemPatientFacilityID)
        TemBoolean = PrintingPlainText(!consultant1x, !consultant1y, UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text))))
        TemBoolean = PrintingPlainText(!patient1x, !patient1y, TemPatient)
        TemBoolean = PrintingPlainText(!appointon1x, !appointon1y, Format(ListDates.Text, DefaultShortDate))
        TemBoolean = PrintingPlainText(!at1x, !at1y, TemAppointmentTime)
        TemBoolean = PrintingPlainText(!drsfee1x, !drsfee1y, Format(TemDoctorFee, "0.00"))
        TemBoolean = PrintingPlainText(!total1x, !total1y, Format(TemDoctorFee + TemInstitutionFee, "0.00"))
        TemBoolean = PrintingPlainText(!hospchg1x, !hospchg1y, Format(TemInstitutionFee, "0.00"))
        TemBoolean = PrintingPlainText(!receptionist1x, !receptionist1y, UserName)
        TemBoolean = PrintingPlainText(!roomno1x, !roomno1y, ListRoomNo.Text)
        If txtPaymentMethod.Text = "Agent" Then
            TemBoolean = PrintingPlainText(!agentcode1x, !agentcode1y, "(" & txtAgentCode.Text & ")")
            TemBoolean = PrintingPlainText(!agentrefno1x, !agentrefno1y, txtAgentRefNo.Text)
        End If
        TemBoolean = PrintingPlainText(!date2x, !date2y, Date)
        'TemBoolean = PrintingPlainText(!refno2x, !refno2y, TemPatientFacilityID)
        TemBoolean = PrintingPlainText(!consultant2x, !consultant2y, UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text))))
        TemBoolean = PrintingPlainText(!patient2x, !patient2y, TemPatient)
        TemBoolean = PrintingPlainText(!appointon2x, !appointon2y, Format(ListDates.Text, DefaultShortDate))
        TemBoolean = PrintingPlainText(!at2x, !at2y, TemAppointmentTime)
        TemBoolean = PrintingPlainText(!drsfee2x, !drsfee2y, Format(TemDoctorFee, "0.00"))
        TemBoolean = PrintingPlainText(!total2x, !total2y, Format(TemDoctorFee + TemInstitutionFee, "0.00"))
        TemBoolean = PrintingPlainText(!hospchg2x, !hospchg2y, Format(TemInstitutionFee, "0.00"))
        TemBoolean = PrintingPlainText(!receptionist2x, !receptionist2y, UserName)
        TemBoolean = PrintingPlainText(!roomno2x, !roomno2y, ListRoomNo.Text)
        If txtPaymentMethod.Text = "Agent" Then
            TemBoolean = PrintingPlainText(!agentcode2x, !agentcode2y, "(" & txtAgentCode.Text & ")")
            TemBoolean = PrintingPlainText(!agentrefno2x, !agentrefno2y, txtAgentRefNo.Text)
        End If
        Printer.Font.Size = 16
        Printer.Font.Bold = True
        Printer.Font.Name = "Arial"
        TemBoolean = PrintingPlainText(!appono1x, !appono1y, TemDaySerial)
        TemBoolean = PrintingPlainText(!appono2x, !appono2y, TemDaySerial)
        .Close
    End With
    Printer.EndDoc
End Sub

Private Sub UpdatePatientCredit()
With DataEnvironment1.rssqlTem7
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientmaindetails where patient_ID = " & TemPatientID
    .Open
    If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!Credit) Then
            !Credit = !Credit - TemDoctorFee - TemInstitutionFee
        Else
            !Credit = 0 - TemDoctorFee - TemInstitutionFee
        End If
    .Update
    .Close
End With
End Sub

Private Sub BillPrint()
    Dim TemRows As Long

With Printer
        
        .Font = "Bernard MT Condensed"
        Printer.Print
        .FontSize = 14
        Printer.Print Tab(2); InstitutionName
        .FontSize = 12
        Printer.Print Tab(3); InstitutionAddress
        Printer.Print Tab(3); InstitutionTelephone
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        
        .FontName = "Courier"
        .FontSize = 10
        Printer.Print
        
        Dim TemTab1 As Long
        Dim TemTab2 As Long
        Dim TemTab3 As Long
        Dim TemTab4 As Long
        Dim TemTab5 As Long
        Dim TemTab6 As Long
        
        TemTab1 = 2
        TemTab2 = 6
        TemTab3 = 20
        TemTab4 = 25
        TemTab5 = 36
        TemTab6 = 16
        
        Printer.Print Tab(TemTab1); "Patient";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); TemPatient
        Printer.Print
        Printer.Print Tab(TemTab1); "Consultant";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text)))
        Printer.Print
        Printer.Print Tab(TemTab1); "Appo. Date ";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); Format(ListDates.Text, DefaultShortDate)
        
        Printer.Print Tab(TemTab1); "Appo. Time";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); TemAppointmentTime
        
        Printer.Print Tab(TemTab1); "Appo. No.";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); TemDaySerial
        Printer.Print Tab(TemTab1); "Appo. ID";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); TemPatientFacilityID
        Printer.Print
        Printer.Print Tab(TemTab1); "Doctor Fee";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3 + 8 - Len(Format(TemDoctorFee, "0.00"))); Format(TemDoctorFee, "0.00")
        Printer.Print Tab(TemTab1); "Hospital Fee";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3 + 8 - Len(Format(TemInstitutionFee, "0.00"))); Format(TemInstitutionFee, "0.00")
        Printer.Print Tab(TemTab1); "Total Fee";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3 + 8 - Len(Format(TemDoctorFee + TemInstitutionFee, "0.00"))); Format(TemDoctorFee + TemInstitutionFee, "0.00")
        Printer.Print
        Printer.Print Tab(TemTab2); "--------------------"
        Printer.Print Tab(TemTab2); UserName
        Printer.Print Tab(TemTab2); Format(Date, DefaultLongDate)
                
        .EndDoc
    End With
End Sub

Private Sub BillPrint2()


    Dim TemRows As Long

With Printer

        Printer.Font = "Arial Black"
        Printer.Print
        
        Printer.FontSize = 11
        Printer.Print Tab(2); InstitutionName;
        Printer.Print Tab(51); InstitutionName
        
        Printer.FontSize = 9
        Printer.Print Tab(3); InstitutionAddress;
        Printer.Print Tab(64); InstitutionAddress
        
        Printer.Print Tab(3); InstitutionTelephone;
        Printer.Print Tab(64); InstitutionTelephone
        
        Printer.FontName = "Courier"
        Printer.FontSize = 10
        Printer.Print
        
        Dim TemTab1 As Long
        Dim TemTab2 As Long
        Dim TemTab3 As Long
        Dim TemTab4 As Long
        Dim TemTab5 As Long
        Dim TemTab6 As Long
        Dim TemTab7 As Long
        Dim TemTab8 As Long
        Dim TemTab9 As Long
        Dim TemTab10 As Long
        Dim TemTab11 As Long
        Dim TemTab12 As Long
        
        TemTab1 = 2
        TemTab2 = 6
        TemTab3 = 20
        TemTab4 = 25
        TemTab5 = 36
        TemTab6 = 16
        
        Dim Displace As Long
        
        Displace = 73
        
        TemTab7 = 2 + Displace
        TemTab8 = 16 + Displace
        TemTab9 = 20 + Displace
        TemTab10 = 25 + Displace
        TemTab11 = 36 + Displace
        TemTab12 = 16 + Displace
        
        Printer.Font.Bold = True
        Printer.Font.Underline = True
        Printer.Print Tab(TemTab3);
        Printer.Print Tab(TemTab9);
        Printer.Font.Bold = False
        Printer.Font.Underline = False
        
        Printer.Print Tab(TemTab1); "Patient"; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); TemPatient;
        'd
        Printer.Print Tab(TemTab7); "Patient";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); TemPatient
        
        Printer.Print
        Printer.Print Tab(TemTab1); "Consultant"; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text)));
        'd
        Printer.Print Tab(TemTab7); "Consultant";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text)))
        Printer.Print
        Printer.Print Tab(TemTab1); "Appo. Date "; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); Format(ListDates.Text, DefaultLongDate);
        'd
        Printer.Print Tab(TemTab7); "Appo. Date ";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); Format(ListDates.Text, DefaultLongDate)
        
        Printer.Print Tab(TemTab1); "Appo. Time"; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); TemAppointmentTime;
        'd
        Printer.Print Tab(TemTab7); "Appo. Time";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); TemAppointmentTime
        
        Printer.Print Tab(TemTab1); "Appo. No."; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); TemDaySerial;
        
        Printer.Print Tab(TemTab7); "Appo. No.";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); TemDaySerial
        
        Printer.Print Tab(TemTab1); "Room No."; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); ListRoomNo.Text;
        'd
        
        Printer.Print Tab(TemTab7); "Room No.";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); ListRoomNo.Text
        
        
        Printer.Print Tab(TemTab1); "Appo. ID"; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); TemPatientFacilityID;
        'd
        
        Printer.Print Tab(TemTab7); "Appo. ID";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); TemPatientFacilityID
        
        Printer.Print
        
        If txtPaymentMethod = "Cash" Then
        
            Printer.Print Tab(TemTab1); "Doctor Fee";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3 + 8 - Len(Format(TemDoctorFee, "0.00"))); Format(TemDoctorFee, "0.00");
            
            'd
            Printer.Print Tab(TemTab7); "Doctor Fee";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9 + 8 - Len(Format(TemDoctorFee, "0.00"))); Format(TemDoctorFee, "0.00")
            
            
            Printer.Print Tab(TemTab1); "Hospital Fee";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3 + 8 - Len(Format(TemInstitutionFee, "0.00"))); Format(TemInstitutionFee, "0.00");
            
            Printer.Print Tab(TemTab7); "Hospital Fee";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9 + 8 - Len(Format(TemInstitutionFee, "0.00"))); Format(TemInstitutionFee, "0.00")
            
            Printer.Print Tab(TemTab1); "Total Fee";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3 + 8 - Len(Format(TemDoctorFee + TemInstitutionFee, "0.00"))); Format(TemDoctorFee + TemInstitutionFee, "0.00");
            'd
            
            Printer.Print Tab(TemTab7); "Total Fee";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9 + 8 - Len(Format(TemDoctorFee + TemInstitutionFee, "0.00"))); Format(TemDoctorFee + TemInstitutionFee, "0.00")
        
            Printer.Print Tab(TemTab1); "Payment Method";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3); "Cash";
            'd
            
            Printer.Print Tab(TemTab7); "Payment Method";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9); "Cash"
        
        
        ElseIf txtPaymentMethod.Text = "Agent" Then
            Printer.Print Tab(TemTab1); "Doctor Fee";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3 + 8 - Len(Format(TemDoctorFee, "0.00"))); Format(TemDoctorFee, "0.00");
            
            'd
            Printer.Print Tab(TemTab7); "Doctor Fee";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9 + 8 - Len(Format(TemDoctorFee, "0.00"))); Format(TemDoctorFee, "0.00")
            
            
            Printer.Print Tab(TemTab1); "Hospital Fee";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3 + 8 - Len(Format(TemInstitutionFee, "0.00"))); Format(TemInstitutionFee, "0.00");
            
            Printer.Print Tab(TemTab7); "Hospital Fee";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9 + 8 - Len(Format(TemInstitutionFee, "0.00"))); Format(TemInstitutionFee, "0.00")
            
            Printer.Print Tab(TemTab1); "Total Fee";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3 + 8 - Len(Format(TemDoctorFee + TemInstitutionFee, "0.00"))); Format(TemDoctorFee + TemInstitutionFee, "0.00");
            'd
            
            Printer.Print Tab(TemTab7); "Total Fee";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9 + 8 - Len(Format(TemDoctorFee + TemInstitutionFee, "0.00"))); Format(TemDoctorFee + TemInstitutionFee, "0.00")
        
            Printer.Print Tab(TemTab1); "Payment Method";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3); "Agent";
            'd
            
            Printer.Print Tab(TemTab7); "Payment Method";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9); "Agent"
        
        ElseIf txtPaymentMethod.Text = "Cash" Then
            Printer.Print Tab(TemTab1); "Doctor Fee";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3 + 8 - Len(Format(0, "0.00"))); Format(0, "0.00");
            
            'd
            Printer.Print Tab(TemTab7); "Doctor Fee";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9 + 8 - Len(Format(0, "0.00"))); Format(0, "0.00")
            
            
            Printer.Print Tab(TemTab1); "Hospital Fee";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3 + 8 - Len(Format(0, "0.00"))); Format(0, "0.00");
            
            Printer.Print Tab(TemTab7); "Hospital Fee";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9 + 8 - Len(Format(0, "0.00"))); Format(0, "0.00")
            
            Printer.Print Tab(TemTab1); "Total Fee";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3 + 8 - Len(Format(0 + 0, "0.00"))); Format(0 + 0, "0.00");
            'd
            
            Printer.Print Tab(TemTab7); "Total Fee";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9 + 8 - Len(Format(0 + 0, "0.00"))); Format(TemDoctorFee + TemInstitutionFee, "0.00")
        
            Printer.Print Tab(TemTab1); "Payment Method";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3); "Credit";
            'd
            
            Printer.Print Tab(TemTab7); "Payment Method";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9); "Credit"
        
        End If
        
        Printer.Print
        Printer.Print
        
        Printer.Print Tab(TemTab2); "--------------------";
        Printer.Print Tab(TemTab8); "--------------------"
        
        Printer.Print Tab(TemTab2); UserName;
        Printer.Print Tab(TemTab8); UserName
        
        Printer.Print Tab(TemTab2); Format(Date, DefaultLongDate);
        'D
        Printer.Print Tab(TemTab8); Format(Date, DefaultLongDate)
        Printer.EndDoc
    End With

End Sub




Private Sub bttnAllDoctors_Click()
    
    
    
    
    
    
    
    If PartialRepayments = True Then
        Const PreSHape = "SHAPE {"
        Const Sql = "SELECT tblPatientFacility.*, tblDoctor.DoctorListedName, tblPatientMainDetails.FirstName, tblTitle.Title FROM tblTitle RIGHT JOIN ((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblDoctor ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID Where "
        Const PostSHape = "(((tblPatientFacility.HospitalFacility_ID) = 10)) }  AS cmmdTotalDoctorFee COMPUTE cmmdTotalDoctorFee, SUM(cmmdTotalDoctorFee.'PersonalDue') AS DocDue, SUM(cmmdTotalDoctorFee.'InstitutionDue') AS HosDue, SUM(cmmdTotalDoctorFee.'TotalDue') AS TotDue, ANY(cmmdTotalDoctorFee.'DoctorListedName') AS DoctorNameToDisplay, ANY(cmmdTotalDoctorFee.'Title') AS DoctorTitleToDisplay BY 'DoctorListedName' "
        
        With DataEnvironment1
            If .rscmmdTotalDoctorFee_Grouping.State = 1 Then .rscmmdTotalDoctorFee_Grouping.Close
            .Commands!cmmdTotalDoctorFee_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & MonthView2.Value & "' and " & PostSHape
            .cmmdTotalDoctorFee_Grouping
        End With
        With DataReportAllDoctors
            If HospitalDetails = True Then
                .Sections("ReportHeader").Controls.Item("InstitutionName").Caption = InstitutionName
                .Sections("ReportHeader").Controls.Item("InstitutionAddress").Caption = InstitutionAddress
                .Sections("ReportHeader").Controls.Item("lbldate").Caption = Format(Date, DefaultLongDate)
                .Sections("ReportFooter").Controls.Item("ad1").Caption = LongAd
            Else
                .Sections("ReportHeader").Controls.Item("InstitutionName").Caption = Empty
                .Sections("ReportHeader").Controls.Item("InstitutionAddress").Caption = Empty
                .Sections("ReportHeader").Controls.Item("lbldate").Caption = Format(Date, DefaultLongDate)
                .Sections("ReportFooter").Controls.Item("ad1").Caption = LongAd
            End If
            .Show
        End With
    Else
        Const PreSHape1 = "SHAPE {"
        Const Sql1 = " SELECT tblPatientFacility.*, tblDoctor.DoctorListedName, tblPatientMainDetails.FirstName, tblTitle.Title FROM tblTitle RIGHT JOIN ((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblDoctor ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID Where "
        Const PostSHape1 = " (((tblPatientFacility.HospitalFacility_ID) = 10))}  AS NewDoctorView COMPUTE NewDoctorView, COUNT(NewDoctorView.'PatientFacility_ID') AS ValiedVisits, SUM(NewDoctorView.'PersonalDue') AS TotalDoctorDue, ANY(NewDoctorView.'Title') AS DoctorTitle BY 'DoctorListedName' "
        
        With DataEnvironment1
            If .rsNewDoctorView_Grouping.State = 1 Then .rsNewDoctorView_Grouping.Close
            If PayToDoctor = True Then
                .Commands!NewDoctorView_Grouping.CommandText = PreSHape1 & Sql1 & " appointmentdate = '" & MonthView2.Value & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and " & PostSHape1
            Else
                .Commands!NewDoctorView_Grouping.CommandText = PreSHape1 & Sql1 & " appointmentdate = '" & MonthView2.Value & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and patientabsent = 0 and " & PostSHape1
            End If
            .NewDoctorView_Grouping
        End With
        With DataReportAllDoctorsNew
            If HospitalDetails = True Then
                .Sections("ReportHeader").Controls.Item("InstitutionName").Caption = InstitutionName
                .Sections("ReportHeader").Controls.Item("InstitutionAddress").Caption = InstitutionAddress
                .Sections("ReportHeader").Controls.Item("lbldate").Caption = Format(Date, DefaultLongDate)
                .Sections("ReportFooter").Controls.Item("ad1").Caption = LongAd
            Else
                .Sections("ReportHeader").Controls.Item("InstitutionName").Caption = Empty
                .Sections("ReportHeader").Controls.Item("InstitutionAddress").Caption = Empty
                .Sections("ReportHeader").Controls.Item("lbldate").Caption = Format(Date, DefaultLongDate)
                .Sections("ReportFooter").Controls.Item("ad1").Caption = LongAd
            End If
            .Show
        End With
    End If
    
End Sub

Private Sub bttnAllPatients_Click()
    Const PreSHape = "SHAPE {"
    Const Sql = "SELECT tblPatientFacility.*, tblDoctor.DoctorListedName, tblTitle.Title FROM tblTitle RIGHT JOIN (tblPatientFacility LEFT JOIN tblDoctor ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID Where "
    Const PostSHape = "(((tblPatientFacility.HospitalFacility_ID) = 10))}  AS cmmdAllDoctorPatients COMPUTE cmmdAllDoctorPatients, COUNT(cmmdAllDoctorPatients.'PatientFacility_ID') AS TotalPatientCount, sum(cmmdAllDoctorPatients.'CancelledNull') AS TotalCancellations, SUM(cmmdAllDoctorPatients.'RefundNull') AS TotalRefunds, SUM(cmmdAllDoctorPatients.'FullyPaidNull') AS TotalFullyPaid, sum(cmmdAllDoctorPatients.'PatientAbsentNull') AS TotalAbsent, ANY(cmmdAllDoctorPatients.'Title') AS DoctorTitle BY 'DoctorListedName'"
    With DataEnvironment1

        If .rscmmdAllDoctorPatients_Grouping.State = 1 Then .rscmmdAllDoctorPatients_Grouping.Close

        If DetailedCount = False Then
            If PayToDoctor = True Then
                .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & MonthView2.Value & "' and " & PostSHape
            Else
                .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & MonthView2.Value & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and patientabsent = 0 and " & PostSHape

            End If
            .cmmdAllDoctorPatients_Grouping
            DataReportAllPatients.Sections("Section1").Controls.Item("lbl1").Visible = False
            DataReportAllPatients.Sections("Section1").Controls.Item("lbl2").Visible = True
            DataReportAllPatients.Sections("Section1").Controls.Item("lbl3").Visible = False
            DataReportAllPatients.Sections("Section1").Controls.Item("lbl4").Visible = False
            DataReportAllPatients.Sections("Section1").Controls.Item("lbl10").Visible = False
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl5").Visible = False
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl6").Visible = True
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl7").Visible = False
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl8").Visible = False
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl9").Visible = False
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl6").Caption = "Total Patients"
            DataReportAllPatients.Sections("ReportFooter").Controls.Item("Function1").Visible = False
            DataReportAllPatients.Sections("ReportFooter").Controls.Item("Function2").Visible = True
            DataReportAllPatients.Sections("ReportFooter").Controls.Item("Function3").Visible = False
            DataReportAllPatients.Sections("ReportFooter").Controls.Item("Function4").Visible = False
            DataReportAllPatients.Sections("ReportFooter").Controls.Item("Function5").Visible = False
        Else
            .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & Date & "' and " & PostSHape
            .cmmdAllDoctorPatients_Grouping
            DataReportAllPatients.Sections("Section1").Controls.Item("lbl1").Visible = True
            DataReportAllPatients.Sections("Section1").Controls.Item("lbl2").Visible = True
            DataReportAllPatients.Sections("Section1").Controls.Item("lbl3").Visible = True
            DataReportAllPatients.Sections("Section1").Controls.Item("lbl4").Visible = True
            DataReportAllPatients.Sections("Section1").Controls.Item("lbl10").Visible = True
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl5").Visible = True
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl6").Visible = True
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl7").Visible = True
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl8").Visible = True
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl9").Visible = True
            DataReportAllPatients.Sections("PageHeader").Controls.Item("lbl6").Caption = "Fully Paid"
            DataReportAllPatients.Sections("ReportFooter").Controls.Item("Function1").Visible = True
            DataReportAllPatients.Sections("ReportFooter").Controls.Item("Function2").Visible = True
            DataReportAllPatients.Sections("ReportFooter").Controls.Item("Function3").Visible = True
            DataReportAllPatients.Sections("ReportFooter").Controls.Item("Function4").Visible = True
            DataReportAllPatients.Sections("ReportFooter").Controls.Item("Function5").Visible = True
        End If
    End With
    With DataReportAllPatients
        If HospitalDetails = True Then
            .Sections("ReportHeader").Controls.Item("InstitutionName").Caption = InstitutionName
            .Sections("ReportHeader").Controls.Item("InstitutionAddress").Caption = InstitutionAddress
            .Sections("ReportHeader").Controls.Item("lbldate").Caption = Format(Date, DefaultLongDate)
            .Sections("ReportFooter").Controls.Item("ad1").Caption = LongAd
        Else
            .Sections("ReportHeader").Controls.Item("InstitutionName").Caption = Empty
            .Sections("ReportHeader").Controls.Item("InstitutionAddress").Caption = Empty
            .Sections("ReportHeader").Controls.Item("lbldate").Caption = Format(Date, DefaultLongDate)
            .Sections("ReportFooter").Controls.Item("ad1").Caption = LongAd
        End If
        Set .DataSource = DataEnvironment1
        .Show
    End With

End Sub

Private Sub bttnCancellation_Click()
    Dim TemResponce As Integer
    
    With DataEnvironment1.rssqlTem7
    


        If .State = 1 Then .Close
        .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(txtBookingID.Text)
        .Open
        
        If OptionRepayAgent.Value = False And OptionRepayPatient.Value = False And !PaymentMode = "Agent" Then
            TemResponce = MsgBox("You have not selected wether to repay the patient or the agent. Please select one.", vbQuestion, "Repay to whom?")
            OptionRepayPatient.SetFocus
            Exit Sub
        End If
        
        
        If .RecordCount = 0 Then
            TemResponce = MsgBox("There is no such a booking ID in the database. Please recheck", vbCritical, "ID Not found")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        
        If !HospitalFacility_ID <> 10 Then
            TemResponce = MsgBox("There booking ID is not for a channeling. Please recheck", vbCritical, "ID Not for channeling")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        
        If UserAuthority = AuthorityUser Then
            If !paidtostaff = True Then
                TemResponce = MsgBox("The money is already paid to the doctor. Therefore no refund can be done by a user. An accountant can pay if it is essential", vbCritical, "Already paid to the doctor")
                txtBookingID.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            End If
        Else
            If !paidtostaff = True Then
                TemResponce = MsgBox("The money is already paid to the doctor. Are you sure you want to refund ?", vbCritical + vbYesNo, "Already paid to the doctor")
                If TemResponce = vbNo Then
                    txtBookingID.SetFocus
                    SendKeys "{home}+{end}"
                    Exit Sub
                End If
            End If
        End If
        
        If !Cancelled = True Then
            TemResponce = MsgBox("The booking is already cancelled. You can't cancel it again", vbCritical, "Already cancelled")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        
        If !Refund = True Then
            TemResponce = MsgBox("The booking has already repaied. You can't cancel it", vbCritical, "Repaied")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        
        If !FullyPaid = 0 Then
            TemResponce = MsgBox("The patient has not completed the payment. You can't cancel it", vbCritical, "Repaied")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    
        Dim TemAgentId As Long
        TemAgentId = !Agent_ID
    
        If Val(lblPreviousTotalRepayC.Caption) + Val(txtRepayTotalC.Text) > Val(lblTotalPaidC.Caption) Then
            TemResponce = MsgBox("You can't repay an amount grater than that paid initially by the patient", vbCritical, "Exceeds Payment")
'            txtStaffRepayC.SetFocus
            Exit Sub
        End If
        
        
    End With

    With DataEnvironment1.rssqlTem
        If OptionRepayPatient.Value = True Then
                If .State = 1 Then .Close
                .Source = "select * from tblpatientrepay"
                If .State = 0 Then .Open
                .AddNew
                !patient_ID = TemPatientID
                !HospitalFacility_ID = 10
                !repayUser_ID = UserID
                !repaydate = Date
                !RepayTime = Time
                !StaffRepay = Val(txtStaffRepayC.Text)
                !InstitutionRepay = Val(txtInstitutionRepayC.Text)
                !OtherRepay = Val(txtOtherRepayC.Text)
                !TotalRepay = Val(txtRepayTotalC.Text)
                !Staff_ID = Val(ListConsultantIDs.Text)
                
                If Trim(txtCancellationComments.Text) = "" Then
                    !RepayComments = "Cancellation"
                Else
                    !RepayComments = txtCancellationComments.Text
                End If
                
                !patientfacility_ID = TemPatientFacilityID
                !RefundToAgent = False
                !RefundToPatient = 1
                .Update
                .Close
                .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(ListPatientFacilityIDs.Text)
                
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                    If IsNull(!Personalrefund) Then
                        !personaldue = !PersonalFee - Val(txtStaffRepayC.Text)
                        !Personalrefund = Val(txtStaffRepayC.Text)
                    Else
                        !personaldue = !PersonalFee - (Val(!Personalrefund) + Val(txtStaffRepayC.Text))
                        !Personalrefund = Val(!Personalrefund) + Val(txtStaffRepayC.Text)
                    End If
                    
                    If IsNull(!institutionrefund) Then
                        !institutiondue = !InstitutionFee - Val(txtInstitutionRepayC.Text)
                        !institutionrefund = Val(txtInstitutionRepayC.Text)
                    Else
                        !institutiondue = !InstitutionFee - (Val(!institutionrefund) + Val(txtInstitutionRepayC.Text))
                        !institutionrefund = Val(!institutionrefund) + Val(txtInstitutionRepayC.Text)
                    End If
                    
                    If IsNull(!otherrefund) Then
                        !otherdue = !otherfee - Val(txtOtherRepayC.Text)
                        !otherrefund = Val(txtOtherRepayC.Text)
                    Else
                        !otherdue = !otherfee - (Val(!otherrefund) + Val(txtOtherRepayC.Text))
                        !otherrefund = Val(!otherrefund) + Val(txtOtherRepayC.Text)
                    End If
                    
                    If IsNull(!totalrefund) Then
                        !TotalDue = !totalfee - Val(txtRepayTotalC.Text)
                        !totalrefund = Val(txtRepayTotalC.Text)
                    Else
                        !TotalDue = !totalfee - (Val(!totalrefund) + Val(txtRepayTotalC.Text))
                        !totalrefund = Val(!totalrefund) + Val(txtRepayTotalC.Text)
                    End If
                    
                    If Trim(txtCancellationComments.Text) = "" Then
                        !RepayComments = "Cancellation"
                    Else
                        !RepayComments = txtCancellationComments.Text
                    End If
                    
                    !repaydate = Date
                    !RepayTime = Time
                    !Cancelled = True
                    !cancellednull = 1
                    !repayUser_ID = UserID
                    !RefundToPatient = 1
                    !RefundToAgent = False
                    .Update
                    .Close
        ElseIf OptionRepayAgent.Value = True Then
                If .State = 1 Then .Close
                .Source = "select * from tblpatientrepay"
                If .State = 0 Then .Open
                .AddNew
                !patient_ID = TemPatientID
                !HospitalFacility_ID = 10
                !repayUser_ID = UserID
                !repaydate = Date
                !RepayTime = Time
                !StaffRepay = Val(txtStaffRepayC.Text)
                !InstitutionRepay = Val(txtInstitutionRepayC.Text)
                !OtherRepay = Val(txtOtherRepayC.Text)
                !TotalRepay = Val(txtRepayTotalC.Text)
                !Staff_ID = Val(ListConsultantIDs.Text)
                
                If Trim(txtCancellationComments.Text) = "" Then
                    !RepayComments = "Cancellation"
                Else
                    !RepayComments = txtCancellationComments.Text
                End If
                
                !patientfacility_ID = TemPatientFacilityID
                !RefundToPatient = False
                !RefundToAgent = 1
                .Update
                .Close
                .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(ListPatientFacilityIDs.Text)
                
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                    
                    If IsNull(!Personalrefund) Then
                        !personaldue = !PersonalFee - Val(txtStaffRepayC.Text)
                        !Personalrefund = Val(txtStaffRepayC.Text)
                    Else
                        !personaldue = !PersonalFee - (Val(!Personalrefund) + Val(txtStaffRepayC.Text))
                        !Personalrefund = Val(!Personalrefund) + Val(txtStaffRepayC.Text)
                    End If
                    
                    If IsNull(!institutionrefund) Then
                        !institutiondue = !InstitutionFee - Val(txtInstitutionRepayC.Text)
                        !institutionrefund = Val(txtInstitutionRepayC.Text)
                    Else
                        !institutiondue = !InstitutionFee - (Val(!institutionrefund) + Val(txtInstitutionRepayC.Text))
                        !institutionrefund = Val(!institutionrefund) + Val(txtInstitutionRepayC.Text)
                    End If
                    
                    If IsNull(!otherrefund) Then
                        !otherdue = !otherfee - Val(txtOtherRepayC.Text)
                        !otherrefund = Val(txtOtherRepayC.Text)
                    Else
                        !otherdue = !otherfee - (Val(!otherrefund) + Val(txtOtherRepayC.Text))
                        !otherrefund = Val(!otherrefund) + Val(txtOtherRepayC.Text)
                    End If
                    
                    If IsNull(!totalrefund) Then
                        !TotalDue = !totalfee - Val(txtRepayTotalC.Text)
                        !totalrefund = Val(txtRepayTotalC.Text)
                    Else
                        !TotalDue = !totalfee - (Val(!totalrefund) + Val(txtRepayTotalC.Text))
                        !totalrefund = Val(!totalrefund) + Val(txtRepayTotalC.Text)
                    End If
                    
                    If Trim(txtCancellationComments.Text) = "" Then
                        !RepayComments = "Cancellation"
                    Else
                        !RepayComments = txtCancellationComments.Text
                    End If
                    
                    !repaydate = Date
                    !RepayTime = Time
                    !Cancelled = True
                    !cancellednull = 1
                    !repayUser_ID = UserID
                    !RefundToPatient = False
                    !RefundToAgent = 1
                    .Update
                .Close
                If .State = 1 Then .Close
                .Source = "SELECT tblinstitutions.* from tblinstitutions where institution_ID =" & TemAgentId
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                !InstitutionCredit = !InstitutionCredit + Val(txtRepayTotalC.Text)
                .Update
                .Close
        Else
                If .State = 1 Then .Close
                .Source = "select * from tblpatientrepay"
                If .State = 0 Then .Open
                .AddNew
                !patient_ID = TemPatientID
                !HospitalFacility_ID = 10
                !repayUser_ID = UserID
                !repaydate = Date
                !RepayTime = Time
                !StaffRepay = Val(txtStaffRepayC.Text)
                !InstitutionRepay = Val(txtInstitutionRepayC.Text)
                !OtherRepay = Val(txtOtherRepayC.Text)
                !TotalRepay = Val(txtRepayTotalC.Text)
                !Staff_ID = Val(ListConsultantIDs.Text)
                
                If Trim(txtCancellationComments.Text) = "" Then
                    !RepayComments = "Cancellation"
                Else
                    !RepayComments = txtCancellationComments.Text
                End If
                
                !patientfacility_ID = TemPatientFacilityID
                !RefundToPatient = 1
                !RefundToAgent = False
                .Update
                .Close
                .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(ListPatientFacilityIDs.Text)
                
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                    
                    If IsNull(!Personalrefund) Then
                        !personaldue = !PersonalFee - Val(txtStaffRepayC.Text)
                        !Personalrefund = Val(txtStaffRepayC.Text)
                    Else
                        !personaldue = !PersonalFee - (Val(!Personalrefund) + Val(txtStaffRepayC.Text))
                        !Personalrefund = Val(!Personalrefund) + Val(txtStaffRepayC.Text)
                    End If
                    
                    If IsNull(!institutionrefund) Then
                        !institutiondue = !InstitutionFee - Val(txtInstitutionRepayC.Text)
                        !institutionrefund = Val(txtInstitutionRepayC.Text)
                    Else
                        !institutiondue = !InstitutionFee - (Val(!institutionrefund) + Val(txtInstitutionRepayC.Text))
                        !institutionrefund = Val(!institutionrefund) + Val(txtInstitutionRepayC.Text)
                    End If
                    
                    If IsNull(!otherrefund) Then
                        !otherdue = !otherfee - Val(txtOtherRepayC.Text)
                        !otherrefund = Val(txtOtherRepayC.Text)
                    Else
                        !otherdue = !otherfee - (Val(!otherrefund) + Val(txtOtherRepayC.Text))
                        !otherrefund = Val(!otherrefund) + Val(txtOtherRepayC.Text)
                    End If
                    
                    If IsNull(!totalrefund) Then
                        !TotalDue = !totalfee - Val(txtRepayTotalC.Text)
                        !totalrefund = Val(txtRepayTotalC.Text)
                    Else
                        !TotalDue = !totalfee - (Val(!totalrefund) + Val(txtRepayTotalC.Text))
                        !totalrefund = Val(!totalrefund) + Val(txtRepayTotalC.Text)
                    End If
                    
                    If Trim(txtCancellationComments.Text) = "" Then
                        !RepayComments = "Cancellation"
                    Else
                        !RepayComments = txtCancellationComments.Text
                    End If
                    
                    !repaydate = Date
                    !RepayTime = Time
                    !Cancelled = True
                    !cancellednull = 1
                    !repayUser_ID = UserID
                    !RefundToPatient = 1
                    !RefundToAgent = False
                    .Update
                .Close
        
        End If
    End With
    
    Call FormatGridPatients
    Call ListDatesAndSecessions_Click
    
End Sub

Private Sub bttnCashSettle_Click()

Dim TemResponce As Integer

    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(txtBookingID.Text)
        .Open
        If .RecordCount = 0 Then
            TemResponce = MsgBox("There is no such a booking ID in the database. Please recheck", vbCritical, "ID Not found")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        If !HospitalFacility_ID <> 10 Then
            TemResponce = MsgBox("There booking ID is not for a channeling. Please recheck", vbCritical, "ID Not for channeling")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        If !FullyPaid = 1 Then
            TemResponce = MsgBox("The money is fully paid. You can't pay again", vbCritical, "Already cancelled")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
            !totalfee = Val(lblTotalFeeToPay.Caption)
            !TotalDue = Val(lblTotalFeeToPay.Caption)
            !PersonalFee = Val(lblDoctorFeeToPay.Caption)
            !personaldue = Val(lblDoctorFeeToPay.Caption)
            !InstitutionFee = Val(lblHospitalFeeToPay.Caption)
            !institutiondue = Val(lblHospitalFeeToPay.Caption)
            !totalfeetopay = 0
            !PersonalFeeToPay = 0
            !InstitutionFeeToPay = 0
            !FullyPaid = 1
            !fullypaidnull = 1
            !SettleCashDate = Date
            !SettleCashTime = Time
            !CreditSettleUser_ID = UserID
            .Update
            TemBillId = !PatientBill_ID
        .Close
    End With
    Call UpdatePatientbill
    If OptionSettleCreditPrint.Value = True Then
        Call SetBillPrinter
        Call SetBillPaper
    Else
    
    End If
    
    Call FormatGridPatients
    Call ListDatesAndSecessions_Click

End Sub

Private Sub UpdatePatientbill()
'
'With DataEnvironment1.rssqlTem15
'
'    If .State = 1 Then .Close
'    .Source = "Select * From tblPatientBill Where (PatientBill_ID = " & TemBillId & ")"
'    .Open
'
'
'    If .RecordCount = 0 Then Exit Sub
'    !Credit = Val(!Credit) - Val(lblTotalFeeToPay.Caption)
'    .Update
'    If .State = 1 Then .Close
'
'End With
'
'
End Sub

Private Sub bttnChangeName_Click()
    Dim TemResponce As Integer
    If AllowNameChange = False Then
        TemResponce = MsgBox("You have not allowed to change names", vbCritical, "Not Allowed")
        txtNameChange.SetFocus
        Exit Sub
    End If
    If Trim(txtNameChange.Text) = "" Then
        TemResponce = MsgBox("You have not enter a name", vbCritical, "No name")
        txtNameChange.SetFocus
        Exit Sub
    End If
    If Trim(txtNameChange.Text) = Trim(txtBookedPatientName.Text) Then
        TemResponce = MsgBox("You have entered the very same name, So can't change", vbCritical, "No name")
        txtNameChange.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "select * from tblpatientmaindetails where patient_ID = " & Val(txtBookedPatientID.Text)
        .Open
        If .RecordCount = 0 Then Exit Sub
        !FirstName = Trim(txtNameChange.Text)
        .Update
        .Close
    End With
    txtBookedPatientName.Text = txtNameChange.Text
    ListDatesAndSecessions_Click
    txtNameChange.Text = Empty
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDoctorView_Click()
Dim TemResponce As Long

If ListSpecialities.ListIndex < 0 Or (IsNumeric(ListSpecialityIDs.Text) = False And ListSpecialityIDs.Text <> "All") Then
    TemResponce = MsgBox("You have not selected a speciality", vbCritical, "No COnsultant")
    ListSpecialities.SetFocus
    Exit Sub
End If


If ListConsultants.ListIndex < 0 Or IsNumeric(ListConsultantIDs.Text) = False Then
    TemResponce = MsgBox("You have not selected a consultant", vbCritical, "No COnsultant")
    ListConsultants.SetFocus
    Exit Sub
End If

If ListDatesAndSecessions.ListIndex < 0 Or (IsNumeric(ListSecessionIDs.Text) = False And ListSecessionIDs.Text <> "All") Then
    TemResponce = MsgBox("You have not selected a secession", vbCritical, "No Date & Secession")
    ListDatesAndSecessions.SetFocus
    Exit Sub
End If
    
    With DataEnvironment1.rssqlDoctorView
        If .State = 1 Then .Close
        If ListSecessionIDs.Text = "All" Then
            .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where staff_ID = " & Val(ListConsultantIDs.Text) & " and appointmentdate = '" & MonthView1.Value & "' and hospitalfacility_ID = 10 order by dayserial"
        Else
            .Source = "SELECT tblPatientFacility.*  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where staff_ID = " & Val(ListConsultantIDs.Text) & " and appointmentdate = '" & MonthView1.Value & "' and secession = " & Val(ListSecessionIDs.Text) & " and hospitalfacility_ID = 10 order by dayserial"
        End If
        .Open
    End With
    With DataReportDoctorView
        If HospitalDetails = True Then
            .Sections.Item("ReportHeader10").Controls.Item("RptName").Caption = InstitutionName
            .Sections.Item("ReportHeader10").Controls.Item("RptAddress").Caption = InstitutionAddress
            .Sections.Item("Section2").Controls.Item("lbldoctorname").Caption = "Consultant : " & FindLDoctorFromID(Val(ListConsultantIDs.Text))
            .Sections.Item("Section2").Controls.Item("lbldatesecession").Caption = "Date : " & Format(MonthView1.Value, DefaultLongDate) & "  Secession : " & FindSecessionFromID(Val(ListSecessionIDs.Text))
            .Sections.Item("Section5").Controls.Item("lblad1").Caption = LongAd
        Else
            .Sections.Item("ReportHeader10").Controls.Item("RptName").Caption = Empty
            .Sections.Item("ReportHeader10").Controls.Item("RptAddress").Caption = Empty
            .Sections.Item("Section2").Controls.Item("lbldoctorname").Caption = "Consultant : " & FindLDoctorFromID(Val(ListConsultantIDs.Text))
            .Sections.Item("Section2").Controls.Item("lbldatesecession").Caption = "Date : " & Format(MonthView1.Value, DefaultLongDate) & "  Secession : " & FindSecessionFromID(Val(ListSecessionIDs.Text))
            .Sections.Item("Section5").Controls.Item("lblad1").Caption = LongAd
        End If
        Set .DataSource = DataEnvironment1.rssqlDoctorView
        .Show
    End With
End Sub

Private Sub bttnMarkAbsent_Click()
    Dim TemResponce As Integer
    If ListPatientFacilities.ListIndex < 0 Or IsNumeric(ListPatientFacilityIDs.Text) = False Then
        TemResponce = MsgBox("You have not selected a patient to mark as absent", vbCritical, "Patient?")
        ListPatientFacilities.SetFocus
        Exit Sub
    End If
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "select * from tblpatientfacility where patientfacility_ID =" & txtBookingID.Text
        .Open
        If .RecordCount = 0 Then .Close: Exit Sub
        If !patientabsent = True Then
            TemResponce = MsgBox("This patient is already marked as absent", vbInformation, "Already Marked")
            .Close
            Exit Sub
        End If
        !patientabsent = True
        !PatientAbsentNull = 1
        .Update
        .Close
    End With
    Call ListDatesAndSecessions_Click
End Sub

Private Sub bttnMarkPresent_Click()
    Dim TemResponce As Integer
    If ListPatientFacilities.ListIndex < 0 Or IsNumeric(ListPatientFacilityIDs.Text) = False Then
        TemResponce = MsgBox("You have not selected a patient to mark as absent", vbCritical, "Patient?")
        ListPatientFacilities.SetFocus
        Exit Sub
    End If
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "select * from tblpatientfacility where patientfacility_ID =" & txtBookingID.Text
        .Open
        If .RecordCount = 0 Then .Close: Exit Sub
        If !patientabsent = 0 Then
            TemResponce = MsgBox("This patient is already marked as present", vbInformation, "Present")
            .Close
            Exit Sub
        End If
        !patientabsent = 0
        !PatientAbsentNull = 0
        .Update
        .Close
    End With
    Call ListDatesAndSecessions_Click
End Sub

Private Sub bttnNurseView_Click()
Dim TemResponce As Long

If ListSpecialities.ListIndex < 0 Or (IsNumeric(ListSpecialityIDs.Text) = False And ListSpecialityIDs.Text <> "All") Then
    TemResponce = MsgBox("You have not selected a speciality", vbCritical, "No COnsultant")
    ListSpecialities.SetFocus
    Exit Sub
End If


If ListConsultants.ListIndex < 0 Or IsNumeric(ListConsultantIDs.Text) = False Then
    TemResponce = MsgBox("You have not selected a consultant", vbCritical, "No COnsultant")
    ListConsultants.SetFocus
    Exit Sub
End If

If ListDatesAndSecessions.ListIndex < 0 Or (IsNumeric(ListSecessionIDs.Text) = False And ListSecessionIDs.Text <> "All") Then
    TemResponce = MsgBox("You have not selected a secession", vbCritical, "No Date & Secession")
    ListDatesAndSecessions.SetFocus
    Exit Sub
End If





    With DataEnvironment1.rssqlNurseView
        If .State = 1 Then .Close
        If ListSecessionIDs.Text = "All" Then
            .Source = "SELECT tblPatientFacility.*, tblInstitutions.InstitutionName, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionCode FROM tblInstitutions RIGHT JOIN (tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID  where staff_ID = " & Val(ListConsultantIDs.Text) & " and appointmentdate = '" & MonthView1.Value & "' and tblPatientFacility.hospitalfacility_ID = 10 order by dayserial "
        Else
            .Source = "SELECT tblPatientFacility.*, tblInstitutions.InstitutionName, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionCode FROM tblInstitutions RIGHT JOIN (tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID  where staff_ID = " & Val(ListConsultantIDs.Text) & " and appointmentdate = '" & MonthView1.Value & "' and secession = " & Val(ListSecessionIDs.Text) & " and tblPatientFacility.hospitalfacility_ID = 10 order by dayserial "
        End If
        .Open
    End With
    With DataReportNurseView
        If HospitalDetails = True Then
            .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = InstitutionName
            .Sections.Item("Section4").Controls.Item("lblInstitutionAddress").Caption = InstitutionAddress
            .Sections.Item("Section4").Controls.Item("lblinstitutiontelephone").Caption = "Nurse View"
            .Sections.Item("Section2").Controls.Item("lbldoctorname").Caption = "Consultant : " & FindLDoctorFromID(Val(ListConsultantIDs.Text))
            .Sections.Item("Section2").Controls.Item("lbldatesecession").Caption = "Date : " & Format(ListDates.Text, DefaultLongDate) & "   Secession : " & FindSecessionFromID(Val(ListSecessionIDs.Text))
            .Sections.Item("Section5").Controls.Item("lblad1").Caption = LongAd
        Else
            .Sections.Item("Section4").Controls.Item("lblInstitutionname").Caption = Empty
            .Sections.Item("Section4").Controls.Item("lblInstitutionAddress").Caption = Empty
            .Sections.Item("Section4").Controls.Item("lblinstitutiontelephone").Caption = "Nurse View"
            .Sections.Item("Section2").Controls.Item("lbldoctorname").Caption = "Consultant : " & FindLDoctorFromID(Val(ListConsultantIDs.Text))
            .Sections.Item("Section2").Controls.Item("lbldatesecession").Caption = "Date : " & Format(ListDates.Text, DefaultLongDate) & "   Secession : " & FindSecessionFromID(Val(ListSecessionIDs.Text))
            .Sections.Item("Section5").Controls.Item("lblad1").Caption = LongAd
        End If
        Set .DataSource = DataEnvironment1.rssqlNurseView
        .Show
    End With
End Sub

Private Sub bttnRefund_Click()
Dim TemResponce As Integer

    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(txtBookingID.Text)
        .Open
        If .RecordCount = 0 Then
            TemResponce = MsgBox("There is no such a booking ID in the database. Please recheck", vbCritical, "ID Not found")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        If !HospitalFacility_ID <> 10 Then
            TemResponce = MsgBox("There booking ID is not for a channeling. Please recheck", vbCritical, "ID Not for channeling")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        If UserAuthority = AuthorityUser Then
            If !paidtostaff = True Then
                TemResponce = MsgBox("The money is already paid to the doctor. Therefore no refund can be done by a user. An accountant can pay if it is essential", vbCritical, "Already paid to the doctor")
                txtBookingID.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            End If
        Else
            If !paidtostaff = True Then
                TemResponce = MsgBox("The money is already paid to the doctor. Are you sure you want to refund ?", vbCritical + vbYesNo, "Already paid to the doctor")
                If TemResponce = vbNo Then
                    txtBookingID.SetFocus
                    SendKeys "{home}+{end}"
                    Exit Sub
                End If
            End If
        End If
        If !Cancelled = True Then
            TemResponce = MsgBox("The booking is already cancelled. You can't cancel it again", vbCritical, "Already cancelled")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        If !Refund = True Then
            TemResponce = MsgBox("The booking has already repaied. You can't cancel it", vbCritical, "Repaied")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        If !FullyPaid = 0 Then
            TemResponce = MsgBox("The patient has not completed the payment. You can't cancel it", vbCritical, "Repaied")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        If Val(lblPreviousTotalRepayR.Caption) + Val(txtRepayTotalR.Text) > Val(lblTotalPaidR.Caption) Then
            TemResponce = MsgBox("You can't repay an amount grater than that paid initially by the patient", vbCritical, "Exceeds Payment")
            txtStaffRepayR.SetFocus
            Exit Sub
        End If
    End With
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "select * from tblpatientrepay"
        If .State = 0 Then .Open
        .AddNew
        !patient_ID = TemPatientID
        !HospitalFacility_ID = 10
        !repayUser_ID = UserID
        !repaydate = Date
        !RepayTime = Time
        !StaffRepay = Val(txtStaffRepayR.Text)
        !InstitutionRepay = Val(txtInstitutionRepayR.Text)
        !OtherRepay = Val(txtOtherRepayR.Text)
        !TotalRepay = Val(txtRepayTotalR.Text)
        !Staff_ID = Val(ListConsultantIDs.Text)
        If Trim(txtRefundComments.Text) = "" Then
            !RepayComments = "Refund"
        Else
            !RepayComments = txtRefundComments.Text
        End If
        !patientfacility_ID = TemPatientFacilityID
        !RefundToPatient = 1
        .Update
        .Close
        .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & TemPatientFacilityID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
            If IsNull(!Personalrefund) Then
                !personaldue = !PersonalFee - Val(txtStaffRepayR.Text)
                !Personalrefund = Val(txtStaffRepayR.Text)
            Else
                !personaldue = !PersonalFee - (Val(!Personalrefund) + Val(txtStaffRepayR.Text))
                !Personalrefund = Val(!Personalrefund) + Val(txtStaffRepayR.Text)
            End If
            If IsNull(!institutionrefund) Then
                !institutiondue = !InstitutionFee - Val(txtInstitutionRepayR.Text)
                !institutionrefund = Val(txtInstitutionRepayR.Text)
            Else
                !institutiondue = !InstitutionFee - (Val(!institutionrefund) + Val(txtInstitutionRepayR.Text))
                !institutionrefund = Val(!institutionrefund) + Val(txtInstitutionRepayR.Text)
            End If
            If IsNull(!otherrefund) Then
                !otherdue = !otherfee - Val(txtOtherRepayR.Text)
                !otherrefund = Val(txtOtherRepayR.Text)
            Else
                !otherdue = !otherfee - (Val(!otherrefund) + Val(txtOtherRepayR.Text))
                !otherrefund = Val(!otherrefund) + Val(txtOtherRepayR.Text)
            End If
            If IsNull(!totalrefund) Then
                !TotalDue = !totalfee - Val(txtRepayTotalR.Text)
                !totalrefund = Val(txtRepayTotalR.Text)
            Else
                !TotalDue = !totalfee - (Val(!totalrefund) + Val(txtRepayTotalR.Text))
                !totalrefund = Val(!totalrefund) + Val(txtRepayTotalR.Text)
            End If
            If Trim(txtRefundComments.Text) = "" Then
                !RepayComments = "Refund"
            Else
                !RepayComments = txtRefundComments.Text
            End If
            !repaydate = Date
            !RepayTime = Time
            !Cancelled = False
            !Refund = True
            !refundnull = 1
            !repayUser_ID = UserID
            !RefundToPatient = 1
        
            .Update
        .Close
    End With
    
    Call FormatGridPatients
    Call ListDatesAndSecessions_Click

End Sub


Private Sub bttnReprint_Click()
    Dim TemRows As Long
    Dim TemResponce As Integer
    
    With DataEnvironment1.rssqlTem7
    
        If .State = 1 Then .Close
        .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(txtBookingID.Text)
        .Open
        
        If .RecordCount = 0 Then
            TemResponce = MsgBox("There is no such a booking ID in the database. Please recheck", vbCritical, "ID Not found")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        
        If !HospitalFacility_ID <> 10 Then
            TemResponce = MsgBox("There booking ID is not for a channeling. Please recheck", vbCritical, "ID Not for channeling")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        If UserAuthority = AuthorityUser Then
            If !paidtostaff = True Then
                TemResponce = MsgBox("The money is already paid to the doctor. Therefore you can't issue a copy of the receipt. An accountant can pay if it is essential", vbCritical, "Already paid to the doctor")
                txtBookingID.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            End If
        Else
            If !paidtostaff = True Then
                TemResponce = MsgBox("The money is already paid to the doctor. Are you sure you want to print a copy of the bill ?", vbCritical + vbYesNo, "Already paid to the doctor")
                If TemResponce = vbNo Then
                    txtBookingID.SetFocus
                    SendKeys "{home}+{end}"
                    Exit Sub
                End If
            End If
        End If
        If !Cancelled = True Then
            TemResponce = MsgBox("The booking is cancelled. You can print the bill again.", vbCritical, "Already cancelled")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        If !Refund = True Then
            TemResponce = MsgBox("The booking has repaied. You can print a bill again", vbCritical, "Repaied")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        
        If !FullyPaid = 0 Then
            TemResponce = MsgBox("The patient has not completed the payment. You can't cancel it", vbCritical, "Repaied")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    
    End With
    
    Call SetBillPrinter
    Call SetBillPaper
    
    
End Sub


Private Sub ComboPatientName_Change()
    Call FillPatientSearchGrid
End Sub

Private Sub FillPatientSearchGrid()
Dim NowROw As Long
With DataEnvironment1.rssqlTem11
    If .State = 1 Then .Close
    .Source = "SELECT tblDoctor.Doctor_ID, tblDoctor.DoctorListedName, tblPatientFacility.*, tblPatientMainDetails.Patient_ID, tblPatientMainDetails.FirstName FROM (tblDoctor RIGHT JOIN tblPatientFacility ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientMainDetails.FirstName ='" & ComboPatientName.Text & "') and appointmentdate = '" & DTPickerFindPatientDate.Value & "' order by patientfacility_id"
    .Open
    FormatPatientSearchGrid
    NowROw = 0
    If .RecordCount = 0 Then .Close: Exit Sub
    Do While .EOF = False
        NowROw = NowROw + 1
        gridPatient.Rows = NowROw + 1
        gridPatient.Row = NowROw
        
        gridPatient.col = 0
        gridPatient.CellAlignment = 7
        gridPatient.Text = !patientfacility_ID
        
        gridPatient.col = 1
        gridPatient.CellAlignment = 1
        gridPatient.Text = !FirstName
        
        gridPatient.col = 2
        gridPatient.CellAlignment = 1
        gridPatient.Text = FindLDoctorFromID(!Doctor_ID)
        
        gridPatient.col = 3
        gridPatient.Text = Format(!BookingDate, DefaultShortDate)
        gridPatient.CellAlignment = 7
        
        gridPatient.col = 4
        gridPatient.Text = Format(!AppointmentDate, DefaultShortDate)
        gridPatient.CellAlignment = 7
        
        gridPatient.col = 5
        gridPatient.CellAlignment = 4
        If !PaymentMode = "Agent" Then
            gridPatient.Text = FindAgentFromID(!Agent_ID)
        ElseIf !PaymentMode = "Cash" Then
            gridPatient.Text = "Cash"
        ElseIf !PaymentMode = "Credit" Then
            gridPatient.Text = "Credit"
        End If
        
        .MoveNext
    Loop
    If .State = 1 Then .Close
End With
End Sub

Private Sub ComboPatientName_Click()
    Call FillPatientSearchGrid
End Sub


Private Sub FormatPatientSearchGrid()

With gridPatient
    .Clear
    
    .Rows = 1
    .Cols = 6
    
    .ColWidth(0) = 700
    .ColWidth(2) = 2400
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1800

    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(4) + .ColWidth(2) + .ColWidth(3) + .ColWidth(5) + 100)
    .Row = 0
    
    .col = 0
    .CellAlignment = 4
    .Text = "ID"
    
    .col = 1
    .CellAlignment = 4
    .Text = "Patient Name"
    
    .col = 2
    .CellAlignment = 4
    .Text = "Consultant"
    
    .col = 3
    .CellAlignment = 4
    .Text = "Booking"
    
    .col = 4
    .CellAlignment = 4
    .Text = "Appointment"
    
    .col = 5
    .CellAlignment = 4
    .Text = "Agent"
    
    
    
    
End With
End Sub

Private Sub DTPickerFindPatientDate_Change()
    Call FormatPatientSearchGrid
    Call FillPatientName
End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.WindowState = 2

    If SetPrinter = False Then
        Unload Me
        Exit Sub
    End If
    If UserAuthority <> AuthorityOwner Then
        MonthView1.Enabled = False
    End If
    
End Sub

Private Sub FillPatientName()
    With DataEnvironment1.rssqlTem18
        If .State = 1 Then .Close
        .Open "SELECT tblPatientFacility.*, tblPatientMainDetails.* FROM tblPatientFacility LEFT OUTER JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID  where appointmentdate = '" & DTPickerFindPatientDate.Value & "' Order By patientfacility_ID"
        ComboPatientName.Clear
        If .RecordCount = 0 Then Exit Sub
        ComboPatientName.Visible = False
        While .EOF = False
            ComboPatientName.AddItem Format(!FirstName, "")
            .MoveNext
        Wend
        ComboPatientName.Visible = True
End With
End Sub



Private Function SetPrinter() As Boolean
SetPrinter = False
Dim MyPrinter As Printer

For Each MyPrinter In Printers
    If MyPrinter.DeviceName = BillPrinterName Then
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






Private Sub gridPatient_Click()
gridPatient.col = 0
    If IsNumeric(gridPatient.Text) = False Then Exit Sub
    txtSearchBookingID.Text = gridPatient.Text
    Call bttnSearch_Click
'gridPatient.Col = 0
'gridPatient.ColSel = gridPatient.Cols - 1
FormatPatientSearchGrid
SSTab2.Tab = 0

End Sub

Private Sub FindAgentName()
gridPatient.col = 5
If IsNumeric(gridPatient.Text) = False Then Exit Sub
With DataEnvironment1.rssqlTem13
    If .State = 1 Then .Close
    .Open "Select * From  tblInstitutions Where (Institution_Id =" & gridPatient.Text & ")"
    If .RecordCount = 0 Then Exit Sub
    
    lblAgentName.Caption = !InstitutionName
    
    If .State = 1 Then .Close

End With

End Sub





Private Function GetBookedNumber(BookingDate As Date, SecessionID As Long) As Long
With DataEnvironment1.rssqlTem5
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientfacility where hospitalfacility_ID = " & 10 & " and AppointmentDate = '" & BookingDate & "' and Secession = " & SecessionID
    .Open
    GetBookedNumber = .RecordCount
    If .State = 1 Then .Close
End With
End Function


Private Sub ListConsultants_GotFocus()
    BoxConsultant.BackColor = BttnBackColour ' vbRed
End Sub

Private Sub ListConsultants_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    Call ClearPatientDetails
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
    SSTab2.SetFocus
    KeyCode = Empty
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
    SSTab2.SetFocus
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







Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
MonthView2.Value = MonthView1.Value

    Dim TemNum As Long
    Dim DateFound As Boolean
    Dim Tem
    
For TemNum = 0 To ListDates.ListCount - 1
    ListDates.ListIndex = TemNum
    If IsDate(ListDates.Text) Then
        If DateClicked = ListDates.Text Then
            DateFound = True
            TemNum = ListDates.ListCount - 1
        End If
    End If
Next

If DateFound = False Then
    Beep
Else
    ListDatesAndSecessions.ListIndex = ListDates.ListIndex
    ListDatesAndSecessions_Click
End If
End Sub

Private Sub MonthView2_GotFocus()
    BoxDates.BackColor = BttnBackColour ' vbRed
End Sub

Private Sub MonthView2_LostFocus()
    BoxDates.BackColor = FrameBackColour ' vbRed
End Sub

Private Sub txtInstitutionRepayC_Change()
    txtRepayTotalC.Text = Format((Val(txtStaffRepayC.Text) + Val(txtInstitutionRepayC.Text) + Val(txtOtherRepayC.Text)), "0.00")
End Sub

Private Sub txtInstitutionRepayR_Change()
    txtRepayTotalR.Text = Format((Val(txtStaffRepayR.Text) + Val(txtInstitutionRepayR.Text) + Val(txtOtherRepayR.Text)), "0.00")
End Sub

Private Sub txtNameChange_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then bttnChangeName_Click


End Sub

Private Sub txtOtherRepayC_Change()
    txtRepayTotalC.Text = Format((Val(txtStaffRepayC.Text) + Val(txtInstitutionRepayC.Text) + Val(txtOtherRepayC.Text)), "0.00")
End Sub

Private Sub txtOtherRepayR_Change()
    txtRepayTotalR.Text = Format((Val(txtStaffRepayR.Text) + Val(txtInstitutionRepayR.Text) + Val(txtOtherRepayR.Text)), "0.00")
End Sub

Private Sub bttnSearch_Click()
    Call SearchBookingID
    SSTab2.Tab = 0
    txtSearchBookingID.Text = Empty
End Sub




Private Sub txtSearchAgentRefNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then bttnAgentRefSearch_Click
End Sub

Private Sub txtSearchBookingID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SearchBookingID
End Sub


Private Sub SearchBookingID()
Dim TemResponce As Integer

With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    .Source = "Select * from tblpatientfacility where patientfacility_ID = " & Val(txtSearchBookingID.Text)
    .Open
    If .RecordCount = 0 Then
        TemResponce = MsgBox("There is no such booking ID. Please re-check", vbCritical, "No such ID")
        .Close
        Exit Sub
    End If
    Call ListAllConsultants
   
    Dim TemNum As Long
    
    If ListConsultants.ListCount = 0 Then
        TemResponce = MsgBox("The consultant is deleted", vbCritical, "Consultant Deleted")
        Exit Sub
    End If
    
    MonthView2.Value = !AppointmentDate
        
    Dim ConsultantFound As Boolean
    ConsultantFound = False
    For TemNum = 0 To ListConsultantIDs.ListCount - 1
        ListConsultantIDs.ListIndex = TemNum
        If Val(ListConsultantIDs.Text) = !Staff_ID Then
            ListConsultants.ListIndex = TemNum
            ListConsultants_Click
            TemNum = ListConsultantIDs.ListCount
            ConsultantFound = True
        End If
    Next
    If ConsultantFound = False Then
        TemResponce = MsgBox("The consultant the patient booked is deleted", vbCritical, "Deleted")
        Exit Sub
    End If
    
    If ListDatesAndSecessions.ListCount = 0 Then
        TemResponce = MsgBox("The booking date for the patient is deleted", vbCritical, "Deleted")
        Exit Sub
    End If
    
    
    Dim DateFound As Boolean
    Dim FoundSecession As Long
    DateFound = False
    
    For TemNum = 0 To ListSecessionIDs.ListCount - 1
        ListSecessionIDs.ListIndex = TemNum
            ListSecessionIDs.ListIndex = TemNum
            If ListSecessionIDs.Text = !Secession Then
                FoundSecession = TemNum
                DateFound = True
                TemNum = ListSecessionIDs.ListCount - 1
            End If

    Next
    
    If DateFound = True Then
        ListDatesAndSecessions.ListIndex = FoundSecession
        ListDatesAndSecessions_Click
    Else
        ListDatesAndSecessions.ListIndex = 0
        ListDatesAndSecessions_Click
    End If
    
    If ListPatientFacilities.ListCount = 0 Then Exit Sub
    
    For TemNum = 0 To ListPatientFacilities.ListCount - 1
        ListPatientFacilityIDs.ListIndex = TemNum
        If Val(ListPatientFacilityIDs.Text) = Val(txtSearchBookingID.Text) Then
            ListPatientFacilities.ListIndex = TemNum
            ListPatientFacilities_Click
            TemNum = ListPatientFacilities.ListCount
        End If
    Next

End With

End Sub


Private Sub txtStaffRepayC_Change()
txtRepayTotalC.Text = Format((Val(txtStaffRepayC.Text) + Val(txtInstitutionRepayC.Text) + Val(txtOtherRepayC.Text)), "0.00")
End Sub

Private Sub txtStaffRepayR_Change()
txtRepayTotalR.Text = Format((Val(txtStaffRepayR.Text) + Val(txtInstitutionRepayR.Text) + Val(txtOtherRepayR.Text)), "0.00")
End Sub
