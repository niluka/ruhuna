VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOldChannelingMS 
   Caption         =   "Channeling"
   ClientHeight    =   10230
   ClientLeft      =   -135
   ClientTop       =   645
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
   Icon            =   "frmOldChannelingMS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   15240
   Begin VB.TextBox txtMaxCounter 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   3000
      TabIndex        =   202
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtTemAgentID 
      Height          =   375
      Left            =   9840
      TabIndex        =   149
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   4350
      ItemData        =   "frmOldChannelingMS.frx":0442
      Left            =   5760
      List            =   "frmOldChannelingMS.frx":0444
      TabIndex        =   2
      ToolTipText     =   "List of Date, Secession, Maximum number per secession, Starting Time and already given numbers of the selected consultant"
      Top             =   360
      Width           =   3855
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
      ItemData        =   "frmOldChannelingMS.frx":0446
      Left            =   9600
      List            =   "frmOldChannelingMS.frx":0448
      TabIndex        =   3
      ToolTipText     =   "List of number, patient, paid or not, cancelled or refunded, agent code and present or absent"
      Top             =   360
      Width           =   5655
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
      ItemData        =   "frmOldChannelingMS.frx":044A
      Left            =   3000
      List            =   "frmOldChannelingMS.frx":044C
      TabIndex        =   1
      ToolTipText     =   "List of Consultants of selected speciality"
      Top             =   360
      Width           =   2655
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
      ItemData        =   "frmOldChannelingMS.frx":044E
      Left            =   240
      List            =   "frmOldChannelingMS.frx":0450
      TabIndex        =   0
      ToolTipText     =   "List of Specialities"
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox ListSecessionStartingTime 
      Height          =   3660
      Left            =   13320
      TabIndex        =   65
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListPatientFacilityIDs 
      Height          =   3900
      Left            =   13920
      TabIndex        =   63
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox ListConsultantIDs 
      Height          =   3900
      Left            =   13320
      TabIndex        =   62
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox ListSecessionIDs 
      Height          =   3900
      Left            =   13440
      TabIndex        =   61
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListDates 
      Height          =   3660
      Left            =   13200
      TabIndex        =   60
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListSpecialityIDs 
      Height          =   3900
      Left            =   13440
      TabIndex        =   59
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   240
      Left            =   2640
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8880
      Value           =   1  'Checked
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   4215
      Left            =   5280
      TabIndex        =   17
      Top             =   4920
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   8
      Tab             =   7
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Booking"
      TabPicture(0)   =   "frmOldChannelingMS.frx":0452
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FramePatientDetails"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Reprint"
      TabPicture(1)   =   "frmOldChannelingMS.frx":046E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameReprints"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cancel"
      TabPicture(2)   =   "frmOldChannelingMS.frx":048A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameCancellations"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Refund"
      TabPicture(3)   =   "frmOldChannelingMS.frx":04A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameRefunds"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Settle "
      TabPicture(4)   =   "frmOldChannelingMS.frx":04C2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameSettleCredit"
      Tab(4).Control(1)=   "bttnAgentBookingValidation"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Change"
      TabPicture(5)   =   "frmOldChannelingMS.frx":04DE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Search"
      TabPicture(6)   =   "frmOldChannelingMS.frx":04FA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame5"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Views"
      TabPicture(7)   =   "frmOldChannelingMS.frx":0516
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "Frame7"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      Begin VB.Frame FrameReprints 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   177
         Top             =   360
         Width           =   9615
         Begin btButtonEx.ButtonEx bttnReprint 
            Height          =   375
            Left            =   5160
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Appearance      =   3
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
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   189
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   188
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   187
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label25 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   186
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   185
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   184
            Top             =   2400
            Width           =   1935
         End
         Begin VB.Label lblPaymentMethod 
            BackStyle       =   0  'Transparent
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
            Left            =   77640
            TabIndex        =   183
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label lblDoctorFeePaid 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   182
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblHospitalFeePaid 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   181
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblOtherFeePaid 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   180
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblTotalFeePaid 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   179
            Top             =   1920
            Width           =   1335
         End
      End
      Begin VB.Frame FrameRefunds 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   151
         Top             =   360
         Width           =   9615
         Begin VB.Frame Frame3 
            Height          =   975
            Left            =   7320
            TabIndex        =   157
            Top             =   1440
            Width           =   2175
            Begin VB.OptionButton OptionRefundPrint 
               Caption         =   "Print"
               Height          =   255
               Left            =   75120
               TabIndex        =   159
               Top             =   240
               Width           =   1935
            End
            Begin VB.OptionButton OptionRefundDoNotPrint 
               Caption         =   "Do not print"
               Height          =   255
               Left            =   120
               TabIndex        =   158
               Top             =   600
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin VB.TextBox txtRefundComments 
            Height          =   375
            Left            =   2040
            MultiLine       =   -1  'True
            TabIndex        =   156
            Top             =   2520
            Width           =   5175
         End
         Begin VB.TextBox txtStaffRepayR 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   155
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtInstitutionRepayR 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   154
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtOtherRepayR 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   153
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtRepayTotalR 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   152
            Top             =   1920
            Width           =   1335
         End
         Begin btButtonEx.ButtonEx bttnRefund 
            Height          =   375
            Left            =   7680
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   2520
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
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   176
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   175
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   174
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   173
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblStaffFeePaidR 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   172
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblInstitutionFeePaidR 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   171
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblOtherFeePaidR 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   170
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblTotalPaidR 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   169
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   168
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblPreviousStaffRepayR 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   167
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblPreviousInstitutionRepayR 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   166
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblPreviousOtherRepayR 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   165
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblPreviousTotalRepayR 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   164
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   163
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   162
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   161
            Top             =   240
            Width           =   1575
         End
      End
      Begin btButtonEx.ButtonEx bttnAgentBookingValidation 
         Height          =   375
         Left            =   -68400
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "ValidateAgent Booking"
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
      Begin VB.Frame Frame7 
         Height          =   3615
         Left            =   120
         TabIndex        =   139
         Top             =   360
         Width           =   9615
         Begin btButtonEx.ButtonEx bttnAllSecessionPatients 
            Height          =   375
            Left            =   3120
            TabIndex        =   147
            Top             =   1800
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
         Begin MSComCtl2.MonthView MonthView1 
            Height          =   2820
            Left            =   6480
            TabIndex        =   140
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   4974
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            StartOfWeek     =   58982401
            CurrentDate     =   39446
         End
         Begin btButtonEx.ButtonEx bttnNurseView 
            Height          =   375
            Left            =   360
            TabIndex        =   141
            Top             =   840
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
            Left            =   360
            TabIndex        =   142
            Top             =   1320
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
         Begin btButtonEx.ButtonEx bttnAllPatients 
            Height          =   375
            Left            =   3120
            TabIndex        =   143
            Top             =   840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "All &Patients"
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
            Left            =   3120
            TabIndex        =   144
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "All D&octors"
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
         Begin VB.Frame Frame8 
            Caption         =   "Selected Secession"
            Height          =   2295
            Left            =   120
            TabIndex        =   145
            Top             =   480
            Width           =   2535
            Begin btButtonEx.ButtonEx bttnSecession 
               Height          =   375
               Left            =   240
               TabIndex        =   148
               Top             =   1320
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   661
               Appearance      =   3
               Caption         =   "Secession View"
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
            Begin btButtonEx.ButtonEx btnActiveView 
               Height          =   375
               Left            =   240
               TabIndex        =   201
               Top             =   1800
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   661
               Appearance      =   3
               Caption         =   "&Active View"
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
         Begin VB.Frame Frame9 
            Caption         =   "Today's"
            Height          =   1935
            Left            =   2880
            TabIndex        =   146
            Top             =   480
            Width           =   2535
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   134
         Top             =   360
         Width           =   9615
         Begin VB.TextBox txtPhoneChange 
            Height          =   360
            Left            =   2040
            TabIndex        =   199
            Top             =   2880
            Width           =   3135
         End
         Begin VB.TextBox txtNameChange 
            Height          =   360
            Left            =   2040
            TabIndex        =   135
            Top             =   2400
            Width           =   3135
         End
         Begin btButtonEx.ButtonEx bttnChangeName 
            Height          =   855
            Left            =   5280
            TabIndex        =   136
            Top             =   2400
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   1508
            Appearance      =   3
            Caption         =   "Change"
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
            Left            =   240
            TabIndex        =   137
            Top             =   360
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
            Left            =   240
            TabIndex        =   138
            Top             =   840
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
         Begin VB.Label Label63 
            Caption         =   "New Phone No"
            Height          =   255
            Left            =   120
            TabIndex        =   200
            Top             =   2880
            Width           =   1935
         End
         Begin VB.Label Label62 
            Caption         =   "New Name"
            Height          =   255
            Left            =   120
            TabIndex        =   198
            Top             =   2400
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   122
         Top             =   360
         Width           =   9615
         Begin VB.ComboBox ComboPatientName 
            Height          =   360
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox txtSearchBookingID 
            Height          =   375
            Left            =   6960
            TabIndex        =   124
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtSearchAgentRefNo 
            Height          =   375
            Left            =   6960
            TabIndex        =   123
            Top             =   720
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPickerFindPatientDate 
            Height          =   375
            Left            =   840
            TabIndex        =   126
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MM yyyy"
            Format          =   58982403
            CurrentDate     =   39470
         End
         Begin MSFlexGridLib.MSFlexGrid gridPatient 
            Height          =   2295
            Left            =   120
            TabIndex        =   127
            Top             =   1200
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   4048
            _Version        =   393216
         End
         Begin btButtonEx.ButtonEx bttnSearch 
            Height          =   375
            Left            =   8280
            TabIndex        =   128
            Top             =   240
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
         Begin btButtonEx.ButtonEx bttnAgentRefNoSearch 
            Height          =   375
            Left            =   8280
            TabIndex        =   129
            Top             =   720
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
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   133
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Booking ID"
            Height          =   255
            Left            =   5040
            TabIndex        =   132
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Left            =   120
            TabIndex        =   131
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent Referrance No."
            Height          =   255
            Left            =   5040
            TabIndex        =   130
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.Frame FramePatientDetails 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   96
         Top             =   360
         Width           =   9615
         Begin VB.TextBox txtBookedPatientContactNo 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   195
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox txtAppTime 
            Height          =   360
            Left            =   7800
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtAppDate 
            Height          =   360
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtConsultant 
            Height          =   360
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtBookingDate 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   3120
            Width           =   3135
         End
         Begin VB.TextBox txtAgentRefNo 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   2160
            Width           =   3135
         End
         Begin VB.TextBox txtBookedPatientName 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtBookedPatientID 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   97
            Top             =   240
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.TextBox txtBookingID 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox txtPaymentMethod 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox txtBookingUser 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   2640
            Width           =   3135
         End
         Begin VB.TextBox txtCancelRefund 
            Height          =   840
            Left            =   6000
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox txtCreditSettle 
            Height          =   1440
            Left            =   6000
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   2160
            Width           =   3495
         End
         Begin VB.TextBox txtAgentAndCode 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1680
            Width           =   3135
         End
         Begin VB.TextBox txtAgentCode 
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   117
            Top             =   1680
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact No"
            Height          =   255
            Left            =   120
            TabIndex        =   196
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "App. Time"
            Height          =   255
            Left            =   4920
            TabIndex        =   121
            Top             =   720
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "Date/Time"
            Height          =   255
            Left            =   4920
            TabIndex        =   120
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Consultant"
            Height          =   255
            Left            =   4920
            TabIndex        =   119
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "Booking Date"
            Height          =   255
            Left            =   120
            TabIndex        =   118
            Top             =   3120
            Width           =   1695
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent Ref. No"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Agent"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Name"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "Booking ID"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Method"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "Booking User"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   2640
            Width           =   2415
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Cancel / Refund"
            Height          =   615
            Left            =   4920
            TabIndex        =   99
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Settling"
            Height          =   615
            Left            =   4920
            TabIndex        =   98
            Top             =   2160
            Width           =   1095
         End
      End
      Begin VB.Frame FrameCancellations 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   76
         Top             =   360
         Width           =   9615
         Begin VB.Frame Frame2 
            Height          =   975
            Left            =   7320
            TabIndex        =   110
            Top             =   1440
            Visible         =   0   'False
            Width           =   2175
            Begin VB.OptionButton OptionDoNotPrintCancel 
               Caption         =   "Do not print"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   600
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton OptionPrintCancel 
               Caption         =   "Print"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.OptionButton OptionRepayAgent 
            Caption         =   "Repay Agent"
            Height          =   255
            Left            =   7560
            TabIndex        =   34
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.OptionButton OptionRepayPatient 
            Caption         =   "Repay Patient"
            Height          =   255
            Left            =   7560
            TabIndex        =   35
            Top             =   840
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtStaffRepayC 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   30
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtInstitutionRepayC 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   31
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtOtherRepayC 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            TabIndex        =   32
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtRepayTotalC 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   5880
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtCancellationComments 
            Height          =   375
            Left            =   2040
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   2520
            Width           =   5175
         End
         Begin btButtonEx.ButtonEx bttnCancellation 
            Height          =   375
            Left            =   7680
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Appearance      =   3
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
            BackStyle       =   0  'Transparent
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
            TabIndex        =   87
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   92
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   91
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   89
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   88
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblInstitutionFeePaidC 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   86
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblOtherFeePaidC 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   85
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblTotalPaidC 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   84
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   83
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label lblPreviousStaffRepayC 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   82
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblPreviousInstitutionRepayC 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   81
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblPreviousOtherRepayC 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   80
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblPreviousTotalRepayC 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   79
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   78
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   77
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   90
            Top             =   1440
            Width           =   2055
         End
      End
      Begin VB.Frame FrameSettleCredit 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   66
         Top             =   360
         Width           =   9615
         Begin VB.Frame Frame4 
            Height          =   975
            Left            =   7320
            TabIndex        =   111
            Top             =   1200
            Width           =   2175
            Begin VB.OptionButton OptionSettleCreditPrint 
               Caption         =   "Print"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton OptionSettleCreditDoNotPrint 
               Caption         =   "Do not print"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   600
               Width           =   1935
            End
         End
         Begin btButtonEx.ButtonEx bttnCashSettle 
            Height          =   375
            Left            =   6480
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   2280
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            Appearance      =   3
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
            BackStyle       =   0  'Transparent
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
            TabIndex        =   75
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   74
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
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
            TabIndex        =   73
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lblDoctorFeeToPay 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   72
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblHospitalFeeToPay 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   71
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblOtherFeeToPay 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   70
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblTotalFeeToPay 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            TabIndex        =   69
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
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
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            TabIndex        =   67
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label lblAgentName 
         Height          =   375
         Left            =   -74760
         TabIndex        =   95
         Top             =   1680
         Width           =   2655
      End
   End
   Begin VB.Frame FramePatient 
      Caption         =   "Add Patient"
      Height          =   4575
      Left            =   120
      TabIndex        =   46
      Top             =   4800
      Width           =   5055
      Begin VB.CheckBox chkScan 
         Caption         =   "&Scan"
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkForigner 
         Caption         =   "&Forigner"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin btButtonEx.ButtonEx bttnAddPatient 
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Add"
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   3735
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   6588
         _Version        =   393216
         Tab             =   2
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "&Cash"
         TabPicture(0)   =   "frmOldChannelingMS.frx":0532
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "FrameCash"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "A&gent"
         TabPicture(1)   =   "frmOldChannelingMS.frx":054E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FrameAgent"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Cr&edit"
         TabPicture(2)   =   "frmOldChannelingMS.frx":056A
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame1"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame1 
            Caption         =   "Credit"
            Height          =   3255
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   4455
            Begin VB.TextBox txtCreditContactNo 
               Height          =   360
               Left            =   1680
               MaxLength       =   35
               TabIndex        =   193
               Top             =   1920
               Width           =   2655
            End
            Begin VB.TextBox txtCreditPatientName 
               Height          =   360
               Left            =   1680
               MaxLength       =   35
               TabIndex        =   16
               Top             =   1440
               Width           =   2655
            End
            Begin MSDataListLib.DataCombo cmbTStaff 
               Height          =   360
               Left            =   1080
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   720
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin VB.CheckBox chkThroughAgent 
               Caption         =   "Through Staff"
               Height          =   240
               Left            =   120
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   360
               Value           =   1  'Checked
               Width           =   2895
            End
            Begin MSDataListLib.DataCombo cmbTStaffCode 
               Height          =   360
               Left            =   120
               TabIndex        =   197
               TabStop         =   0   'False
               Top             =   720
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin VB.Label Label60 
               BackStyle       =   0  'Transparent
               Caption         =   "Contact No."
               Height          =   255
               Left            =   120
               TabIndex        =   194
               Top             =   1920
               Width           =   2415
            End
            Begin VB.Label Label38 
               BackStyle       =   0  'Transparent
               Caption         =   "Patient Na&me"
               Height          =   255
               Left            =   120
               TabIndex        =   108
               Top             =   1440
               Width           =   2415
            End
            Begin VB.Label lblCredit 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   1680
               TabIndex        =   58
               Top             =   2760
               Width           =   2655
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Amount  : (Rs.)"
               Height          =   375
               Left            =   240
               TabIndex        =   57
               Top             =   2760
               Width           =   1455
            End
         End
         Begin VB.Frame FrameAgent 
            Caption         =   "Agent"
            Height          =   3255
            Left            =   -74880
            TabIndex        =   48
            Top             =   360
            Width           =   4455
            Begin MSDataListLib.DataCombo DataComboAgent 
               Height          =   360
               Left            =   1680
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   720
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin VB.TextBox txtAgentName 
               Height          =   375
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   109
               Top             =   720
               Width           =   2655
            End
            Begin VB.TextBox txtAgentPatientName 
               Height          =   360
               Left            =   1680
               MaxLength       =   35
               TabIndex        =   9
               Top             =   1680
               Width           =   2655
            End
            Begin VB.TextBox txtAgentRef 
               Height          =   360
               Left            =   1680
               TabIndex        =   8
               Top             =   1200
               Width           =   2655
            End
            Begin MSDataListLib.DataCombo DataComboAgentCode 
               Height          =   360
               Left            =   1680
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   240
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   635
               _Version        =   393216
               MatchEntry      =   -1  'True
               Style           =   2
               ListField       =   ""
               BoundColumn     =   ""
               Text            =   ""
               Object.DataMember      =   ""
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "&Patient Name"
               Height          =   255
               Left            =   120
               TabIndex        =   107
               Top             =   1680
               Width           =   2415
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Agent  Name   :"
               Height          =   255
               Left            =   120
               TabIndex        =   105
               Top             =   720
               Width           =   3135
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Ref. No."
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   1200
               Width           =   1695
            End
            Begin VB.Label txtAgentBalance 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   2520
               TabIndex        =   54
               Top             =   2760
               Width           =   1815
            End
            Begin VB.Label lblAgentAmount 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   2520
               TabIndex        =   52
               Top             =   2280
               Width           =   1815
            End
            Begin VB.Label Label29 
               BackStyle       =   0  'Transparent
               Caption         =   "&Agent Code     :"
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   240
               Width           =   3135
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "Agent &Balance : (Rs.)"
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   2760
               Width           =   2775
            End
            Begin VB.Label Label31 
               BackStyle       =   0  'Transparent
               Caption         =   "A&mount           : (Rs.)"
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   2280
               Width           =   2775
            End
         End
         Begin VB.Frame FrameCash 
            Caption         =   "Cash"
            Height          =   3255
            Left            =   -74880
            TabIndex        =   47
            Top             =   360
            Width           =   4455
            Begin VB.TextBox txtCashContactNo 
               Height          =   360
               Left            =   2040
               MaxLength       =   25
               TabIndex        =   190
               Top             =   720
               Width           =   2295
            End
            Begin VB.TextBox txtCashPatientName 
               Height          =   360
               Left            =   2040
               MaxLength       =   25
               TabIndex        =   4
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label Label59 
               BackStyle       =   0  'Transparent
               Caption         =   "Contact No."
               Height          =   255
               Left            =   120
               TabIndex        =   191
               Top             =   720
               Width           =   2415
            End
            Begin VB.Label Label28 
               BackStyle       =   0  'Transparent
               Caption         =   "Patient Na&me"
               Height          =   255
               Left            =   120
               TabIndex        =   106
               Top             =   240
               Width           =   2415
            End
            Begin VB.Label lblCashDue 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   2040
               TabIndex        =   50
               Top             =   1320
               Width           =   2295
            End
            Begin VB.Label Label35 
               BackStyle       =   0  'Transparent
               Caption         =   "Amount  :   (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   51
               Top             =   1320
               Width           =   3495
            End
         End
      End
      Begin VB.TextBox txtPatientName 
         Height          =   360
         Left            =   1440
         MaxLength       =   35
         TabIndex        =   44
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtContactNo 
         Height          =   360
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   192
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "&Patient Name"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   2415
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   13080
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   9240
      Width           =   2055
      _ExtentX        =   3625
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
      Height          =   3900
      Left            =   13560
      TabIndex        =   64
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListRoomNo 
      Height          =   2220
      Left            =   4080
      TabIndex        =   94
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "No.         Name         Payment    Paid        Can/Ref       P/Ab"
      Height          =   255
      Left            =   9840
      TabIndex        =   115
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Date           Sess.           Start     Booked"
      Height          =   255
      Left            =   5880
      TabIndex        =   114
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Consultant"
      Height          =   255
      Left            =   3000
      TabIndex        =   113
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Speciality"
      Height          =   255
      Left            =   240
      TabIndex        =   112
      Top             =   120
      Width           =   2535
   End
   Begin VB.Shape BoxPatients 
      BackStyle       =   1  'Opaque
      Height          =   4650
      Left            =   9600
      Top             =   120
      Width           =   5775
   End
   Begin VB.Shape BoxDates 
      BackStyle       =   1  'Opaque
      Height          =   4650
      Left            =   5760
      Top             =   120
      Width           =   3855
   End
   Begin VB.Shape BoxConsultant 
      BackStyle       =   1  'Opaque
      Height          =   4650
      Left            =   2880
      Top             =   120
      Width           =   2895
   End
   Begin VB.Shape BoxSpeciality 
      BackStyle       =   1  'Opaque
      Height          =   4650
      Left            =   120
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmOldChannelingMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
    Dim temSQL As String
    
 
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
    
    Dim TemSDoctorFee As Double
    Dim TemSFDoctorFee As Double
    Dim TemSADoctorFee As Double
    
    Dim TemSInstitutionFee As Double
    Dim TemSFInstitutionFee As Double
    Dim TemSAInstitutionFee As Double
    
    Dim TemSOtherFee As Double
    Dim TemSFOtherFee As Double
    Dim TemSAOtherFee As Double
    
    
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
'    Dim TemDoctorID As Long
    Dim TemAppointmentDate As Date
    Dim TemAppointmentTime As Date
    Dim TemDaySerial As Long
    
    Dim TemAgentBookingID As Long
    Dim TemSecessionStartingTime As Date
    Dim TemUsualDuration As Long
    Dim TemPatient As String
    Dim TemContactNo As String
    Dim TemConsultant As String
    Dim TemNonCancelledVisits As Long
    Dim TemBillId As Long
    
    Dim TemPreviousDate As Date
    Dim TemTextForList As String


Private Sub btnActiveView_Click()
Dim TemResponce As Long
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

NullToSero

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

If ListDatesAndSecessions.ListIndex < 0 Or IsNumeric(ListSecessionIDs.Text) = False Or IsDate(ListDates.Text) = False Then
    TemResponce = MsgBox("You have not selected a Date and secession", vbCritical, "No Date & Secession")
    ListDatesAndSecessions.SetFocus
    Exit Sub
End If

    With DataEnvironment1.rssqlDoctorView
        If .State = 1 Then .Close
        If PayToDoctor = True Then
            .Source = "SELECT tblPatientFacility.*, ( personalfee + PersonalFeeToPay - Personalrefund  ) as NetPersonalPayment  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where staff_ID = " & Val(ListConsultantIDs.Text) & " and Cancelled = 0 AND appointmentdate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " order by dayserial"
        Else
            .Source = "SELECT tblPatientFacility.*  , ( personalfee + PersonalFeeToPay  - Personalrefund  ) as NetPersonalPayment  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where staff_ID = " & Val(ListConsultantIDs.Text) & " and Cancelled = 0 AND appointmentdate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " and patientabsent = 0 order by dayserial"
        End If
        .Open
    End With
    With DataReportDoctorView1
        If HospitalDetails = True Then
            .Sections.Item("ReportHeader10").Controls.Item("RptName").Caption = InstitutionName
            .Sections.Item("ReportHeader10").Controls.Item("RptAddress").Caption = InstitutionAddress
'            .Sections.Item("Section4").Controls.Item("lblinstitutiontelephone").Caption = "Doctor View"
            .Sections.Item("Section2").Controls.Item("lbldoctorname").Caption = "Consultant : " & FindLDoctorFromID(Val(ListConsultantIDs.Text))
            .Sections.Item("Section2").Controls.Item("lbldatesecession").Caption = "Date : " & Format(MonthView1.Value, DefaultLongDate) & "  Secession : " & FindSecessionFromID(Val(ListSecessionIDs.Text))
            .Sections.Item("Section5").Controls.Item("lblad1").Caption = LongAd
        Else
            .Sections.Item("ReportHeader10").Controls.Item("RptName").Caption = Empty
            .Sections.Item("ReportHeader10").Controls.Item("RptAddress").Caption = Empty
'            .Sections.Item("Section4").Controls.Item("lblinstitutiontelephone").Caption = "Doctor View"
            .Sections.Item("Section2").Controls.Item("lbldoctorname").Caption = "Consultant : " & FindLDoctorFromID(Val(ListConsultantIDs.Text))
            .Sections.Item("Section2").Controls.Item("lbldatesecession").Caption = "Date : " & Format(MonthView1.Value, DefaultLongDate) & "  Secession : " & FindSecessionFromID(Val(ListSecessionIDs.Text))
            .Sections.Item("Section5").Controls.Item("lblad1").Caption = LongAd
        End If
        Set .DataSource = DataEnvironment1.rssqlDoctorView
        .Show
    End With

End Sub

Private Sub bttnAddPatient_Click()
    Dim TemResponce  As Integer
    Dim rsTem As New ADODB.Recordset
    
    FindSecessionDetails

     

    If Not IsNumeric(ListConsultantIDs.Text) Then
        TemResponce = MsgBox("You have not selected a name of the doctor", vbCritical, "No doctor")
        ListConsultants.SetFocus
        Exit Sub
    End If
    
    If SSTab1.Tab = 2 Then
        If chkThroughAgent.Value = 1 Then
        
        Else
            If Trim(txtCreditContactNo.Text) = "" Then
                TemResponce = MsgBox("You have to enter the contact number for all non-staff phone bookings", vbCritical, "Contact No?")
                txtCreditContactNo.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If Not IsDate(ListDates.Text) Then
        TemResponce = MsgBox("You have not selected a date", vbCritical, "No date")
        ListDatesAndSecessions.SetFocus
        Exit Sub
    End If
    
    If AgentNameForCreditBookings = True And SSTab1.Tab = 2 And chkThroughAgent.Value = 1 And Trim(txtPatientName.Text) = "" Then
        If Not IsNumeric(cmbTStaff.BoundText) Then
            TemResponce = MsgBox("You have not selected a staff member", vbCritical, "Agent?")
            cmbTStaffCode.SetFocus
            Exit Sub
        End If
        DataComboAgent.Text = cmbTStaff.Text
        txtPatientName.Text = cmbTStaff.Text & " (" & cmbTStaffCode.Text & ")"
    End If
    
    
    If Trim(txtPatientName.Text) = "" Then
        TemResponce = MsgBox("You have not entered a name of the patient to add", vbCritical, "No Name")
        Select Case SSTab1.Tab
            Case 0:     txtCashPatientName.SetFocus
            Case 1:     txtAgentPatientName.SetFocus
            Case 2:     txtCreditPatientName.SetFocus
            Case Else:  txtPatientName.SetFocus
        End Select
        Exit Sub
    ElseIf InStr(txtPatientName.Text, ";") Then
        TemResponce = MsgBox("You have entered an invalid name", vbCritical, "No Name")
        Select Case SSTab1.Tab
            Case 0:     txtCashPatientName.SetFocus
            Case 1:     txtAgentPatientName.SetFocus
            Case 2:     txtCreditPatientName.SetFocus
            Case Else:  txtPatientName.SetFocus
        End Select
        Exit Sub
    Else
        TemPatient = txtPatientName.Text
        TemContactNo = txtContactNo.Text
    End If
    
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT Count(tblpatientfacility.patientfacility_ID) as SecessionPatients from tblpatientfacility where hospitalfacility_ID = 10 and  AppointmentDate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " and cancelled <> 1 "
        .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
        If IsNull(!SecessionPatients) = False Then
            TemNonCancelledVisits = !SecessionPatients
             If SecessionMax <> 0 Then
                If TemNonCancelledVisits >= Val(ListSecessionMax.Text) Then
                    TemResponce = MsgBox("Adding this patient will increase the maximum number for the consultant. Do you still want to add the patient?", vbYesNo, "Exceed Maximum")
                    If TemResponce = vbNo Then .Close: Exit Sub
                End If
            End If
        End If
    End With
    
    
    If CanSettlePayment = False Then Exit Sub
        
    TemAgentRefNo = Trim(txtAgentRef.Text)
    
'    If AutomaticCapitalization = True Then
'        On Error Resume Next
'
'
'        Dim TemLeftName As String
'        Dim TemRightName As String
'        Dim TemResultName As String
'        Dim TemRemainingName As String
'        If InStr(1, txtPatientName.Text, " ") = 0 Then
'            TemResultName = UCase(Left(txtPatientName.Text, 1)) & LCase(Right(txtPatientName.Text, Len(txtPatientName.Text) - 1))
'            txtPatientName.Text = TemResultName
'        Else
'            TemRemainingName = txtPatientName.Text
'            While InStr(1, TemRemainingName, " ")
'                TemLeftName = Left(TemRemainingName, InStr(1, TemRemainingName, " ") - 1)
'                TemRightName = Right(TemRemainingName, Len(TemRemainingName) - (InStr(1, TemRemainingName, " ")))
'                TemRemainingName = TemRightName
'                TemResultName = TemResultName & " " & UCase(Left(TemLeftName, 1)) & LCase(Right(TemLeftName, Len(TemLeftName) - 1))
'            Wend
'                TemResultName = TemResultName & " " & UCase(Left(TemRemainingName, 1)) & LCase(Right(TemRemainingName, Len(TemRemainingName) - 1))
'                txtPatientName.Text = TemResultName
'        End If
'    Else
'            TemPatient = UCase(TemPatient)
'    End If
'
'    On Error GoTo 0
    
    txtPatientName.Text = UCase(Trim(txtPatientName.Text))
    
    If chkForigner.Value = 1 And chkScan.Value = 1 Then
        TemDoctorFee = TemSFDoctorFee
        TemInstitutionFee = TemSFInstitutionFee
        TemOtherFee = TemSFOtherFee
        If AddForeignerSuffix = True Then
            txtPatientName.Text = txtPatientName.Text & " (Foreigner)"
        End If
    ElseIf chkForigner.Value = 1 And chkScan.Value = 0 Then
        TemDoctorFee = TemFDoctorFee
        TemInstitutionFee = TemFInstitutionFee
        TemOtherFee = TemFOtherFee
    ElseIf SSTab1.Tab = 1 And chkScan.Value = 1 Then
        TemDoctorFee = TemSADoctorFee
        TemInstitutionFee = TemSAInstitutionFee
        TemOtherFee = TemSAOtherFee
        If AddForeignerSuffix = True Then
            txtPatientName.Text = txtPatientName.Text & " (S)"
        End If
    ElseIf SSTab1.Tab = 1 And chkScan.Value = 0 Then
        TemDoctorFee = TemADoctorFee
        TemInstitutionFee = TemAInstitutionFee
        TemOtherFee = TemAOtherFee
    ElseIf chkScan.Value = 1 Then
        TemDoctorFee = TemSADoctorFee
        TemInstitutionFee = TemSInstitutionFee
        TemOtherFee = TemSOtherFee
        txtPatientName.Text = txtPatientName.Text & " (S)"
    Else
        
    End If
    
    
    
    Dim TemTextForDisplay
    
    If AskBeforeAdding = True Then
        TemTextForDisplay = "Patient Name  " & vbTab & vbTab & ":" & vbTab & Trim(txtPatientName.Text) & vbNewLine
        Select Case SSTab1.Tab
            Case 0:
                        TemTextForDisplay = TemTextForDisplay & "Payment Method" & vbTab & vbTab & ":" & vbTab & "Cash" & vbNewLine
                        TemTextForDisplay = TemTextForDisplay & "Amount        " & vbTab & vbTab & ":" & vbTab & Format(TemDoctorFee + TemInstitutionFee + TemOtherFee, "0.00") & vbNewLine
            Case 1:
                        TemTextForDisplay = TemTextForDisplay & "Payment Method" & vbTab & vbTab & ":" & vbTab & "Agent" & vbNewLine
                        TemTextForDisplay = TemTextForDisplay & "Amount        " & vbTab & vbTab & ":" & vbTab & Format(TemDoctorFee + TemInstitutionFee + TemOtherFee, "0.00") & vbNewLine
                        TemTextForDisplay = TemTextForDisplay & "Agent         " & vbTab & vbTab & ":" & vbTab & DataComboAgent.Text & vbNewLine
                        TemTextForDisplay = TemTextForDisplay & "Agent Code    " & vbTab & vbTab & ":" & vbTab & DataComboAgentCode.Text & vbNewLine
            Case 2:
                        TemTextForDisplay = TemTextForDisplay & "Payment Method" & vbTab & vbTab & ":" & vbTab & "Credit" & vbNewLine
                        TemTextForDisplay = TemTextForDisplay & "Amount Paid   " & vbTab & vbTab & ":" & vbTab & "0.00" & vbNewLine
        End Select
        TemTextForDisplay = TemTextForDisplay & "Doctor Name   " & vbTab & vbTab & ":" & vbTab & ListConsultants.Text & vbNewLine
        TemTextForDisplay = TemTextForDisplay & "Date          " & vbTab & vbTab & ":" & vbTab & Format(ListDates.Text, DefaultLongDate) & vbNewLine
        TemTextForDisplay = TemTextForDisplay & "Secession     " & vbTab & vbTab & ":" & vbTab & FindSecessionFromID(ListSecessionIDs.Text) & vbNewLine
        TemResponce = MsgBox(TemTextForDisplay, vbQuestion + vbYesNo, "Add?")
        If TemResponce = vbNo Then
            ClearForNewPatient
            If AfterAddPatient = True Then
                Dim TemListConsultantID As Long
                Dim TemListSecessionID As Long
                TemListConsultantID = ListConsultants.ListIndex
                TemListSecessionID = ListDatesAndSecessions.ListIndex
                ListDatesAndSecessions_Click
                Call FillGridPatients
            ElseIf AfterAddSpeciality = True Then
                If ListSpecialities.ListCount > 0 Then
                    ListSpecialities.ListIndex = 0
                    ListSpecialities_Click
                End If
                ListSpecialities.SetFocus
            ElseIf AfterAddConsultant = True Then
                If ListConsultants.ListCount > 0 Then
                    ListConsultants.ListIndex = 0
                    ListConsultants_Click
                End If
                ListConsultants.SetFocus
            ElseIf AfterAddDates = True Then
                If ListDatesAndSecessions.ListCount > 0 Then
                    ListDatesAndSecessions.ListIndex = 0
                    ListDatesAndSecessions_Click
                End If
                ListDatesAndSecessions_Click
            End If
            Exit Sub
        End If
    End If
    
    Call AddPatient
    
'    Call AddToBill
    
    If AddToPatientFacility = False Then Exit Sub
    
    If SSTab1.Tab = 1 Then
        If AgentBookingValidation = False Then
            UpdateAgentCredit
            UpdateAgentFacility
        End If
        If AgentBillNumber = True Then Call UpdateAgentBill
    ElseIf SSTab1.Tab = 2 Then
        UpdatePatientCredit
    End If
        
        DisplayDetails
'        BillPrint
        
        If chkPrint.Value = 1 Then
            Call SetBillPrinter
            Call SetBillPaper
        Else
        
        End If
        
        
        
        ClearForNewPatient
                
        If AfterAddPatient = True Then
            TemListConsultantID = ListConsultants.ListIndex
            TemListSecessionID = ListDatesAndSecessions.ListIndex
            ListDatesAndSecessions_Click
            Call FillGridPatients
        ElseIf AfterAddSpeciality = True Then
            If ListSpecialities.ListCount > 0 Then
                ListSpecialities.ListIndex = 0
                ListSpecialities_Click
            End If
            ListSpecialities.SetFocus
        ElseIf AfterAddConsultant = True Then
            If ListConsultants.ListCount > 0 Then
                ListConsultants.ListIndex = 0
                ListConsultants_Click
            End If
            ListConsultants.SetFocus
        ElseIf AfterAddDates = True Then
            If ListDatesAndSecessions.ListCount > 0 Then
                ListDatesAndSecessions.ListIndex = 0
                ListDatesAndSecessions_Click
            End If
            ListDatesAndSecessions_Click
        End If
        

End Sub

Private Sub UpdateAgentBill()
    Dim tr As Integer
    
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT * FROM tblAgentRef WHERE tblAgentRef.AgentRefNo =" & Val(txtAgentRef.Text)
        .Open
        If .RecordCount = 0 Then
            tr = MsgBox("An Error occured when updating agent bill numbers. Make sure there are no two operators attending the booking for the same agency at once.", vbCritical, "Error")
            Exit Sub
        End If
        
        !Booked = True
        .Update
        If .State = 1 Then .Close
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
        If OnePrintForAgents = True And SSTab1.Tab = 1 Then
        
        Else
            'Printer.Line (100, 100)-(Printer.ScaleWidth - 100, Printer.ScaleHeight - 100)
            TemBoolean = PrintingPlainText(!date1x, !date1y, Format(Date, DefaultShortDate))
            TemBoolean = PrintingPlainText(!time1X, !time1y, Time)
            'TemBoolean = PrintingPlainText(!refno1x, !refno1y, TemPatientFacilityID)
            TemBoolean = PrintingPlainText(!consultant1x, !consultant1y, UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text))))
            TemBoolean = PrintingPlainText(!patient1x, !patient1y, TemPatient)
            TemBoolean = PrintingPlainText(!phone1x, !phone1y, TemContactNo)
            
            TemBoolean = PrintingPlainText(!appointon1x, !appointon1y, Format(ListDates.Text, DefaultShortDate))
            TemBoolean = PrintingPlainText(!at1x, !at1y, TemAppointmentTime)
            TemBoolean = PrintingPlainText(!drsfee1x, !drsfee1y, Format(TemDoctorFee, "0.00"))
            TemBoolean = PrintingPlainText(!total1x, !total1y, Format(TemDoctorFee + TemInstitutionFee, "0.00"))
            TemBoolean = PrintingPlainText(!hospchg1x, !hospchg1y, Format(TemInstitutionFee, "0.00"))
            TemBoolean = PrintingPlainText(!receptionist1x, !receptionist1y, UserName)
            TemBoolean = PrintingPlainText(!roomno1x, !roomno1y, ListRoomNo.Text)
            If SSTab1.Tab = 1 Then
                Printer.Font.Size = 9
                TemBoolean = PrintingPlainText(!agentcode1x, !agentcode1y, "(" & DataComboAgentCode.Text & ")")
                TemBoolean = PrintingPlainText(!agentrefno1x, !agentrefno1y, txtAgentRef.Text)
                Printer.Font.Size = 10
            End If
        End If
        TemBoolean = PrintingPlainText(!date2x, !date2y, Format(Date, DefaultShortDate))
        TemBoolean = PrintingPlainText(!time2X, !time2y, Time)
        'TemBoolean = PrintingPlainText(!refno2x, !refno2y, TemPatientFacilityID)
        TemBoolean = PrintingPlainText(!consultant2x, !consultant2y, UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text))))
        Printer.FontBold = True
        TemBoolean = PrintingPlainText(!patient2x, !patient2y, TemPatient)
        Printer.FontBold = False
        TemBoolean = PrintingPlainText(!phone2x, !phone2y, TemContactNo)
        TemBoolean = PrintingPlainText(!appointon2x, !appointon2y, Format(ListDates.Text, DefaultShortDate))
        TemBoolean = PrintingPlainText(!at2x, !at2y, TemAppointmentTime)
        TemBoolean = PrintingPlainText(!drsfee2x, !drsfee2y, Format(TemDoctorFee, "0.00"))
        TemBoolean = PrintingPlainText(!total2x, !total2y, Format(TemDoctorFee + TemInstitutionFee, "0.00"))
        TemBoolean = PrintingPlainText(!hospchg2x, !hospchg2y, Format(TemInstitutionFee, "0.00"))
        TemBoolean = PrintingPlainText(!receptionist2x, !receptionist2y, UserName)
        TemBoolean = PrintingPlainText(!roomno2x, !roomno2y, "ROOM " & ListRoomNo.Text)
        If SSTab1.Tab = 1 Then
            Printer.Font.Size = 9
            TemBoolean = PrintingPlainText(!agentcode2x, !agentcode2y, "(" & DataComboAgentCode.Text & ")")
            TemBoolean = PrintingPlainText(!agentrefno2x, !agentrefno2y, txtAgentRef.Text)
            Printer.Font.Size = 10
        End If
        Printer.Font.Size = 16
        Printer.Font.Bold = True
        Printer.Font.Name = "Arial"
        If OnePrintForAgents = True And SSTab1.Tab = 1 Then
        
        Else
            TemBoolean = PrintingPlainText(!appono1x, !appono1y, TemDaySerial)
        End If
        TemBoolean = PrintingPlainText(!appono2x, !appono2y, TemDaySerial)
        .Close
    End With
    Printer.EndDoc
End Sub

Private Sub SetBillPrinter1()
    CSetPrinter.SetPrinterAsDefault (BillPrinterName)
End Sub

Private Sub SetBillPaper1()
Dim TemResponce As Long
Dim RetVal As Integer
RetVal = SelectForm(BillPaperName, Me.hwnd)
Select Case RetVal
    Case FORM_NOT_SELECTED   ' 0
        TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
    Case FORM_SELECTED   ' 1
        Call SelectPrint1
    Case FORM_ADDED   ' 2
        TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
End Select
End Sub

Private Sub SelectPrint1()
        If PrintingOnBlankPaper = True Then
            BillPrint21
        ElseIf PrintingOnPrintedPaper = True Then
            BillPrint31
        End If
End Sub

Private Sub BillPrint31()
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
        TemBoolean = PrintingPlainText(!time1X, !time1y, Time)
'        TemBoolean = PrintingPlainText(!refno1x, !refno1y, TemPatientFacilityID)
        TemBoolean = PrintingPlainText(!consultant1x, !consultant1y, UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text))))
        TemBoolean = PrintingPlainText(!patient1x, !patient1y, TemPatient)
        TemBoolean = PrintingPlainText(!phone1x, !phone1y, TemContactNo)
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
        TemBoolean = PrintingPlainText(!time2X, !time2y, Time)
'        TemBoolean = PrintingPlainText(!refno2x, !refno2y, TemPatientFacilityID)
        TemBoolean = PrintingPlainText(!consultant2x, !consultant2y, UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text))))
        TemBoolean = PrintingPlainText(!patient2x, !patient2y, TemPatient)
        TemBoolean = PrintingPlainText(!phone2x, !phone2y, TemContactNo)
        
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


Private Sub AddToBill()
'With DataEnvironment1.rssqlTem5
'    If .State = 1 Then .Close
'    .Source = "SELECT * from tblpatientbill"
'    .Open
'    .AddNew
'    !patient_ID = TemPatientID
'    !Date = Date
'    !NetTotal = TemDoctorFee + TemInstitutionFee
'    !GrossTotal = TemDoctorFee + TemInstitutionFee
'    Select Case SSTab1.Tab
'    Case 0:
'        !PaymentMethod = "Cash"
'        !Cash = TemDoctorFee + TemInstitutionFee
'    Case 1:
'        !PaymentMethod = "Agent"
'        !AgentAmount = TemDoctorFee + TemInstitutionFee
'        If IsNumeric(DataComboAgent.BoundText) = True Then !Agent_ID = DataComboAgent.BoundText
'    Case 2:
'        If IsNumeric(cmbTStaff.BoundText) Then !creditstaff_ID = cmbTStaff.BoundText
'        !PaymentMethod = "Credit"
'        !Credit = TemDoctorFee + TemInstitutionFee
'    End Select
'        !user_ID = UserID
'        !BillSuccess = True
'    If chkPrint.Value = 1 And SSTab1.Tab = 2 Then
'        !billprinted = False
'    ElseIf chkPrint.Value = 1 Then
'        !billprinted = True
'    Else
'        !billprinted = False
'    End If
'    .Update
'    TemBillId = !PatientBill_ID
'    .Close
'End With
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

Private Sub ClearForNewPatient()
    txtPatientName.Text = Empty
    txtContactNo.Text = Empty
    
    chkForigner.Value = 0
    chkScan.Value = 0
    chkScan.Visible = False
    If ChangeToCash = True Then SSTab1.Tab = 0
    
    txtCashPatientName.Text = Empty
    txtCashContactNo.Text = Empty
    txtCreditPatientName.Text = Empty
    txtCreditContactNo.Text = Empty
    txtAgentPatientName.Text = Empty
    txtAgentRef.Text = Empty
    
    
    If ClearAgentDetails = True Then
        DataComboAgent.Text = Empty
    End If
    
End Sub



Private Function DisplayDetails() As Boolean
    DisplayDetails = True
    Dim TemResponce
    Dim temText As String
    temText = temText & "Patient Name" & vbTab & " : " & vbTab & TemPatient & vbNewLine
    temText = temText & "Appointment ID" & vbTab & " : " & vbTab & TemPatientFacilityID & vbNewLine
    temText = temText & "Appointment Time" & vbTab & " : " & vbTab & TemAppointmentTime & vbNewLine & vbNewLine
    If UserAuthority = AuthorityAdministrator Or UserAuthority = AuthorityOwner Then
        temText = temText & "Appointment No" & vbTab & " : " & vbTab & TemDaySerial & vbNewLine & vbNewLine
    End If
    If SSTab1.Tab = 1 Then
        temText = temText & "Agent Referance No." & vbTab & " : " & vbTab & TemAgentRefNo
    End If
    TemResponce = MsgBox(temText, vbInformation, "Booking Details")
End Function

Private Sub UpdateAgentFacility()
    With DataEnvironment1.rssqlTem7
        If .State = 1 Then .Close
        .Source = "Select * from tblagentbooking"
        .Open
        .AddNew
        !Agent_ID = Val(txtTemAgentID.Text)
        !patientfacility_ID = TemPatientFacilityID
        !BookingDate = Date
        !BookingTime = Time
        !patient_ID = TemPatientID
        !AppointmentDate = ListDates.Text
        !AgentRefNo = Trim(txtAgentRef.Text)
        .Update
        TemAgentBookingID = !AgentBooking_ID
        .Close
    End With
End Sub

Private Sub UpdateAgentCredit()
    With DataEnvironment1.rssqlTem7
        If .State = 1 Then .Close
        .Source = "SELECT tblinstitutions.* from tblinstitutions where institution_ID =" & DataComboAgent.BoundText
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        
        ' To Check
        
        !InstitutionCredit = !InstitutionCredit - Val(TemDoctorFee + TemInstitutionFee)
        
        ' End Check
        
        .Update
        .Close
    End With
End Sub


Private Sub AddPatient()
'    Dim loRs
'    Dim temSql As String
'    Dim rsTem As New ADODB.Recordset
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "Select * from tblpatientmaindetails"
'        .Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
'        .AddNew
'        !FirstName = UCase(txtPatientName.Text)
'        .Update
'        temSql = "SELECT @@IDENTITY AS NewID"
'        .Close
'        .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
'        TemPatientID = !NewID
'        .Close
'    End With

    Dim loRs
    Dim temSQL As String
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SET NOCOUNT ON;" & _
        "INSERT INTO tblpatientmaindetails (firstname, phone) VALUES ('" & UCase(txtPatientName.Text) & "', '" & txtContactNo.Text & "');" & _
        "SELECT @@IDENTITY AS NewID;"
        Set loRs = cnnChannelling.Execute(temSQL)
        TemPatientID = loRs.Fields("NewID").Value
    End With


'    Dim loRs
'    Dim temSql As String
'    Dim rsTem As New ADODB.Recordset
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "Select * from tblpatientmaindetails"
'
'        .Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
'        .AddNew
'        !firstname = UCase(txtPatientName.Text)
'        .Update
'        TemPatientID = !patient_ID
'        .Close
'    End With


End Sub

Private Function SecessionSerialNO() As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT Max(tblpatientfacility.dayserial) as C from tblpatientfacility where AppointmentDate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " "
        .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
        If IsNull(!C) = False Then
            SecessionSerialNO = !C + 1
        Else
            SecessionSerialNO = 1
        End If
        .Close
    End With
End Function

Private Function AddToPatientFacility() As Boolean





'    Dim rsTem As New ADODB.Recordset
'
'    AddToPatientFacility = False
'
'    Call FindAppointmentTime
'
'    With rsTem
'        If .State = 1 Then .Close
'        'temSql = "SELECT tblpatientfacility.* from tblpatientfacility where hospitalfacility_ID = 10 and AppointmentDate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " order by dayserial"
'        temSql = "SELECT * from tblpatientfacility"
'        If .State = 0 Then .Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
'        .AddNew
'        !user_ID = UserID
'        !patientid = TemPatientID
'        !HospitalFacility_ID = 10
'        !FacilityCatogery = Doctor
'        !PatientBill_ID = TemBillId
'        !Staff_ID = Val(ListConsultantIDs.Text)
'        !BookingDate = Date
'        !bookingtime = Time
'        !AppointmentDate = ListDates.Text
'        !Secession = Val(ListSecessionIDs.Text)
'        !appointmenttime = TemAppointmentTime
'        If SSTab1.Tab = 2 Then
'            !FullyPaid = 0
'        ElseIf SSTab1.Tab = 1 And AgentBookingValidation = True Then
'            !FullyPaid = 0
'        Else
'            !FullyPaid = 1
'            !fullypaidnull = 1
'        End If
'        !cancelled = False
'        !resultsuccess = True
'        If SSTab1.Tab = 0 Then
'            !personalfee = TemDoctorFee
'            !personaldue = TemDoctorFee
'            !institutionfee = TemInstitutionFee
'            !institutiondue = TemInstitutionFee
'            !otherfee = 0
'            !otherdue = 0
'            !totalfee = TemDoctorFee + TemInstitutionFee
'            !TotalDue = TemDoctorFee + TemInstitutionFee
'            !PersonalFeeToPay = 0
'            !InstitutionFeeToPay = 0
'            !otherfeetopay = 0
'            !totalfeetopay = 0
'            !PaymentMode = "Cash"
'            !paymentmethod_ID = 1
'            !FullyPaid = 1
'            !fullypaidnull = 1
'        ElseIf SSTab1.Tab = 1 Then
'            If AgentBookingValidation = False Then
'                !personalfee = TemDoctorFee
'                !personaldue = TemDoctorFee
'                !institutionfee = TemInstitutionFee
'                !institutiondue = TemInstitutionFee
'                !otherfee = 0
'                !otherdue = 0
'                !totalfee = TemDoctorFee + TemInstitutionFee
'                !TotalDue = TemDoctorFee + TemInstitutionFee
'                !PersonalFeeToPay = 0
'                !InstitutionFeeToPay = 0
'                !otherfeetopay = 0
'                !totalfeetopay = 0
'                !PaymentMode = "Agent"
'                !paymentmethod_ID = 2
'                !Agent_ID = Val(DataComboAgent.BoundText)
'                !FullyPaid = 1
'                !fullypaidnull = 1
'                !AgentRefNo = Trim(txtAgentRef.Text)
'            Else
'                !personalfee = 0
'                !personaldue = 0
'                !institutionfee = 0
'                !institutiondue = 0
'                !otherfee = 0
'                !otherdue = 0
'                !totalfee = 0
'                !TotalDue = 0
'                !PersonalFeeToPay = TemDoctorFee
'                !InstitutionFeeToPay = TemInstitutionFee
'                !otherfeetopay = 0
'                !totalfeetopay = TemDoctorFee + TemInstitutionFee
'                !PaymentMode = "Agent"
'                !paymentmethod_ID = 2
'                !Agent_ID = Val(DataComboAgent.BoundText)
'                !FullyPaid = 0
'                !AgentRefNo = Trim(txtAgentRef.Text)
'            End If
'        ElseIf SSTab1.Tab = 2 Then
'            !personalfee = 0
'            !personaldue = 0
'            !institutionfee = 0
'            !institutiondue = 0
'            !otherfee = 0
'            !otherdue = 0
'            !totalfee = 0
'            !TotalDue = 0
'            !PersonalFeeToPay = TemDoctorFee
'            !InstitutionFeeToPay = TemInstitutionFee
'            !otherfeetopay = 0
'            !totalfeetopay = TemDoctorFee + TemInstitutionFee
'            !PaymentMode = "Credit"
'            !paymentmethod_ID = 4
'            If IsNumeric(cmbTStaff.BoundText) Then !creditstaff_ID = cmbTStaff.BoundText
'            !FullyPaid = 0
'        End If
'        If chkPrint.Value = 1 And SSTab1.Tab = 2 Then
'            !billprinted = False
'        ElseIf chkPrint.Value = 1 Then
'            !billprinted = True
'        Else
'            !billprinted = False
'        End If
'        TemDaySerial = SecessionSerialNO
'        !DaySerial = TemDaySerial
'        .Update
'        temSql = "SELECT @@IDENTITY AS NewID"
'        .Close
'        .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
'        TemPatientFacilityID = !NewID
'        .Close
'    End With
'Call FillGridPatients
'AddToPatientFacility = True


    Dim temFields As String

    temFields = "Agent_ID ,AgentRefNo ,AppointmentDate ,appointmenttime ,billprinted ,BookingDate ,bookingtime ,cancelled ,creditstaff_ID ,DaySerial ,FacilityCatogery ,FullyPaid ,fullypaidnull ,HospitalFacility_ID ,institutiondue ,institutionfee ,InstitutionFeeToPay ,IsScan ,otherdue ,otherfee ,otherfeetopay ,PatientBill_ID ,patientid ,paymentmethod_ID ,PaymentMode ,personaldue ,personalfee ,PersonalFeeToPay ,resultsuccess ,Secession ,Staff_ID ,TotalDue ,totalfee ,totalfeetopay ,user_ID "

    Dim rsTem1 As New ADODB.Recordset

    AddToPatientFacility = False
    
    Call FindAppointmentTime

    With rsTem1
        If .State = 1 Then .Close
        'temSql = "SELECT tblpatientfacility.* from tblpatientfacility where hospitalfacility_ID = 10 and AppointmentDate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " order by dayserial"
        'temSql = "SELECT " & temFields & " from tblpatientfacility"
        temSQL = "SELECT * from tblpatientfacility where patientfacility_ID = 0"
        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
        
        
        
        .AddNew
        !user_ID = UserID
        !patientid = TemPatientID
        !HospitalFacility_ID = 10
        !FacilityCatogery = Doctor
        !PatientBill_ID = TemBillId
        !Staff_ID = Val(ListConsultantIDs.Text)
        !BookingDate = Date
        !BookingTime = Time
        !AppointmentDate = ListDates.Text
        !Secession = Val(ListSecessionIDs.Text)
        !appointmenttime = TemAppointmentTime
        If SSTab1.Tab = 2 Then
            !FullyPaid = 0
        ElseIf SSTab1.Tab = 1 And AgentBookingValidation = True Then
            !FullyPaid = 0
        Else
            !FullyPaid = 1
            !fullypaidnull = 1
        End If
        !Cancelled = False
        !resultsuccess = True
        If SSTab1.Tab = 0 Then
            !PersonalFee = TemDoctorFee
            !personaldue = TemDoctorFee
            !InstitutionFee = TemInstitutionFee
            !institutiondue = TemInstitutionFee
            !otherfee = 0
            !otherdue = 0
            !totalfee = TemDoctorFee + TemInstitutionFee
            !TotalDue = TemDoctorFee + TemInstitutionFee
            !PersonalFeeToPay = 0
            !InstitutionFeeToPay = 0
            !otherfeetopay = 0
            !totalfeetopay = 0
            !PaymentMode = "Cash"
            !PaymentMethod_ID = 1
            !FullyPaid = 1
            !fullypaidnull = 1
        ElseIf SSTab1.Tab = 1 Then
'            If AgentBookingValidation = False Then
                !PersonalFee = TemDoctorFee
                !personaldue = TemDoctorFee
                !InstitutionFee = TemInstitutionFee
                !institutiondue = TemInstitutionFee
                !otherfee = 0
                !otherdue = 0
                !totalfee = TemDoctorFee + TemInstitutionFee
                !TotalDue = TemDoctorFee + TemInstitutionFee
                !PersonalFeeToPay = 0
                !InstitutionFeeToPay = 0
                !otherfeetopay = 0
                !totalfeetopay = 0
                !PaymentMode = "Agent"
                !PaymentMethod_ID = 2
                !Agent_ID = Val(DataComboAgent.BoundText)
                !FullyPaid = 1
                !fullypaidnull = 1
                !AgentRefNo = Trim(txtAgentRef.Text)
'            Else
'                !personalfee = 0
'                !personaldue = 0
'                !institutionfee = 0
'                !institutiondue = 0
'                !otherfee = 0
'                !otherdue = 0
'                !totalfee = 0
'                !TotalDue = 0
'                !PersonalFeeToPay = TemDoctorFee
'                !InstitutionFeeToPay = TemInstitutionFee
'                !otherfeetopay = 0
'                !totalfeetopay = TemDoctorFee + TemInstitutionFee
'                !PaymentMode = "Agent"
'                !PaymentMethod_ID = 2
'                !Agent_ID = Val(DataComboAgent.BoundText)
'                !FullyPaid = 0
'                !AgentRefNo = Trim(txtAgentRef.Text)
'            End If
        ElseIf SSTab1.Tab = 2 Then
            !PersonalFee = 0
            !personaldue = 0
            !InstitutionFee = 0
            !institutiondue = 0
            !otherfee = 0
            !otherdue = 0
            !totalfee = 0
            !TotalDue = 0
            
'            If TemDoctorFee = 0 Or TemInstitutionFee = 0 Then ListDatesAndSecessions_Click
            
            !PersonalFeeToPay = TemDoctorFee
            !InstitutionFeeToPay = TemInstitutionFee
            !otherfeetopay = 0
            !totalfeetopay = TemDoctorFee + TemInstitutionFee
            !PaymentMode = "Credit"
            !PaymentMethod_ID = 4
            If IsNumeric(cmbTStaff.BoundText) Then !CreditStaff_ID = Val(cmbTStaff.BoundText)
            !FullyPaid = 0
        End If
        If chkPrint.Value = 1 And SSTab1.Tab = 2 Then
            !billprinted = False
        ElseIf chkPrint.Value = 1 Then
            !billprinted = True
        Else
            !billprinted = False
        End If
        TemDaySerial = SecessionSerialNO
        !DaySerial = TemDaySerial
        If chkScan.Value = 1 Then
            !IsScan = True
        End If
        .Update
        TemPatientFacilityID = !patientfacility_ID
        .Close
    End With
Call FillGridPatients
AddToPatientFacility = True


End Function


Private Function CanSettlePayment() As Boolean
    Dim TemResponce  As Integer
    CanSettlePayment = False
    
    
    
    
    Select Case SSTab1.Tab
    
    Case 0:
            If Val(lblCashDue.Caption) <= 0 Then
                FindSecessionDetails
            End If
    
    
    Case 1:
            If Val(lblAgentAmount.Caption) <= 0 Then
                FindSecessionDetails
            End If
    
    Case 2:
            If Val(lblCredit.Caption) <= 0 Then
                FindSecessionDetails
            End If
    
    End Select
    

    If PaymentCash <> 1 And SSTab1.Tab = 0 Then
        TemResponce = MsgBox("You are not allowed to do cash bookings using this computer. Please select another payment method or ask the administration to change the preferances so that you are allowed to accept cash", vbInformation, "Cash NOT allowed")
        SSTab1.SetFocus
        Exit Function
    End If
    
    If PaymentCredit <> 1 And SSTab1.Tab = 2 Then
        TemResponce = MsgBox("You are not allowed to do credit bookings using this computer. Please select another payment method or ask the administration to change the preferances so that you are allowed to do credit bookings", vbInformation, "Credit NOT allowed")
        SSTab1.SetFocus
        Exit Function
    End If
    
    If PaymentAgent <> 1 And SSTab1.Tab = 1 Then
        TemResponce = MsgBox("You are not allowed to do agent bookings using this computer. Please select another payment method or ask the administration to change the preferances so that you are allowed to do agent bookings", vbInformation, "Agent NOT allowed")
        SSTab1.SetFocus
        Exit Function
    End If
    
    
    Select Case SSTab1.Tab
    
    Case 0:
    
    
    Case 1:
        If Not IsNumeric(DataComboAgent.BoundText) Then
            TemResponce = MsgBox("You have not selected an agent", vbInformation, "Agent")
            If DataComboAgent.Visible = True And DataComboAgent.Enabled = True Then
                DataComboAgent.SetFocus
            Else
                DataComboAgentCode.SetFocus
            End If
            Exit Function
        End If
        
        If AgentEssential = True Then
            If Trim(txtAgentRef.Text) = "" Then
                TemResponce = MsgBox("You have not entered the agent referance No.", vbCritical, "Agent referance No")
                txtAgentRef.SetFocus
                Exit Function
            End If
        End If

        If TemAgentCredit - (TemDoctorFee + TemInstitutionFee) < (0 - TemAgentMaxCredit) Then
            TemResponce = MsgBox("This bill will lead to increase the credit limit of the agent. If you want to proceed, increase the credit limit or adviced the agent to settle cash", vbInformation, "Credit Limit")
            If DataComboAgent.Visible = True Then
                DataComboAgent.SetFocus
            Else
                DataComboAgentCode.SetFocus
            End If
            Exit Function
        End If
    
    
    
'        Dim tr As Integer
'
'        With DataEnvironment1.rssqlTem
'            If .State = 1 Then .Close
'            .Source = "SELECT * FROM tblAgentRef WHERE tblAgentRef.AgentRefNo =" & Val(txtAgentRef.Text)
'            .Open
'            If .RecordCount = 0 Then
'                tr = MsgBox("An Error occured when updating agent bill numbers. Make sure there are no two operators attending the booking for the same agency at once.", vbCritical, "Error")
'                Exit Sub
'            End If
'
'            !booked = True
'            .Update
'            If .State = 1 Then .Close
'        End With

    
    
    
        If chkForigner = 1 Then
            TemDoctorFee = TemFDoctorFee
            TemInstitutionFee = TemFInstitutionFee
            TemOtherFee = TemFOtherFee
        End If
    
    
    Case 2:
        If chkThroughAgent.Value = 0 Then
            If CanChannellByPhone(Val(ListConsultantIDs.Text)) = False Then
                TemResponce = MsgBox("Telephone bookings are NOT allowed for this consultant", vbInformation, "No Telephone Bookings")
                Exit Function
            End If
        End If
    
    End Select
    CanSettlePayment = True
End Function





Private Sub bttnAgentBookingValidation_Click()






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
'            !SettleCashDate = Date
'            !SettleCashTime = Time
'            !CreditSettleUser_ID = UserID
            .Update
            TemBillId = !PatientBill_ID
        .Close
    End With
    
    If OptionSettleCreditPrint.Value = True Then
        Call SetBillPrinter1
        Call SetBillPaper1
    Else
    
    End If
    
    Call FormatGridPatients
    Call ListDatesAndSecessions_Click


    With DataEnvironment1.rssqlTem7
        If .State = 1 Then .Close
        .Source = "SELECT tblinstitutions.* from tblinstitutions where institution_ID =" & Val(txtTemAgentID.Text)
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        !InstitutionCredit = !InstitutionCredit - (Val(lblDoctorFeeToPay.Caption) + Val(lblHospitalFeeToPay.Caption))
        .Update
        .Close
    End With

            UpdateAgentFacility


End Sub

Private Sub bttnAllDoctors_Click()
    If PartialRepayments = True Then
        Const PreSHape = "SHAPE {"
        Const Sql = "SELECT tblPatientFacility.*, tblDoctor.DoctorListedName, tblPatientMainDetails.FirstName, tblTitle.Title FROM tblTitle RIGHT JOIN ((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblDoctor ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID Where "
        Const PostSHape = "(((tblPatientFacility.HospitalFacility_ID) = 10)) }  AS cmmdTotalDoctorFee COMPUTE cmmdTotalDoctorFee, SUM(cmmdTotalDoctorFee.'PersonalDue') AS DocDue, SUM(cmmdTotalDoctorFee.'InstitutionDue') AS HosDue, SUM(cmmdTotalDoctorFee.'TotalDue') AS TotDue, ANY(cmmdTotalDoctorFee.'DoctorListedName') AS DoctorNameToDisplay, ANY(cmmdTotalDoctorFee.'Title') AS DoctorTitleToDisplay BY 'DoctorListedName' "
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

        With DataEnvironment1
            If .rscmmdTotalDoctorFee_Grouping.State = 1 Then .rscmmdTotalDoctorFee_Grouping.Close
            .Commands!cmmdTotalDoctorFee_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & Date & "' and " & PostSHape
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
                .Commands!NewDoctorView_Grouping.CommandText = PreSHape1 & Sql1 & " appointmentdate = '" & Date & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and " & PostSHape1
            Else
                .Commands!NewDoctorView_Grouping.CommandText = PreSHape1 & Sql1 & " appointmentdate = '" & Date & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and patientabsent = 0 and " & PostSHape1
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

CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1
    
        If .rscmmdAllDoctorPatients_Grouping.State = 1 Then .rscmmdAllDoctorPatients_Grouping.Close
        
        If DetailedCount = False Then
            If PayToDoctor = True Then
                .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & Date & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and " & PostSHape
            Else
                .Commands!cmmdAllDoctorPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & Date & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and patientabsent = 0 and " & PostSHape
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

Private Sub bttnAllSecessionPatients_Click()
    Const PreSHape = "SHAPE {"
    Const Sql = "SELECT tblPatientFacility.*, tblDoctor.DoctorListedName, tblFacilitySecession.SecessionName, tblTitle.Title , tblPatientMainDetails.FirstName FROM tblTitle RIGHT JOIN (tblDoctor RIGHT JOIN (tblFacilitySecession RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblFacilitySecession.FacilitySecession_ID = tblPatientFacility.Secession) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID Where "
    Const PostSHape = "(((tblPatientFacility.HospitalFacility_ID)=10))}  AS AllSecessionPatients COMPUTE AllSecessionPatients, ANY(AllSecessionPatients.'DoctorListedName') AS SecessionDoctorName, ANY(AllSecessionPatients.'SecessionName') AS ThisSecessionName, SUM(AllSecessionPatients.'CancelledNull') AS AllCancelled, SUM(AllSecessionPatients.'RefundNull') AS AllRefunds, SUM(AllSecessionPatients.'PatientAbsentNull') AS AllAbsent, SUM(AllSecessionPatients.'FullyPaidNull') AS AllFullyPaid, COUNT(AllSecessionPatients.'PatientFacility_ID') AS AllPatients, ANY(AllSecessionPatients.'Title') AS DoctorTitle BY 'DoctorListedName','SecessionName' "

CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

    With DataEnvironment1
        If .rsAllSecessionPatients_Grouping.State = 1 Then .rsAllSecessionPatients_Grouping.Close
        
        If DetailedCount = False Then
            If PayToDoctor = True Then
                .Commands!AllSecessionPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & Date & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and " & PostSHape
            Else
                .Commands!AllSecessionPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & Date & "' and fullypaid = 1 and cancelled = 0 and refund = 0 and patientabsent = 0 and " & PostSHape
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
                .Commands!AllSecessionPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & Date & "' and  " & PostSHape
            Else
                .Commands!AllSecessionPatients_Grouping.CommandText = PreSHape & Sql & " appointmentdate = '" & Date & "' and patientabsent = 0 and " & PostSHape
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
    TemResponce = MsgBox("Are You Sure You want to Cancel this Patient  ?", vbCritical + vbYesNo, "Cancellation")
    If TemResponce = vbNo Then Exit Sub
    
    With DataEnvironment1.rssqlTem7
    

        If .State = 1 Then .Close
        .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(txtBookingID.Text)
        .Open
        
        If !PaymentMode = "Agent" Then
            If OptionRepayAgent.Value = False And OptionRepayPatient.Value = False Then
                TemResponce = MsgBox("You have not selected wether to repay the patient or the agent. Please select one.", vbQuestion, "Repay to whom?")
                OptionRepayPatient.SetFocus
                Exit Sub
            End If
        Else
            OptionRepayAgent.Value = False
            OptionRepayPatient.Value = False
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
        
'        If !FullyPaid = 0 Then
'            TemResponce = MsgBox("The patient has not completed the payment. You can't cancel it", vbCritical, "Repaied")
''            txtBookingID.SetFocus
'            SendKeys "{home}+{end}"
'            Exit Sub
'        End If
    
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
                !repaytime = Time
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
                    !repaytime = Time
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
                !repaytime = Time
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
                    !repaytime = Time
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
                !repaytime = Time
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
                    !repaytime = Time
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

    If Val(lblTotalFeeToPay.Caption) <= 0 Then
        TemResponce = MsgBox("There has been an error. Please cancel this booking")
        Exit Sub
    End If

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
            
            
            TemDoctorFee = Val(lblDoctorFeeToPay.Caption)
            TemInstitutionFee = Val(lblHospitalFeeToPay.Caption)

            
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
        Call SetBillPrinter1
        Call SetBillPaper1
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
    
    If ListPatientFacilities.ListIndex < 0 Or IsNumeric(ListPatientFacilityIDs.Text) = False Then
    TemResponce = MsgBox("You have not selected a patient to change Name", vbCritical, "Patient?")
    ListPatientFacilities.SetFocus
    Exit Sub
    End If
    
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
        .Source = "select * from tblPatientFacility where PatientFacility_ID = " & Val(txtBookingID.Text)
        .Open
        
        If !FullyPaid = 1 Then
        TemResponce = MsgBox("The patient completed the payment. You can't Change the Name", vbCritical, "Can't Change")
        txtBookingID.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
        End If

        TemResponce = MsgBox("Are You Sure You want to Change Patient Name  ?", vbCritical + vbYesNo, "Change Name")
        If TemResponce = vbNo Then Exit Sub
    
        If .State = 1 Then .Close
        .Source = "select * from tblpatientmaindetails where patient_ID = " & Val(txtBookedPatientID.Text)
        .Open
        If .RecordCount = 0 Then Exit Sub
        !FirstName = Trim(txtNameChange.Text)
        !Phone = Trim(txtPhoneChange.Text)
        .Update
        .Close
    End With
    txtBookedPatientName.Text = txtNameChange.Text
    txtBookedPatientContactNo.Text = txtPhoneChange.Text
    ListDatesAndSecessions_Click
    txtNameChange.Text = Empty
    txtPhoneChange.Text = Empty
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub



Private Sub bttnDoctorView_Click()
Dim TemResponce As Long
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

NullToSero

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

If ListDatesAndSecessions.ListIndex < 0 Or IsNumeric(ListSecessionIDs.Text) = False Or IsDate(ListDates.Text) = False Then
    TemResponce = MsgBox("You have not selected a Date and secession", vbCritical, "No Date & Secession")
    ListDatesAndSecessions.SetFocus
    Exit Sub
End If

    With DataEnvironment1.rssqlDoctorView
        If .State = 1 Then .Close
        If PayToDoctor = True Then
            .Source = "SELECT tblPatientFacility.*, ( personalfee + PersonalFeeToPay - Personalrefund  ) as NetPersonalPayment  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where staff_ID = " & Val(ListConsultantIDs.Text) & " and appointmentdate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " order by dayserial"
        Else
            .Source = "SELECT tblPatientFacility.*  , ( personalfee + PersonalFeeToPay  - Personalrefund  ) as NetPersonalPayment  , tblPatientMainDetails.FirstName FROM  tblPatientFacility LEFT OUTER JOIN    tblPatientMainDetails ON    tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID where staff_ID = " & Val(ListConsultantIDs.Text) & " and appointmentdate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " and patientabsent = 0 order by dayserial"
        End If
        .Open
    End With
    With DataReportDoctorView1
        If HospitalDetails = True Then
            .Sections.Item("ReportHeader10").Controls.Item("RptName").Caption = InstitutionName
            .Sections.Item("ReportHeader10").Controls.Item("RptAddress").Caption = InstitutionAddress
'            .Sections.Item("Section4").Controls.Item("lblinstitutiontelephone").Caption = "Doctor View"
            .Sections.Item("Section2").Controls.Item("lbldoctorname").Caption = "Consultant : " & FindLDoctorFromID(Val(ListConsultantIDs.Text))
            .Sections.Item("Section2").Controls.Item("lbldatesecession").Caption = "Date : " & Format(MonthView1.Value, DefaultLongDate) & "  Secession : " & FindSecessionFromID(Val(ListSecessionIDs.Text))
            .Sections.Item("Section5").Controls.Item("lblad1").Caption = LongAd
        Else
            .Sections.Item("ReportHeader10").Controls.Item("RptName").Caption = Empty
            .Sections.Item("ReportHeader10").Controls.Item("RptAddress").Caption = Empty
'            .Sections.Item("Section4").Controls.Item("lblinstitutiontelephone").Caption = "Doctor View"
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
        If !paidtostaff = True Then
            TemResponce = MsgBox("This patient fee is already paid to the doctor. You can't make present or absent after paying the doctor", vbInformation, "Present")
            .Close
            Exit Sub
        End If
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
    
    TemResponce = MsgBox("Are You Sure You want to Marke As Present this Patient  ?", vbCritical + vbYesNo, " Marke as Present")
    If TemResponce = vbNo Then Exit Sub

    
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "select * from tblpatientfacility where patientfacility_ID =" & txtBookingID.Text
        .Open
        If .RecordCount = 0 Then .Close: Exit Sub
        If !paidtostaff = True Then
            TemResponce = MsgBox("This patient fee is already paid to the doctor. You can't make present or absent after paying the doctor", vbInformation, "Present")
            .Close
            Exit Sub
        End If
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
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

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

If ListDatesAndSecessions.ListIndex < 0 Or IsNumeric(ListSecessionIDs.Text) = False Or IsDate(ListDates.Text) = False Then
    TemResponce = MsgBox("You have not selected a Date and secession", vbCritical, "No Date & Secession")
    ListDatesAndSecessions.SetFocus
    Exit Sub
End If

    With DataEnvironment1.rssqlNurseView
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientFacility.*, tblInstitutions.InstitutionName, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionCode FROM tblInstitutions RIGHT JOIN (tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID  where staff_ID = " & Val(ListConsultantIDs.Text) & " and appointmentdate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " order by dayserial "
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

    TemResponce = MsgBox("Are You Sure You want to Refund Doctor Fee to Patient  ?", vbCritical + vbYesNo, "Refund")
    If TemResponce = vbNo Then Exit Sub

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
        !repaytime = Time
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
            !repaytime = Time
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
    
    TemResponce = MsgBox("Are You Sure You want to Reprint of this Patients  ?", vbCritical + vbYesNo, "Reprint")
    If TemResponce = vbNo Then Exit Sub
    
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
            TemResponce = MsgBox("The booking has repaied. You can't print a bill Copy", vbCritical, "Repaied")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        
        If !FullyPaid = 0 Then
            TemResponce = MsgBox("The patient has not completed the payment. You can't print a bill Copy", vbCritical, "Not Paid")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    
    End With
    
    Call SetBillPrinter1
    Call SetBillPaper1
    
    
End Sub





Private Sub bttnSecession_Click()
If IsDate(ListDates.Text) = False Then Exit Sub
With DataEnvironment1.rsSecessionView_Grouping
If .State = 1 Then .Close

    .Open " SHAPE {SELECT tblPatientMainDetails.Patient_ID, tblPatientFacility.Secession, tblPatientMainDetails.FirstName, tblFacilitySecession.SecessionName, tblFacilitySecession.StartingTime, tblPatientFacility.DaySerial, tblPatientFacility.PatientFacility_ID FROM (( tblPatientFacility LEFT OUTER JOIN tblFacilitySecession ON tblPatientFacility.Secession = tblFacilitySecession.FacilitySecession_ID) LEFT OUTER JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID) WHERE ((tblPatientFacility.Staff_ID =  " & Val(ListConsultantIDs.Text) & ") and (tblPatientFacility.AppointmentDate = '" & ListDates.Text & "')) Order by tblFacilitySecession.StartingTime, tblPatientFacility.DaySerial }  AS SecessionView COMPUTE SecessionView, ANY(SecessionView.'SecessionName') AS SecessionNameValue BY 'StartingTime'"


'    .Open " SHAPE {SELECT tblPatientMainDetails.Patient_ID, tblPatientFacility.Secession, tblPatientMainDetails.FirstName, tblFacilitySecession.SecessionName, tblFacilitySecession.StartingTime, tblPatientFacility.DaySerial, tblPatientFacility.PatientFacility_ID FROM (( tblPatientFacility LEFT OUTER JOIN tblFacilitySecession ON tblPatientFacility.Secession = tblFacilitySecession.FacilitySecession_ID) LEFT OUTER JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID) WHERE ((tblPatientFacility.Staff_ID =  " & Val(ListConsultantIDs.Text) & ") and (tblPatientFacility.AppointmentDate = '" & ListDates.Text & "') and (tblPatientFacility.FullyPaid = 1 )) Order by tblFacilitySecession.StartingTime, tblPatientFacility.DaySerial }  AS SecessionView COMPUTE SecessionView, ANY(SecessionView.'SecessionName') AS SecessionNameValue BY 'StartingTime'"
   
   
   ' .Open " SHAPE {SELECT tblPatientMainDetails.Patient_ID, tblPatientFacility.Secession, tblPatientMainDetails.FirstName, tblFacilitySecession.SecessionName, tblFacilitySecession.StartingTime, tblPatientFacility.DaySerial FROM ((tblPatientFacility LEFT OUTER JOIN tblFacilitySecession ON tblPatientFacility.Secession = tblFacilitySecession.FacilitySecession_ID) LEFT OUTER JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID) WHERE ((tblPatientFacility.Staff_ID = " & Val(ListConsultantIDs.Text) & ") and (tblPatientFacility.AppointmentDate = '" & ListDates.Text & "'))}  AS SecessionView COMPUTE SecessionView, ANY(SecessionView.'SecessionName') AS SecessionNameValue BY 'StartingTime'"
'    .Open " SHAPE {SELECT tblPatientMainDetails.Patient_ID, tblPatientFacility.Secession, tblPatientMainDetails.FirstName, tblFacilitySecession.SecessionName, tblFacilitySecession.StartingTime, tblPatientFacility.DaySerial FROM ((tblPatientFacility LEFT OUTER JOIN tblFacilitySecession ON tblPatientFacility.Secession = tblFacilitySecession.FacilitySecession_ID) LEFT OUTER JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID) WHERE ((tblPatientFacility.Staff_ID = " & Val(ListConsultantIDs.Text) & ") and (tblPatientFacility.AppointmentDate = '" & ListDates.Text & "'))}  AS SecessionView COMPUTE SecessionView, ANY(SecessionView.'Secession') AS SecessionNameValue BY 'StartingTime'"
     
    dtrSecessionView.Sections.Item("ReportHeader").Controls.Item("lblName").Caption = InstitutionName
    dtrSecessionView.Sections.Item("ReportHeader").Controls.Item("lblAddress").Caption = InstitutionAddress
    dtrSecessionView.Sections.Item("ReportHeader").Controls.Item("lblReport").Caption = "Patients for All Seccession"
    
    dtrSecessionView.Sections.Item("PageHeader").Controls.Item("lblDoctor").Caption = "Consultant : " & FindLDoctorFromID(Val(ListConsultantIDs.Text))
    dtrSecessionView.Sections.Item("PageHeader").Controls.Item("lblDate").Caption = Format(Date, "dd /mmmm /yyyy")
    dtrSecessionView.Sections.Item("PageFooter").Controls.Item("lblad").Caption = LongAd
    dtrSecessionView.Show
End With

End Sub

Private Sub chkForigner_Click()
        If chkForigner.Value = 0 And chkScan.Value = 1 Then
            lblAgentAmount.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
        ElseIf chkForigner.Value = 0 And chkScan.Value = 0 Then
            lblAgentAmount.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
        
        ElseIf chkForigner.Value = 1 And chkScan.Value = 1 Then
            lblAgentAmount.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
        ElseIf chkForigner.Value = 1 And chkScan.Value = 0 Then
            lblAgentAmount.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
        
        
        ElseIf SSTab1.Tab = 1 And chkScan.Value = 1 Then
            lblAgentAmount.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
        
        ElseIf SSTab1.Tab = 1 And chkScan.Value = 0 Then
            lblAgentAmount.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
        
        
        Else
            If chkScan.Value = 1 Then
                lblAgentAmount.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            Else
                lblAgentAmount.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            End If
        End If
        
        
        Select Case SSTab1.Tab
        
        Case 0
        txtCashPatientName.SetFocus
        Case 1
        txtAgentPatientName.SetFocus
        Case 2
        txtCreditPatientName.SetFocus
        End Select
End Sub

Private Sub chkScan_Click()
        If chkForigner.Value = 0 And chkScan.Value = 1 Then
            lblAgentAmount.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
        
        ElseIf chkForigner.Value = 0 And chkScan.Value = 0 Then
            lblAgentAmount.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
        
        ElseIf chkForigner.Value = 1 And chkScan.Value = 1 Then
            lblAgentAmount.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
        
        ElseIf chkForigner.Value = 1 And chkScan.Value = 0 Then
            lblAgentAmount.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
        
        
        ElseIf SSTab1.Tab = 1 And chkScan.Value = 1 Then
            lblAgentAmount.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
        
        ElseIf SSTab1.Tab = 1 And chkScan.Value = 0 Then
            lblAgentAmount.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
        
        
        Else
            If chkScan.Value = 1 Then
                lblAgentAmount.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            Else
                lblAgentAmount.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            End If
        End If
        
        
        Select Case SSTab1.Tab
        
        Case 0
        txtCashPatientName.SetFocus
        Case 1
        txtAgentPatientName.SetFocus
        Case 2
        txtCreditPatientName.SetFocus
        End Select

End Sub

Private Sub chkThroughAgent_Click()
    If chkThroughAgent.Value = 1 Then
        cmbTStaff.Enabled = True
        cmbTStaffCode.Enabled = True
        cmbTStaff.Visible = True
        cmbTStaffCode.Visible = True
        
    Else
        cmbTStaff.Visible = False
        cmbTStaffCode.Visible = False
        
        cmbTStaff.Text = Empty
        cmbTStaff.Enabled = False
        cmbTStaffCode.Text = Empty
        cmbTStaffCode.Enabled = False
    End If
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

Private Sub Command1_Click()
    Dim rsTem1 As New ADODB.Recordset
    
    With rsTem1
        If .State = 1 Then .Close
        'temSql = "SELECT tblpatientfacility.* from tblpatientfacility where hospitalfacility_ID = 10 and AppointmentDate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " order by dayserial"
        temSQL = "SELECT PatientFacility_ID from tblpatientfacility"
        .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
                
        ' tem
        Dim temFields As String
        
        temFields = ""
        
        temFields = temFields & "Agent_ID ,AgentRefNo "
        temFields = temFields & " AppointmentDate ,appointmenttime ,billprinted ,BookingDate ,bookingtime "
        'temFields = temFields & " cancelled ,creditstaff_ID ,DaySerial ,FacilityCatogery "
        'temFields = temFields & "FullyPaid ,fullypaidnull ,HospitalFacility_ID ,institutiondue ,institutionfee ,InstitutionFeeToPay ,IsScan ,otherdue ,otherfee ,otherfeetopay ,"
        'temFields = temFields & "PatientBill_ID ,patientid ,paymentmethod_ID ,PaymentMode ,personaldue ,personalfee ,PersonalFeeToPay ,resultsuccess ,Secession ,Staff_ID ,TotalDue ,totalfee ,totalfeetopay ,user_ID "


        If .State = 1 Then .Close
        'temSql = "SELECT tblpatientfacility.* from tblpatientfacility where hospitalfacility_ID = 10 and AppointmentDate = '" & ListDates.Text & "' and secession = " & Val(ListSecessionIDs.Text) & " order by dayserial"
        temSQL = "SELECT  " & temFields & " from tblpatientfacility"
        
        temSQL = "Select * from qryPatientFacility"
        temSQL = "sp2"
        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
        
    End With
    Dim Cmd As ADODB.Command
    
    Set Cmd = New ADODB.Command
     Cmd.ActiveConnection = cnnChannelling
     Cmd.CommandType = adCmdStoredProc
     Cmd.CommandText = "sp2"

'    Cmd.Parameters.Append Cmd.CreateParameter_
'    ("empid", adVarChar, adParamInput, 6, str_empid)

    Set rsTem1 = Cmd.Execute
    
End Sub

Private Sub DataComboAgent_Change()
On Error GoTo ErrorHandler
    txtAgentName.Text = DataComboAgent.Text
    Dim TemResponce  As Integer
    If Not IsNumeric(DataComboAgent.BoundText) Then Exit Sub
    
    DataComboAgentCode.BoundText = DataComboAgent.BoundText
    
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblinstitutions.* from tblinstitutions where Institution_ID = " & DataComboAgent.BoundText
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!InstitutionCredit) Then
            TemAgentCredit = !InstitutionCredit
        Else
            TemAgentCredit = Empty
        End If
        txtAgentBalance.Caption = Format(TemAgentCredit, "0.00")
        If Not IsNull(!InstitutionMaxCredit) Then
            TemAgentMaxCredit = !InstitutionMaxCredit
        Else
            TemAgentMaxCredit = 0
        End If
                
                
        If (0 - TemAgentMaxCredit) > TemAgentCredit Then
            TemResponce = MsgBox("This agent has already exceeded the credit limit, Increase the credit limit or ask the agent to pay some credit", vbInformation, "Exceed Credit Limit")
            DataComboAgent.Text = Empty
            If DataComboAgent.Visible = True Then
                DataComboAgent.SetFocus
            Else
                DataComboAgentCode.SetFocus
            End If
        End If
        If !InstitutionBlackListed = True Then
            TemResponce = MsgBox("This agent is black listed, Select another agent or discuss with the management to remove from the Black List", vbInformation, "Black Listed Patient")
            DataComboAgent.Text = Empty
            If DataComboAgent.Visible = True Then
                DataComboAgent.SetFocus
            Else
                DataComboAgentCode.SetFocus
            End If
        End If
        .Close
        
        If AgentBillNumber = True Then
            If .State = 1 Then .Close
            .Source = "SELECT tblAgentRef.AgentRefNo FROM tblAgentRef WHERE (((tblAgentRef.Agent_ID)=" & DataComboAgent.BoundText & ") AND ((tblAgentRef.Booked)=False)) ORDER BY tblAgentRef.AgentRefNo"
            .Open
            If .RecordCount = 0 Then
                TemResponce = MsgBox("There are no Bills issued for this agent. You can ask them to get a Bill Book or ask the owner to change the preferances to allow inserting any bill number", vbCritical, "No Bill Numbers")
                DataComboAgent.SetFocus
                Exit Sub
            End If
            .MoveFirst
            txtAgentRef.Text = !AgentRefNo
            If .State = 1 Then .Close
        End If
        
    End With
Exit Sub

ErrorHandler:
Exit Sub

End Sub

Private Sub DataComboAgent_Click(Area As Integer)
'    DataComboAgent_Change
End Sub

Private Sub DataComboAgentCode_Change()
    On Error GoTo ErrorHandler
    If IsNumeric(DataComboAgentCode.BoundText) = False Then Exit Sub
    DataComboAgent.BoundText = DataComboAgentCode.BoundText
    Exit Sub
    
ErrorHandler:
    Exit Sub

End Sub

Private Sub FormatPatientSearchGrid()

With gridPatient
    .Clear
    
    .Rows = 1
    .Cols = 6
    
    .ColWidth(0) = 320
    .ColWidth(2) = 2000
    .ColWidth(3) = 1400
    .ColWidth(4) = 1200
    .ColWidth(5) = 2000

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

Private Sub DataComboAgentCode_Click(Area As Integer)
'    DataComboAgent_Change
'    On Error GoTo ErrorHandler
'    If IsNumeric(DataComboAgentCode.BoundText) = False Then Exit Sub
'    DataComboAgent.BoundText = DataComboAgentCode.BoundText
'    Exit Sub
'
'ErrorHandler:
'    Exit Sub
End Sub

Private Sub DataComboAgentCode_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHandler
If IsNumeric(DataComboAgentCode.BoundText) = False Then Exit Sub
DataComboAgent.BoundText = DataComboAgentCode.BoundText
If KeyAscii = 13 Then txtAgentRef.SetFocus
Exit Sub

ErrorHandler:
Exit Sub

End Sub

Private Sub cmbTStaff_Change()
'    txtCreditPatientName.Text = cmbTStaff.Text
    cmbTStaffCode.BoundText = cmbTStaff.BoundText
End Sub


Private Sub cmbTStaffCode_Change()
    On Error Resume Next
    cmbTStaff.BoundText = cmbTStaffCode.BoundText
End Sub

Private Sub cmbTStaff_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCreditPatientName.SetFocus
        SendKeys "{home}+{end}"
    End If
End Sub

Private Sub cmbTStaffCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(cmbTStaffCode.BoundText) = True Then
            txtCreditPatientName.SetFocus
            SendKeys "{home}+{end}"
        Else
            cmbTStaff.SetFocus
        End If
    End If
End Sub

Private Sub DTPickerFindPatientDate_Change()
    Call FormatPatientSearchGrid
    Call FillPatientName
End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.WindowState = 2
SSTab1.Tab = 0
    
    If SetPrinter = False Then
        Unload Me
        Exit Sub
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


Private Sub FillAgentCombos()
    If AgentCashOnly = True Then
        With DataEnvironment1
            DataComboAgentCode.RowMember = Empty
            DataComboAgentCode.ListField = Empty
            DataComboAgentCode.BoundColumn = Empty
            If .rssqlTemAgents2.State = 1 Then .rssqlTemAgents2.Close
            .Commands!SqlTemAgentS2.CommandText = "SELECT tblInstitutions.institutioncode , tblInstitutions.institution_ID From tblInstitutions ORDER BY tblInstitutions.InstitutionCode"
            .SqlTemAgentS2
            Set DataComboAgentCode.RowSource = DataEnvironment1
            DataComboAgentCode.RowMember = "sqlTemAgents2"
            DataComboAgentCode.ListField = "InstitutionCode"
            DataComboAgentCode.BoundColumn = "Institution_ID"
        End With
        With DataEnvironment1
            DataComboAgent.RowMember = Empty
            DataComboAgent.ListField = Empty
            DataComboAgent.BoundColumn = Empty
            If .rssqlTemAgents1.State = 1 Then .rssqlTemAgents1.Close
            .Commands!sqlTemAgents1.CommandText = "SELECT tblInstitutions.institutionname , tblinstitutions.institution_ID From tblInstitutions ORDER BY tblInstitutions.institutionname"
            .sqlTemAgents1
            Set DataComboAgent.RowSource = DataEnvironment1
            DataComboAgent.RowMember = "sqlTemAgents1"
            DataComboAgent.ListField = "InstitutionName"
            DataComboAgent.BoundColumn = "Institution_ID"
            
            
            
            cmbTStaff.RowMember = Empty
            cmbTStaff.ListField = Empty
            cmbTStaff.BoundColumn = Empty
            If .rssqlTemStaff1.State = 1 Then .rssqlTemStaff1.Close
            .Commands!sqlTemStaff1.CommandText = "SELECT tblstaff.stafflistedname , tblstaff.Staff_ID From tblStaff ORDER BY tblStaff.StaffListedName "
            .sqlTemStaff1
            Set cmbTStaff.RowSource = DataEnvironment1
            cmbTStaff.RowMember = "sqlTemStaff1"
            cmbTStaff.ListField = "stafflistedname"
            cmbTStaff.BoundColumn = "Staff_ID"
            
            
            cmbTStaffCode.RowMember = Empty
            cmbTStaffCode.ListField = Empty
            cmbTStaffCode.BoundColumn = Empty
            If .rssqlTemStaff2.State = 1 Then .rssqlTemStaff2.Close
            .Commands!sqlTemStaff2.CommandText = "SELECT tblstaff.staffCode , tblstaff.Staff_ID From tblStaff ORDER BY tblStaff.StaffCode "
            .sqlTemStaff2
            Set cmbTStaffCode.RowSource = DataEnvironment1
            cmbTStaffCode.RowMember = "sqlTemStaff2"
            cmbTStaffCode.ListField = "staffCode"
            cmbTStaffCode.BoundColumn = "Staff_ID"
            
            
        End With
    Else
        With DataEnvironment1
            DataComboAgentCode.RowMember = Empty
            DataComboAgentCode.ListField = Empty
            DataComboAgentCode.BoundColumn = Empty
            If .rssqlTemAgents2.State = 1 Then .rssqlTemAgents2.Close
            .Commands!SqlTemAgentS2.CommandText = "SELECT tblInstitutions.institutioncode , tblInstitutions.institution_ID From tblInstitutions where cashagent = 1 ORDER BY tblInstitutions.InstitutionCode"
            .SqlTemAgentS2
            Set DataComboAgentCode.RowSource = DataEnvironment1
            DataComboAgentCode.RowMember = "sqlTemAgents2"
            DataComboAgentCode.ListField = "InstitutionCode"
            DataComboAgentCode.BoundColumn = "Institution_ID"
        End With
        With DataEnvironment1
            DataComboAgent.RowMember = Empty
            DataComboAgent.ListField = Empty
            DataComboAgent.BoundColumn = Empty
            If .rssqlTemAgents1.State = 1 Then .rssqlTemAgents1.Close
            .Commands!sqlTemAgents1.CommandText = "SELECT tblInstitutions.institutionname , tblinstitutions.institution_ID From tblInstitutions where cashagent = 1 ORDER BY tblInstitutions.institutionname"
            .sqlTemAgents1
            Set DataComboAgent.RowSource = DataEnvironment1
            DataComboAgent.RowMember = "sqlTemAgents1"
            DataComboAgent.ListField = "InstitutionName"
            DataComboAgent.BoundColumn = "Institution_ID"
        End With
        With DataEnvironment1
            cmbTStaff.RowMember = Empty
            cmbTStaff.ListField = Empty
            cmbTStaff.BoundColumn = Empty
            If .rssqlTemStaff1.State = 1 Then .rssqlTemStaff1.Close
            .Commands!sqlTemStaff1.CommandText = "SELECT tblstaff.stafflistedname , tblstaff.Staff_ID From tblStaff ORDER BY tblStaff.StaffListedName "
            .sqlTemStaff1
            Set cmbTStaff.RowSource = DataEnvironment1
            cmbTStaff.RowMember = "sqlTemStaff1"
            cmbTStaff.ListField = "stafflistedname"
            cmbTStaff.BoundColumn = "Staff_ID"
        
        
        
        
                    cmbTStaffCode.RowMember = Empty
            cmbTStaffCode.ListField = Empty
            cmbTStaffCode.BoundColumn = Empty
            If .rssqlTemStaff2.State = 1 Then .rssqlTemStaff2.Close
            .Commands!sqlTemStaff2.CommandText = "SELECT tblstaff.staffCode , tblstaff.Staff_ID From tblStaff ORDER BY tblStaff.StaffCode "
            .sqlTemStaff2
            Set cmbTStaffCode.RowSource = DataEnvironment1
            cmbTStaffCode.RowMember = "sqlTemStaff2"
            cmbTStaffCode.ListField = "staffCode"
            cmbTStaffCode.BoundColumn = "Staff_ID"

        
        
        End With
    End If
End Sub

Private Sub Form_Load()
    
    txtMaxCounter.Text = Val(GetSetting(App.EXEName, Me.Name, txtMaxCounter.Name, AdvanceBookingDays))
    
    
    Call FormatGridSpeciality
    Call FormatGridConsultants
    Call FormatGridDates
    Call FormatGridPatients
    Call FillSpeciality
    Call Setcolours
    Call FillAgentCombos
    
    Dim ingRet As Long
    
    Dim TabDates(1) As Long
    Dim TabDatesSecessions(4) As Long
    Dim TabPatientFacilities(6) As Long
    
    'No, Pt, FullyPaid, Remarks,
    TabDates(0) = 48
    TabDates(1) = 166
    
    If CanSelectAgent = False Then
        DataComboAgent.Visible = False
    Else
        DataComboAgent.Visible = True
    End If
    
    TabDatesSecessions(0) = 9 * 4
    TabDatesSecessions(1) = 18 * 4
    TabDatesSecessions(2) = 23 * 4
    TabDatesSecessions(3) = 28 * 4
    TabDatesSecessions(4) = 33 * 4
    
    TabPatientFacilities(0) = 3 * 4
    TabPatientFacilities(1) = 15 * 4
    TabPatientFacilities(2) = 20 * 4
    TabPatientFacilities(3) = 28 * 4
    TabPatientFacilities(4) = 29 * 4
'    TabPatientFacilities(5) = 33 * 4
'    TabPatientFacilities(6) = 41 * 4
    
    ingRet = SendMessage(ListDates.hwnd, LB_SETTABSTOPS, 2, TabDates(0))
    ingRet = SendMessage(ListPatientFacilities.hwnd, LB_SETTABSTOPS, 7, TabPatientFacilities(0))
    ingRet = SendMessage(ListDatesAndSecessions.hwnd, LB_SETTABSTOPS, 5, TabDatesSecessions(0))
    
    DTPickerFindPatientDate.Value = Date

    If AllowAbsent = False Then
        bttnMarkAbsent.Visible = False
        bttnMarkPresent.Visible = False
    Else
        bttnMarkAbsent.Visible = True
        bttnMarkPresent.Visible = True
    End If
    
    If DisplayPrintChkBox = True Then
        chkPrint.Value = 1
        chkPrint.Visible = True
    Else
        chkPrint.Value = 1
        chkPrint.Visible = False
    End If
    
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
     
    If PaymentCash = 1 Then
        SSTab1.TabEnabled(0) = True
    End If
    
    If PaymentCredit = 1 Then
        SSTab1.TabEnabled(2) = True
    End If
    
    If PaymentAgent = 1 Then
        SSTab1.TabEnabled(1) = True
    End If
    
    If AgentBookingValidation = True Then
        bttnAgentBookingValidation.Visible = True
    Else
        bttnAgentBookingValidation.Visible = False
    End If
    
    Call FillPatientName
    
    
    If UserAuthority <> AuthorityAdministrator Or UserAuthority <> AuthorityOwner Then
        Label20.Visible = False
        txtSearchBookingID.Visible = False
        bttnSearch.Visible = False
        Label46.Visible = False
        txtBookingID.Visible = False
        
    End If
    
    chkThroughAgent.Value = GetSetting(App.EXEName, Me.Name, chkThroughAgent.Name, 1)
    
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

Private Sub FormatGridSpeciality()
    ListSpecialities.Clear
    ListSpecialityIDs.Clear
End Sub

Private Sub FormatGridConsultants()
    ListConsultants.Clear
    ListConsultantIDs.Clear
End Sub

Private Sub FormatGridDates()
    ListDates.Clear
    ListDatesAndSecessions.Clear
    ListSecessionIDs.Clear
    ListSecessionMax.Clear
    ListSecessionStartingTime.Clear
    ListRoomNo.Clear
End Sub

Private Sub FormatGridPatients()
    ListPatientFacilities.Clear
    ListPatientFacilityIDs.Clear
    
    FrameCancellations.Enabled = False
    FrameRefunds.Enabled = False
    FrameReprints.Enabled = False
    FrameSettleCredit.Enabled = False
    
    FramePatient.Enabled = True
    
    
End Sub


Private Sub ClearPatientDetails()
    
    txtAgentBalance.Caption = Empty
    txtBookedPatientID.Text = Empty
    txtBookedPatientName.Text = Empty
    txtBookedPatientContactNo.Text = Empty
    txtNameChange.Text = Empty
    txtPhoneChange.Text = Empty
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
    txtAppDate.Text = Empty
    txtAppTime.Text = Empty
    txtBookingDate.Text = Empty
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
    txtSearchAgentRefNo.Text = Empty
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




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, chkThroughAgent.Name, chkThroughAgent.Value
    SaveSetting App.EXEName, Me.Name, txtMaxCounter.Name, Val(txtMaxCounter.Text)
End Sub

Private Sub gridPatient_Click()
gridPatient.col = 0
    If IsNumeric(gridPatient.Text) = False Then Exit Sub
    txtSearchBookingID.Text = gridPatient.Text
    Call bttnSearch_Click
'gridPatient.Col = 0
'gridPatient.ColSel = gridPatient.Cols - 1
SSTab2.Tab = 0
txtSearchBookingID.Text = Empty
gridPatient.Clear
FormatPatientSearchGrid

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



Private Sub ListConsultants_Click()
    ClearPatientDetails
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    Call FormatGridDates
    Call FormatGridPatients
    TemPatientFacilityID = 0
    
    TemDoctorFee = 0
    TemFDoctorFee = 0
    TemADoctorFee = 0
    
    TemInstitutionFee = 0
    TemFInstitutionFee = 0
    TemADoctorFee = 0
    
    TemOtherFee = 0
    TemFOtherFee = 0
    TemAOtherFee = 0
    

    TemSDoctorFee = 0
    TemSFDoctorFee = 0
    TemSADoctorFee = 0
    
    TemSInstitutionFee = 0
    TemSFInstitutionFee = 0
    TemSADoctorFee = 0
    
    TemSOtherFee = 0
    TemSFOtherFee = 0
    TemSAOtherFee = 0
    
    
'    TemDoctorID = 0
    TemAppointmentDate = Empty
    TemAppointmentTime = Empty
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
'    TemDoctorID = Val(ListConsultantIDs.Text)
    Call FillDates
End Sub

Private Sub FillDates()
        
    ListDatesAndSecessions.Visible = False:     Me.MousePointer = vbHourglass:
        
    Call FormatGridDates
    
    Dim TemCounter As Long
    Dim TemBookingDate As Date
    Dim TemDateCounter As Long
    Dim NowROw As Long
    
    With DataEnvironment1.rssqlTem5
        If .State = 1 Then .Close
        .Source = "SELECT tblfacilitysecession.* from tblfacilitysecession where hospitalfacility_ID = 10 and staff_ID = " & Val(ListConsultantIDs.Text)
        If .State = 0 Then .Open
        If .RecordCount = 0 Then .Close: ListDatesAndSecessions.Visible = True:     Me.MousePointer = vbDefault: Exit Sub
        .Close
    End With
    
    TemCounter = 0
    TemDateCounter = 0
    NowROw = 0
    TemPreviousDate = Date - 2
    
    Dim temMaxCounter As Long
    temMaxCounter = Val(txtMaxCounter.Text)
    If temMaxCounter = 0 Then temMaxCounter = AdvanceBookingDays
    While TemCounter < temMaxCounter
        TemBookingDate = DateAdd("d", TemDateCounter, Date - temMaxCounter)
        TemDateCounter = TemDateCounter + 1
        
        With DataEnvironment1.rssqlTem4
            If .State = 1 Then .Close
            .Source = "Select * from tblfacilitysecession where hospitalfacility_ID =  10  and staff_ID = " & Val(ListConsultantIDs.Text) & " and AlteredDate = '" & TemBookingDate & "' order by StartingTime"
            .Open
            If .RecordCount <> 0 Then
                If !fulldayleave = False Then
                    TemCounter = TemCounter + 1
                    While .EOF = False
                        If TemPreviousDate = TemBookingDate Then
                            TemTextForList = Space(8)
                        Else
                            TemTextForList = Format(TemBookingDate, DefaultShortDate)
                        End If
                        
                        TemTextForList = TemTextForList & Space(2) & Left(!SecessionName, 4)
                        
                        
                        
'                        If !Maximum <> 0 Then
'                            TemTextForList = TemTextForList & vbTab & !Maximum
'                        Else
'                            TemTextForList = TemTextForList & vbTab & "**"
'                        End If
                        
                        
                        
                        
                        
                        'TemTextForList = TemTextForList & vbTab & Format(!startingtime, "hh:mm AMPM")
                        
                        TemTextForList = TemTextForList & Space(2) & Format(!startingtime, "hh:mm AMPM")
                        
                        TemTextForList = TemTextForList & Space(2) & GetBookedNumber(TemBookingDate, !facilitysecession_ID)
                        ListDates.AddItem TemBookingDate
                        ListSecessionIDs.AddItem !facilitysecession_ID
                        ListDatesAndSecessions.AddItem TemTextForList
                        ListSecessionMax.AddItem !Maximum
                        ListSecessionStartingTime.AddItem !startingtime
                        ListRoomNo.AddItem !roomno
                        TemPreviousDate = TemBookingDate
                        .MoveNext
                    Wend
                End If
                .Close
            Else
                If .State = 1 Then .Close
                .Source = "Select * from tblfacilitysecession where hospitalfacility_ID = 10 and staff_ID = " & Val(ListConsultantIDs.Text) & " and SecessionWeekday = " & Weekday(TemBookingDate) & " order by StartingTime"
                .Open
                If .RecordCount <> 0 Then
                    TemCounter = TemCounter + 1
                    While .EOF = False
                        If TemPreviousDate = TemBookingDate Then
                            TemTextForList = Space(8)
                        Else
                            TemTextForList = Format(TemBookingDate, DefaultShortDate)
                        End If
                        
                        TemTextForList = TemTextForList & Space(2) & Left(!SecessionName, 4)
                        
'                        If !Maximum <> 0 Then
'                            TemTextForList = TemTextForList & vbTab & !Maximum
'                        Else
'                            TemTextForList = TemTextForList & vbTab & "**"
'                        End If
                        
                        'TemTextForList = TemTextForList & vbTab & Format(!startingtime, "hh:mm AMPM")
                        
                        TemTextForList = TemTextForList & Space(2) & Format(!startingtime, "hh:mm AMPM")
                        
                        TemTextForList = TemTextForList & Space(2) & GetBookedNumber(TemBookingDate, !facilitysecession_ID)
                        
                        ListDates.AddItem TemBookingDate
                        ListSecessionIDs.AddItem !facilitysecession_ID
                        ListDatesAndSecessions.AddItem TemTextForList
                        ListSecessionMax.AddItem !Maximum
                        ListSecessionStartingTime.AddItem !startingtime
                        If Not IsNull(!roomno) Then
                            ListRoomNo.AddItem !roomno
                        Else
                            ListRoomNo.AddItem ""
                        End If
                        TemPreviousDate = TemBookingDate
                        .MoveNext
                    Wend
                End If
            End If
        End With
    
    Wend
    
    ListDatesAndSecessions.Visible = True
    Me.MousePointer = vbDefault
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
    If ListConsultants.ListIndex < 0 And ListConsultants.ListCount > 0 Then ListConsultants.ListIndex = 0
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
ElseIf KeyCode = vbKeyUp Then ' Or vbKeyDown Then
    If ListConsultants.ListIndex > 0 Then ListConsultants.ListIndex = ListConsultants.ListIndex - 1
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
'    Call ClearPatientDetails
'    Call FormatGridDates
'    Call FormatGridPatients
'    TemDoctorFee = 0
'    TemFDoctorFee = 0
'    TemInstitutionFee = 0
'    TemFInstitutionFee = 0
'    TemOtherFee = 0
''    TemDoctorID = 0
'    TemAppointmentDate = Empty
'    TemAppointmentTime = Empty
''    TwoSecessions = True
'    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
''    TemDoctorID = Val(ListConsultantIDs.Text)
'    Call FillDates
'    ListDatesAndSecessions.SetFocus
    KeyCode = Empty
ElseIf KeyCode = vbKeyDown Then
    If ListConsultants.ListIndex < ListConsultants.ListCount - 1 Then ListConsultants.ListIndex = ListConsultants.ListIndex + 1
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
'    Call ClearPatientDetails
'    Call FormatGridDates
'    Call FormatGridPatients
'    TemDoctorFee = 0
'    TemFDoctorFee = 0
'    TemInstitutionFee = 0
'    TemFInstitutionFee = 0
'    TemOtherFee = 0
''    TemDoctorID = 0
'    TemAppointmentDate = Empty
'    TemAppointmentTime = Empty
''    TwoSecessions = True
'    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
''    TemDoctorID = Val(ListConsultantIDs.Text)
'    Call FillDates
'    ListDatesAndSecessions.SetFocus
    KeyCode = Empty
End If

End Sub


Private Sub ListConsultants_LostFocus()
    BoxConsultant.BackColor = FrameBackColour ' vbRed
End Sub

Private Sub ListDatesAndSecessions_Click()
    ListDates.ListIndex = ListDatesAndSecessions.ListIndex
    ListSecessionIDs.ListIndex = ListDatesAndSecessions.ListIndex
    ListSecessionMax.ListIndex = ListDatesAndSecessions.ListIndex
    ListSecessionStartingTime.ListIndex = ListDatesAndSecessions.ListIndex
    ListRoomNo.ListIndex = ListDatesAndSecessions.ListIndex
    TemAppointmentDate = ListDates.Text
    
    MonthView1.Value = ListDates.Text
    
    Call ClearPatientDetails

    Call FormatGridPatients
    
    If Not IsDate(ListDates.Text) Then Exit Sub
    If Not IsNumeric(ListSecessionIDs.Text) Then Exit Sub
    If Not IsDate(ListSecessionStartingTime.Text) Then Exit Sub
    
'    TemSecession = Val(ListSecessionIDs.Text)
    SecessionMax = Val(ListSecessionMax.Text)
    TemSecessionStartingTime = Val(ListSecessionStartingTime.Text)
    
    Call FindSecessionDetails
    Call FillGridPatients
        
    DTPickerFindPatientDate.Value = ListDates.Text

    If TemSADoctorFee + TemSAInstitutionFee + TemSAOtherFee + TemSDoctorFee + TemSFDoctorFee + TemSFInstitutionFee + TemSFOtherFee + TemSInstitutionFee + TemSOtherFee > 0 Then
        chkScan.Value = 0
        chkScan.Visible = True
    Else
        chkScan.Value = 0
        chkScan.Visible = False
    End If

End Sub

Private Sub FindSecessionDetails()
With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    .Source = "Select * from tblfacilitysecession where FacilitySecession_ID = " & ListSecessionIDs.Text
    .Open
    If .RecordCount = 0 Then Exit Sub
    
        TemDoctorFee = Val(Format(!LocaldoctorFee, "0.00"))
        TemFDoctorFee = Val(Format(!Foreigndoctorfee, "0.00"))
        TemADoctorFee = Val(Format(!agentDoctorFee, "0.00"))
        
        TemInstitutionFee = Val(Format(!LocalHospitalFee, "0.00"))
        TemFInstitutionFee = Val(Format(!ForeignHospitalFee, "0.00"))
        TemAInstitutionFee = Val(Format(!AgentHospitalFee, "0.00"))
        
        TemOtherFee = Val(Format(!LocalOtherFee, "0.00"))
        TemFOtherFee = Val(Format(!ForeignOtherFee, "0.00"))
        TemAOtherFee = Val(Format(!AgentOtherFee, "0.00"))
        
        TemSDoctorFee = Val(Format(!SLocaldoctorFee, "0.00"))
        TemSFDoctorFee = Val(Format(!SForeigndoctorfee, "0.00"))
        TemSADoctorFee = Val(Format(!SAgentDoctorFee, "0.00"))
        
        TemSInstitutionFee = Val(Format(!SLocalHospitalFee, "0.00"))
        TemSFInstitutionFee = Val(Format(!SForeignHospitalFee, "0.00"))
        TemSAInstitutionFee = Val(Format(!SAgentHospitalFee, "0.00"))
        
        TemSOtherFee = Val(Format(!SLocalOtherFee, "0.00"))
        TemSFOtherFee = Val(Format(!SForeignOtherFee, "0.00"))
        TemSAOtherFee = Val(Format(!SAgentOtherFee, "0.00"))
        
        
        TemSecessionStartingTime = !startingtime
        TemUsualDuration = !usualduration
        TemCanByPassOrder = !CanByPassOrder
        TemCalculateAppointment = !calculateappointment
        SecessionMax = !Maximum
        
        If chkScan.Value = 0 Then
            If chkForigner.Value = 0 Then
                lblAgentAmount.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            ElseIf chkForigner.Value = 1 Then
                lblAgentAmount.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
            ElseIf SSTab1.Tab = 1 Then
                lblAgentAmount.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
            Else
                lblAgentAmount.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            End If
        Else
            If chkForigner.Value = 0 Then
                lblAgentAmount.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            ElseIf chkForigner.Value = 1 Then
                lblAgentAmount.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
            ElseIf SSTab1.Tab = 1 Then
                lblAgentAmount.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
            Else
                lblAgentAmount.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
                lblCashDue.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
                lblCredit.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            End If
        End If
    .Close
    End With

End Sub

Private Sub FindAppointmentTime()
'    If TemUsualDuration = 0 Then Exit Sub
    If TemSecessionStartingTime = TimeSerial(0, 0, 0) Then Exit Sub
    TemAppointmentTime = TimeSerial(Hour(TemSecessionStartingTime), Minute(TemSecessionStartingTime) + (TemUsualDuration * TemNonCancelledVisits), 0)
End Sub

Private Sub FillGridPatients()
    Dim TemTextForList As String
    Call ClearPatientDetails

    Call FormatGridPatients
        With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = "SELECT * from tblpatientfacility where hospitalfacility_ID = 10 and Staff_ID = " & Val(ListConsultantIDs.Text) & " and AppointmentDate = '" & ListDates.Text & "' and Secession = " & Val(ListSecessionIDs.Text) & " order by DaySerial"
        .Open
        If .RecordCount = 0 Then Exit Sub
        While Not .EOF
            TemTextForList = !DaySerial & Space(2) & Left(FindPatientByID(!patientid) & Space(13), 13)
            
            
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
            
            If !Cancelled = True Then
                TemTextForList = TemTextForList & Space(2) & "Cancel"
            ElseIf !Refund = True Then
                TemTextForList = TemTextForList & Space(2) & "Refund"
            Else
                TemTextForList = TemTextForList & Space(2) & Space(6)
            End If
            
            
            If !patientabsent = True Then
                TemTextForList = TemTextForList & vbTab & "Ab"
            Else
                TemTextForList = TemTextForList & vbTab & " "
            End If
            ListPatientFacilities.AddItem TemTextForList
            ListPatientFacilityIDs.AddItem !patientfacility_ID
            .MoveNext
        Wend
    End With

End Sub

Private Sub ListDatesAndSecessions_GotFocus()
    BoxDates.BackColor = BttnBackColour ' vbRed
End Sub

Private Sub ListDatesAndSecessions_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If ListDatesAndSecessions.ListIndex < 0 And ListDatesAndSecessions.ListCount > 1 Then ListDatesAndSecessions.ListIndex = 0
        Select Case SSTab1.Tab
            Case 0:     txtCashPatientName.SetFocus
            Case 1:     DataComboAgentCode.SetFocus
            Case 2:
                        If cmbTStaffCode.Enabled = True Then
                            cmbTStaffCode.SetFocus
                        Else
                            txtCreditPatientName.SetFocus
                        End If
                        
            Case Else:  txtPatientName.SetFocus
        End Select
    KeyCode = Empty
ElseIf KeyCode = vbKeyRight Then
    If ListDatesAndSecessions.ListIndex < 0 And ListDatesAndSecessions.ListCount > 1 Then ListDatesAndSecessions.ListIndex = 0
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
    TemContactNo = FindPatientContactNoByID(!patientid)
    TemPatientFacilityID = !patientfacility_ID
    TemAppointmentDate = Format(!AppointmentDate, DefaultLongDate)
    TemAppointmentTime = !appointmenttime
    TemDaySerial = !DaySerial
    
    TemDoctorFee = !PersonalFee
    TemInstitutionFee = !InstitutionFee
    TemOtherFee = !otherfee
    
    txtBookedPatientName.Text = TemPatient
    txtBookedPatientContactNo.Text = TemContactNo
    txtNameChange.Text = TemPatient
    txtPhoneChange.Text = TemContactNo
    txtBookedPatientID.Text = TemPatientID
    txtBookingID.Text = TemPatientFacilityID
    txtPaymentMethod.Text = !PaymentMode
    txtBookingUser.Text = FindStaffFromID(!user_ID)
    txtConsultant.Text = ListConsultants.Text
    txtAppDate.Text = Format(TemAppointmentDate, DefaultLongDate)
    txtAppTime.Text = TemAppointmentTime
    txtBookingDate.Text = Format(!BookingDate, DefaultLongDate) & " " & Format(!BookingTime, "hh:mm AMPM")
    
    If Not IsNull(!PersonalFee) Then
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
        lblOtherFeePaid.Caption = Format(!otherfee, "0.00")
        lblOtherFeePaidR.Caption = Format(!otherfee, "0.00")
        lblOtherFeePaidC.Caption = Format(!otherfee, "0.00")
    Else
        lblOtherFeePaid.Caption = "0.00"
        lblOtherFeePaidR.Caption = "0.00"
        lblOtherFeePaidC.Caption = Format(0, "0.00")
    End If
        
        
    If Not IsNull(!totalfee) Then
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
        
        If IsNull(!CreditStaff_ID) = False Then
            If !CreditStaff_ID <> 0 Then
                txtCreditSettle.Text = txtCreditSettle.Text + "Staff Booking for " & FindStaffFromID(!CreditStaff_ID) & ". "
            End If
            If IsNull(!CreditSettleUser_ID) Or !CreditSettleUser_ID = 0 Then
                txtCreditSettle.Text = txtCreditSettle.Text + "The patient has to pay Rs." & Format(!totalfeetopay, "0.00")
                FrameCancellations.Enabled = False
                FrameRefunds.Enabled = False
                FrameSettleCredit.Enabled = True
            Else
                txtCreditSettle.Text = txtCreditSettle.Text + "The patient had settled it by paying Rs." & Format(!totalfee, "0.00") & " to " & FindStaffFromID(!CreditSettleUser_ID)
                FrameCancellations.Enabled = True
                FrameRefunds.Enabled = True
                FrameSettleCredit.Enabled = False
            End If
        Else
            If IsNull(!CreditSettleUser_ID) Or !CreditSettleUser_ID = 0 Then
                txtCreditSettle.Text = "Telephone Booking. The patient has to pay Rs." & Format(!totalfeetopay, "0.00")
                FrameCancellations.Enabled = False
                FrameRefunds.Enabled = False
                FrameSettleCredit.Enabled = True
            Else
                txtCreditSettle.Text = "Telephone Booking. The patient had settled it by paying Rs." & Format(!totalfee, "0.00") & " to " & FindStaffFromID(!CreditSettleUser_ID)
                FrameCancellations.Enabled = True
                FrameRefunds.Enabled = True
                FrameSettleCredit.Enabled = False
            End If
            
        End If
        
    ElseIf !PaymentMode = "Agent" Then
        txtAgentAndCode.Text = FindAgentFromID(!Agent_ID)
        If !FullyPaid = True Then
        
            txtCreditSettle.Text = "Agent Booking"
            If AgentBookingValidation = True Then
                txtTemAgentID.Text = !Agent_ID
                txtCreditSettle.Text = txtCreditSettle.Text & ". Confirmed"
            Else
                txtTemAgentID = Empty
            End If
        Else
            txtCreditSettle.Text = "Agent Booking. Yet to confirm"
        End If
        FrameCancellations.Enabled = True
        FrameRefunds.Enabled = True
        FrameSettleCredit.Enabled = False
        txtAgentCode.Text = FindAgentCodeFromID(!Agent_ID)
        If Not IsNull(!AgentRefNo) Then
            txtAgentRefNo.Text = !AgentRefNo
        Else
            txtAgentRefNo.Text = Empty
        End If
    Else
        txtCreditSettle.Text = "No credit issues"
        FrameCancellations.Enabled = True
        FrameRefunds.Enabled = True
        FrameSettleCredit.Enabled = False
    End If
    
    If !Cancelled = True Then
        txtCancelRefund.Text = "Cancelled on " & Format(!repaydate, DefaultLongDate) & " by " & FindStaffFromID(!repayUser_ID) & ". Rs. " & Format(!totalrefund, "0.00") & " was repaied."
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
        txtCancelRefund.Text = "Refunded on " & Format(!repaydate, DefaultLongDate) & " by " & FindStaffFromID(!repayUser_ID) & ". Rs. " & Format(!totalrefund, "0.00") & " was repaied."
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



Private Sub ListPatientFacilities_GotFocus()
    BoxPatients.BackColor = BttnBackColour ' vbRed
End Sub


Private Sub ListPatientFacilities_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    Select Case SSTab1.Tab
        Case 0: txtCashPatientName.SetFocus
        Case 1: DataComboAgentCode.SetFocus
        Case 2: chkThroughAgent.SetFocus
    End Select
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
    If ListSpecialities.ListIndex < 0 And ListSpecialities.ListCount > 0 Then ListSpecialities.ListIndex = 0
    ListConsultants.SetFocus
    KeyCode = Empty
Else
End If
End Sub

Private Sub ListSpecialities_LostFocus()
     BoxSpeciality.BackColor = FrameBackColour ' - 2147483633
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
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

Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Tab = 1 And AgentBillNumber = True Then
    txtAgentRef.Locked = True
Else
    txtAgentRef.Locked = False
End If

If SSTab1.Tab = 2 Then
    chkPrint.Value = 0
Else
    chkPrint.Value = 1
End If

    If chkScan.Value = 0 Then
        If chkForigner.Value = 0 Then
            lblAgentAmount.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
        ElseIf chkForigner.Value = 1 Then
            lblAgentAmount.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "0.00")
        ElseIf SSTab1.Tab = 1 Then
            lblAgentAmount.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemADoctorFee + TemAInstitutionFee), "0.00")
        Else
            lblAgentAmount.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemDoctorFee + TemInstitutionFee), "0.00")
        End If
    Else
        If chkForigner.Value = 0 Then
            lblAgentAmount.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
        ElseIf chkForigner.Value = 1 Then
            lblAgentAmount.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemSFDoctorFee + TemSFInstitutionFee), "0.00")
        ElseIf SSTab1.Tab = 1 Then
            lblAgentAmount.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemSADoctorFee + TemSAInstitutionFee), "0.00")
        Else
            lblAgentAmount.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            lblCashDue.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
            lblCredit.Caption = Format((TemSDoctorFee + TemSInstitutionFee), "0.00")
        End If
    End If
End Sub




Private Sub txtAgentBalance_Click()
    DataComboAgent_Change
End Sub

Private Sub txtAgentName_Click()
    DataComboAgent_Change
End Sub

Private Sub txtAgentPatientName_Change()
    txtPatientName.Text = txtAgentPatientName.Text
End Sub

Private Sub txtAgentPatientName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then bttnAddPatient_Click
End Sub

Private Sub txtAgentRef_Click()
    DataComboAgent_Change
End Sub

Private Sub txtAgentRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAgentPatientName.SetFocus
    
End Sub

Private Sub txtCashPatientName_Change()
    txtPatientName.Text = txtCashPatientName.Text
End Sub

Private Sub txtCashContactNo_Change()
    txtContactNo.Text = txtCashContactNo.Text
End Sub

Private Sub txtCashPatientName_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then bttnAddPatient_Click
    If KeyAscii = 13 Then txtCashContactNo.SetFocus
    
End Sub

Private Sub txtCashContactNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then bttnAddPatient_Click
End Sub

Private Sub txtCreditPatientName_Change()
    txtPatientName.Text = txtCreditPatientName.Text
End Sub

Private Sub txtCreditContactNo_Change()
    txtContactNo.Text = txtCreditContactNo.Text
End Sub

Private Sub txtCreditPatientName_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then bttnAddPatient_Click
    If KeyAscii = 13 Then
        txtCreditContactNo.SetFocus
    End If
End Sub

Private Sub txtCreditContactNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then bttnAddPatient_Click
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

Private Sub txtPhoneChange_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then bttnChangeName_Click


End Sub

Private Sub txtOtherRepayC_Change()
    txtRepayTotalC.Text = Format((Val(txtStaffRepayC.Text) + Val(txtInstitutionRepayC.Text) + Val(txtOtherRepayC.Text)), "0.00")
End Sub

Private Sub txtOtherRepayR_Change()
    txtRepayTotalR.Text = Format((Val(txtStaffRepayR.Text) + Val(txtInstitutionRepayR.Text) + Val(txtOtherRepayR.Text)), "0.00")
End Sub


Private Sub txtPatientName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bttnAddPatient_Click
End Sub


Private Sub bttnSearch_Click()
    Call SearchBookingID
    SSTab2.Tab = 0
    txtSearchBookingID.Text = Empty
    'ComboPatientName.Text = Empty
    gridPatient.Clear
End Sub


Private Sub txtSearchAgentRefNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then bttnAgentRefNoSearch_Click
End Sub

Private Sub txtSearchBookingID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SearchBookingID
End Sub


Private Sub bttnAgentRefNoSearch_Click()
    Call SearchAgentReferranceNo
    SSTab2.Tab = 0
    txtSearchBookingID.Text = Empty
    'ComboPatientName.Text = Empty
    gridPatient.Clear
End Sub


Private Sub SearchAgentReferranceNo()

Dim TemResponce As Integer

With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    .Source = "Select * from tblpatientfacility where AgentRefNo = '" & txtSearchAgentRefNo.Text & "'"
    .Open
    If .RecordCount = 0 Then
        TemResponce = MsgBox("There is no such referrance number. Please re-check", vbCritical, "No such No")
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
    DateFound = False
    For TemNum = 0 To ListDates.ListCount - 1
        ListDates.ListIndex = TemNum
        If ListDates.Text = !AppointmentDate Then
            ListSecessionIDs.ListIndex = TemNum
            If ListSecessionIDs.Text = !Secession Then
                ListDatesAndSecessions.ListIndex = TemNum
                ListDatesAndSecessions_Click
                TemNum = ListDates.ListCount
                DateFound = True
            End If
        End If
    Next
    If DateFound = False Then
        TemResponce = MsgBox("The booking date is in the past, You can search patients with the appointment dates today onwards. If you want to locate the patient goto Locate Patients screen", vbCritical, "Past Appointment")
        Exit Sub
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
    
'    ListSpecialities.ListIndex = 0
'    ListSpecialityIDs.ListIndex = 0
'
'    ListSpecialities_Click
'
    Call ListAllConsultants
    
    Dim TemNum As Long
    
    If ListConsultants.ListCount = 0 Then
        TemResponce = MsgBox("The consultant is deleted", vbCritical, "Consultant Deleted")
        Exit Sub
    End If
        
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
    DateFound = False
    For TemNum = 0 To ListDates.ListCount - 1
        ListDates.ListIndex = TemNum
        If ListDates.Text = !AppointmentDate Then
            ListSecessionIDs.ListIndex = TemNum
            If ListSecessionIDs.Text = !Secession Then
                ListDatesAndSecessions.ListIndex = TemNum
                ListDatesAndSecessions_Click
                TemNum = ListDates.ListCount
                DateFound = True
            End If
        End If
    Next
    
    If DateFound = False Then
        TemResponce = MsgBox("The booking date is in the past, You can search patients with the appointment dates today onwards. If you want to locate the patient goto Locate Patients screen", vbCritical, "Past Appointment")
        Exit Sub
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































Private Sub Setcolours()
    Select Case ColourScheme
    Case 1:
        BttnBackColour = 5341695
        BttnForeColour = 1314458
        FrmBackColour = 11066623
        FrmForeColour = 1314458
        FrameBackColour = 11066623
        FrameForeColour = 1314458
        TxtBackColour = 9881851
        TxtForeColour = 1314458
        LblBackColour = 11066623
        LblForeColour = 1314458
        GridBackColor = 9881855
        GridBackColorBkg = 10474239
        GridBackColorFixed = 8566015
        GridBackColorSel = 5341695
        GridForeColor = 1314458
        GridForeColorFixed = 11944
        GridForeColorSel = 3014824
    Case 2:
        BttnBackColour = 14803300
        BttnForeColour = 5539362
        FrmBackColour = 16766120
        FrmForeColour = 5539362
        FrameBackColour = 16766120
        FrameForeColour = 5539362
        TxtBackColour = 16760450
        TxtForeColour = 5539362
        LblBackColour = 16766120
        LblForeColour = 5539362
        GridBackColor = 16760450
        GridBackColorBkg = 16771260
        GridBackColorFixed = 16105620
        GridBackColorSel = 16737380
        GridForeColor = 5539362
        GridForeColorFixed = 5539362
        GridForeColorSel = 16765588
    Case 3:
        BttnBackColour = 51455
        BttnForeColour = 942490
        FrmBackColour = 11070719
        FrmForeColour = 942490
        FrameBackColour = 11070719
        FrameForeColour = 942490
        TxtBackColour = 11528439
        TxtForeColour = 1314458
        LblBackColour = 11070719
        LblForeColour = 942490
        GridBackColor = 16760450
        GridBackColorBkg = 16771260
        GridBackColorFixed = 16105620
        GridBackColorSel = 16737380
        GridForeColor = 5539362
        GridForeColorFixed = 5539362
        GridForeColorSel = 16765588
    End Select
    bttnAddPatient.BackColor = BttnBackColour
    bttnAddPatient.ForeColor = BttnForeColour
    bttnReprint.BackColor = BttnBackColour
    bttnReprint.ForeColor = BttnForeColour
    bttnCancellation.BackColor = BttnBackColour
    bttnCancellation.ForeColor = BttnForeColour
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
    bttnRefund.BackColor = BttnBackColour
    bttnRefund.ForeColor = BttnForeColour
    bttnRefund.BackColor = BttnBackColour
    bttnRefund.ForeColor = BttnForeColour
    bttnAllPatients.BackColor = BttnBackColour
    bttnAllPatients.ForeColor = BttnForeColour
    bttnCashSettle.BackColor = BttnBackColour
    bttnCashSettle.ForeColor = BttnForeColour
    bttnMarkAbsent.BackColor = BttnBackColour
    bttnMarkAbsent.ForeColor = BttnForeColour
    bttnMarkPresent.BackColor = BttnBackColour
    bttnMarkPresent.ForeColor = BttnForeColour
    bttnChangeName.BackColor = BttnBackColour
    bttnChangeName.ForeColor = BttnForeColour
    bttnAllSecessionPatients.BackColor = BttnBackColour
    bttnAllSecessionPatients.ForeColor = BttnForeColour
    bttnNurseView.BackColor = BttnBackColour
    bttnNurseView.ForeColor = BttnForeColour
    bttnDoctorView.BackColor = BttnBackColour
    bttnDoctorView.ForeColor = BttnForeColour
    bttnSearch.BackColor = BttnBackColour
    bttnSearch.ForeColor = BttnForeColour
    bttnAllDoctors.BackColor = BttnBackColour
    bttnAllDoctors.ForeColor = BttnForeColour
    Frame6.BackColor = FrameBackColour
    Frame6.ForeColor = FrameForeColour
    Frame7.BackColor = FrameBackColour
    Frame7.ForeColor = FrameForeColour
    FrameCash.BackColor = FrmBackColour
    FrameCash.ForeColor = FrmForeColour
    FrameAgent.BackColor = FrameBackColour
    FrameAgent.ForeColor = FrameForeColour
    Frame1.BackColor = FrameBackColour
    Frame1.ForeColor = FrameForeColour
    Frame5.BackColor = FrameBackColour
    Frame5.ForeColor = FrameForeColour
    bttnAgentRefNoSearch.BackColor = BttnBackColour
    bttnAgentRefNoSearch.ForeColor = BttnForeColour
    FrameSettleCredit.BackColor = FrameBackColour
    FrameSettleCredit.ForeColor = FrameForeColour
    FramePatientDetails.BackColor = FrameBackColour
    FramePatientDetails.ForeColor = FrameForeColour
    FrameReprints.BackColor = FrameBackColour
    FrameReprints.ForeColor = FrameForeColour
    FrameCancellations.BackColor = FrameBackColour
    FrameCancellations.ForeColor = FrameForeColour
    FrameRefunds.BackColor = FrameBackColour
    FrameRefunds.ForeColor = FrameForeColour
    Frame1.BackColor = FrameBackColour
    Frame1.ForeColor = FrameForeColour
    chkThroughAgent.BackColor = LblBackColour
    chkThroughAgent.ForeColor = LblForeColour
    Frame3.BackColor = FrameBackColour
    Frame3.ForeColor = FrameForeColour
    OptionRefundPrint.BackColor = FrameBackColour
    OptionRefundPrint.ForeColor = FrameForeColour
    OptionRefundDoNotPrint.BackColor = FrameBackColour
    OptionRefundDoNotPrint.ForeColor = FrameForeColour
    OptionRepayAgent.BackColor = FrameBackColour
    OptionRepayAgent.ForeColor = FrameForeColour
    OptionRepayPatient.BackColor = FrameBackColour
    OptionRepayPatient.ForeColor = FrameForeColour
    Frame4.BackColor = FrameBackColour
    Frame4.ForeColor = FrameForeColour
    OptionSettleCreditPrint.BackColor = FrameBackColour
    OptionSettleCreditPrint.ForeColor = FrameForeColour
    OptionSettleCreditDoNotPrint.BackColor = FrameBackColour
    OptionSettleCreditDoNotPrint.ForeColor = FrameForeColour
    Frame2.BackColor = FrameBackColour
    Frame2.ForeColor = FrameForeColour
    Frame8.BackColor = FrameBackColour
    Frame8.ForeColor = FrameForeColour
    Frame9.BackColor = FrameBackColour
    Frame9.ForeColor = FrameForeColour
    OptionPrintCancel.BackColor = FrameBackColour
    OptionPrintCancel.ForeColor = FrameForeColour
    OptionDoNotPrintCancel.BackColor = FrameBackColour
    OptionDoNotPrintCancel.ForeColor = FrameForeColour
    Label1.BackColor = LblBackColour
    Label1.ForeColor = LblForeColour
    Label10.BackColor = LblBackColour
    Label10.ForeColor = LblForeColour
    Label11.BackColor = LblBackColour
    Label11.ForeColor = LblForeColour
    Label12.BackColor = LblBackColour
    Label12.ForeColor = LblForeColour
    Label13.BackColor = LblBackColour
    Label13.ForeColor = LblForeColour
    Label14.BackColor = LblBackColour
    Label14.ForeColor = LblForeColour
    Label15.BackColor = LblBackColour
    Label15.ForeColor = LblForeColour
    Label16.BackColor = LblBackColour
    Label16.ForeColor = LblForeColour
    Label2.BackColor = LblBackColour
    Label2.ForeColor = LblForeColour
    Label18.BackColor = LblBackColour
    Label18.ForeColor = LblForeColour
    Label3.BackColor = LblBackColour
    Label3.ForeColor = LblForeColour
    Label4.BackColor = LblBackColour
    Label4.ForeColor = LblForeColour
    Label17.BackColor = LblBackColour
    Label17.ForeColor = LblForeColour
    Label4.BackColor = LblBackColour
    Label4.ForeColor = LblForeColour
    Label5.BackColor = LblBackColour
    Label5.ForeColor = LblForeColour
    Label6.BackColor = LblBackColour
    Label6.ForeColor = LblForeColour
    Label7.BackColor = LblBackColour
    Label7.ForeColor = LblForeColour
    Label8.BackColor = LblBackColour
    Label8.ForeColor = LblForeColour
    Label9.BackColor = LblBackColour
    Label9.ForeColor = LblForeColour
    Label1.BackColor = LblBackColour
    Label1.ForeColor = LblForeColour
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
        Printer.Print Tab(TemTab3); Format(ListDates.Text, DefaultLongDate)
        
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
        Printer.Print Tab(TemTab2); Format(Date, DefaultShortDate)
                
        .EndDoc
    End With
End Sub

Private Sub BillPrint2()


    Dim TemRows As Long

With Printer

        Printer.Font = "Arial Black"
'        Printer.Print
        
        Printer.FontSize = 11
        Printer.Print Tab(2); InstitutionName;
        Printer.Print Tab(54); InstitutionName
        
'        Printer.FontSize = 9
'        Printer.Print Tab(3); InstitutionAddress;
'        Printer.Print Tab(64); InstitutionAddress
        
'        Printer.Print Tab(3); InstitutionTelephone;
'        Printer.Print Tab(64); InstitutionTelephone
        
        Printer.FontName = "Courier"
        Printer.FontSize = 8
'        Printer.Print
        
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
        
        Displace = 88
        
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
'        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); TemPatient;
        'd
        Printer.Print Tab(TemTab7); "Patient";
'        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); TemPatient
        
'        Printer.Print
        Printer.Print Tab(TemTab1); "Consultant"; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text)));
        'd
        Printer.Print Tab(TemTab7); "Consultant";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); UCase(FindLDoctorFromID(Val(ListConsultantIDs.Text)))
'        Printer.Print
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
        
'        Printer.Print
        
        If SSTab1.Tab = 0 Then
        
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
        
            Printer.Print Tab(TemTab1); "Payment";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3); "Cash";
            'd

            Printer.Print Tab(TemTab7); "Payment";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9); "Cash"
        
        
        ElseIf SSTab1.Tab = 1 Then
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
        
            Printer.Print Tab(TemTab1); "Payment";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3); "Agent";
            'd

            Printer.Print Tab(TemTab7); "Payment";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9); "Agent"
        
        ElseIf SSTab1.Tab = 2 Then
        
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
        
            Printer.Print Tab(TemTab1); "Payment";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3); "Credit";
            'd

            Printer.Print Tab(TemTab7); "Payment";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9); "Credit"
        
        End If
        
        Printer.Print
        Printer.Print
        
        Printer.Print Tab(TemTab2); "--------------------";
        Printer.Print Tab(TemTab8); "--------------------"
        
        Printer.Print Tab(TemTab2); UserName;
        Printer.Print Tab(TemTab8); UserName
        
        Printer.Print Tab(TemTab2); Time;
        Printer.Print Tab(TemTab8); Time
        
        Printer.Print Tab(TemTab2); Format(Date, DefaultShortDate);
        Printer.Print Tab(TemTab8); Format(Date, DefaultShortDate)
        
        Printer.EndDoc
    End With

End Sub



Private Sub BillPrint21()


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
        
        If txtPaymentMethod.Text = "Cash" Then
        
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
        
        ElseIf txtPaymentMethod.Text = "Credit" Then
        
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
        
        Printer.Print Tab(TemTab2); Time;
        Printer.Print Tab(TemTab8); Time
        
        Printer.Print Tab(TemTab2); Format(Date, DefaultShortDate);
        Printer.Print Tab(TemTab8); Format(Date, DefaultShortDate)
        
        Printer.EndDoc
    End With

End Sub




