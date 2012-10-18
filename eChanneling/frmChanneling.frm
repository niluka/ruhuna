VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChannelingMS 
   Caption         =   "Channeling"
   ClientHeight    =   8610
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
   ScaleHeight     =   8610
   ScaleWidth      =   15240
   Begin VB.ListBox ListPatientFacilities 
      Height          =   4380
      ItemData        =   "frmChanneling.frx":0000
      Left            =   10320
      List            =   "frmChanneling.frx":0002
      TabIndex        =   3
      Top             =   360
      Width           =   4695
   End
   Begin VB.ListBox ListDatesAndSecessions 
      Height          =   4380
      ItemData        =   "frmChanneling.frx":0004
      Left            =   6000
      List            =   "frmChanneling.frx":0006
      TabIndex        =   2
      Top             =   360
      Width           =   3975
   End
   Begin VB.ListBox ListConsultants 
      Height          =   4380
      ItemData        =   "frmChanneling.frx":0008
      Left            =   3120
      List            =   "frmChanneling.frx":000A
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox ListSpecialities 
      Height          =   4380
      ItemData        =   "frmChanneling.frx":000C
      Left            =   240
      List            =   "frmChanneling.frx":000E
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox ListSecessionStartingTime 
      Height          =   4620
      Left            =   13920
      TabIndex        =   95
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListPatientFacilityIDs 
      Height          =   4620
      Left            =   14280
      TabIndex        =   80
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox ListConsultantIDs 
      Height          =   4620
      Left            =   14280
      TabIndex        =   79
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox ListSecessionIDs 
      Height          =   4620
      Left            =   13920
      TabIndex        =   78
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListDates 
      Height          =   4620
      Left            =   13920
      TabIndex        =   77
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox ListSpecialityIDs 
      Height          =   4620
      Left            =   14040
      TabIndex        =   76
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   240
      Left            =   4200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5280
      Value           =   1  'Checked
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   3495
      Left            =   7320
      TabIndex        =   7
      Top             =   5040
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Cancellations"
      TabPicture(0)   =   "frmChanneling.frx":0010
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameCancellations"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Refunds"
      TabPicture(1)   =   "frmChanneling.frx":002C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameRefunds"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Reprints"
      TabPicture(2)   =   "frmChanneling.frx":0048
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameReprints"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Calancer"
      TabPicture(3)   =   "frmChanneling.frx":0064
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MonthView1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Settle Credit"
      TabPicture(4)   =   "frmChanneling.frx":0080
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameSettleCredit"
      Tab(4).ControlCount=   1
      Begin VB.Frame FrameSettleCredit 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   83
         Top             =   360
         Width           =   7455
         Begin btButtonEx.ButtonEx ButtonEx1 
            Height          =   375
            Left            =   5160
            TabIndex        =   84
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
            TabIndex        =   93
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
            TabIndex        =   92
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
            TabIndex        =   91
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label41 
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
            TabIndex        =   89
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label40 
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
            TabIndex        =   88
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label39 
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
            TabIndex        =   87
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label38 
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   90
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame FrameReprints 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   7455
         Begin btButtonEx.ButtonEx bttnReprint 
            Height          =   375
            Left            =   5160
            TabIndex        =   29
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
            TabIndex        =   75
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
            TabIndex        =   74
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
            TabIndex        =   73
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label33 
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
            TabIndex        =   72
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label32 
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
            TabIndex        =   71
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label28 
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
            TabIndex        =   70
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label27 
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
            TabIndex        =   69
            Top             =   480
            Width           =   1335
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
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   66
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
            TabIndex        =   65
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame FrameRefunds 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   7455
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   2520
            MultiLine       =   -1  'True
            TabIndex        =   61
            Top             =   2520
            Width           =   2895
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4080
            TabIndex        =   59
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4080
            TabIndex        =   58
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4080
            TabIndex        =   57
            Top             =   480
            Width           =   1335
         End
         Begin btButtonEx.ButtonEx bttnRefund 
            Height          =   375
            Left            =   5520
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
            TabIndex        =   64
            Top             =   2400
            Width           =   1935
         End
         Begin VB.Label Label20 
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
            TabIndex        =   63
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label19 
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
            TabIndex        =   10
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label18 
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
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label17 
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
            TabIndex        =   12
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label16 
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
            TabIndex        =   13
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label15 
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
            Left            =   4200
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label14 
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
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label13 
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
            TabIndex        =   16
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label12 
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
            TabIndex        =   17
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label9 
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
            TabIndex        =   62
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame FrameCancellations 
         Height          =   3015
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   7455
         Begin VB.TextBox txtRepayComments 
            Height          =   375
            Left            =   2520
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   2520
            Width           =   2895
         End
         Begin VB.TextBox txtRepayTotal 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtOtherRepay 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4080
            TabIndex        =   45
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtInstitutionRepay 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4080
            TabIndex        =   44
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtStaffRepay 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4080
            TabIndex        =   43
            Top             =   480
            Width           =   1335
         End
         Begin btButtonEx.ButtonEx bttnCancelVisit 
            Height          =   375
            Left            =   5520
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            Caption         =   "Cancel Visit"
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
         Begin VB.Label Label11 
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
            TabIndex        =   18
            Top             =   2400
            Width           =   1935
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
            Left            =   360
            TabIndex        =   19
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label lblTotalPaid 
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
            TabIndex        =   56
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
            Left            =   2520
            TabIndex        =   55
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblInstitutionFeePaid 
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
            TabIndex        =   54
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lblStaffFeePaid 
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
            TabIndex        =   53
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label8 
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
            Left            =   4200
            TabIndex        =   52
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label7 
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
            TabIndex        =   51
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
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
            TabIndex        =   50
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label5 
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
            TabIndex        =   49
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label4 
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
            TabIndex        =   48
            Top             =   480
            Width           =   1935
         End
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2820
         Left            =   -70440
         TabIndex        =   40
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   49872897
         CurrentDate     =   39446
      End
   End
   Begin VB.Frame FramePatient 
      Caption         =   "Add Patient"
      Height          =   3495
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   7095
      Begin VB.CheckBox chkForigner 
         Caption         =   "Forigner"
         Height          =   255
         Left            =   5400
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtPatientName 
         Height          =   360
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin btButtonEx.ButtonEx bttnAddPatient 
         Height          =   375
         Left            =   5400
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3000
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
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   3836
         _Version        =   393216
         Tab             =   2
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Cash"
         TabPicture(0)   =   "frmChanneling.frx":009C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "FrameCash"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Agent"
         TabPicture(1)   =   "frmChanneling.frx":00B8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FrameAgent"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Credit"
         TabPicture(2)   =   "frmChanneling.frx":00D4
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame1"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame1 
            Caption         =   "Credit"
            Height          =   1695
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   6495
            Begin VB.CheckBox chkThroughAgent 
               Caption         =   "Through Agent"
               Height          =   240
               Left            =   120
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   240
               Value           =   1  'Checked
               Width           =   2895
            End
            Begin MSDataListLib.DataCombo DataComboThroughAgent 
               Bindings        =   "frmChanneling.frx":00F0
               Height          =   360
               Left            =   3120
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   240
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   635
               _Version        =   393216
               Style           =   2
               ListField       =   "InstitutionName"
               BoundColumn     =   "Institution_ID"
               Text            =   ""
               Object.DataMember      =   "sqlInstitutions"
            End
            Begin VB.Label lblCredit 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   3720
               TabIndex        =   42
               Top             =   1080
               Width           =   2655
            End
            Begin VB.Label Label2 
               Caption         =   "Amount  :                                    (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   41
               Top             =   1080
               Width           =   3495
            End
         End
         Begin VB.Frame FrameAgent 
            Caption         =   "Agent"
            Height          =   1695
            Left            =   -74880
            TabIndex        =   26
            Top             =   360
            Width           =   6495
            Begin MSDataListLib.DataCombo DataComboAgent 
               Bindings        =   "frmChanneling.frx":010F
               Height          =   360
               Left            =   3120
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   240
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   635
               _Version        =   393216
               Style           =   2
               ListField       =   "InstitutionName"
               BoundColumn     =   "Institution_ID"
               Text            =   ""
               Object.DataMember      =   "sqlInstitutions"
            End
            Begin VB.Label Label29 
               Caption         =   "&Agent              :"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   360
               Width           =   3135
            End
            Begin VB.Label Label30 
               Caption         =   "Agent &Balance : (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   38
               Top             =   1200
               Width           =   2775
            End
            Begin VB.Label txtAgentBalance 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   3120
               TabIndex        =   37
               Top             =   1200
               Width           =   3255
            End
            Begin VB.Label Label31 
               Caption         =   "A&mount           : (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   36
               Top             =   720
               Width           =   2775
            End
            Begin VB.Label lblAgentAmount 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   3120
               TabIndex        =   35
               Top             =   720
               Width           =   3255
            End
         End
         Begin VB.Frame FrameCash 
            Caption         =   "Cash"
            Height          =   1695
            Left            =   -74880
            TabIndex        =   24
            Top             =   360
            Width           =   6495
            Begin VB.Label Label35 
               Caption         =   "Amount  :                                    (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   34
               Top             =   720
               Width           =   3495
            End
            Begin VB.Label lblCashDue 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   3720
               TabIndex        =   33
               Top             =   720
               Width           =   2655
            End
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Patient Name"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   2415
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   13560
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   9600
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   4620
      Left            =   13920
      TabIndex        =   94
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape BoxPatients 
      BackStyle       =   1  'Opaque
      Height          =   4650
      Left            =   10200
      Top             =   240
      Width           =   4935
   End
   Begin VB.Shape BoxDates 
      BackStyle       =   1  'Opaque
      Height          =   4650
      Left            =   5880
      Top             =   240
      Width           =   4215
   End
   Begin VB.Shape BoxConsultant 
      BackStyle       =   1  'Opaque
      Height          =   4650
      Left            =   3000
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape BoxSpeciality 
      BackStyle       =   1  'Opaque
      Height          =   4650
      Left            =   120
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmChannelingMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemStaffFacilityID As Long
Dim TwoSecessions As Boolean
Dim TemDoctorFee As Double
Dim TemFDoctorFee As Double
Dim TemInstitutionFee As Double
Dim TemFInstitutionFee As Double
Dim TemOtherFee As Double
Dim SecessionMax As Long
Dim TemSecession  As Integer
Dim TemAgentCredit As Double
Dim TemPatientID As Long
Dim TemAgentMaxCredit As Double
Dim TemDoctorID As Long
Dim TemAppointmentDate As Date
Dim TemAppointmentTime As Date
Dim TemDaySerial As Long
Dim TemPatientFacilityID As Long
Dim TemAgentBookingID As Long
Dim TemSecessionStartingTime As Date
Dim TemUsualDuration As Long
Dim TemPatient As String
Dim TemConsultant As String
Dim TemNonCancelledVisits As Long
Dim TemBillID As Long
Dim GridSpecialityFilled As Boolean
Dim GridConsultantsFilled As Boolean
Dim GridDatesFilled As Boolean
Dim GridPatientsFilled As Boolean


Private Sub bttnAddPatient_Click()
    Dim TemResponce  As Integer
    

    If Not IsNumeric(ListConsultantIDs.Text) Then
        TemResponce = MsgBox("You have not selected a name of the doctor", vbCritical, "No doctor")
        ListConsultants.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(ListDates.Text) Then
        TemResponce = MsgBox("You have not selected a date", vbCritical, "No date")
        ListDatesAndSecessions.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPatientName.Text) = "" Then
        TemResponce = MsgBox("You have not entered a name of the patient to add", vbCritical, "No Name")
        txtPatientName.SetFocus
        Exit Sub
    Else
        TemPatient = txtPatientName.Text
    End If
    
    If CanSettlePayment = False Then Exit Sub
    
    Call AddPatient
    Call AddToBill
    
    If AddToPatientFacility = False Then Exit Sub
    
    If SSTab1.Tab = 1 Then
        UpdateAgentCredit
        UpdateAgentFacility
    ElseIf SSTab1.Tab = 2 Then
        UpdatePatientCredit
    End If
        
        DisplayDetails
'        BillPrint
        BillPrint2
        ClearForNewPatient
    
End Sub

Private Sub AddToBill()
With DataEnvironment1.rssqlTem5
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientbill"
    .Open
    .AddNew
    !patient_ID = TemPatientID
    !Date = Date
    !NetTotal = TemDoctorFee + TemInstitutionFee
    !GrossTotal = TemDoctorFee + TemInstitutionFee
    Select Case SSTab1.Tab
    Case 0:
        !paymentmethod = "Cash"
        !Cash = TemDoctorFee + TemInstitutionFee
    Case 1:
        !paymentmethod = "Agent"
        !AgentAmount = TemDoctorFee + TemInstitutionFee
        If IsNumeric(DataComboAgent.BoundText) = True Then !agent_ID = DataComboAgent.BoundText
    Case 2:
        If IsNumeric(DataComboThroughAgent.BoundText) Then !creditagent_ID = DataComboThroughAgent.BoundText
        !paymentmethod = "Credit"
        !credit = TemDoctorFee + TemInstitutionFee
    End Select
        !User_ID = UserID
        !BillSuccess = True
    If chkPrint.Value = 1 And SSTab1.Tab = 2 Then
        !billprinted = False
    ElseIf chkPrint.Value = 1 Then
        !billprinted = True
    Else
        !billprinted = False
    End If
    .Update
    TemBillID = !PatientBill_ID
    .Close
End With
End Sub

Private Sub UpdatePatientCredit()
With DataEnvironment1.rssqlTem7
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientmaindetails where patient_ID = " & TemPatientID
    .Open
    If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!credit) Then
            !credit = !credit - TemDoctorFee - TemInstitutionFee
        Else
            !credit = 0 - TemDoctorFee - TemInstitutionFee
        End If
    .Update
    .Close
End With
End Sub

Private Sub ClearForNewPatient()
    txtPatientName.Text = Empty
    chkForigner.Value = 0
    DataComboAgent.Text = Empty
    SSTab1.Tab = 0
End Sub

Private Sub BillPrint()
If chkPrint.Value <> 1 Then Exit Sub
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
        Printer.Print Tab(TemTab3); UCase(FindDoctorFromID(TemDoctorID))
        Printer.Print
        Printer.Print Tab(TemTab1); "Appo. Date ";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); Format(TemAppointmentDate, "dd mmmm yyyy")
        
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
        Printer.Print Tab(TemTab2); Format(Date, "dd mmmm yyyy")
                
        .EndDoc
    End With
End Sub

Private Sub BillPrint2()

If chkPrint.Value <> 1 Then Exit Sub

    Dim TemRows As Long

With Printer

        
        
        .Font = "Arial Black"
        Printer.Print
        
        .FontSize = 11
        Printer.Print Tab(2); InstitutionName;
        Printer.Print Tab(51); InstitutionName
        
        .FontSize = 9
        Printer.Print Tab(3); InstitutionAddress;
        Printer.Print Tab(64); InstitutionAddress
        
        Printer.Print Tab(3); InstitutionTelephone;
        Printer.Print Tab(64); InstitutionTelephone
'        Printer.Print
'        Printer.Print
'        Printer.Print
'        Printer.Print
'        Printer.Print
'        Printer.Print
'        Printer.Print
        
        .FontName = "Courier"
        .FontSize = 10
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
        Printer.Print Tab(TemTab3); UCase(FindDoctorFromID(TemDoctorID));
        'd
        Printer.Print Tab(TemTab7); "Consultant";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); UCase(FindDoctorFromID(TemDoctorID))
        Printer.Print
        Printer.Print Tab(TemTab1); "Appo. Date "; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); Format(TemAppointmentDate, "dd mmmm yyyy");
        'd
        Printer.Print Tab(TemTab7); "Appo. Date ";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); Format(TemAppointmentDate, "dd mmmm yyyy")
        
'        Printer.Print Tab(TemTab1); "Appo. Time"; ;
'        Printer.Print Tab(TemTab6); " : "; ;
'        Printer.Print Tab(TemTab3); TemAppointmentTime;
'        'd
'        Printer.Print Tab(TemTab7); "Appo. Time";
'        Printer.Print Tab(TemTab8); " : ";
'        Printer.Print Tab(TemTab9); TemAppointmentTime
        
        Printer.Print Tab(TemTab1); "Appo. No."; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); TemDaySerial;
        
        Printer.Print Tab(TemTab7); "Appo. No.";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); TemDaySerial
        
        
        Printer.Print Tab(TemTab1); "Appo. ID"; ;
        Printer.Print Tab(TemTab6); " : "; ;
        Printer.Print Tab(TemTab3); TemPatientFacilityID;
        'd
        
        Printer.Print Tab(TemTab7); "Appo. ID";
        Printer.Print Tab(TemTab8); " : ";
        Printer.Print Tab(TemTab9); TemPatientFacilityID
        
        Printer.Print
        
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
        
            Printer.Print Tab(TemTab1); "Payment Method";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3); "Cash";
            'd
            
            Printer.Print Tab(TemTab7); "Total Fee";
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
        
            Printer.Print Tab(TemTab1); "Payment Method";
            Printer.Print Tab(TemTab6); " : ";
            Printer.Print Tab(TemTab3); "Cash";
            'd
            
            Printer.Print Tab(TemTab7); "Total Fee";
            Printer.Print Tab(TemTab8); " : ";
            Printer.Print Tab(TemTab9); "Cash"
        
        ElseIf SSTab2.Tab = 2 Then
        
        End If
        
        Printer.Print
        Printer.Print
        
        Printer.Print Tab(TemTab2); "--------------------";
        Printer.Print Tab(TemTab8); "--------------------"
        
        Printer.Print Tab(TemTab2); UserName;
        Printer.Print Tab(TemTab8); UserName
        
        Printer.Print Tab(TemTab2); Format(Date, "dd mmmm yyyy");
        'D
        Printer.Print Tab(TemTab8); Format(Date, "dd mmmm yyyy")
        .EndDoc
    End With
End Sub
Private Function DisplayDetails() As Boolean
    DisplayDetails = True
    Dim TemResponce
    Dim TemText As String
    TemText = TemText & "Patient Name        :  " & TemPatient & vbNewLine
    TemText = TemText & "Patient ID          :  " & TemPatientID & vbNewLine
    TemText = TemText & "Appointment Time    :  " & TemAppointmentTime & vbNewLine & vbNewLine
    TemText = TemText & "Appointment No      :  " & TemDaySerial & vbNewLine & vbNewLine
    If SSTab1.Tab = 1 Then
        TemText = TemText & "Agent Referance No. :  " & TemAgentBookingID
    End If
    TemResponce = MsgBox(TemText, vbInformation, "Booking Details")
End Function

Private Sub UpdateAgentFacility()
    With DataEnvironment1.rssqlTem7
        If .State = 1 Then .Close
        .Source = "Select * from tblagentbooking"
        .Open
        .AddNew
        !agent_ID = DataComboAgent.BoundText
        !PatientFacility_ID = TemPatientFacilityID
        !BookingDate = Date
        !patient_ID = TemPatientID
        !AppointmentDate = TemAppointmentDate
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
        !InstitutionCredit = !InstitutionCredit - Val(TemDoctorFee + TemInstitutionFee)
        .Update
        .Close
    End With
End Sub


Private Sub AddPatient()
    With DataEnvironment1.rssqlTem3
        If .State = 1 Then .Close
        .Source = "select * from tblpatientmaindetails"
        .Open
        .AddNew
        !firstname = txtPatientName.Text
        .Update
        TemPatientID = !patient_ID
        .Close
    End With
End Sub


Private Function AddToPatientFacility() As Boolean

AddToPatientFacility = False

    With DataEnvironment1.rssqlTem5
        If .State = 1 Then .Close
        .Source = "SELECT tblpatientfacility.* from tblpatientfacility where AppointmentDate = #" & TemAppointmentDate & "# and FacilityStaff_ID = " & TemStaffFacilityID & " and secession = " & TemSecession & " and staff_ID = " & TemDoctorID & " and cancelled <> true order by dayserial"
        If .State = 0 Then .Open
        
        TemNonCancelledVisits = .RecordCount
        
        If .State = 1 Then .Close
        
        .Source = "SELECT tblpatientfacility.* from tblpatientfacility where AppointmentDate = #" & TemAppointmentDate & "# and FacilityStaff_ID = " & TemStaffFacilityID & " and secession = " & TemSecession & " and staff_ID = " & TemDoctorID & " order by dayserial"
        
        If .State = 0 Then .Open
        
        If .RecordCount = 0 Then
            TemDaySerial = 1
        Else
            .MoveLast
            TemDaySerial = 1 + !DaySerial
        End If
        
        If SecessionMax <> 0 Then
            If TemNonCancelledVisits > SecessionMax Then
                Dim TemResponce  As Integer
                TemResponce = MsgBox("Adding this patient will increase the maximum number for the consultant. Do you still want to add the patient?", vbYesNo, "Exceed Maximum")
                If TemResponce = vbNo Then Exit Function
            End If
        End If
        
        .AddNew
        !User_ID = UserID
        !patientid = TemPatientID
        !hospitalfacility_ID = 10
        !facilitystaff_ID = TemStaffFacilityID
        !FacilityCatogery = Doctor
        !PatientBill_ID = TemBillID
        !staff_ID = TemDoctorID
        !BookingDate = Date
        !AppointmentDate = TemAppointmentDate
        !secession = TemSecession
        !DaySerial = TemDaySerial
        !appointmenttime = TemAppointmentTime
        
        If SSTab1.Tab = 2 Then
            !fullypaid = False
        Else
            !fullypaid = True
        End If
        
        
        !cancelled = False
        !resultsuccess = True
        
        If SSTab1.Tab = 0 Then
            !Personalfee = TemDoctorFee
            !institutionfee = TemInstitutionFee
            !otherfee = 0
            !totalfee = TemDoctorFee + TemInstitutionFee
            !Personalfeetopay = 0
            !institutionfeetopay = 0
            !otherfeetopay = 0
            !totalfeetopay = 0
            !paymentmode = "Cash"
            !paymentmethod_ID = 1
        ElseIf SSTab1.Tab = 1 Then
            !Personalfee = TemDoctorFee
            !institutionfee = TemInstitutionFee
            !otherfee = 0
            !totalfee = TemDoctorFee + TemInstitutionFee
            !Personalfeetopay = 0
            !institutionfeetopay = 0
            !otherfeetopay = 0
            !totalfeetopay = 0
            !paymentmode = "Agent"
            !paymentmethod_ID = 2
            !agent_ID = Val(DataComboAgent.BoundText)
        ElseIf SSTab1.Tab = 2 Then
            !Personalfee = 0
            !institutionfee = 0
            !otherfee = 0
            !totalfee = 0
            !Personalfeetopay = TemDoctorFee
            !institutionfeetopay = TemInstitutionFee
            !otherfeetopay = 0
            !totalfeetopay = TemDoctorFee + TemInstitutionFee
            !paymentmode = "Credit"
            !paymentmethod_ID = 4
            If IsNumeric(DataComboThroughAgent.BoundText) Then !creditagent_ID = DataComboThroughAgent.BoundText
        End If
        
        If chkPrint.Value = 1 And SSTab1.Tab = 2 Then
            !billprinted = False
        ElseIf chkPrint.Value = 1 Then
            !billprinted = True
        Else
            !billprinted = False
        End If
        
        .Update
        TemPatientFacilityID = !PatientFacility_ID
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
    
    Case 1:
        If Not IsNumeric(DataComboAgent.BoundText) Then
            TemResponce = MsgBox("You have not selected an agent", vbInformation, "Agent")
            DataComboAgent.SetFocus
            Exit Function
        End If

        If TemAgentCredit - (TemDoctorFee + TemInstitutionFee) < (0 - TemAgentMaxCredit) Then
            TemResponce = MsgBox("This bill will lead to increase the credit limit of the agent. If you want to proceed, increase the credit limit or adviced the agent to settle cash", vbInformation, "Credit Limit")
            DataComboAgent.SetFocus
            Exit Function
        End If
    
        If chkForigner = 1 Then
            TemDoctorFee = TemFDoctorFee
            TemInstitutionFee = TemFInstitutionFee
        End If
    
    
    End Select
    CanSettlePayment = True
End Function



Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub chkForigner_Click()
    If TemStaffFacilityID <> 0 Then Call FindCharges
End Sub

Private Sub DataComboAgent_Click(Area As Integer)


    Dim TemResponce  As Integer
    If Not IsNumeric(DataComboAgent.BoundText) Then Exit Sub
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
        txtAgentBalance.Caption = Format(TemAgentCredit, "#0.00")
        If Not IsNull(!InstitutionMaxCredit) Then
            TemAgentMaxCredit = !InstitutionMaxCredit
        Else
            TemAgentMaxCredit = 0
        End If
                
                
        If (0 - TemAgentMaxCredit) > TemAgentCredit Then
            TemResponce = MsgBox("This agent has already exceeded the credit limit, Increase the credit limit or ask the agent to pay some credit", vbInformation, "Exceed Credit Limit")
            DataComboAgent.Text = Empty
            DataComboAgent.SetFocus
        End If
        If !InstitutionBlackListed = True Then
            TemResponce = MsgBox("This agent is black listed, Select another agent or discuss with the management to remove from the Black List", vbInformation, "Black Listed Patient")
            DataComboAgent.Text = Empty
            DataComboAgent.SetFocus
        End If
        .Close
    End With


End Sub

Private Sub Form_Activate()
On Error Resume Next
Me.WindowState = 2
SSTab1.Tab = 0

End Sub

Private Sub Form_Load()
    If SetPrinter = False Then
        Unload Me
        Exit Sub
    End If
    Call FormatGridSpeciality
    Call FormatGridConsultants
    Call FormatGridDates
    Call FormatGridPatients
    Call FillSpeciality

Dim ingRet As Long

Dim TabDates(1) As Long
Dim TabDatesSecessions(4) As Long
Dim TabPatientFacilities(3) As Long

'No, Pt, FullyPaid, Remarks,
TabDates(0) = 48
TabDates(1) = 166

TabDatesSecessions(0) = 20
TabDatesSecessions(1) = 30
TabDatesSecessions(2) = 40
TabDatesSecessions(3) = 50
TabDatesSecessions(4) = 60


TabPatientFacilities(0) = 20
TabPatientFacilities(1) = 40
TabPatientFacilities(2) = 60
TabPatientFacilities(3) = 80

ingRet = SendMessage(ListDates.hwnd, LB_SETTABSTOPS, 2, TabDates(0))
ingRet = SendMessage(ListPatientFacilities.hwnd, LB_SETTABSTOPS, 2, TabPatientFacilities(0))
ingRet = SendMessage(ListPatientFacilities.hwnd, LB_SETTABSTOPS, 2, TabDatesSecessions(0))





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
        frmPreferances.Show
        frmPreferances.ZOrder 0
        frmPreferances.SSTab1.Tab = 1
        frmPreferances.ComboBillPrinter.SetFocus
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


Private Sub FillSpeciality()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblspeciality order by speciality "
    .Open
    ListSpecialities.AddItem "All"
    ListSpecialityIDs.AddItem "All"
    If .RecordCount <> 0 Then
        While Not .EOF
            ListSpecialities.AddItem !speciality
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
    .Source = "SELECT tblfacilitystaff.* , tbldoctor.* FROM tblfacilitystaff left join tbldoctor on tblfacilitystaff.staff_ID = tbldoctor.doctor_ID  where tblfacilitystaff.HospitalFacility_ID = 10 order by doctorlistedname"
    .Open
    If .RecordCount = 0 Then Exit Sub
    While Not .EOF
        ListConsultants.AddItem !doctorlistedname
        ListConsultantIDs.AddItem !facilitystaff_ID
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub ListSelectedConsultants()
    Call FormatGridConsultants
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = "SELECT tblfacilitystaff.* , tbldoctor.* FROM tblfacilitystaff left join tbldoctor on tblfacilitystaff.staff_ID = tbldoctor.doctor_ID  where tblfacilitystaff.HospitalFacility_ID = 10 and doctorspeciality_ID = " & Val(ListSpecialityIDs.Text) & " order by doctorlistedname"
        .Open
        If .RecordCount = 0 Then Exit Sub
        While Not .EOF
            ListConsultants.AddItem !doctorlistedname
            ListConsultantIDs.AddItem !facilitystaff_ID
            .MoveNext
        Wend
        .Close
    End With
End Sub


Private Sub ListConsultants_Click()
MsgBox "OK"

'    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
'
'    Call FormatGridDates
'    Call FormatGridPatients
'
'    TemStaffFacilityID = 0
'    TemDoctorFee = 0
'    TemFDoctorFee = 0
'    TemInstitutionFee = 0
'    TemFInstitutionFee = 0
'    TemOtherFee = 0
'    TemDoctorID = 0
'    TemAppointmentDate = Empty
'    TemAppointmentTime = Empty
'    TwoSecessions = True
'
'    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
'    TemStaffFacilityID = Val(ListConsultantIDs.Text)
'
'    Call FillDates
'    Call FindStaffFacilityDetails
'
'
End Sub

Private Sub FillDates()
    
    Call FormatGridDates
    
    Dim TemCounter As Long
    Dim TemBookingDate As Date
    Dim TemDateCounter As Long
    Dim NowRow As Long
    
    Dim MorningAvailable As Boolean
    Dim EveningAvailable As Boolean
    
    Dim TemPreviousDate As Date
    
    Dim TemTextForList As String
    
    With DataEnvironment1.rssqlTem5
        If .State = 1 Then .Close
        .Source = "SELECT tblfacilitystaff.* from tblfacilitystaff where facilitystaff_ID = " & TemStaffFacilityID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then .Close: Exit Sub
        .Close
    End With
    
    TemCounter = 0
    TemDateCounter = 0
    NowRow = 0
    
    While TemCounter < AdvanceBookingDays And TemDateCounter < 31
        TemBookingDate = DateAdd("d", TemDateCounter, Date)
        MorningAvailable = False
        EveningAvailable = False
        
        If TwoSecessions = True Then
            If FacilityAvailable(TemBookingDate, True, False) = True Then
                MorningAvailable = True
                If TemPreviousDate = TemBookingDate Then
                    TemTextForList = vbTab
                Else
                    TemTextForList = Format(TemBookingDate, "dd MMM yyyy")
                End If
                TemTextForList = TemTextForList & vbTab & "Morning"
                If SecessionMax <> 0 Then
                    TemTextForList = TemTextForList & vbTab & SecessionMax
                Else
                    TemTextForList = TemTextForList & vbTab & "No Limit"
                End If
                TemTextForList = TemTextForList & vbTab & TemSecessionStartingTime
                TemTextForList = TemTextForList & vbTab & GetBookedNumber(TemBookingDate, True, False)
                ListDates.AddItem TemBookingDate
                ListSecessionIDs.AddItem MorningSecession
                ListDatesAndSecessions.AddItem TemTextForList
                ListSecessionMax.AddItem SecessionMax
                ListSecessionStartingTime.AddItem TemSecessionStartingTime
                TemPreviousDate = TemBookingDate
            End If
            If FacilityAvailable(TemBookingDate, False, True) = True Then
                EveningAvailable = True
                If TemPreviousDate = TemBookingDate Then
                    TemTextForList = vbTab
                Else
                    TemTextForList = Format(TemBookingDate, "dd MMM yyyy")
                End If
                TemTextForList = TemTextForList & vbTab & "Evening"
                If SecessionMax <> 0 Then
                    TemTextForList = TemTextForList & vbTab & SecessionMax
                Else
                    TemTextForList = TemTextForList & vbTab & "No Limit"
                End If
                TemTextForList = TemTextForList & vbTab & TemSecessionStartingTime
                TemTextForList = TemTextForList & vbTab & GetBookedNumber(TemBookingDate, False, True)
                ListDates.AddItem TemBookingDate
                ListSecessionIDs.AddItem EveningSecession
                ListDatesAndSecessions.AddItem TemTextForList
                ListSecessionMax.AddItem SecessionMax
                ListSecessionStartingTime.AddItem TemSecessionStartingTime
                TemPreviousDate = TemBookingDate
            End If
            If MorningAvailable = True Or EveningAvailable = True Then TemCounter = TemCounter + 1
        Else
            If FacilityAvailable(TemBookingDate, False, False) = True Then
                If TemPreviousDate = TemBookingDate Then
                    TemTextForList = vbTab
                Else
                    TemTextForList = Format(TemBookingDate, "dd MMM yyyy")
                End If
                TemTextForList = TemTextForList & vbTab & "Day"
                If SecessionMax <> 0 Then
                    TemTextForList = TemTextForList & vbTab & SecessionMax
                Else
                    TemTextForList = TemTextForList & vbTab & "No Limit"
                End If
                TemTextForList = TemTextForList & vbTab & TemSecessionStartingTime
                TemTextForList = TemTextForList & vbTab & GetBookedNumber(TemBookingDate, False, False)
                ListDates.AddItem TemBookingDate
                ListSecessionIDs.AddItem NoReleventSecession
                ListDatesAndSecessions.AddItem TemTextForList
                ListSecessionMax.AddItem SecessionMax
                ListSecessionStartingTime.AddItem TemSecessionStartingTime
                TemPreviousDate = TemBookingDate
            End If
        End If
        TemDateCounter = TemDateCounter + 1
    Wend
End Sub

Private Sub FindStaffFacilityDetails()
With DataEnvironment1.rssqlTem5
    If .State = 1 Then Close
    .Source = "SELECT * from tblfacilitystaff where facilitystaff_ID = " & TemStaffFacilityID
    .Open
    If .RecordCount = 0 Then Exit Sub
    
        TemDoctorFee = !usualpersonalFee
        TemFDoctorFee = !foreignerpersonalfee
        TemInstitutionFee = !UsualInstitutionFee
        TemFInstitutionFee = !ForeignerinstitutionFee
        TemDoctorID = !staff_ID
        TemUsualDuration = !usualduration
        
        If chkForigner.Value = 0 Then
            lblAgentAmount.Caption = Format((TemDoctorFee + TemInstitutionFee), "#0.00")
            lblCashDue.Caption = Format((TemDoctorFee + TemInstitutionFee), "#0.00")
            lblCredit.Caption = Format((TemDoctorFee + TemInstitutionFee), "#0.00")
        Else
            lblAgentAmount.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "#0.00")
            lblCashDue.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "#0.00")
            lblCredit.Caption = Format((TemFDoctorFee + TemFInstitutionFee), "#0.00")
        End If
        
    .Close
End With
End Sub


Private Function GetBookedNumber(BookingDate As Date, TemMorningSecession As Boolean, TemEveningSecession As Boolean) As Long
With DataEnvironment1.rssqlTem5
If TemMorningSecession = True Then
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientfacility where FacilityStaff_ID = " & TemStaffFacilityID & " and AppointmentDate = #" & BookingDate & "# and Secession = " & MorningSecession & " and cancelled = false"
    .Open
    GetBookedNumber = .RecordCount
ElseIf TemEveningSecession = True Then
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientfacility where FacilityStaff_ID = " & TemStaffFacilityID & " and AppointmentDate = #" & BookingDate & "# and Secession = " & EveningSecession & " and cancelled = false"
    .Open
    GetBookedNumber = .RecordCount
Else
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientfacility where FacilityStaff_ID = " & TemStaffFacilityID & " and AppointmentDate = #" & BookingDate & "# and Secession = " & NoReleventSecession & " and cancelled = false"
    .Open
    GetBookedNumber = .RecordCount

End If


    If .State = 1 Then .Close
End With
End Function


Private Sub ListConsultants_GotFocus()
    BoxConsultant.BackColor = BttnBackColour ' vbRed
End Sub

Private Sub ListConsultants_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    Call FormatGridDates
    Call FormatGridPatients
    TemStaffFacilityID = 0
    TemDoctorFee = 0
    TemFDoctorFee = 0
    TemInstitutionFee = 0
    TemFInstitutionFee = 0
    TemOtherFee = 0
    TemDoctorID = 0
    TemAppointmentDate = Empty
    TemAppointmentTime = Empty
    TwoSecessions = True
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    TemStaffFacilityID = Val(ListConsultantIDs.Text)
    Call FillDates
    Call FindStaffFacilityDetails
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

Private Sub ListDatesAndSecessions_Click()
    ListDates.ListIndex = ListDatesAndSecessions.ListIndex
    ListSecessionIDs.ListIndex = ListDatesAndSecessions.ListIndex
    ListSecessionMax.ListIndex = ListDatesAndSecessions.ListIndex
    ListSecessionStartingTime.ListIndex = ListDatesAndSecessions.ListIndex
    
    TemAppointmentDate = ListDates.Text
    
    MonthView1.Value = TemAppointmentDate
    
    Call FormatGridPatients
    
    If Not IsDate(ListDates.Text) Then Exit Sub
    If Not IsNumeric(ListSecessionIDs.Text) Then Exit Sub
    If Not IsDate(ListSecessionStartingTime.Text) Then Exit Sub
    
    TemSecession = Val(ListSecessionIDs.Text)
    SecessionMax = Val(ListSecessionMax.Text)
    TemSecessionStartingTime = Val(ListSecessionStartingTime.Text)
    
    Call FillGridPatients
    Call FindAppointmentTime
    

End Sub



Private Sub FindAppointmentTime()
    If TemUsualDuration = 0 Then Exit Sub
    If TemSecessionStartingTime = TimeSerial(0, 0, 0) Then Exit Sub
    TemAppointmentTime = TimeSerial(Hour(TemSecessionStartingTime), Minute(TemSecessionStartingTime) + TemUsualDuration, 0)
End Sub

Private Sub FillGridPatients()
    Dim TemTextForList As String
    
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = "SELECT * from tblpatientfacility where FacilityStaff_ID = " & TemStaffFacilityID & " and AppointmentDate = #" & TemAppointmentDate & "# and Secession = " & TemSecession & " order by DaySerial"
        .Open
        If .RecordCount = 0 Then Exit Sub
        While Not .EOF
            TemTextForList = !DaySerial & vbTab & FindPatientByID(!patientid)
            ListPatientFacilities.AddItem TemTextForList
            ListPatientFacilityIDs.AddItem !PatientFacility_ID
            .MoveNext
        Wend
    End With

End Sub

Private Sub ListDatesAndSecessions_GotFocus()
    BoxDates.BackColor = BttnBackColour ' vbRed
End Sub

Private Sub ListDatesAndSecessions_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
    txtPatientName.SetFocus
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

Private Sub ListPatientFacilities_Click()
    ListPatientFacilityIDs.ListIndex = ListPatientFacilities.ListIndex
    
    If IsNumeric(ListPatientFacilityIDs.Text) Then
        FrameCancellations.Enabled = True
        FrameRefunds.Enabled = True
        FrameReprints.Enabled = True
        FrameSettleCredit.Enabled = True
        TemPatientFacilityID = Val(ListPatientFacilityIDs.Text)
        Call GetPatientDetails
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
        
' **************************

' **************************

        .Close
    End With
End Sub

Private Sub ClearPatientDetails()

End Sub


Private Sub ListPatientFacilities_GotFocus()
    BoxPatients.BackColor = BttnBackColour ' vbRed
End Sub


Private Sub ListPatientFacilities_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    SSTab2.SetFocus
    KeyCode = Nothing
ElseIf KeyCode = vbright Then
    ListDatesAndSecessions.SetFocus
    KeyCode = Nothing
Else

End If
End Sub

Private Sub ListPatientFacilities_LostFocus()
    BoxPatients.BackColor = FrameBackColour ' vbRed
End Sub

Private Sub ListSpecialities_Click()
    ListSpecialityIDs.ListIndex = ListSpecialities.ListIndex
    If ListSpecialities.Text = "All" Then
        ListAllConsultants
    ElseIf ListSpecialities.Text <> "All" And IsNumeric(ListSpecialityIDs.Text) = True Then
        ListSelectedConsultants
    Else
        FormatGridConsultants
    End If
End Sub



Private Function FacilityAvailable(BookingDate As Date, TemMorningSecession As Boolean, TemEveningSecession As Boolean) As Boolean
    Dim TemResponce  As Integer
    FacilityAvailable = False
    With DataEnvironment1.rssqlTem6
        If .State = 1 Then .Close
        .Source = "SELECT * from tblfacilitystaff where FacilityStaff_ID = " & TemStaffFacilityID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Function
        
        Select Case Weekday(BookingDate)
        Case vbMonday:
            If TemMorningSecession = True Then
                SecessionMax = !FacilityMondayMNo
                TemSecessionStartingTime = !FacilityMondayMStarting
            ElseIf TemEveningSecession = True Then
                SecessionMax = !FacilityMondayENo
                TemSecessionStartingTime = !FacilityMondayEStarting
            Else
                SecessionMax = !mondaymax
            End If
        Case vbTuesday:
            If TemMorningSecession = True Then
                SecessionMax = !FacilitytuesdayMNo
                TemSecessionStartingTime = !FacilitytuesdayMStarting
            ElseIf TemEveningSecession = True Then
                SecessionMax = !FacilitytuesdayENo
                TemSecessionStartingTime = !FacilitytuesdayEStarting
            Else
                SecessionMax = !Tuesdaymax
            End If
        Case vbWednesday:
            If TemMorningSecession = True Then
                SecessionMax = !FacilitywednesdayMNo
                TemSecessionStartingTime = !FacilitywednesdayMStarting
            ElseIf TemEveningSecession = True Then
                SecessionMax = !FacilitywednesdayENo
                TemSecessionStartingTime = !FacilitywednesdayEStarting
            Else
                SecessionMax = !Wednesdaymax
            End If
        Case vbThursday:
            If TemMorningSecession = True Then
                SecessionMax = !FacilitythursdayMNo
                TemSecessionStartingTime = !FacilitythursdayMStarting
            ElseIf TemEveningSecession = True Then
                SecessionMax = !FacilitythursdayENo
                TemSecessionStartingTime = !FacilitythursdayEStarting
            Else
                SecessionMax = !Thursdaymax
            End If
        Case vbFriday:
            If TemMorningSecession = True Then
                SecessionMax = !FacilityfridayMNo
                TemSecessionStartingTime = !FacilityfridayMStarting
            ElseIf TemEveningSecession = True Then
                SecessionMax = !FacilityfridayENo
                TemSecessionStartingTime = !FacilityfridayEStarting
            Else
                SecessionMax = !Fridaymax
            End If
        Case vbSaturday:
            If TemMorningSecession = True Then
                SecessionMax = !FacilitysaturdayMNo
                TemSecessionStartingTime = !FacilitysaturdayMStarting
            ElseIf TemEveningSecession = True Then
                SecessionMax = !FacilitysaturdayENo
                TemSecessionStartingTime = !FacilitysaturdayEStarting
            Else
                SecessionMax = !Saturdaymax
            End If
        Case vbSunday:
            If TemMorningSecession = True Then
                SecessionMax = !FacilitysundayMNo
                TemSecessionStartingTime = !FacilitysundayMStarting
            ElseIf TemEveningSecession = True Then
                SecessionMax = !FacilitysundayENo
                TemSecessionStartingTime = !FacilitysundayEStarting
            Else
                SecessionMax = !Sundaymax
            End If
        End Select
        
        Select Case Weekday(BookingDate)
        Case vbMonday:
            If !FullDayLeaveMonday = True Then
                Exit Function
            End If
            If !FacilityMondayM = False And TemMorningSecession = True Then
                Exit Function
            End If
            If !FacilityMondayE = False And TemEveningSecession = True Then
                Exit Function
            End If
        Case vbTuesday:
            If !FullDayLeaveTuesday = True Then
                Exit Function
            End If
            If !FacilityTuesdayM = False And TemMorningSecession = True Then
                Exit Function
            End If
            If !FacilityTuesdayE = False And TemEveningSecession = True Then
                Exit Function
            End If
        Case vbWednesday:
            If !FullDayLeaveWednesday = True Then
                Exit Function
            End If
            If !FacilityWednesdayM = False And TemMorningSecession = True Then
                Exit Function
            End If
            If !FacilityWednesdayE = False And TemEveningSecession = True Then
                Exit Function
            End If
        Case vbThursday:
            If !FullDayLeaveThursday = True Then
                Exit Function
            End If
            If !FacilityThursdayM = False And TemMorningSecession = True Then
                Exit Function
            End If
            If !FacilityThursdayE = False And TemEveningSecession = True Then
                Exit Function
            End If
        Case vbFriday:
            If !FullDayLeaveFriday = True Then
                Exit Function
            End If
            If !FacilityFridayM = False And TemMorningSecession = True Then
                Exit Function
            End If
            If !FacilityFridayE = False And TemEveningSecession = True Then
                Exit Function
            End If
        Case vbSaturday:
            If !FullDayLeaveSaturday = True Then
                Exit Function
            End If
            If !FacilitySaturdayM = False And TemMorningSecession = True Then
                Exit Function
            End If
            If !FacilitySaturdayE = False And TemEveningSecession = True Then
                Exit Function
            End If
        Case vbSunday:
            If !FullDayLeaveSunday = True Then
                Exit Function
            End If
            If !FacilitySundayM = False And TemMorningSecession = True Then
                Exit Function
            End If
            If !FacilitySundayE = False And TemEveningSecession = True Then
                Exit Function
            End If
        End Select
    .Close
    .Source = "SELECT * from tblfacilitystaffleave where (FacilityStaff_ID = " & TemStaffFacilityID & ") and (FacilityStaffLeaveDate = #" & BookingDate & "#)"
    If .State = 0 Then .Open
    
    If .RecordCount = 0 Then
        FacilityAvailable = True
        Exit Function
    Else
        If TemMorningSecession = True Then
            SecessionMax = !morningMax
            TemSecessionStartingTime = !morningStarting
        ElseIf TemEveningSecession = True Then
            SecessionMax = !eveningmax
            TemSecessionStartingTime = !eveningstarting
        Else
            SecessionMax = !daymax
        End If
    End If
    
    If !fulldayleave = True Then
        Exit Function
    End If
    If !Morning = False And TemMorningSecession = True Then
        Exit Function
    End If
    If !Evening = False And TemEveningSecession = True Then
        Exit Function
    End If
    .Close
    End With
    FacilityAvailable = True
End Function


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

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Dim TemNum As Long
    Dim DateFound As Boolean
    Dim Tem
    
For TemNum = 1 To ListDates.ListCount - 1
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
    ListDates.ListIndex = 0
Else
    ListDatesAndSecessions.ListIndex = ListDates.ListIndex
    ListDatesAndSecessions_Click
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Tab = 2 Then
    chkPrint.Value = 0
Else
    chkPrint.Value = 1
End If

End Sub

