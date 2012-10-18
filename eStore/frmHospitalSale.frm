VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHospitalSale 
   Caption         =   "Hospital Issue"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10470
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   1680
      TabIndex        =   101
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtTotalStock 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   98
      Top             =   480
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   495
      Left            =   11280
      TabIndex        =   22
      Top             =   7440
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   76873731
      CurrentDate     =   39691
   End
   Begin VB.TextBox txtTotalCost 
      Height          =   375
      Left            =   9360
      TabIndex        =   92
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   12000
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   11280
      TabIndex        =   21
      Top             =   7920
      Width           =   3855
      Begin VB.OptionButton optOne 
         Caption         =   "1"
         Enabled         =   0   'False
         Height          =   240
         Left            =   2160
         TabIndex        =   24
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton OptTwo 
         Caption         =   "2"
         Enabled         =   0   'False
         Height          =   240
         Left            =   3120
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.TextBox txtItemCost 
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCostRate 
      Height          =   375
      Left            =   10440
      TabIndex        =   19
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtSaleProfit 
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtBHTProfit 
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCategoryProfit 
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSPrice 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin btButtonEx.ButtonEx bttnSettle 
      Height          =   615
      Left            =   11280
      TabIndex        =   26
      Top             =   8640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Appearance      =   3
      BorderColor     =   12583104
      Caption         =   "Se&ttle"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GridBatch 
      Height          =   1575
      Left            =   6840
      TabIndex        =   28
      Top             =   840
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2778
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   495
      Left            =   13920
      TabIndex        =   8
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   12583104
      Caption         =   "&Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtcCatogery 
      Height          =   465
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   820
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItem 
      Height          =   465
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   820
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2895
      Left            =   240
      TabIndex        =   11
      Top             =   7440
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   5106
      _Version        =   393216
      Tab             =   1
      TabHeight       =   617
      ForeColor       =   12583104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Total"
      TabPicture(0)   =   "frmHospitalSale.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label24"
      Tab(0).Control(1)=   "Label23"
      Tab(0).Control(2)=   "Label22"
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(4)=   "dtcSale"
      Tab(0).Control(5)=   "txtNTotal"
      Tab(0).Control(6)=   "txtDiscount"
      Tab(0).Control(7)=   "txtGTotal"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Payment"
      TabPicture(1)   =   "frmHospitalSale.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frameUnit"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameStaff"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frameOutPatient"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frameInPatient"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "frameCreditCard"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "frameCash"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "frameCredit"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "frameCheque"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "frmHospitalSale.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkRequest"
      Tab(2).Control(1)=   "dtcDepartment"
      Tab(2).Control(2)=   "dtcCheckedStaff"
      Tab(2).Control(3)=   "dtcIssueStaff"
      Tab(2).Control(4)=   "Label20"
      Tab(2).Control(5)=   "Label21"
      Tab(2).ControlCount=   6
      Begin VB.Frame frameCheque 
         Caption         =   "Cheque"
         Height          =   2415
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   5535
         Begin VB.TextBox txtChequeNo 
            Height          =   375
            Left            =   1320
            TabIndex        =   49
            Top             =   1320
            Width           =   4095
         End
         Begin MSComCtl2.DTPicker dtpChequeDate 
            Height          =   495
            Left            =   1320
            TabIndex        =   48
            Top             =   1800
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   873
            _Version        =   393216
            CalendarForeColor=   12583104
            CalendarTitleForeColor=   12583104
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   76873731
            CurrentDate     =   39551
         End
         Begin MSDataListLib.DataCombo dtcBranch 
            Height          =   465
            Left            =   1320
            TabIndex        =   50
            Top             =   840
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   820
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcBank 
            Height          =   465
            Left            =   1320
            TabIndex        =   51
            Top             =   360
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   820
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label34 
            Caption         =   "Branch"
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label35 
            Caption         =   "Bank"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label36 
            Caption         =   "No"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label37 
            Caption         =   "Date"
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   1800
            Width           =   1575
         End
      End
      Begin VB.Frame frameCredit 
         Caption         =   "Credit"
         Height          =   2175
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   3495
         Begin VB.TextBox txtCreditDue 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            TabIndex        =   45
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label30 
            Caption         =   "Due"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame frameCash 
         Caption         =   "Cash"
         Height          =   2175
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   3495
         Begin VB.TextBox txtDue 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            TabIndex        =   40
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtCashPaid 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            TabIndex        =   39
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            TabIndex        =   38
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label25 
            Caption         =   "Due"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label26 
            Caption         =   "&Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label27 
            Caption         =   "Change"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.TextBox txtGTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -67200
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -67200
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtNTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -67200
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CheckBox chkRequest 
         Caption         =   "Make a request"
         Height          =   375
         Left            =   -74760
         TabIndex        =   32
         Top             =   2040
         Visible         =   0   'False
         Width           =   3255
      End
      Begin MSDataListLib.DataCombo dtcDepartment 
         Height          =   465
         Left            =   -70080
         TabIndex        =   31
         Top             =   1920
         Visible         =   0   'False
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   820
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcSale 
         Height          =   465
         Left            =   -74760
         TabIndex        =   13
         Top             =   1080
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   820
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcCheckedStaff 
         Height          =   360
         Left            =   -70080
         TabIndex        =   35
         Top             =   1080
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   820
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcIssueStaff 
         Height          =   360
         Left            =   -70080
         TabIndex        =   36
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   820
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame frameCreditCard 
         Caption         =   "Credit Card"
         Height          =   2415
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   5535
         Begin VB.TextBox txtCreditCardNo 
            Height          =   375
            Left            =   1320
            TabIndex        =   60
            Top             =   1320
            Width           =   4095
         End
         Begin VB.TextBox txtCreditCode 
            Height          =   375
            Left            =   1320
            TabIndex        =   57
            Top             =   1800
            Width           =   4095
         End
         Begin MSDataListLib.DataCombo dtcCardBank 
            Height          =   465
            Left            =   1320
            TabIndex        =   58
            Top             =   840
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   820
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcCreditCard 
            Height          =   465
            Left            =   1320
            TabIndex        =   59
            Top             =   360
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   820
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label31 
            Caption         =   "No"
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label32 
            Caption         =   "Card"
            Height          =   375
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label33 
            Caption         =   "Bank"
            Height          =   375
            Left            =   120
            TabIndex        =   62
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label28 
            Caption         =   "Code"
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   1800
            Width           =   1575
         End
      End
      Begin VB.Frame frameInPatient 
         Caption         =   "Indoor Patient"
         Height          =   2415
         Left            =   5760
         TabIndex        =   68
         Top             =   360
         Width           =   5055
         Begin VB.TextBox txtTemCreditCustomerBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   1440
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.TextBox txtPatient 
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   960
            Width           =   3615
         End
         Begin VB.TextBox txtBHTBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   1440
            Width           =   3615
         End
         Begin MSDataListLib.DataCombo dtcBHT 
            Height          =   465
            Left            =   1320
            TabIndex        =   69
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   820
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblHealthSchemeSupplier 
            Height          =   375
            Left            =   1320
            TabIndex        =   97
            Top             =   1920
            Width           =   3615
         End
         Begin VB.Label Label38 
            Caption         =   "BHT"
            Height          =   375
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label39 
            Caption         =   "Patient"
            Height          =   495
            Left            =   120
            TabIndex        =   74
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label40 
            Caption         =   "Balance"
            Height          =   495
            Left            =   120
            TabIndex        =   73
            Top             =   1440
            Width           =   1575
         End
      End
      Begin VB.Frame frameOutPatient 
         Caption         =   "Out Patient"
         Height          =   2415
         Left            =   5760
         TabIndex        =   76
         Top             =   360
         Width           =   5055
         Begin VB.TextBox txtCreditCustomerBalance 
            Height          =   375
            Left            =   1320
            TabIndex        =   77
            Top             =   960
            Width           =   3615
         End
         Begin MSDataListLib.DataCombo dtcCreditCustomer 
            Height          =   465
            Left            =   1320
            TabIndex        =   78
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   820
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label42 
            Caption         =   "Balance"
            Height          =   375
            Left            =   120
            TabIndex        =   80
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label43 
            Caption         =   "Name"
            Height          =   375
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame frameStaff 
         Caption         =   "Staff Issue"
         Height          =   2415
         Left            =   5760
         TabIndex        =   81
         Top             =   360
         Width           =   5055
         Begin VB.TextBox txtStaffBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   960
            Width           =   3615
         End
         Begin VB.TextBox txtTemStaffCredit 
            Height          =   375
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   1680
            Visible         =   0   'False
            Width           =   2535
         End
         Begin MSDataListLib.DataCombo dtcStaffCustomer 
            Height          =   465
            Left            =   1320
            TabIndex        =   84
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   820
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label41 
            Caption         =   "Balance"
            Height          =   495
            Left            =   120
            TabIndex        =   86
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label44 
            Caption         =   "Staff"
            Height          =   375
            Left            =   120
            TabIndex        =   85
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame frameUnit 
         Caption         =   "Select the Unit"
         Height          =   2415
         Left            =   5760
         TabIndex        =   65
         Top             =   360
         Width           =   5055
         Begin MSDataListLib.DataCombo dtcUnit 
            Height          =   465
            Left            =   960
            TabIndex        =   66
            Top             =   480
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   820
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label46 
            Caption         =   "Unit"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Label Label14 
         Caption         =   "Total"
         Height          =   375
         Left            =   -68760
         TabIndex        =   91
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Discount"
         Height          =   495
         Left            =   -68760
         TabIndex        =   90
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label23 
         Caption         =   "Net Total"
         Height          =   375
         Left            =   -68760
         TabIndex        =   89
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Sal&e Catogery"
         Height          =   375
         Left            =   -74880
         TabIndex        =   12
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label20 
         Caption         =   "Issued By"
         Height          =   495
         Left            =   -74880
         TabIndex        =   88
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "Checked By"
         Height          =   495
         Left            =   -74880
         TabIndex        =   87
         Top             =   1200
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   4335
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   495
      Left            =   13920
      TabIndex        =   10
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   12583104
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo dtcCode 
      Height          =   465
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   820
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   13440
      TabIndex        =   27
      Top             =   8640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Appearance      =   3
      BorderColor     =   12583104
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCategory 
      Height          =   375
      Left            =   1920
      TabIndex        =   102
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lblDristributor 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   7680
      TabIndex        =   100
      Top             =   6960
      Width           =   7335
   End
   Begin VB.Label Label3 
      Caption         =   "Total Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   99
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "&Quantity"
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Rate"
      Height          =   375
      Left            =   10440
      TabIndex        =   96
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Price"
      Height          =   375
      Left            =   12000
      TabIndex        =   95
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "&Category"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblIUnit 
      Height          =   375
      Left            =   8400
      TabIndex        =   94
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblDisplayTotal 
      Caption         =   "Cash Rs. 0.00"
      Height          =   375
      Left            =   240
      TabIndex        =   93
      Top             =   6960
      Width           =   10935
   End
   Begin VB.Label Label29 
      Caption         =   "&Item"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label45 
      Caption         =   "C&ode"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "frmHospitalSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCategory As New ADODB.Recordset
    Dim rsCode As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsUnit As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    Dim rsTemPrice As New ADODB.Recordset
    Dim rsTemOrder As New ADODB.Recordset
    Dim rsTemSaleBill As New ADODB.Recordset
    Dim rsTemSale As New ADODB.Recordset
    Dim rsTemBatch As New ADODB.Recordset
    Dim rsTemPatient As New ADODB.Recordset
    Dim rsTemCC As New ADODB.Recordset
    Dim rsTemCash As New ADODB.Recordset
    Dim rsTemCredit As New ADODB.Recordset
    Dim rsTemCheque As New ADODB.Recordset
    Dim rsTemCustomer As New ADODB.Recordset
    Dim rsTemDistributor As New ADODB.Recordset

    Dim rsBanks As New ADODB.Recordset
    Dim rsCities As New ADODB.Recordset
    Dim rsCreditCards As New ADODB.Recordset
    Dim rsSale As New ADODB.Recordset
    Dim rsTemStaff As New ADODB.Recordset
    Dim rsBHT As New ADODB.Recordset
    Dim rsPatients As New ADODB.Recordset
    Dim rsStore As New ADODB.Recordset
    Dim temSql As String
    Dim NewItem As New Item
    Dim NewSale As New Sale
    
    
    
    Dim rsDI As New ADODB.Recordset
    Dim TemDI As Long
    Dim rsTemDistributor1 As New ADODB.Recordset
    

    Dim LastVisibleRow As Long
    
    Dim TemSaleBillID As Long

    Dim CsetPrinter As New cSetDfltPrinter
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean


Private Sub bttnAdd_Click()
    If CanAdd = False Then Exit Sub
    With GridItem
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .CellAlignment = 1
        .Text = .Row
        .Col = 1
        .CellAlignment = 1
        .Text = dtcItem.Text
        .Col = 2
        .CellAlignment = 1
        .Text = GridBatch.TextMatrix(GridBatch.Row, 0)
        .Col = 3
        .CellAlignment = 7
        If NewSale.Unit = True Then
            .Text = Format(Val(txtCostRate.Text), "0.00") & " per " & NewItem.IUnit
        Else
            .Text = Format(Val(txtRate.Text), "0.00") & " per " & NewItem.IUnit
        End If
        .Col = 4
        .CellAlignment = 7
        .Text = txtQty.Text & " " & NewItem.IUnit
        .Col = 5
        .CellAlignment = 7
        .Text = Format(Val(txtPrice.Text), "0.00")
        .Col = 6
        .Text = Val(dtcItem.BoundText)
        .Col = 7
        .Text = GridBatch.TextMatrix(GridBatch.Row, 4)
        .Col = 9
        .CellAlignment = 7
        .Text = Format(Val(txtRate.Text), "0.00")
        .Col = 8
        .CellAlignment = 7
        .Text = txtQty.Text
        .Col = 10
        .CellAlignment = 7
        .Text = Val(txtItemCost.Text)
        .Col = 11
        .CellAlignment = 7
        .Text = dtcCatogery.Text
        .Col = 12
        .CellAlignment = 7
        .Text = dtcCatogery.BoundText
        .Col = 13
        .Text = lblIUnit.Caption
        .Col = 14
        .Text = Val(txtCategoryProfit.Text)
        .Col = 15
        .Text = Val(txtSaleProfit.Text)
        .Col = 16
        .Text = Val(txtBHTProfit.Text)
        .Col = 17
        .Text = Val(txtSPrice.Text)
        
'   0   No
'   1   Item
'   2   Batch
'   3   Rate
'   4   Amount
'   5   Price
'   6   ItemID
'   7   BatchID
'   8   AMount
'   9   Rate
'   10  Cost
'   11  Category
'   12  CatogoryID
'   13  IUnit

'   14  CategoryProfit
'   15  SaleProfit
'   16  BHTProfit
'   17  Real Price
        
        CalculateTotal
        ClearAddValues
        FormatSelectStock
        CalculateDiscount
    End With
   ' If GridItem.Rows > 9 Then GridItem.TopRow = GridItem.Rows - 9
    If GridItem.RowIsVisible(GridItem.Row) = False Then
        GridItem.TopRow = GridItem.Rows - LastVisibleRow
    Else
        LastVisibleRow = LastVisibleRow + 1
    End If
    bttnDelete.Enabled = False
    dtcCatogery.Text = Empty
    dtcCatogery.SetFocus
End Sub

Private Sub ClearAddValues()
    txtQty.Text = Empty
    txtRate.Text = Empty
    txtPrice.Text = Empty
    txtItemCost.Text = Empty
    dtcItem.Text = Empty
    dtcCode.Text = Empty
    txtCostRate.Text = Empty
    lblDristributor.Caption = Empty
'    dtcCatogery.Text = Empty

'
'    dtcBHT.Text = Empty
'    dtcCreditCustomer.Text = Empty
'    dtcStaffCustomer.Text = Empty
End Sub

Private Sub CalculateTotal()
    Dim i As Integer
    Dim Total As Double
    Dim Cost As Double
    With GridItem
        For i = 1 To .Rows - 1
            Total = Total + Val(.TextMatrix(i, 5))
            Cost = Cost + Val(.TextMatrix(i, 10))
        Next
    End With
    txtGTotal.Text = Format(Total, "0.00")
    txtTotalCost.Text = Cost
End Sub

Private Sub CalculateNetTotal()
    txtNTotal.Text = Format(Val(txtGTotal.Text) - Val(txtDiscount.Text), "0.00")
End Sub

Private Function CanAdd() As Boolean
    CanAdd = False
    Dim tr As Integer
        If IsNumeric(dtcItem.BoundText) = False Then
            tr = MsgBox("You have not entered the item to add", vbCritical, "Item?")
            dtcItem.SetFocus
            Exit Function
        End If
        If IsNumeric(txtQty.Text) = False Or Val(txtQty.Text) = 0 Then
            tr = MsgBox("You have not entered the quentity", vbCritical, "Quentity?")
            txtQty.SetFocus
            Exit Function
        End If
        If txtRate.Visible = True Then
            If IsNumeric(txtRate.Text) = False Or Val(txtRate.Text) = 0 Then
                tr = MsgBox("You have not entered the rate", vbCritical, "Rate")
                txtRate.SetFocus
                SendKeys "{home}+{end}"
                Exit Function
            End If
        ElseIf txtCostRate.Visible = True Then
            If IsNumeric(txtCostRate.Text) = False Or Val(txtCostRate.Text) = 0 Then
                tr = MsgBox("You have not entered the rate", vbCritical, "Rate")
                txtCostRate.SetFocus
                SendKeys "{home}+{end}"
                Exit Function
            End If
        End If
        
        If CalculateStock(dtcItem.BoundText, , UserStoreID).Amount <= 0 Then
            tr = MsgBox("There are no stocks", vbCritical, "No Stocks")
            dtcCode.SetFocus
            Exit Function
        End If
        
        Dim x As Integer
        For x = 1 To GridItem.Rows - 1
            If GridItem.TextMatrix(x, 7) = GridBatch.TextMatrix(GridBatch.Row, 4) Then
                tr = MsgBox("One batch can be selected only once for a bill!", vbCritical, "Same Batch Twice")
                GridBatch.SetFocus
                Exit Function
            End If
        Next
        
        If QtyOK = False Then Exit Function
    CanAdd = True
End Function

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDelete_Click()
    With GridItem
        If .Rows <= 1 Then Exit Sub
        If .Rows = 2 Then
            FormatItemGrid
        Else
            .RemoveItem (.Row)
        End If
        Call CalculateTotal
        Call CalculateDiscount
        bttnDelete.Enabled = False
        Dim i As Integer
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
    End With
    If GridItem.Rows > 9 Then GridItem.TopRow = GridItem.Rows - 9
End Sub

Private Sub bttnSettle_Click()
    Dim TemOutPatientID As Long
    Dim temBHTID As Long
    Dim TemCreditCardID As Long
    Dim TemCashID As Long
    Dim TemCreditID As Long
    Dim TemChequeID As Long
    Dim TemOtherID As Long
    Dim i As Integer
    
    Dim MyTemStock As Stock
    
    txtDue.Text = txtNTotal.Text
    
    dtcCreditCustomer.Text = UCase(dtcCreditCustomer.Text)
    
    
    
    If CanSettle = False Then Exit Sub
    
    
    With GridItem
        For i = 1 To .Rows - 1
            MyTemStock = CalculateStock(Val(.TextMatrix(i, 6)), Val(.TextMatrix(i, 7)), UserStoreID)
            If MyTemStock.Amount < Val(.TextMatrix(i, 8)) Then
                MsgBox "There are no adequate stocks to sale" & vbNewLine & "Item : " & vbTab & GridItem.TextMatrix(i, 1) & vbNewLine & "Batch : " & vbTab & GridItem.TextMatrix(i, 9) & vbNewLine & "Current Stock : " & vbTab & MyTemStock.Amount & vbNewLine & "Sale quentity : " & vbTab & GridItem.TextMatrix(i, 14)
                Exit Sub
            End If
        Next
    End With
    
    
    If NewSale.OutPatient = True Then
        If IsNumeric(dtcCreditCustomer.BoundText) = True Then
            TemOutPatientID = dtcCreditCustomer.BoundText
        ElseIf dtcCreditCustomer.Text <> Empty Then
            TemOutPatientID = WritePatient
        Else
            TemOutPatientID = 1
            dtcCreditCustomer.BoundText = 1
        End If
    End If
    TemSaleBillID = SaleBillID
    If NewSale.CreditCard = True Then TemCreditCardID = ReceiveCreditCard(TemSaleBillID)
    If NewSale.Cash = True Then TemCashID = ReceiveCash(TemSaleBillID)
    If NewSale.Cheque = True Then ReceiveCheque (TemSaleBillID)
    If NewSale.Credit = True Then ReceiveCredit (TemSaleBillID)
    If NewSale.Other = True Then ReceiveOther (TemSaleBillID)
    If NewSale.Credit = True Then
        If NewSale.OutPatient = True Then
            With rsTemCustomer
                If .State = 1 Then .Close
                temSql = "SELECT * from tblPatientMainDetails where patientID = " & dtcCreditCustomer.BoundText
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !Credit = !Credit - Val(txtGTotal.Text)
                    .Update
                End If
                .Close
            End With
        ElseIf NewSale.InPatient = True Then
            With rsTemCustomer
                If .State = 1 Then .Close
                temSql = "SELECT * from tblBHT where BHTID = " & dtcBHT.BoundText
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !Balance = !Balance - Val(txtGTotal.Text)
                    .Update
                End If
                .Close
            End With
        ElseIf NewSale.Staff = True Then
            With rsTemCustomer
                If .State = 1 Then .Close
                temSql = "SELECT * from tblStaff where StaffID = " & dtcStaffCustomer.BoundText
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !Credit = !Credit - Val(txtGTotal.Text)
                    .Update
                End If
                .Close
            End With
        End If
    End If
    
    
    
    If rsTemSale.State = 1 Then rsTemSale.Close
    temSql = "SELECT tblSale.* FROM tblSale Where SaleBillID = 0"
    rsTemSale.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
    With GridItem
        For i = 1 To .Rows - 1
            If ConsumeStocks(UserStoreID, Val(.TextMatrix(i, 7)), Val(.TextMatrix(i, 8))) = True Then
                rsTemSale.AddNew
                rsTemSale!SaleBillID = TemSaleBillID
                rsTemSale!CategoryID = Val(dtcSale.BoundText)
                rsTemSale!ItemID = Val(.TextMatrix(i, 6))
                rsTemSale!BatchID = Val(.TextMatrix(i, 7))
                rsTemSale!StoreID = UserStoreID
                rsTemSale!Date = Date  'dtpDate.Value
                rsTemSale!Time = Now
                rsTemSale!StaffID = UserID
                If IsNumeric(dtcCheckedStaff.BoundText) = True Then rsTemSale!CheckedStaffID = dtcCheckedStaff.BoundText
                rsTemSale!Amount = Val(.TextMatrix(i, 8))
                rsTemSale!Rate = Val(.TextMatrix(i, 9))
                rsTemSale!GrossPrice = Val(.TextMatrix(i, 5))
                rsTemSale!Discount = Val(.TextMatrix(i, 5)) * NewSale.SaleDiscountPercent / 100
                rsTemSale!DiscountPercent = NewSale.SaleDiscountPercent
                rsTemSale!Price = rsTemSale!GrossPrice - rsTemSale!Discount
                rsTemSale!Cost = Val(.TextMatrix(i, 10))
                If NewSale.OutPatient = True Then
                    rsTemSale!BilledOutPatientID = TemOutPatientID
                ElseIf NewSale.InPatient = True Then
                    rsTemSale!BilledBHTID = dtcBHT.BoundText
                ElseIf NewSale.Staff = True Then
                    rsTemSale!BilledStaffID = dtcStaffCustomer.BoundText
                ElseIf NewSale.Unit = True Then
                    rsTemSale!BilledUnitID = Val(dtcUnit.BoundText)
                End If
                If NewSale.Cash = True Then
                    rsTemSale!PaymentMethodID = 1
                    rsTemSale!PaymentMethod = "Cash"
                ElseIf NewSale.Credit = True Then
                    rsTemSale!PaymentMethodID = 4
                    rsTemSale!PaymentMethod = "Credit"
                ElseIf NewSale.Cheque = True Then
                    rsTemSale!PaymentMethodID = 5
                    rsTemSale!PaymentMethod = "Cheque"
                ElseIf NewSale.CreditCard = True Then
                    rsTemSale!PaymentMethodID = 3
                    rsTemSale!PaymentMethod = "Credit Card"
                ElseIf NewSale.Other = True Then
                    rsTemSale!PaymentMethodID = 8
                    rsTemSale!PaymentMethod = "Other"
                End If
                rsTemSale.Update
            End If
        Next i
    End With
    With rsTemSaleBill
        If .State = 1 Then .Close
        temSql = "SELECT * from tblSaleBill where SaleBillID = " & TemSaleBillID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If NewSale.OutPatient = True Then
                !BilledOutPatientID = TemOutPatientID
            ElseIf NewSale.InPatient = True Then
                !BilledBHTID = dtcBHT.BoundText
            ElseIf NewSale.Staff = True Then
                !BilledStaffID = dtcStaffCustomer.BoundText
            ElseIf NewSale.Unit = True Then
                !BilledUnitID = Val(dtcUnit.BoundText)
            End If
            If NewSale.Cash = True Then
                !PaymentMethodID = 1
                !PaymentMethod = "Cash"
                !ReceivedCashID = TemCashID
            ElseIf NewSale.Credit = True Then
                !PaymentMethodID = 4
                !PaymentMethod = "Credit"
                !ReceivedCreditID = TemCreditID
            ElseIf NewSale.Cheque = True Then
                !PaymentMethodID = 5
                !PaymentMethod = "Cheque"
                !ReceivedChequeID = TemChequeID
            ElseIf NewSale.CreditCard = True Then
                !PaymentMethodID = 3
                !PaymentMethod = "Credit Card"
                !receivedCreditcardID = TemCreditCardID
            ElseIf NewSale.Other = True Then
                !PaymentMethodID = 8
                !PaymentMethod = "Other"
                !receivedCreditcardID = TemOtherID
            End If
            !NetCost = Val(txtTotalCost.Text)
            .Update
        End If
        .Close
    End With

'    Call SetBillPrinter
'    Call SetBillPaper
    
    
    If NewSale.OutPatient = True Then
        Call POSPrint
    Else
        Call RuhunaPrint
    End If
    
    Call ClearBillValues
    Call FormatItemGrid
    
    MsgBox "Bill Number : " & TemSaleBillID

    SSTab2.Tab = 0
'    dtcCode.SetFocus

    On Error Resume Next

    dtcCatogery.Text = Empty
    dtcCatogery.SetFocus
End Sub


Private Sub SetBillPrinter()
    CsetPrinter.SetPrinterAsDefault (BillPrinterName)
End Sub

Private Sub SetImpactPrinter()
    CsetPrinter.SetPrinterAsDefault (BillPrinterName)
End Sub

Private Sub SetPOSPrinter()
    'CSetPrinter.SetPrinterAsDefault (PPmt/)
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

Private Sub POSPrint()
    'On Error GoTo eh
    CsetPrinter.SetPrinterAsDefault (PrescreptionPrinterName)

    PrinterName = Printer.DeviceName
    
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    
    
    CsetPrinter.SetPrinterAsDefault (PrescreptionPrinterName)
    
        
    Dim MyPrinter As VB.Printer
    For Each MyPrinter In VB.Printers
        If MyPrinter.DeviceName = (PrescreptionPrinterName) Then
            Set Printer = MyPrinter
        End If
    Next
    
    If SelectForm(PrescreptionPaperName, Me.hwnd) = 1 Then
        Dim i As Integer
        Dim Tab1 As Integer
        Dim Tab2 As Integer
        Dim Tab3 As Integer
        Dim Tab4 As Integer
        Dim Teb5 As Interior
        
        Dim SmallestFontSize As Integer
    
        Tab1 = 0
        Tab2 = 4
        Tab3 = 28
        Tab4 = 20
    
    
        SmallestFontSize = 9
    
        With Printer
            '.Font = "Tahoma"
            
            '.Font = "Microsoft Sans Serif"
            
            '.Font.Name = "Niagara Engraved"
            
            .Font.Name = "Times New Roman"
            
            .Font.Bold = False
            
            .Font.Size = SmallestFontSize + 3
            Printer.Print dtcSale.Text
            
            .FontSize = SmallestFontSize + 4
            Printer.Print
            Printer.Print Tab(Tab1); HospitalName
            
            .FontSize = SmallestFontSize
            
            Printer.Print Tab(Tab1); HospitalAddress
            Printer.Print Tab(Tab1); HospitalDescreption
            Printer.Print
            
            
            
            
            Printer.Font.Size = SmallestFontSize
            Printer.Print
            Printer.Print Tab(Tab1); "Bill No." & TemSaleBillID
            Printer.Print Tab(Tab1); "Date : "; Format(Date, "dd MM yy"); Tab(Tab1 + 25); "Time : " & Time
            
            Printer.Print Tab(Tab1); "---------------------------------------------------------------"
            If NewSale.OutPatient = True Then
                Printer.Print Tab(Tab1); "Patient : "; dtcCreditCustomer.Text
            ElseIf NewSale.InPatient = True Then
                Printer.Print Tab(Tab1); "Indoor Patient : "; txtPatient.Text
            ElseIf NewSale.Staff = True Then
                Printer.Print Tab(Tab1); "Staff member : "; dtcStaffCustomer.Text
            End If
                Printer.Print Tab(Tab1); "---------------------------------------------------------------"
'            Printer.Print
    
            .FontSize = SmallestFontSize
    '        .Font = "Lucida Console"
        End With
        
        Tab1 = 0
        Tab2 = 29
        Tab3 = 34
        Tab4 = 42
        
        Printer.Print Tab(Tab1); "---------------------------------------------------------------"
        
        Printer.Print ; Tab(Tab1); Left("Descreption" & Space(100), Tab2 - Tab1 - 1);
        Printer.Print ; Tab(Tab2); Right((Space(4)) & "Qty", 4);
        Printer.Print ; Tab(Tab3); Right((Space(7)) & "Rate", 7);
        Printer.Print ; Tab(Tab4); Right((Space(11)) & "Value", 9)
        
        Printer.Print Tab(Tab1); "---------------------------------------------------------------"
        
        
        With GridItem
            For i = 1 To .Rows - 1
                Printer.Print ; Tab(Tab1); Left(.TextMatrix(i, 1) & Space(100), Tab2 - Tab1 - 1);
                Printer.Print ; Tab(Tab2); Right((Space(4)) & .TextMatrix(i, 8), 4);
                Printer.Print ; Tab(Tab3); Right((Space(7)) & Format(Val(.TextMatrix(i, 9)), "#,##0.00"), 7);
                Printer.Print ; Tab(Tab4); Right((Space(11)) & Format(Val(.TextMatrix(i, 5)), "#,##0.00"), 9)
            Next i
        End With
        
        '   0   No
        '   1   Item
        '   2   Batch
        '   3   Rate
        '   4   Amount
        '   5   Price
        '   6   ItemID
        '   7   BatchID
        '   8   AMount
        '   9   Rate
        
        
        With Printer
            
'            Printer.Print
'            Printer.Print
            Printer.Print Tab(Tab1); "---------------------------------------------------------------"
'            Printer.Print
            
            
            .Font.Size = SmallestFontSize
            Tab4 = 20
            
            Printer.Print Tab(Tab1); "Gross Total"; Tab(Tab4); Right((Space(11)) & Format(Val((txtGTotal.Text)), "#,##0.00"), 9)

            If Val(txtDiscount.Text) > 0 Then
                Printer.Print Tab(Tab1); "Discount"; Tab(Tab4); Right((Space(11)) & Format(Val((txtDiscount.Text)), "#,##0.00"), 9)
                .FontSize = SmallestFontSize + 3
                .FontBold = True
                Printer.Print Tab(Tab1); "Net Total"; Tab(Tab4); Right((Space(11)) & Format(Val((txtNTotal.Text)), "#,##0.00"), 9)
                .FontSize = SmallestFontSize
                .FontBold = False
            End If
    
'            Printer.Print Tab(Tab1); "Paid"; Tab(Tab4); Right((Space(10)) & (txtCashPaid.Text), 10)
'            Printer.Print Tab(Tab1); "Balance"; Tab(Tab4); Right((Space(10)) & (txtBalance.Text), 10)
            .FontSize = SmallestFontSize + 2
            Printer.Print Tab(Tab1); "---------------------------------------------------------------"
            Printer.Print
            Printer.Print Tab(Tab1); "Operated By " & UserName
            Printer.Print "Returns are accepted only within 3 days"
            Printer.Print Tab(Tab1); "." ' "---------------------------------------------------------------"
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
'            Printer.Print "."
            
            
            '.FontSize = 10
'            Printer.Print
'            Printer.Print
'            Printer.Print
'            Printer.Print Tab(Tab1); HospitalName
'            .FontSize = 8
'            Printer.Print Tab(Tab1); HospitalDescreption
'            Printer.Print Tab(Tab1); HospitalAddress
'            Printer.Print
            
            
            
            Printer.Print
            .EndDoc
        End With
    
        '   0   No
        '   1   Item
        '   2   Batch
        '   3   Rate
        '   4   Amount
        '   5   Price
        '   6   ItemID
        '   7   BatchID
        '   8   AMount
        '   9   Rate


    End If
    
End Sub

Private Sub SelectPrint()
    If LCase(Left(Trim(HospitalName), 1)) = "m" Then
        MataraPrint
    ElseIf LCase(Left(Trim(HospitalName), 1)) = "r" Then
        RuhunaPrint
    ElseIf LCase(Left(Trim(HospitalName), 1)) = "c" Then
        CooperativePrint
    Else
    
    End If
End Sub

Private Sub RuhunaPrint()
    
    On Error GoTo eh
    
    CsetPrinter.SetPrinterAsDefault (BillPrinterName)

    PrinterName = Printer.DeviceName
    
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    
    
    CsetPrinter.SetPrinterAsDefault (BillPrinterName)
    
        
    Dim MyPrinter As VB.Printer
    For Each MyPrinter In VB.Printers
        If MyPrinter.DeviceName = BillPrinterName Then
            Set Printer = MyPrinter
        End If
    Next


    
If SelectForm(BillPaperName, Me.hwnd) = 1 Then
    Dim i As Integer
    Dim Tab1 As Integer
    Dim Tab2 As Integer
    Dim Tab3 As Integer
    Dim Tab4 As Integer
    Dim Tab5 As Integer
    Dim Tab6 As Integer
    Dim Tab7 As Integer
    Dim Tab8 As Integer
    Dim Tab9 As Integer
    
    Dim MyPageCount As Integer
    Dim MyLineNumber As Integer
    
    
    Tab1 = 4
    Tab2 = 15
    Tab3 = 36
    Tab4 = 20
    Tab5 = 50
    Tab6 = 55
    Tab7 = 70
    Tab8 = 23
    Tab9 = 65
    With Printer
        .TrackDefault = False
        .PaperBin = vbPRBNTractor
        .FontSize = 12
        .Font = "Lucida Console"
        
        MyPageCount = 1
        MyLineNumber = 0
        Printer.CurrentX = 1400 * 0.3
        
        Printer.Print
'        If NewSale.OutPatient = True Then
'            If NewSale.CreditCard = True Then
'                Printer.Print Tab(Tab8 + 10); "Credit Card Invoice"
'            ElseIf NewSale.Cash = True Then
'                Printer.Print Tab(Tab8 + 10); "Cash Invoice"
'            End If
'            Printer.Print
'        End If
        
        Dim temText As String
        
        temText = dtcSale.Text
        If NewSale.Unit = True Then temText = temText & " " & dtcUnit.Text
        
        Printer.Print Tab(Tab8 + 5); temText
        
        
        
        .FontSize = 12
        .Font = "Lucida Console"
        Printer.Print Tab(4); "RUHUNU HOSPITAL (PVT) LTD "
        .FontSize = 10
        .Font = "Lucida Console"

        If NewSale.OutPatient = True Then
            Printer.Print Tab(Tab1); "Karapitiya, Galle." & "Tel: 091-2234059-60, Fax:091-2234061"
        End If
        Printer.Print
        
        .FontSize = 10
        .Font = "Lucida Console"
        
        Dim TemString As String
        If NewSale.OutPatient = True Then
            TemString = "OP"
        ElseIf NewSale.InPatient = True Then
            TemString = "IP"
        ElseIf NewSale.Staff = True Then
            TemString = "SP"
        End If
        Printer.Print Tab(Tab1); "Issue No -    "; TemSaleBillID & "-" & TemString; "       Date : "; Format(Date, "dd MM yy"); Tab(Tab6); "Time : "; Time
        If NewSale.OutPatient = True Then
            Printer.Print Tab(Tab1); "Patient : "; dtcCreditCustomer.Text
        ElseIf NewSale.InPatient = True Then
            Printer.Print Tab(Tab1); "Indoor Patient : "; txtPatient.Text; "         BHT : "; dtcBHT.Text & "  " & lblHealthSchemeSupplier.Caption
        ElseIf NewSale.Staff = True Then
            Printer.Print Tab(Tab1); "Staff member : "; dtcStaffCustomer.Text
        ElseIf NewSale.Unit = True Then
            Printer.Print Tab(Tab1); "Unit         : "; dtcUnit.Text
        End If
        
            Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
        Printer.Print Tab(Tab1); "Item Name"; Tab(Tab3 + 5); "Qty"; Tab(Tab5); Right(Space(12) & "Price", 9); Tab(Tab9); Right(Space(12) & "Value", 13)
            Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
    End With
    Tab1 = 4
    Tab2 = 15
    Tab3 = 36
    Tab4 = 20
    Tab5 = 50
    Tab6 = 55
    Tab7 = 70
    Tab9 = 65
    With GridItem
        For i = 1 To .Rows - 1
'        If MyPageCount = 1 And MyLineNumber > 8 Then
'            'Printer.Print
'            Printer.Print Tab(70); "Page No. " & MyPageCount
'            Printer.NewPage
'            MyPageCount = MyPageCount + 1
'            Printer.CurrentX = 1440 * 0.5
'            MyLineNumber = 1
'        ElseIf MyPageCount > 1 And MyLineNumber > 11 Then
'            'Printer.Print
'            Printer.Print Tab(70); "Page No. " & MyPageCount
'            Printer.NewPage
'            MyPageCount = MyPageCount + 1
'            Printer.CurrentX = 1440 * 0.5
'            MyLineNumber = 1
'        Else
'            MyLineNumber = MyLineNumber + 1
'        End If
                
            Printer.Print Tab(Tab1); Left(.TextMatrix(i, 1), 30);
            Printer.Print Tab(Tab3); Right(Space(10) & (.TextMatrix(i, 8)), 10);
            Printer.Print Tab(Tab5); Right(Space(12) & Format(.TextMatrix(i, 9), "0.00"), 9);
            Printer.Print Tab(Tab7); Right(Space(12) & Format(.TextMatrix(i, 5), "0.00"), 8)
        
        
        Next i
    End With
    With Printer
        Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
        Dim NewTab1 As Integer
        Dim NewTab2 As Integer
        Dim NewTab3 As Integer
        NewTab1 = 40
        NewTab2 = 68
        Printer.Print
        Printer.Print Tab(NewTab1); "Gross Total "; Tab(NewTab2); Right((Space(9)) & Format(Val((txtGTotal.Text)), "#,##0.00"), 10)
        Printer.Print Tab(NewTab1); "Discount    "; Tab(NewTab2); Right((Space(9)) & Format(Val((txtDiscount.Text)), "#,##0.00"), 10)
        
        Printer.Font.Bold = True
        Printer.Print Tab(NewTab1 - 5); "Net Total   "; Tab(NewTab2 - 5); Right((Space(9)) & Format(Val((txtNTotal.Text)), "#,##0.00"), 10)
        Printer.Font.Bold = False
        
        Printer.Print
        Printer.Print
        Printer.Print Tab(Tab1); "Operate by "; UserName  ' ; Tab(Tab5); "Issued by "; dtcIssueStaff
        If NewSale.OutPatient = True Then
            Printer.Print Tab(Tab1); "Returns are acceptted only within 3 days"
        End If
'        Printer.Print vbNewLine
'        Printer.Print vbNewLine
        .EndDoc
    End With
    


End If

Exit Sub

eh:
    MsgBox "Printer Error"

End Sub

Private Sub CooperativePrint()

End Sub

Private Sub MataraPrint()

    Dim i As Integer
    Dim Tab1 As Integer
    Dim Tab2 As Integer
    Dim Tab3 As Integer
    Dim Tab4 As Integer

    Tab1 = 0
    Tab2 = 4
    Tab3 = 28
    Tab4 = 20

    With Printer

        .FontSize = 12
'        .Font = "Arial Black"
                Printer.Print
        Printer.Print Tab(Tab1); "MATARA NURSING HOME (PVT) LTD"
        .FontSize = 10
'        .Font = "Arial Black"
        Printer.Print
        Printer.Print Tab(Tab1); "Anagarika Dharmapala Mawath, Matara"
        Printer.Print Tab(Tab1); "041-2222177, 041-5676265"
        Printer.Print
        Printer.Print Tab(Tab1); "Date : "; Format(Date, "dd MM yy")
        Printer.Print Tab(Tab1); "Time : "; Time; Tab(Tab1 + 25); "Bill No." & TemSaleBillID
        Printer.Print Tab(Tab1); "--------------------------------------"
        If NewSale.OutPatient = True Then
            Printer.Print Tab(Tab1); "Patient : "; dtcCreditCustomer.Text
        ElseIf NewSale.InPatient = True Then
            Printer.Print Tab(Tab1); "Indoor Patient : "; txtPatient.Text
        ElseIf NewSale.Staff = True Then
            Printer.Print Tab(Tab1); "Staff member : "; dtcStaffCustomer.Text
        End If
            Printer.Print Tab(Tab1); "--------------------------------------"
        Printer.Print

        .FontSize = 10
'        .Font = "Lucida Console"
    End With
    With GridItem
        For i = 1 To .Rows - 1
            Printer.Print Tab(Tab1); .TextMatrix(i, 8); Tab(Tab2); Left(.TextMatrix(i, 1), 24); Tab(Tab3); Right((Space(10)) & .TextMatrix(i, 5), 10)
        Next i
    End With
    With Printer
        .Font = 12
        Printer.Print
        Printer.Print
        Printer.Print Tab(Tab1); "--------------------------------------"
        Printer.Print
        Printer.Print Tab(Tab1); "Gross Total"; Tab(Tab4); Right((Space(10)) & (txtGTotal.Text), 10)

        If Val(txtDiscount.Text) > 0 Then
            Printer.Print Tab(Tab1); "Discount"; Tab(Tab4); Right((Space(10)) & (txtDiscount.Text), 10)
            Printer.Print Tab(Tab1); "Net Total"; Tab(Tab4); Right((Space(10)) & (txtNTotal.Text), 10)
        End If

        Printer.Print Tab(Tab1); "Paid"; Tab(Tab4); Right((Space(10)) & (txtCashPaid.Text), 10)
        Printer.Print Tab(Tab1); "Balance"; Tab(Tab4); Right((Space(10)) & (txtBalance.Text), 10)

        Printer.Print Tab(Tab1); "--------------------------------------"
        Printer.Print
        Printer.Print Tab(Tab1); "THANK YOU"
        Printer.Print
        Printer.Print Tab(Tab1); "--------------------------------------"
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        .EndDoc
    End With

    '   0   No
    '   1   Item
    '   2   Batch
    '   3   Rate
    '   4   Amount
    '   5   Price
    '   6   ItemID
    '   7   BatchID
    '   8   AMount
    '   9   Rate

End Sub

Private Sub ClearBillValues()
    Call ClearAddValues
    Call FormatItemGrid
    txtGTotal.Text = "0.00"
    txtNTotal.Text = "0.00"
    txtDiscount.Text = "0.00"
    txtCashPaid.Text = "0.00"
    txtTotalCost.Text = Empty
    dtcBHT.Text = Empty
    dtcCreditCustomer.Text = Empty
    dtcStaffCustomer.Text = Empty
    dtcUnit.Text = Empty
    txtPatient.Text = Empty
    txtBHTBalance.Text = Empty
    lblHealthSchemeSupplier.Caption = Empty
    lblDristributor.Caption = Empty
End Sub

Private Function ConsumeStocks(ByVal IStoreIDValue As Long, ByVal BatchIDValue As Long, ByVal Quentity As Double) As Boolean
    Dim tr As Integer
    On Error GoTo eh
    ConsumeStocks = False
    With rsTemBatch
        If .State = 1 Then .Close
        temSql = "SELECT * from tblBatchstock where batchid = " & BatchIDValue & " AND StoreID = " & IStoreIDValue & " ORDER BY tblBatchstock.Stock DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount < 1 Then
            tr = MsgBox("There is no such drug batch", vbCritical, "Error")
            .Close
            Exit Function
        End If
        If !Stock < Quentity Then
            tr = MsgBox("There are no enough stocks in you store to transfer to another store", vbCritical, "No Enough Stocks")
            .Close
            Exit Function
        End If
        !Stock = !Stock - Quentity
        .Update
        .Close
    ConsumeStocks = True
    Exit Function

eh:
    If .State = 1 Then
        .CancelUpdate
        .Close
    End If
    tr = MsgBox("Could not deduct stocks from your store" & vbNewLine & Err.Description, vbCritical, "Error")
    Exit Function
    End With
End Function


Private Function ReceiveCredit(SaleBillID As Long) As Long
    'New Changes
    Exit Function
    
    
    
    With rsTemCredit
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblReceivedCredit where ReceivedCreditID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = Val(dtcIssueStaff.BoundText)
        !ReceivedDate = Date
        !ReceivedTime = Now
        If NewSale.InPatient = True Then
            !ReceivedFromBHTID = Val(dtcBHT.BoundText)
        ElseIf NewSale.OutPatient = True Then
            !ReceivedFromOutPatientID = Val(dtcCreditCustomer.BoundText)
        ElseIf NewSale.Staff = True Then
            !ReceivedFromStaffID = Val(dtcStaffCustomer.BoundText)
        End If
        !Price = Val(txtNTotal.Text)
        !StoreID = UserStoreID
        !SaleBillID = SaleBillID
        .Update
        .Close
        temSql = "SELECT @@IDENTITY AS NewID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        ReceiveCredit = !NewID
        .Close
    End With
End Function

Private Function ReceiveOther(SaleBillID As Long) As Long
    'New Changes
    Exit Function
    
    With rsTemCredit
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblReceivedOther where ReceivedOtherID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = Val(dtcIssueStaff.BoundText)
        !ReceivedDate = Date
        !ReceivedTime = Now
        If NewSale.InPatient = True Then
            !ReceivedFromBHTID = Val(dtcBHT.BoundText)
        ElseIf NewSale.OutPatient = True Then
            !ReceivedFromOutPatientID = Val(dtcCreditCustomer.BoundText)
        ElseIf NewSale.Staff = True Then
            !ReceivedFromStaffID = Val(dtcStaffCustomer.BoundText)
        ElseIf NewSale.Unit = True Then
            !ReceivedFromUnitID = Val(dtcUnit.BoundText)
        End If
        !Price = Val(txtNTotal.Text)
        !StoreID = UserStoreID
        !SaleBillID = SaleBillID
        .Update
        .Close
        temSql = "SELECT @@IDENTITY AS NewID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        ReceiveOther = !NewID
        .Close
    End With
End Function


Private Function ReceiveCheque(SaleBillID As Long) As Long
    'New Changes
    Exit Function
    
    With rsTemCheque
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblReceivedCheque where ReceivedChequeID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = Val(dtcIssueStaff.BoundText)
        !ReceivedDate = Date
        !ReceivedTime = Now
        !bankID = Val(dtcBank.BoundText)
        If IsNumeric(dtcBranch.BoundText) = True Then
            !BranchID = Val(dtcBranch.BoundText)
        End If
        !ChequeDate = Format(dtpChequeDate.Value, "dd MMMMM yyyy")
        !ChequeNo = txtChequeNo.Text
        If NewSale.InPatient = True Then
            !ReceivedFromBHTID = dtcBHT.BoundText
        ElseIf NewSale.OutPatient = True Then
            !ReceivedFromOutPatientID = Val(dtcCreditCustomer.BoundText)
        ElseIf NewSale.Staff = True Then
            !ReceivedFromStaffID = Val(dtcStaffCustomer.BoundText)
        End If
        !StoreID = UserStoreID
        !Price = Val(txtNTotal.Text)
        !SaleBillID = SaleBillID
        .Update
        .Close
        temSql = "SELECT @@IDENTITY AS NewID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        ReceiveCheque = !NewID
        .Close
    End With
End Function


Private Function ReceiveCash(SaleBillID As Long) As Long
    'New Changes
    Exit Function
    
    With rsTemCash
        If .State = 1 Then .Close
        temSql = "SELECT tblReceivedCash.* FROM tblReceivedCash where ReceivedCashID = 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = Val(dtcIssueStaff.BoundText)
        !ReceivedDate = Date
        !ReceivedTime = Now
        If NewSale.InPatient = True Then
            !ReceivedFromBHTID = Val(dtcBHT.BoundText)
        ElseIf NewSale.OutPatient = True Then
            !ReceivedFromOutPatientID = Val(dtcCreditCustomer.BoundText)
        ElseIf NewSale.Staff = True Then
            !ReceivedFromStaffID = Val(dtcStaffCustomer.BoundText)
        End If
        !Price = Val(txtNTotal.Text)
        !StoreID = UserStoreID
        !SaleBillID = SaleBillID
        .Update
        .Close
        temSql = "SELECT @@IDENTITY AS NewID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        ReceiveCash = !NewID
        
        .Close
    End With
End Function


Private Function ReceiveCreditCard(SaleBillID As Long) As Long
    'New Changes
    Exit Function
    
    With rsTemCC
        If .State = 1 Then .Close
        temSql = "SELECT tblReceivedCreditCard.* FROM tblReceivedCreditCard where ReceivedCreditCard = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !CreditCardNo = Val(txtCreditCardNo.Text)
        !ReceivedSTaffID = Val(dtcIssueStaff.BoundText)
        !CardTypeID = Val(dtcCreditCard.BoundText)
        !AuthrizationCode = Val(txtCreditCode.Text)
        !ReceivedSTaffID = Val(dtcIssueStaff.BoundText)
        !ReceivedDate = Date
        !ReceivedTime = Now
        !AuthrizationDate = Date
        !AuthrizationTime = Now
        !AuthrizationStaffID = dtcIssueStaff.BoundText
        If NewSale.InPatient = True Then
            !ReceivedFromBHTID = dtcBHT.BoundText
        ElseIf NewSale.OutPatient = True Then
            !ReceivedFromOutPatientID = dtcCreditCustomer.BoundText
        ElseIf NewSale.Staff = True Then
            !ReceivedFromStaffID = dtcStaffCustomer.BoundText
        End If
        !Price = Val(txtNTotal.Text)
        !StoreID = UserStoreID
        !SaleBillID = SaleBillID
        .Update
        .Close
        temSql = "SELECT @@IDENTITY AS NewID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        ReceiveCreditCard = !NewID
        .Close
    End With
End Function

Private Function WritePatient() As Long
    Dim temPatient As String
    With rsTemPatient
        If .State = 1 Then .Close
        temSql = "SELECT * from tblpatientmaindetails Where PatientID = 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !FirstName = dtcCreditCustomer.Text
        .Update
        .Close
        temSql = "SELECT @@IDENTITY AS NewID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        WritePatient = !NewID
        .Close
    End With
    With dtcCreditCustomer
        Set .RowSource = Nothing
        .ListField = Empty
        .BoundColumn = Empty
    End With
    With rsPatients
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientMainDetails.* FROM tblPatientMainDetails ORDER BY tblPatientMainDetails.FirstName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCreditCustomer
        Set .RowSource = rsPatients
        .ListField = "FirstName"
        .BoundColumn = "PatientID"
        .BoundText = WritePatient
    End With
End Function

Private Function SaleBillID() As Long
    With rsTemSaleBill
        If .State = 1 Then .Close
        temSql = "SELECT * from tblSaleBill Where SaleBillID = 0"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Date = Date ' dtpDate.Value
        !Time = Now
        !StaffID = Val(dtcIssueStaff.BoundText)
        !StoreID = UserStoreID
        !Price = Val(txtGTotal.Text)
        !Discount = Val(txtDiscount.Text)
        !DiscountPercent = ((Val(txtDiscount.Text)) / (Val(txtGTotal.Text))) * 100
        !NetPrice = Val(txtNTotal.Text)
        !TotalMedicineIncome = Val(txtNTotal.Text)
        !SaleCategoryID = Val(dtcSale.BoundText)
        If IsNumeric(dtcCheckedStaff.BoundText) = True Then !CheckedStaffID = Val(dtcCheckedStaff.BoundText)
        .Update
        .Close
        temSql = "SELECT @@IDENTITY AS NewID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        SaleBillID = !NewID
    End With
End Function

Private Function CanSettle() As Boolean
    Dim tr As Integer
    CanSettle = False
    If GridItem.Rows <= 1 Then
        tr = MsgBox("There are no items to sell", vbCritical, "No Items")
        dtcItem.SetFocus
        Exit Function
    End If
    If IsNumeric(dtcSale.BoundText) = False Then
        tr = MsgBox("You have not selected the payment method", vbCritical, "No Items")
        SSTab2.Tab = 0
        dtcSale.SetFocus
        Exit Function
    End If
    
    If NewSale.Cash = True Then
        If UCase(Left(HospitalName, 1)) = "R" Then
                txtCashPaid.Text = Val(txtCashPaid.Text)
        Else
            If IsNumeric(txtCashPaid.Text) = False Then
                tr = MsgBox("You have not entered a valied cash amount", vbCritical, "Cash?")
                SSTab2.Tab = 1
                txtCashPaid.SetFocus
                Exit Function
            End If
            If Val(txtCashPaid.Text) < Val(txtDue.Text) Then
                tr = MsgBox("The amount you pay is not sufficient", vbCritical, "Not sufficient cash")
                SSTab2.Tab = 1
                txtCashPaid.SetFocus
                Exit Function
            End If
        End If
    ElseIf NewSale.Credit = True Then
    
    ElseIf NewSale.Cheque = True Then
        If IsNumeric(dtcBank.BoundText) = False Then
            tr = MsgBox("You have not selected a Bank", vbCritical, "Bank?")
            SSTab2.Tab = 1
            dtcBank.SetFocus
            Exit Function
        End If
        If Trim(txtChequeNo.Text) = "" Then
            tr = MsgBox("You have not entered the cheque number", vbCritical, "Cheque Number?")
            SSTab2.Tab = 1
            txtChequeNo.SetFocus
            Exit Function
        End If
    ElseIf NewSale.CreditCard = True Then
'        If IsNumeric(dtcCreditCard.BoundText) = False Then
'            tr = MsgBox("You have not selected the Credit Card Type", vbCritical, "Card type?")
'            SSTab2.Tab = 1
'            dtcCreditCard.SetFocus
'            Exit Function
'        End If
'        If IsNumeric(dtcCardBank.BoundText) = False Then
'            tr = MsgBox("You have not selected the cadit card issued bank", vbCritical, "Bank?")
'            SSTab2.Tab = 1
'            dtcCardBank.SetFocus
'            Exit Function
'        End If
'        If Trim(txtCreditCardNo.Text) = "" Then
'            tr = MsgBox("You have not entered a valied credit card number", vbCritical, "Card Number?")
'            SSTab2.Tab = 1
'            txtCreditCardNo.SetFocus
'            Exit Function
'        End If
'        If Trim(txtCreditCode.Text) = "" Or IsNumeric(txtCreditCode.Text) = False Then
'            tr = MsgBox("You have not entered a valied autherization code", vbCritical, "Authorization code?")
'            SSTab2.Tab = 1
'            txtCreditCode.SetFocus
'            Exit Function
'        End If
    End If
    
    If Trim(dtcBHT.Text) <> "" Then
        If IsNumeric(dtcBHT.BoundText) = True Then
            If dtcBHT.Text = dtcBHT.BoundText Then
                tr = MsgBox("You have not selected the BHT number", vbCritical, "BHT?")
                SSTab2.Tab = 1
                dtcBHT.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If NewSale.InPatient = True Then
        If IsNumeric(dtcBHT.BoundText) = False Then
            tr = MsgBox("You have not selected the BHT number", vbCritical, "BHT?")
            SSTab2.Tab = 1
            dtcBHT.SetFocus
            Exit Function
        End If
            
    
    ElseIf NewSale.OutPatient = True Then
    
    ElseIf NewSale.Staff = True Then
        If IsNumeric(dtcStaffCustomer.BoundText) = False Then
            tr = MsgBox("You have not selected the staff member to whom the items are issued", vbCritical, "Staff member?")
            SSTab2.Tab = 1
            dtcStaffCustomer.SetFocus
            Exit Function
        End If
    ElseIf NewSale.Unit = True Then
        If IsNumeric(dtcUnit.BoundText) = False Then
            tr = MsgBox("You have not selected the unit")
            SSTab2.Tab = 1
            dtcUnit.SetFocus
            Exit Function
        End If
    End If
    
    CanSettle = True
End Function

Private Sub Command1_Click()
    Call POSPrint
End Sub

Private Sub dtcBHT_Click(Area As Integer)
    Dim TemBHTCredit As Double
    Dim temPatientID As Long
    Dim HSSID As Long
    On Error Resume Next
    If IsNumeric(dtcBHT.BoundText) = False Then Exit Sub
    lblHealthSchemeSupplier.Caption = Empty
    With rsTemStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblBHT where BHTID = " & Val(dtcBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!Balance) Then
                TemBHTCredit = !Balance
            Else
                TemBHTCredit = 0
            End If
            txtTemCreditCustomerBalance.Text = TemBHTCredit
            If TemBHTCredit < 0 Then
                txtBHTBalance.Text = "(" & Format(Abs(TemBHTCredit), "#,##0.00") & ")"
            Else
                txtBHTBalance.Text = Format(TemBHTCredit, "#,##0.00")
            End If
            If Not IsNull(!HealthSchemeSupplierID) Then
                HSSID = !HealthSchemeSupplierID
            Else
                HSSID = 0
            End If
            temPatientID = !PatientID
        End If
        If .State = 1 Then .Close
    End With
    
    If HSSID <> 0 Then
        With rsTemStaff
            If .State = 1 Then .Close
            temSql = "Select * from tblHealthSchemeSuppliers where HealthSchemeSupplierID = " & HSSID
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                If Not IsNull(!HealthSchemeSupplierName) Then
                    lblHealthSchemeSupplier.Caption = !HealthSchemeSupplierName
                Else
                    lblHealthSchemeSupplier.Caption = Empty
                End If
            Else
                lblHealthSchemeSupplier.Caption = Empty
            End If
        End With
    End If
    
    With rsTemPatient
        If .State = 1 Then .Close
        temSql = "SELECT * from tblPatientMainDetails where PatientID = " & temPatientID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtPatient.Text = !FirstName
        End If
        .Close
    End With
    With rsTemPatient
        If .State = 1 Then .Close
        temSql = "SELECT * from tblHealthSchemeSuppliers where HealthSchemeSupplierID = " & HSSID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            
        If .RecordCount > 0 Then
            If Not IsNull(!ProfitMargin) Then
                txtBHTProfit.Text = !ProfitMargin
            Else
                txtBHTProfit.Text = 0
            End If
        Else
            txtBHTProfit.Text = 0
        End If
        .Close
    End With
    ChangeGridRateValues
    CalculateTotal
'    ClearAddValues
'    FormatSelectStock
    CalculateDiscount

End Sub

Private Sub dtcBHT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        bttnSettle_Click
        KeyCode = Empty
    End If
End Sub

Private Sub dtcCatogery_LostFocus()
    If IsNumeric(dtcCatogery.BoundText) Then
        ListSelectedItems
    Else
        ListAllItems
    End If
    dtcItem.Text = Empty
    dtcCode.Text = Empty
'    Dim rsIC As New ADODB.Recordset
'    With rsIC
'        If .State = 1 Then .Close
'        temSql = "Select * from tblItemCategory where ItemCategoryID = " & Val(dtcCatogery.BoundText)
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            lblCategory.Caption = !ItemCategory
'        End If
'        .Close
'    End With
End Sub

Private Sub dtcCode_Change()
    dtcItem.BoundText = dtcCode.BoundText
End Sub


Private Sub dtcCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtcCode.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtQty.SetFocus
    End If
End Sub


Private Sub dtcCode_LostFocus()
    dtcItem_LostFocus
End Sub

Private Sub dtcCreditCustomer_Click(Area As Integer)
    Dim TemCreditCustomerCredit As Double
    If IsNumeric(dtcCreditCustomer.BoundText) = False Then Exit Sub
    With rsTemStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblpatientmaindetails where patientID = " & Val(dtcCreditCustomer.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!Credit) Then
                TemCreditCustomerCredit = !Credit
            Else
                TemCreditCustomerCredit = 0
            End If
            txtTemCreditCustomerBalance.Text = TemCreditCustomerCredit
            If TemCreditCustomerCredit < 0 Then
                txtCreditCustomerBalance.Text = "(" & Format(Abs(TemCreditCustomerCredit), "#,##0.00") & ")"
            Else
                txtCreditCustomerBalance.Text = Format(TemCreditCustomerCredit, "#,##0.00")
            End If
        End If
        If .State = 1 Then .Close
    End With
End Sub


Private Sub dtcCreditCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        bttnSettle_Click
    End If
End Sub

Private Sub dtcItem_Change()
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    FillTotalStock (Val(dtcItem.BoundText))
'    Dim tr As Integer
'    dtcCode.BoundText = dtcItem.BoundText
'    NewItem.ID = dtcItem.BoundText
'    txtCategoryProfit.Text = NewItem.SalesMargin
'    Call FillAddPrice(dtcItem.BoundText)
'    lblIUnit.Caption = NewItem.IUnit
'    Call CalculatePrice
'    Call FillSelectStock(dtcItem.BoundText)
'    DistributorDetails (Val(dtcItem.BoundText))
End Sub

Private Sub SelectCatogery()
    Dim rsTemItem As New ADODB.Recordset
    Dim temID As Long
    temID = dtcItem.BoundText
    With rsTemItem
        If .State = 1 Then .Close
        temSql = "SELECT * from tblItem where ItemID = " & temID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            dtcCatogery.BoundText = !ItemCategoryID
        End If
        .Close
    End With
    dtcItem.BoundText = temID
End Sub

Private Sub FillAddPrice(ByVal ItemID As Long)
    txtRate.Text = Empty
    txtCostRate.Text = Empty
    With rsTemPrice
        If .State = 1 Then .Close
        temSql = "SELECT tblCurrentSalePrice.SPrice FROM tblCurrentSalePrice WHERE tblCurrentSalePrice.ItemID=" & ItemID & " Order By SetDate Desc, SetTime DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtSPrice.Text = !SPrice
            If NewSale.Unit = True Then
                    txtRate.Text = Format(rsTemPrice!SPrice, "0.00")
            Else
                If Val(txtSaleProfit.Text) = 0 And Val(txtBHTProfit.Text) = 0 Then
                    txtRate.Text = Format(rsTemPrice!SPrice, "0.00")
                Else
                    txtRate.Text = Format((!SPrice / (Val(txtCategoryProfit.Text) + 100)) * (Val(txtCategoryProfit.Text) + Val(txtSaleProfit.Text) + Val(txtBHTProfit.Text) + 100), "0.00")
                End If
            End If
        End If
    End With
    With rsTemPrice
        If .State = 1 Then .Close
        temSql = "SELECT tblCurrentPurchasePrice.PPrice FROM tblCurrentPurchasePrice WHERE tblCurrentPurchasePrice.ItemID=" & ItemID & " Order By SetDate Desc, SetTime DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtCostRate.Text = Format(rsTemPrice!PPrice, "0.00")
            If NewSale.Unit = True Then
                txtRate.Text = Format(rsTemPrice!PPrice, "##00.00")
            Else
            
            End If
        End If
    End With
End Sub

Private Sub FormatSelectStock()
    With GridBatch
        .ScrollBars = flexScrollBarVertical
        .Clear
        .Cols = 6
        .Rows = 1
        .Row = 0
        .FixedCols = 0
        .Col = 0
        .CellAlignment = 4
        .Text = "Batch"
        .Col = 1
        .CellAlignment = 4
        .Text = "Stock (" & NewItem.IUnit & ")"
        .Col = 2
        .CellAlignment = 4
        .Text = "Expiary"
        .Col = 3
        .CellAlignment = 4
        .Text = "Location"
        .ColWidth(1) = 1600
        .ColWidth(2) = 1600
        .ColWidth(3) = 1600
        .ColWidth(4) = 1
        .ColWidth(5) = 1
        .ColWidth(0) = .Width - (.ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 100)
    End With
End Sub


Private Sub FillTotalStock(ByVal ItemID As Long)
    Dim TotalStock As Double
    With rsTemStore
        If .State = 1 Then .Close
        temSql = "SELECT sum(tblBatchStock.Stock) as SumOfStock " & _
                    " FROM dbo.tblBatchStock LEFT OUTER JOIN dbo.tblBatch ON dbo.tblBatchStock.BatchID = dbo.tblBatch.BatchID " & _
                    " WHERE tblBatch.ItemID=" & ItemID & " AND tblBatchStock.StoreID=" & UserStoreID & " AND tblBatchStock.Stock > 0 "
                    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
                If Not IsNull(!SumOfStock) Then
                    TotalStock = !SumOfStock
                Else
                    TotalStock = Empty
                End If
        End If
        .Close
    End With
    txtTotalStock.Text = TotalStock
End Sub


Private Sub FillSelectStock(ByVal ItemID As Long)
    Dim TotalStock As Double
    With GridBatch
        .Visible = False
        FormatSelectStock
    End With
    With rsTemStore
        If .State = 1 Then .Close
        temSql = "SELECT tblBatch.*,  tblBatchStock.*, tblLocation.Location " & _
                    " FROM (tblStore RIGHT JOIN (tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) ON tblStore.StoreID = tblBatchStock.StoreID) LEFT JOIN tblLocation ON tblBatchStock.LocationID = tblLocation.LocationID " & _
                    " WHERE tblBatch.ItemID=" & ItemID & " AND tblBatchStock.StoreID=" & UserStoreID & " AND tblBatchStock.Stock > 0 " & _
                    "ORDER BY tblBatch.DOE" 'DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                GridBatch.Rows = GridBatch.Rows + 1
                GridBatch.Row = GridBatch.Rows - 1
                GridBatch.Col = 0
                GridBatch.CellAlignment = 1
                GridBatch.Text = !Batch
                GridBatch.Col = 1
                GridBatch.CellAlignment = 7
                If Not IsNull(!Stock) Then
                    GridBatch.Text = !Stock
                    TotalStock = TotalStock + !Stock
                Else
                    GridBatch.Text = 0
                End If
                GridBatch.Col = 2
                GridBatch.CellAlignment = 1
                GridBatch.Text = Format(!DOE, ShortDateFormat)
                GridBatch.Col = 3
                GridBatch.CellAlignment = 1
                If Not IsNull(!Location) Then
                    GridBatch.Text = !Location
                Else
                    GridBatch.Text = Empty
                End If
                GridBatch.Col = 4
                GridBatch.Text = ![BatchID]
                GridBatch.Col = 5
                GridBatch.Text = ![BatchID]
                
                .MoveNext
            Wend
            GridBatch.Visible = True
            GridBatch.Row = 1
            GridBatch.Col = GridBatch.Cols - 1
            GridBatch.ColSel = 0
        End If
        If GridBatch.Visible = False Then GridBatch.Visible = True
        .Close
    End With
    txtTotalStock.Text = TotalStock
End Sub


Private Sub dtcItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtQty.SetFocus
        SendKeys "{Home}+{end}"
        KeyCode = Empty
    ElseIf KeyCode = vbKeyF2 Then KeyCode = Empty
        frmSelectGeneric.cmbItem.BoundText = Val(dtcItem.BoundText)
        frmSelectGeneric.Show 1
    End If
End Sub

Private Sub dtcItem_LostFocus()
'    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
'    Dim tr As Integer
'    If dtcCatogery.Text = Empty Then SelectCatogery
'    If CalculateStock(dtcItem.BoundText, , UserStoreID).Amount <= 0 Then
'        tr = MsgBox("There are no stocks", vbCritical, "No Stocks")
'        dtcCatogery.Text = Empty
'        dtcItem.SetFocus
'        Exit Sub
'    End If



    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    Dim tr As Integer
    dtcCode.BoundText = dtcItem.BoundText
    NewItem.ID = dtcItem.BoundText
    txtCategoryProfit.Text = NewItem.SalesMargin
    Call FillAddPrice(dtcItem.BoundText)
    lblIUnit.Caption = NewItem.IUnit
    Call CalculatePrice
    Call FillSelectStock(dtcItem.BoundText)
    DistributorDetails (Val(dtcItem.BoundText))



End Sub

Private Sub ChangeGridRateValues()
'   0   No
'   1   Item
'   2   Batch
'   3   Rate
'   4   Amount
'   5   Price
'   6   ItemID
'   7   BatchID
'   8   AMount
'   9   Rate
'   10  Cost
'   11  Category
'   12  CatogoryID
'   13  IUnit
'   14  CategoryProfit
'   15  SaleProfit
'   16  BHTProfit
'   17  Real Price
    Dim i As Integer
    If NewSale.Unit = True Then
        With GridItem
            If .Rows < 2 Then Exit Sub
            For i = 1 To .Rows - 1
                NewItem.ID = Val(.TextMatrix(i, 6))
                .TextMatrix(i, 9) = NewItem.PPrice
                .TextMatrix(i, 5) = Format(Val(.TextMatrix(i, 8)) * NewItem.PPrice, "0.00")
                .TextMatrix(i, 3) = NewItem.PPrice & " Per " & NewItem.IUnit
            Next i
        End With
    Else
        With GridItem
            If .Rows < 2 Then Exit Sub
            For i = 1 To .Rows - 1
                NewItem.ID = Val(.TextMatrix(i, 6))
                If Val(txtSaleProfit.Text) = 0 And Val(txtBHTProfit.Text) = 0 Then
                    .TextMatrix(i, 9) = .TextMatrix(i, 17)
                Else
                    .TextMatrix(i, 9) = (Val(.TextMatrix(i, 17)) / (Val(.TextMatrix(i, 14)) + 100)) * (100 + Val(txtSaleProfit.Text) + Val(.TextMatrix(i, 14)) + Val(txtBHTProfit.Text))
                End If
                .TextMatrix(i, 5) = Format(((Val(.TextMatrix(i, 8))) * (Val(.TextMatrix(i, 9)))), "0.00")
                .TextMatrix(i, 3) = Format(.TextMatrix(i, 9), "0.00") & " Per " & NewItem.IUnit
            Next i
        End With
    End If
    CalculateTotal
    ClearAddValues
    FormatSelectStock
    CalculateDiscount

End Sub

Private Sub dtcSale_Change()
    txtCostRate.Visible = False
    txtRate.Visible = True
    If IsNumeric(dtcSale.BoundText) = False Then Exit Sub
    NewSale.SaleCategoryID = Val(dtcSale.BoundText)
    txtSaleProfit.Text = NewSale.ProfitMargin
    If NewSale.InPatient = False Then txtBHTProfit.Text = 0
    
    If IsNumeric(dtcItem.BoundText) = True Then
        Call FillAddPrice(dtcItem.BoundText)
        Call CalculatePrice
    End If
    
    Call ChangeGridRateValues
    
    lblHealthSchemeSupplier.Caption = Empty
    
    If NewSale.Cash = True Then
        frameCash.Visible = True
        frameCredit.Visible = False
        frameCreditCard.Visible = False
        frameCheque.Visible = False
        lblDisplayTotal.Caption = "Cash sale"
        txtDue.Text = txtNTotal.Text
    ElseIf NewSale.Credit = True Then
        frameCash.Visible = False
        frameCredit.Visible = True
        frameCreditCard.Visible = False
        frameCheque.Visible = False
        lblDisplayTotal.Caption = "Credit sale"
    ElseIf NewSale.Cheque = True Then
        frameCash.Visible = False
        frameCredit.Visible = False
        frameCreditCard.Visible = False
        frameCheque.Visible = True
        lblDisplayTotal.Caption = "Cheque sale"
    ElseIf NewSale.CreditCard = True Then
        frameCash.Visible = False
        frameCredit.Visible = False
        frameCreditCard.Visible = True
        frameCheque.Visible = False
        lblDisplayTotal.Caption = "Credit Card sale"
    ElseIf NewSale.Other = True Then
        frameCash.Visible = False
        frameCredit.Visible = False
        frameCreditCard.Visible = False
        frameCheque.Visible = False
        lblDisplayTotal.Caption = Empty
    End If
    If NewSale.InPatient = True Then
        frameInPatient.Visible = True
        frameOutPatient.Visible = False
        frameStaff.Visible = False
        frameUnit.Visible = False
        lblDisplayTotal.Caption = lblDisplayTotal.Caption & " for In-Hospital Patients"
    ElseIf NewSale.OutPatient = True Then
        frameInPatient.Visible = False
        frameOutPatient.Visible = True
        frameStaff.Visible = False
        frameUnit.Visible = False
        lblDisplayTotal.Caption = lblDisplayTotal.Caption & " for Out-Hospital Patients"
    ElseIf NewSale.Staff = True Then
        frameInPatient.Visible = False
        frameOutPatient.Visible = False
        frameStaff.Visible = True
        frameUnit.Visible = False
        lblDisplayTotal.Caption = lblDisplayTotal.Caption & " for staff members"
    ElseIf NewSale.Unit = True Then
        frameInPatient.Visible = False
        frameOutPatient.Visible = False
        frameStaff.Visible = False
        frameUnit.Visible = True
        lblDisplayTotal.Caption = lblDisplayTotal.Caption & " for Units"
        txtCostRate.Visible = True
        txtRate.Visible = False
    End If
'    SSTab2.Tab = 1
    Call CalculateDiscount
    lblDisplayTotal.Caption = lblDisplayTotal.Caption & " - Rs. " & txtNTotal.Text
End Sub

Private Sub CalculateDiscount()
    txtDiscount.Text = Format(Val(txtGTotal.Text) * (NewSale.SaleDiscountPercent / 100), "0.00")
End Sub

Private Sub dtcSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SSTab2.Tab = 1
        KeyCode = Empty
        If NewSale.Cash = True Then
            txtCashPaid.SetFocus
        ElseIf NewSale.Credit = True Then
            If NewSale.InPatient = True Then
                dtcBHT.SetFocus
            ElseIf NewSale.OutPatient = True Then
                dtcCreditCustomer.SetFocus
            ElseIf NewSale.Staff = True Then
                dtcStaffCustomer.SetFocus
            End If
        ElseIf NewSale.Cheque = True Then
            dtcBank.SetFocus
        ElseIf NewSale.CreditCard = True Then
            dtcCreditCard.SetFocus
        End If
    End If
End Sub

Private Sub dtcStaffCustomer_Change()
    Dim TemStaffCredit As Double
    If IsNumeric(dtcStaffCustomer.BoundText) = False Then Exit Sub
    With rsTemStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblSTaff where staffid = " & Val(dtcStaffCustomer.BoundText)
        
        temSql = "SELECT sum(tblSaleBill.NetPrice-tblSaleBill.ReturnedValue)as AnnualValue " & _
                    "FROM tblReturnBill RIGHT JOIN (tblSaleBill LEFT JOIN tblStaff ON tblSaleBill.BilledStaffID = tblStaff.StaffID) ON tblReturnBill.SaleBillID = tblSaleBill.SaleBillID " & _
                        "WHERE (((tblStaff.StaffID) = " & dtcStaffCustomer.BoundText & ") AND ((tblSaleBill.Date) Between '" & Format(DateSerial(Year(Date), 1, 1), "dd MMMM yyyy") & "' AND '" & Format(Date, "dd MMMM yyyy") & "')  AND ((tblSaleBill.Cancelled)=0) )"

        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!AnnualValue) Then
                TemStaffCredit = !AnnualValue
            Else
                TemStaffCredit = 0
            End If
'            txtTemStaffCredit.Text = TemStaffCredit
'            If TemStaffCredit < 0 Then
'                txtStaffBalance.Text = "(" & Format(Abs(TemStaffCredit), "#,##0.00") & ")"
'            Else
                txtStaffBalance.Text = Format(TemStaffCredit, "#,##0.00")
'            End If
        End If
        If .State = 1 Then .Close
    End With
End Sub

Private Sub dtcStaffCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        bttnSettle_Click
    End If
End Sub

Private Sub dtcUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        bttnSettle_Click
    ElseIf KeyCode = vbKeyEscape Then
        dtcUnit.Text = Empty
    End If
End Sub

Private Sub Form_Activate()
    Me.WindowState = 2
    With rsBHT
        If .State = 1 Then .Close
        temSql = "SELECT tblBHT.* FROM tblBHT WHERE (((tblBHT.Discharge)=0)) ORDER BY tblBHT.BHT"
'        temSql = "SELECT tblBHT.* FROM tblBHT ORDER BY tblBHT.BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        CsetPrinter.SetPrinterAsDefault (PrescreptionPrinterName)
    
        PrinterName = Printer.DeviceName
        
        If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
            ClosePrinter (PrinterHandle)
        End If
        
        
        CsetPrinter.SetPrinterAsDefault (PrescreptionPrinterName)
        
            
        Dim MyPrinter As VB.Printer
        For Each MyPrinter In VB.Printers
            If MyPrinter.DeviceName = (PrescreptionPrinterName) Then
                Set Printer = MyPrinter
            End If
        Next
        
        If SelectForm(PrescreptionPaperName, Me.hwnd) = 1 Then
            Printer.Print
            Printer.EndDoc
        End If
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatItemGrid
    dtcIssueStaff.BoundText = UserID
    dtcIssueStaff.Locked = True
    SSTab2.Tab = 0
    dtcSale.BoundText = Val(GetSetting(App.EXEName, "Options", "SaleCategoryID", 0))
    dtpDate.Value = Date
'    Call FillPrinters
'    On Error Resume Next
'    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
End Sub

Private Sub FormatBatchGrid()
    With GridBatch
        .Cols = 4
        .Rows = 1
        Dim i As Integer
        For i = 0 To .Cols - 1
            .CellAlignment = 4
            Select Case i
                Case 0:
                        .Text = "Batch"
                        .ColWidth(i) = 1500
                Case 1:
                        .Text = "Expiary"
                        .ColWidth(i) = 2500
                Case 2:
                        .ColWidth(i) = 2500
                        .Text = "Location"
                Case Else
                        .ColWidth(i) = 1
                
            End Select
    '   1   Batch
    '   2   Expiary
    '   3   Location
    '   4   BatchID
        Next
    End With
End Sub

Private Sub FormatItemGrid()
    With GridItem
        .Cols = 18
        .Rows = 1
        Dim i As Integer
        For i = 0 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            Select Case i
                Case 0: .Text = "No."
                        .ColWidth(i) = 500
                Case 1: .Text = "Item"
                        .ColWidth(i) = 5700
                Case 2: .Text = "Batch"
                        .ColWidth(i) = 1900
                Case 3: .Text = "Rate"
                        .ColWidth(i) = 2400
                Case 4: .Text = "Amount"
                        .ColWidth(i) = 2200
                Case 5: .Text = "Price"
                        .ColWidth(i) = 2100
                Case Else
                        .ColWidth(i) = 1
            End Select
        Next
        LastVisibleRow = 0
'   0   No
'   1   Item
'   2   Batch
'   3   Rate
'   4   Amount
'   5   Price
'   6   ItemID
'   7   BatchID
'   8   AMount
'   9   Rate
'   10  Cost
'   11  Category
'   12  CatogoryID
'   13  IUnit

'   14  CategoryProfit
'   15  SaleProfit
'   16  BHTProfit


    End With
End Sub

Public Sub FillCombos()
    With rsSale
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleCategory.SaleCategoryID, tblSaleCategory.SaleCategory FROM tblSaleCategory ORDER BY tblSaleCategory.SaleCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcSale
        Set .RowSource = rsSale
        .ListField = "SaleCategory"
        .BoundColumn = "SaleCategoryID"
    End With
    With rsItem
        If .State = 1 Then .Close
        temSql = "SELECT * from tblitem order by display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcItem
        Set .RowSource = rsItem
        .ListField = "display"
        .BoundColumn = "ItemID"
    End With
    With rsItemCategory
        If .State = 1 Then .Close
        temSql = "SELECT * from tblItemCategory order by categoryCode"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCatogery
        Set .RowSource = rsItemCategory
        .ListField = "CategoryCode"
        .BoundColumn = "ItemCategoryID"
    End With
    With rsCode
        If .State = 1 Then .Close
        temSql = "SELECT * from tblitem order by code"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCode
        Set .RowSource = rsCode
        .ListField = "code"
        .BoundColumn = "ItemID"
    End With
    With rsStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblstaff order by listedname"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcIssueStaff
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With dtcCheckedStaff
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With dtcStaffCustomer
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With rsBanks
        If .State = 1 Then .Close
        temSql = "SELECT tblBank.* FROM tblBank ORDER BY tblBank.Bank"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCardBank
        Set .RowSource = rsBanks
        .ListField = "Bank"
        .BoundColumn = "BankID"
    End With
    With dtcBank
        Set .RowSource = rsBanks
        .ListField = "Bank"
        .BoundColumn = "BankID"
    End With
    With rsCreditCards
        If .State = 1 Then .Close
        temSql = "SELECT tblCreditCardType.CreditCardTypeID, tblCreditCardType.CreditCardType FROM tblCreditCardType ORDER BY tblCreditCardType.CreditCardType"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCreditCard
        Set .RowSource = rsCreditCards
        .ListField = "CreditCardType"
        .BoundColumn = "CreditCardTypeID"
    End With
    With rsCities
        If .State = 1 Then .Close
        temSql = "SELECT tblCity.CityId, tblCity.City FROM tblCity ORDER BY tblCity.City"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcBranch
        Set .RowSource = rsCities
        .ListField = "City"
        .BoundColumn = "CityId"
    End With
    With rsBHT
        If .State = 1 Then .Close
        temSql = "SELECT tblBHT.* FROM tblBHT WHERE (((tblBHT.Discharge)=0)) ORDER BY tblBHT.BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
    With rsPatients
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientMainDetails.* FROM tblPatientMainDetails ORDER BY tblPatientMainDetails.FirstName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCreditCustomer
        Set .RowSource = rsPatients
        .ListField = "FirstName"
        .BoundColumn = "PatientID"
    End With
    With rsStore
        If .State = 1 Then .Close
        temSql = "SELECT * from tblStore order by store"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcDepartment
        Set .RowSource = rsStore
        .ListField = "Store"
        .BoundColumn = "StoreID"
    End With
    With rsUnit
        If .State = 1 Then .Close
        temSql = "SELECT * from tblStore order by Store"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcUnit
        Set .RowSource = rsUnit
        .ListField = "Store"
        .BoundColumn = "StoreID"
    End With
End Sub

Private Sub dtcCatogery_Change()
    lblCategory.Caption = ""
    Dim rsIC As New ADODB.Recordset
    With rsIC
        If .State = 1 Then .Close
        temSql = "Select * from tblItemCategory where ItemCategoryID = " & Val(dtcCatogery.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            lblCategory.Caption = !ItemCategory
        End If
        .Close
    End With

'    If IsNumeric(dtcCatogery.BoundText) Then
'        ListSelectedItems
'    Else
'        ListAllItems
'    End If
'    dtcItem.Text = Empty
'    dtcCode.Text = Empty
End Sub


Private Sub ListSelectedItems()
With rsItem
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by display"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "Display"
    .BoundColumn = "ItemID"
End With
With rsCode
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by code"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcCode
    Set .RowSource = rsCode
    .ListField = "Code"
    .BoundColumn = "ItemID"
End With

End Sub

Private Sub ListAllItems()
With rsItem
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem order by display"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "display"
    .BoundColumn = "ItemID"
End With
With rsCode
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem order by code"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcCode
    Set .RowSource = rsCode
    .ListField = "Code"
    .BoundColumn = "ItemID"
End With
End Sub

Private Sub dtcCatogery_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtcCatogery.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcItem.SetFocus
    End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    SaveSetting App.EXEName, Me.Name, "printer", cmbPrinter.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "Options", "SaleCategoryID", dtcSale.BoundText
End Sub

Private Sub GridItem_Click()
    With GridItem
        If .Rows <= 1 Then Exit Sub
        bttnDelete.Enabled = True
        .Col = .Cols - 1
        .ColSel = 0
    End With
End Sub

Private Sub GridItem_DblClick()
    With GridItem
        If .Rows <= 1 Then Exit Sub
        bttnDelete.Enabled = True
        dtcCatogery.Text = Empty
        dtcItem.Text = Empty
        .Col = 6
        dtcItem.BoundText = Val(.Text)
        .Col = 8
        txtQty.Text = Val(.Text)
        .Col = 9
        txtRate.Text = Val(.Text)
        bttnDelete_Click
    End With
    dtcItem.SetFocus
    dtcCode.SetFocus
End Sub

Private Sub txtCashPaid_Change()
    txtBalance.Text = Format((Val(txtCashPaid.Text) - Val(txtDue.Text)), "0.00")
End Sub

Private Sub txtCashPaid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        bttnSettle_Click
    End If
End Sub

Private Sub txtCashPaid_LostFocus()
    txtCashPaid.Text = Format(Val(txtCashPaid.Text), "0.00")
End Sub

Private Sub txtDiscount_Change()
    Call CalculateNetTotal
End Sub

Private Sub txtDue_Change()
    txtBalance.Text = Format((Val(txtCashPaid.Text) - Val(txtDue.Text)), "0.00")
End Sub

Private Sub txtGTotal_Change()
    Call CalculateNetTotal
End Sub

Private Sub txtNTotal_Change()
    txtDue.Text = txtNTotal.Text
    txtCreditDue.Text = txtNTotal.Text
End Sub


Private Sub txtPatient_LostFocus()
    txtPatient.Text = UCase(txtPatient.Text)
End Sub

Private Sub txtQty_Change()
    Call CalculatePrice
End Sub

Private Sub CalculatePrice()
    If NewSale.Unit = True Then
        txtPrice.Text = Format(Val(txtRate.Text) * Val(txtQty.Text), "0.00")
        txtItemCost.Text = Format(Val(txtCostRate.Text) * Val(txtQty.Text), "0.00")
    Else
        txtPrice.Text = Format((Val(txtQty.Text) * Val(txtRate.Text)), "0.00")
        txtItemCost.Text = Format(Val(txtCostRate.Text) * Val(txtQty.Text), "0.00")
    End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        bttnAdd_Click
    End If
End Sub

Private Function QtyOK() As Boolean
    QtyOK = False
    If Not IsNumeric(dtcItem.BoundText) Then Exit Function
    Dim tr As Integer
    Dim temStock As Double
    If dtcCatogery.Text = Empty Then SelectCatogery
    temStock = CalculateStock(dtcItem.BoundText, Val(GridBatch.TextMatrix(GridBatch.Row, 4)), UserStoreID).Amount
    If temStock < Val(txtQty.Text) Then
        tr = MsgBox("There are no Adequate stock. Available quentity is selected", vbCritical, "No Adequate Stocks")
        txtQty.Text = temStock
        txtQty.SetFocus
        SendKeys "{home}+{end}"
        Exit Function
    End If
    QtyOK = True
End Function

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii = vbKeyReturn Then Exit Sub
'    If KeyAscii = vbKeyDelete Then Exit Sub
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = Empty
    End If
End Sub


Private Sub DistributorDetails(ItemID As Long)
    With rsDI
        If .State = 1 Then .Close
        temSql = "SELECT tblItemDistributor.DistributorID FROM tblItemDistributor WHERE (((tblItemDistributor.ItemID)=" & (Val(dtcItem.BoundText)) & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
        TemDI = !DistributorID
        End If
        .Close
    End With
    With rsTemDistributor1
        If .State = 1 Then .Close
        temSql = "SELECT tblDistrubutor.*, tblCity.City FROM tblCity RIGHT JOIN tblDistrubutor ON tblCity.CityId = tblDistrubutor.DistributorCityID Where DistributorId = " & TemDI & ""
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!DistributorName) Then lblDristributor.Caption = !DistributorName
        If .State = 1 Then .Close
    End With
End Sub

'
'Private Sub FillPrinters()
'    Dim MyPrinter As Printer
'    For Each MyPrinter In Printer
'    cmbPrinter.AddItem MyPrinter.DeviceName
'    Next
'End Sub
'
''Dim MyPrinter As Printer
''For Each MyPrinter In Printer
''If MyPrinter.DeviceName = cmbprinter.Text Then
''Set MyPrinter = Printer
''End If
''Next

Public Sub FillPatientCombo()
    With rsBHT
        If .State = 1 Then .Close
        temSql = "SELECT tblBHT.* FROM tblBHT WHERE (((tblBHT.Discharge)=0)) ORDER BY tblBHT.BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With

End Sub
