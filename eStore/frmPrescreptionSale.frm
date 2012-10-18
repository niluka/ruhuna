VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrescreptionSale 
   Caption         =   "Prescreption Sale"
   ClientHeight    =   11010
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
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDuration 
      Height          =   375
      Left            =   1440
      TabIndex        =   92
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtCostRate 
      Height          =   375
      Left            =   9960
      TabIndex        =   85
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtTotalCost 
      Height          =   375
      Left            =   13320
      TabIndex        =   84
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtItemCost 
      Height          =   375
      Left            =   9960
      TabIndex        =   83
      Top             =   1200
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   11280
      TabIndex        =   80
      Top             =   5400
      Width           =   1815
      Begin VB.OptionButton OptTwo 
         Caption         =   "2"
         Height          =   240
         Left            =   1080
         TabIndex        =   82
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optOne 
         Caption         =   "1"
         Height          =   240
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   11280
      TabIndex        =   79
      Top             =   5160
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin btButtonEx.ButtonEx bttnSettle 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   13200
      TabIndex        =   74
      Top             =   5640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Settle"
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
   Begin MSFlexGridLib.MSFlexGrid GridBatch 
      Height          =   975
      Left            =   5400
      TabIndex        =   15
      Top             =   1680
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1720
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   255
      Left            =   8280
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
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
   Begin VB.TextBox txtPrice 
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtRate 
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtQty 
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo dtcCatogery 
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItem 
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   4575
      Left            =   11280
      TabIndex        =   9
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8070
      _Version        =   393216
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
      TabCaption(0)   =   "Total"
      TabPicture(0)   =   "frmPrescreptionSale.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label22"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label23"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label24"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtGTotal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDiscount"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtNTotal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dtcSale"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Payment"
      TabPicture(1)   =   "frmPrescreptionSale.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameInPatient"
      Tab(1).Control(1)=   "frameCash"
      Tab(1).Control(2)=   "frameCredit"
      Tab(1).Control(3)=   "frameCheque"
      Tab(1).Control(4)=   "frameOutPatient"
      Tab(1).Control(5)=   "frameStaff"
      Tab(1).Control(6)=   "frameCreditCard"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "frmPrescreptionSale.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dtcDepartment"
      Tab(2).Control(1)=   "chkRequest"
      Tab(2).Control(2)=   "dtcCheckedStaff"
      Tab(2).Control(3)=   "dtcIssueStaff"
      Tab(2).Control(4)=   "Label21"
      Tab(2).Control(5)=   "Label20"
      Tab(2).ControlCount=   6
      Begin MSDataListLib.DataCombo dtcDepartment 
         Height          =   360
         Left            =   -74520
         TabIndex        =   78
         Top             =   2640
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.CheckBox chkRequest 
         Caption         =   "Make a request"
         Height          =   375
         Left            =   -74760
         TabIndex        =   77
         Top             =   2280
         Width           =   3255
      End
      Begin MSDataListLib.DataCombo dtcSale 
         Height          =   360
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
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
         Left            =   1200
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   1680
         Width           =   2415
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
         Left            =   1200
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1080
         Width           =   2415
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
         Left            =   1200
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   480
         Width           =   2415
      End
      Begin MSDataListLib.DataCombo dtcCheckedStaff 
         Height          =   360
         Left            =   -74520
         TabIndex        =   24
         Top             =   1440
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcIssueStaff 
         Height          =   360
         Left            =   -74520
         TabIndex        =   25
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame frameInPatient 
         Caption         =   "Indoor Patient"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   54
         Top             =   2520
         Width           =   3495
         Begin MSDataListLib.DataCombo dtcBHT 
            Height          =   360
            Left            =   840
            TabIndex        =   60
            Top             =   360
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin VB.TextBox txtBHTBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtPatient 
            Height          =   375
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txtTemCreditCustomerBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   76
            Top             =   1440
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label Label40 
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label39 
            Caption         =   "Patient"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label38 
            Caption         =   "BHT"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame frameCash 
         Caption         =   "Cash"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   28
         Top             =   360
         Width           =   3495
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            TabIndex        =   34
            Top             =   1200
            Width           =   2535
         End
         Begin VB.TextBox txtCashPaid 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            TabIndex        =   33
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtDue 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            TabIndex        =   32
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label27 
            Caption         =   "Change"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label26 
            Caption         =   "Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label25 
            Caption         =   "Due"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame frameCredit 
         Caption         =   "Credit"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   3495
         Begin VB.TextBox txtCreditDue 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            TabIndex        =   36
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label30 
            Caption         =   "Due"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame frameCheque 
         Caption         =   "Cheque"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   45
         Top             =   360
         Width           =   3495
         Begin MSComCtl2.DTPicker dtpChequeDate 
            Height          =   375
            Left            =   720
            TabIndex        =   53
            Top             =   1680
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   20709379
            CurrentDate     =   39551
         End
         Begin VB.TextBox txtChequeNo 
            Height          =   375
            Left            =   720
            TabIndex        =   48
            Top             =   1200
            Width           =   2655
         End
         Begin MSDataListLib.DataCombo dtcBranch 
            Height          =   360
            Left            =   720
            TabIndex        =   46
            Top             =   720
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcBank 
            Height          =   360
            Left            =   720
            TabIndex        =   47
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label37 
            Caption         =   "Date"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label36 
            Caption         =   "No"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label35 
            Caption         =   "Bank"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label34 
            Caption         =   "Branch"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame frameOutPatient 
         Caption         =   "Out Patient"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   61
         Top             =   2520
         Width           =   3495
         Begin VB.TextBox txtCreditCustomerBalance 
            Height          =   375
            Left            =   840
            TabIndex        =   63
            Top             =   840
            Width           =   2535
         End
         Begin MSDataListLib.DataCombo dtcCreditCustomer 
            Height          =   360
            Left            =   840
            TabIndex        =   62
            Top             =   360
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label43 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label42 
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame frameStaff 
         Caption         =   "Staff Issue"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   68
         Top             =   2520
         Width           =   3495
         Begin VB.TextBox txtTemStaffCredit 
            Height          =   375
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   1440
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.TextBox txtStaffBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   840
            Width           =   2535
         End
         Begin MSDataListLib.DataCombo dtcStaffCustomer 
            Height          =   360
            Left            =   840
            TabIndex        =   69
            Top             =   360
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label44 
            Caption         =   "Staff"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label41 
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame frameCreditCard 
         Caption         =   "Credit Card"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   38
         Top             =   360
         Width           =   3495
         Begin VB.TextBox txtCreditCode 
            Height          =   375
            Left            =   720
            TabIndex        =   66
            Top             =   1680
            Width           =   2655
         End
         Begin MSDataListLib.DataCombo dtcCardBank 
            Height          =   360
            Left            =   720
            TabIndex        =   44
            Top             =   720
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcCreditCard 
            Height          =   360
            Left            =   720
            TabIndex        =   43
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.TextBox txtCreditCardNo 
            Height          =   375
            Left            =   720
            TabIndex        =   39
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label28 
            Caption         =   "Code"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label33 
            Caption         =   "Bank"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label32 
            Caption         =   "Card"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label31 
            Caption         =   "No"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.Label Label21 
         Caption         =   "Checked By"
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "Issued By"
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label24 
         Caption         =   "Sale Catogery"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Net Total"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "Discount"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Gross Total"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7011
      _Version        =   393216
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   255
      Left            =   9480
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "&Delete"
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
   Begin MSDataListLib.DataCombo dtcCode 
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcFrequency 
      Height          =   360
      Left            =   1440
      TabIndex        =   86
      Top             =   1680
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcPMessage 
      Height          =   360
      Left            =   1440
      TabIndex        =   87
      Top             =   2160
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcDuration 
      Height          =   360
      Left            =   2520
      TabIndex        =   93
      Top             =   2640
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label9 
      Caption         =   "Duration"
      Height          =   375
      Left            =   120
      TabIndex        =   94
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Code"
      Height          =   375
      Left            =   120
      TabIndex        =   91
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Category"
      Height          =   375
      Left            =   120
      TabIndex        =   90
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Messages"
      Height          =   375
      Left            =   120
      TabIndex        =   89
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Frequency"
      Height          =   375
      Left            =   120
      TabIndex        =   88
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblDisplayTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "Cash Rs. 0.00"
      Height          =   375
      Left            =   11280
      TabIndex        =   75
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Label lblIUnit 
      Height          =   375
      Left            =   8520
      TabIndex        =   14
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Item"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Price"
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Rate"
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Quantity"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmPrescreptionSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCatogery As New ADODB.Recordset
    Dim rsCode As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsFrequency As New ADODB.Recordset
    Dim rsDuration As New ADODB.Recordset
    Dim rsPMessage As New ADODB.Recordset
    Dim rsTemFrequency As New ADODB.Recordset
    Dim rsTemDuration As New ADODB.Recordset
    Dim rsTemPMessage As New ADODB.Recordset
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
    
    Dim rsBanks As New ADODB.Recordset
    Dim rsCities As New ADODB.Recordset
    Dim rsCreditCards As New ADODB.Recordset
    Dim rsSale As New ADODB.Recordset
    Dim rsTemStaff As New ADODB.Recordset
    Dim rsBHT As New ADODB.Recordset
    Dim rsPatients As New ADODB.Recordset
    Dim rsStore As New ADODB.Recordset
    Dim temSQL As String
    Dim NewItem As New Item
    Dim NewSale As New Sale

    Dim CsetPrinter As New cSetDfltPrinter


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
        .Text = Format(Val(txtRate.Text), "#0.00") & " per " & NewItem.IUnit
        .Col = 4
        .CellAlignment = 7
        .Text = txtQty.Text & " " & NewItem.IUnit
        .Col = 5
        .CellAlignment = 7
        .Text = Format(Val(txtPrice.Text), "#0.00")
        .Col = 6
        .Text = Val(dtcItem.BoundText)
        .Col = 7
        .Text = GridBatch.TextMatrix(GridBatch.Row, 4)
        .Col = 9
        .CellAlignment = 7
        .Text = Format(Val(txtRate.Text), "#0.00")
        .Col = 8
        .CellAlignment = 7
        .Text = txtQty.Text
        .Col = 10
        .Text = Val(txtItemCost.Text)
        .Col = 11
        .Text = Val(dtcFrequency.BoundText)
        .Col = 12
        .Text = Val(dtcPMessage.BoundText)
        .Col = 13
        .Text = Val(txtDuration.Text)
        .Col = 14
        .Text = Val(dtcDuration.BoundText)
        
        CalculateTotal
        ClearAddValues
        FormatSelectStock
    End With
    bttnDelete.Enabled = False
    dtcItem.SetFocus
End Sub

Private Sub ClearAddValues()
    txtQty.Text = Empty
    txtRate.Text = Empty
    txtPrice.Text = Empty
    txtItemCost.Text = Empty
    dtcItem.Text = Empty
    dtcCatogery.Text = Empty
    dtcCode.Text = Empty
    txtCostRate.Text = Empty
    dtcFrequency.Text = Empty
    dtcDuration.Text = Empty
    dtcPMessage.Text = Empty
    txtDuration.Text = Empty
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
        If IsNumeric(txtRate.Text) = False Or Val(txtRate.Text) = 0 Then
            tr = MsgBox("You have not entered the rate", vbCritical, "Rate")
            txtRate.SetFocus
            Exit Function
        End If
    
    CanAdd = True
End Function

Private Sub bttnDelete_Click()
    With GridItem
        If .Rows <= 1 Then Exit Sub
        If .Rows = 2 Then
            FormatItemGrid
        Else
            .RemoveItem (.Row)
        End If
        bttnDelete.Enabled = False
        Call CalculateTotal
    End With
End Sub

Private Sub bttnSettle_Click()
    Dim TemSaleBillID As Long
    Dim TemOutPatientID As Long
    Dim TemBHTID As Long
    Dim TemCreditCardID As Long
    Dim TemCashID As Long
    Dim TemCreditID As Long
    Dim TemChequeID As Long
    Dim i As Integer
    
    txtDue.Text = txtNTotal.Text
    If CanSettle = False Then Exit Sub
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
    
    
    
    
    If rsTemSale.State = 1 Then rsTemSale.Close
    temSQL = "SELECT tblSale.* FROM tblSale"
    rsTemSale.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
    With GridItem
        For i = 1 To .Rows - 1
            If ConsumeStocks(UserStoreID, Val(.TextMatrix(i, 7)), Val(.TextMatrix(i, 8))) = True Then
                rsTemSale.AddNew
                rsTemSale!SaleBillID = TemSaleBillID
                rsTemSale!CatogeryID = Val(dtcSale.BoundText)
                rsTemSale!ItemID = Val(.TextMatrix(i, 6))
                rsTemSale!BatchID = Val(.TextMatrix(i, 7))
                rsTemSale!StoreID = UserStoreID
                rsTemSale!Date = Date
                rsTemSale!Time = Time
                rsTemSale!StaffID = UserID
                If IsNumeric(dtcCheckedStaff.BoundText) = True Then rsTemSale!CheckedStaffID = dtcCheckedStaff.BoundText
                rsTemSale!Amount = Val(.TextMatrix(i, 8))
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
                End If
                rsTemSale!Duration = Val(.TextMatrix(1, 13))
                rsTemSale!DurationID = Val(.TextMatrix(1, 14))
                rsTemSale!FrequencyID = Val(.TextMatrix(1, 11))
                rsTemSale!PMessageID = Val(.TextMatrix(1, 12))
                rsTemSale.Update
            End If
        Next i
    End With
    With rsTemSaleBill
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblSaleBill where SaleBillID = " & TemSaleBillID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If NewSale.OutPatient = True Then
                !BilledOutPatientID = TemOutPatientID
            ElseIf NewSale.InPatient = True Then
                !BilledBHTID = dtcBHT.BoundText
            ElseIf NewSale.Staff = True Then
                !BilledStaffID = dtcStaffCustomer.BoundText
            End If
            If NewSale.Cash = True Then
                !PaymentMethodID = 1
                !PaymentMethod = "Cash"
                !receivedCashID = TemCashID
            ElseIf NewSale.Credit = True Then
                !PaymentMethodID = 4
                !PaymentMethod = "Credit"
                !ReceivedCreditID = TemCreditID
            ElseIf NewSale.Cheque = True Then
                !PaymentMethodID = 5
                !PaymentMethod = "Cheque"
                !receivedChequeID = TemChequeID
            ElseIf NewSale.CreditCard = True Then
                !PaymentMethodID = 3
                !PaymentMethod = "Credit Card"
                !receivedCreditcardID = TemCreditCardID
            End If
            !NetCost = Val(txtTotalCost.Text)
            .Update
        End If
        .Close
    End With
        
        If chkPrint.Value = 1 Then
                Call SetBillPrinter
                Call SetBillPaper

            If OptTwo.Value = True Then
                            
             Dim tr As Integer
                tr = MsgBox("Print a Copy?", vbQuestion + vbYesNo, "Print again?")
                If tr = vbYes Then
                    SetBillPrinter
                    SetBillPaper
                End If
            End If
        Else
        
        End If
    
    ClearBillValues
    Call FormatItemGrid
    dtcCatogery.SetFocus
    
End Sub

Private Sub SetBillPrinter()
    CsetPrinter.SetPrinterAsDefault (BillPrinterName)
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
        Printer.Print Tab(Tab1); "Time : "; Time
        Printer.Print Tab(Tab1); "--------------------------------------"
        If NewSale.OutPatient = True Then
            Printer.Print Tab(Tab1); "Patient : "; dtcCreditCustomer.Text
        ElseIf NewSale.OutPatient = True Then
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
        Printer.Print Tab(Tab1); "Discount"; Tab(Tab4); Right((Space(10)) & (txtDiscount.Text), 10)
        Printer.Print Tab(Tab1); "Net Total"; Tab(Tab4); Right((Space(10)) & (txtNTotal.Text), 10)
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
End Sub

Private Function ConsumeStocks(ByVal IStoreIDValue As Long, ByVal BatchIDValue As Long, ByVal Quentity As Double) As Boolean
    Dim tr As Integer
    On Error GoTo eh
    ConsumeStocks = False
    With rsTemBatch
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblBatchstock where batchid = " & BatchIDValue & " AND StoreID = " & IStoreIDValue & " ORDER BY tblBatchstock.Stock DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
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
    With rsTemCredit
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblReceivedCredit"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = dtcIssueStaff.BoundText
        !ReceivedDate = Date
        !ReceivedTime = Time
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
        ReceiveCredit = !ReceivedCreditID
        .Close
    End With
End Function


Private Function ReceiveCheque(SaleBillID As Long) As Long
    With rsTemCheque
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblReceivedCheque"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = dtcIssueStaff.BoundText
        !ReceivedDate = Date
        !ReceivedTime = Time
        !bankID = Val(dtcBank.BoundText)
        If IsNumeric(dtcBranch.BoundText) = True Then
            !BranchID = dtcBranch.BoundText
        End If
        !ChequeDate = Format(dtpChequeDate.Value, "dd MMMMM yyyy")
        !ChequeNo = txtChequeNo.Text
        If NewSale.InPatient = True Then
            !ReceivedFromBHTID = dtcBHT.BoundText
        ElseIf NewSale.OutPatient = True Then
            !ReceivedFromOutPatientID = dtcCreditCustomer.BoundText
        ElseIf NewSale.Staff = True Then
            !ReceivedFromStaffID = dtcStaffCustomer.BoundText
        End If
        !StoreID = UserStoreID
        !Price = Val(txtNTotal.Text)
        !SaleBillID = SaleBillID
        .Update
        ReceiveCheque = !receivedChequeID
        .Close
    End With
End Function


Private Function ReceiveCash(SaleBillID As Long) As Long
    With rsTemCash
        If .State = 1 Then .Close
        temSQL = "SELECT tblReceivedCash.* FROM tblReceivedCash"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = dtcIssueStaff.BoundText
        !ReceivedDate = Date
        !ReceivedTime = Time
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
        ReceiveCash = !receivedCashID
        .Close
    End With
End Function


Private Function ReceiveCreditCard(SaleBillID As Long) As Long
    With rsTemCC
        If .State = 1 Then .Close
        temSQL = "SELECT tblReceivedCreditCard.* FROM tblReceivedCreditCard"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !CreditCardNo = txtCreditCardNo.Text
        !ReceivedSTaffID = dtcIssueStaff.BoundText
        !CardTypeID = dtcCreditCard.BoundText
        !AuthrizationCode = txtCreditCode.Text
        !ReceivedSTaffID = dtcIssueStaff.BoundText
        !ReceivedDate = Date
        !ReceivedTime = Time
        !AuthrizationDate = Date
        !AuthrizationTime = Time
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
        ReceiveCreditCard = !receivedCreditcardID
        .Close
    End With
End Function

Private Function WritePatient() As Long
    Dim temPatient As String
    With rsTemPatient
       If .State = 1 Then .Close
       temSQL = "SELECT * from tblpatientmaindetails"
       .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
       .AddNew
       !FirstName = dtcCreditCustomer.Text
       .Update
       WritePatient = !PatientID
        .Close
    End With
    With dtcCreditCustomer
        Set .RowSource = Nothing
        .ListField = Empty
        .BoundColumn = Empty
    End With
    With rsPatients
        If .State = 1 Then .Close
        temSQL = "SELECT tblPatientMainDetails.* FROM tblPatientMainDetails ORDER BY tblPatientMainDetails.FirstName"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
        temSQL = "SELECT * from tblSaleBill"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Date = Date
        !Time = Time
        !StaffID = dtcIssueStaff.BoundText
        !StoreID = UserStoreID
        !Price = Val(txtGTotal.Text)
        !Discount = Val(txtDiscount.Text)
        !DiscountPercent = ((Val(txtDiscount.Text)) / (Val(txtGTotal.Text))) * 100
        !NetPrice = Val(txtNTotal.Text)
        If IsNumeric(dtcCheckedStaff.BoundText) = True Then !CheckedStaffID = Val(dtcCheckedStaff.BoundText)
        .Update
        SaleBillID = !SaleBillID
        .Close
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
        If IsNumeric(dtcCreditCard.BoundText) = False Then
            tr = MsgBox("You have not selected the Credit Card Type", vbCritical, "Card type?")
            SSTab2.Tab = 1
            dtcCreditCard.SetFocus
            Exit Function
        End If
        If Not IsNumeric(dtcCardBank.BoundText) = False Then
            tr = MsgBox("You have not selected the cadit card issued bank", vbCritical, "Bank?")
            SSTab2.Tab = 1
            dtcCardBank.SetFocus
            Exit Function
        End If
        If Trim(txtCreditCardNo.Text) = "" Then
            tr = MsgBox("You have not entered a valied credit card number", vbCritical, "Card Number?")
            SSTab2.Tab = 1
            txtCreditCardNo.SetFocus
            Exit Function
        End If
        If Trim(txtCreditCode.Text) = "" Or IsNumeric(txtCreditCode.Text) = False Then
            tr = MsgBox("You have not entered a valied autherization code", vbCritical, "Authorization code?")
            SSTab2.Tab = 1
            txtCreditCode.SetFocus
            Exit Function
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
    End If
    
    CanSettle = True
End Function

Private Sub dtcBHT_Click(Area As Integer)
    Dim TemBHTCredit As Double
    If IsNumeric(dtcBHT.BoundText) = False Then Exit Sub
    With rsTemStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblBHT where BHTID = " & Val(dtcBHT.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
        End If
        If .State = 1 Then .Close
    End With

End Sub

Private Sub dtcCode_Change()
    dtcItem.BoundText = dtcCode.BoundText
End Sub


Private Sub dtcCreditCustomer_Click(Area As Integer)
    Dim TemCreditCustomerCredit As Double
    If IsNumeric(dtcCreditCustomer.BoundText) = False Then Exit Sub
    With rsTemStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblpatientmaindetails where patientID = " & Val(dtcCreditCustomer.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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

Private Sub dtcDuration_Change()
    Call CalculateQuentity
End Sub

Private Sub dtcFrequency_Change()
    Call CalculateQuentity
End Sub

Private Sub dtcItem_Change()
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    Dim tr As Integer
    dtcCode.BoundText = dtcItem.BoundText
    NewItem.ID = dtcItem.BoundText
    Call FillAddPrice(dtcItem.BoundText)
    lblIUnit.Caption = NewItem.IUnit
    Call CalculatePrice
    Call FillSelectStock(dtcItem.BoundText)
End Sub

Private Sub FillAddPrice(ByVal ItemID As Long)
    With rsTemPrice
        If .State = 1 Then .Close
        temSQL = "SELECT tblCurrentSalePrice.SPrice FROM tblCurrentSalePrice WHERE tblCurrentSalePrice.ItemID=" & ItemID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
                txtRate.Text = Format(rsTemPrice!SPrice, "##00.00")
        End If
    End With
    With rsTemPrice
        If .State = 1 Then .Close
        temSQL = "SELECT tblCurrentPurchasePrice.PPrice FROM tblCurrentPurchasePrice WHERE tblCurrentPurchasePrice.ItemID=" & ItemID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
                txtCostRate.Text = Format(rsTemPrice!PPrice, "##00.00")
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

Private Sub FillSelectStock(ByVal ItemID As Long)
    With GridBatch
        .Visible = False
        FormatSelectStock
    End With
    With rsTemStore
        If .State = 1 Then .Close
        temSQL = "SELECT tblBatch.*,  tblBatchStock.*, tblLocation.Location " & _
                    " FROM (tblStore RIGHT JOIN (tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) ON tblStore.StoreID = tblBatchStock.StoreID) LEFT JOIN tblLocation ON tblBatchStock.LocationID = tblLocation.LocationID " & _
                    " WHERE tblBatch.ItemID=" & ItemID & " AND tblBatchStock.StoreID=" & UserStoreID & " AND tblBatchStock.Stock > 0 " & _
                    "ORDER BY tblBatch.DOE DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
                Else
                    GridBatch.Text = 0
                End If
                GridBatch.Col = 2
                GridBatch.CellAlignment = 1
                GridBatch.Text = Format(!DOE, ShortDateFormat)
                GridBatch.Col = 3
                GridBatch.CellAlignment = 1
                If Not IsNull(!location) Then
                    GridBatch.Text = !location
                Else
                    GridBatch.Text = Empty
                End If
                GridBatch.Col = 4
                GridBatch.Text = ![tblBatch.BatchID]
                GridBatch.Col = 5
                GridBatch.Text = ![tblBatchStock.BatchID]
                
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
End Sub

Private Sub GetItemDetails(ItemID As Long)
    NewItem.ID = ItemID
'    txtAMP.Text = NewItem.AMP
'    txtAMPP.Text = NewItem.AMPP
'    txtVMP.Text = NewItem.VMP
'    txtVMPP.Text = NewItem.VMPP
'    txtVTM.Text = NewItem.Generic
'    txtDisplay.Text = NewItem.Display
End Sub


Private Sub dtcItem_LostFocus()
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    Dim tr As Integer
    If CalculateStock(dtcItem.BoundText, , UserStoreID).Amount <= 0 Then
        tr = MsgBox("There are no stocks", vbCritical, "No Stocks")
        dtcCatogery.Text = Empty
        dtcItem.SetFocus
        Exit Sub
    End If
End Sub

Private Sub dtcSale_Change()
    If IsNumeric(dtcSale.BoundText) = False Then Exit Sub
    NewSale.SaleCategoryID = Val(dtcSale.BoundText)
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
    End If
    If NewSale.InPatient = True Then
        frameInPatient.Visible = True
        frameOutPatient.Visible = False
        frameStaff.Visible = False
        lblDisplayTotal.Caption = lblDisplayTotal.Caption & " for In-Hospital Patients"
    ElseIf NewSale.OutPatient = True Then
        frameInPatient.Visible = False
        frameOutPatient.Visible = True
        frameStaff.Visible = False
        lblDisplayTotal.Caption = lblDisplayTotal.Caption & " for Out-Hospital Patients"
    ElseIf NewSale.Staff = True Then
        frameInPatient.Visible = False
        frameOutPatient.Visible = False
        frameStaff.Visible = True
        lblDisplayTotal.Caption = lblDisplayTotal.Caption & " for staff members"
    End If
'    SSTab2.Tab = 1
    Call CalculateDiscount
    lblDisplayTotal.Caption = lblDisplayTotal.Caption & " - Rs. " & txtNTotal.Text
End Sub

Private Sub CalculateDiscount()
    txtDiscount.Text = Format((Round(Val(txtGTotal.Text) * (NewSale.SaleDiscountPercent / 100), 0)), "0.00")
End Sub

Private Sub dtcStaffCustomer_Change()
    Dim TemStaffCredit As Double
    If IsNumeric(dtcStaffCustomer.BoundText) = False Then Exit Sub
    With rsTemStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblSTaff where staffid = " & Val(dtcStaffCustomer.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!Credit) Then
                TemStaffCredit = !Credit
            Else
                TemStaffCredit = 0
            End If
            txtTemStaffCredit.Text = TemStaffCredit
            If TemStaffCredit < 0 Then
                txtStaffBalance.Text = "(" & Format(Abs(TemStaffCredit), "#,##0.00") & ")"
            Else
                txtStaffBalance.Text = Format(TemStaffCredit, "#,##0.00")
            End If
        End If
        If .State = 1 Then .Close
    End With
End Sub

Private Sub Form_Activate()
    Me.WindowState = 2
    Call FormatItemGrid
End Sub

Private Sub Form_Load()
    Call FillCombos
    dtcIssueStaff.BoundText = UserID
    dtcIssueStaff.Locked = True
    SSTab2.Tab = 0
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
        .Cols = 15
        .Rows = 1
        Dim i As Integer
        For i = 0 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            Select Case i
                Case 0: .Text = "No."
                        .ColWidth(i) = 500
                Case 1: .Text = "Item"
                        .ColWidth(i) = 4400
                Case 2: .Text = "Batch"
                        .ColWidth(i) = 1500
                Case 3: .Text = "Rate"
                        .ColWidth(i) = 2000
                Case 4: .Text = "Amount"
                        .ColWidth(i) = 1500
                Case 5: .Text = "Price"
                Case Else
                        .ColWidth(i) = 1
            End Select
        Next
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
    End With
End Sub

Private Sub FillCombos()
    With rsSale
        If .State = 1 Then .Close
        temSQL = "SELECT tblSaleCatogery.SaleCatogeryID, tblSaleCatogery.SaleCatogery FROM tblSaleCatogery ORDER BY tblSaleCatogery.SaleCatogery"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcSale
        Set .RowSource = rsSale
        .ListField = "SaleCatogery"
        .BoundColumn = "SaleCatogeryID"
    End With
    With rsItem
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblitem order by display"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcItem
        Set .RowSource = rsItem
        .ListField = "display"
        .BoundColumn = "ItemID"
    End With
    With rsItemCatogery
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblitemcatogery order by itemcatogery"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCatogery
        Set .RowSource = rsItemCatogery
        .ListField = "ItemCatogery"
        .BoundColumn = "ItemCategoryID"
    End With
    With rsCode
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblitem order by code"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCode
        Set .RowSource = rsCode
        .ListField = "code"
        .BoundColumn = "ItemID"
    End With
    With rsStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblstaff order by listedname"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
        temSQL = "SELECT tblBank.* FROM tblBank ORDER BY tblBank.Bank"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
        temSQL = "SELECT tblCreditCardType.CreditCardTypeID, tblCreditCardType.CreditCardType FROM tblCreditCardType ORDER BY tblCreditCardType.CreditCardType"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCreditCard
        Set .RowSource = rsCreditCards
        .ListField = "CreditCardType"
        .BoundColumn = "CreditCardTypeID"
    End With
    With rsCities
        If .State = 1 Then .Close
        temSQL = "SELECT tblCity.CityId, tblCity.City FROM tblCity ORDER BY tblCity.City"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcBranch
        Set .RowSource = rsCities
        .ListField = "City"
        .BoundColumn = "CityId"
    End With
    With rsBHT
        If .State = 1 Then .Close
        temSQL = "SELECT tblBHT.* FROM tblBHT WHERE (((tblBHT.Discharge)=False)) ORDER BY tblBHT.BHT"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
    With rsPatients
        If .State = 1 Then .Close
        temSQL = "SELECT tblPatientMainDetails.* FROM tblPatientMainDetails ORDER BY tblPatientMainDetails.FirstName"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCreditCustomer
        Set .RowSource = rsPatients
        .ListField = "FirstName"
        .BoundColumn = "PatientID"
    End With
    With rsStore
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblStore order by store"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcDepartment
        Set .RowSource = rsStore
        .ListField = "Store"
        .BoundColumn = "StoreID"
    End With
    With rsFrequency
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblFrequency order by Frequency"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcFrequency
        Set .RowSource = rsFrequency
        .ListField = "Frequency"
        .BoundColumn = "FrequencyID"
    End With
    With rsDuration
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblDuration order by Duration"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcDuration
        Set .RowSource = rsDuration
        .ListField = "Duration"
        .BoundColumn = "DurationID"
    End With
    With rsPMessage
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblPMEssage order by PMessage"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcPMessage
        Set .RowSource = rsPMessage
        .ListField = "PMessage"
        .BoundColumn = "PMEssageID"
    End With
End Sub

Private Sub CalculateQuentity()
    If Not IsNumeric(dtcDuration.BoundText) Then Exit Sub
    If Not IsNumeric(dtcFrequency.BoundText) Then Exit Sub
    If Val(txtDuration.Text) = 0 Then Exit Sub
    Dim DInterval As Long
    Dim FInterval As Long
    Dim temQty As Long
    With rsTemDuration
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblDuration where DurationID = " & dtcDuration.BoundText
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!Interval) Then
                DInterval = !Interval
            End If
        End If
        .Close
    End With
    With rsTemFrequency
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblFrequency where FrequencyID = " & dtcFrequency.BoundText
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!Interval) Then
                FInterval = !Interval
            End If
        End If
        .Close
    End With
    temQty = (Val(txtDuration.Text) * DInterval) / (FInterval)
    txtQty.Text = temQty
End Sub

Private Sub dtcCatogery_Change()
    If IsNumeric(dtcCatogery.BoundText) Then
        ListSelectedItems
    Else
        ListAllItems
    End If
    dtcItem.Text = Empty
    dtcCode.Text = Empty
End Sub


Private Sub ListSelectedItems()
With rsItem
    If .State = 1 Then .Close
    temSQL = "SELECT * from tblitem where itemcatogeryID = " & dtcCatogery.BoundText & " order by display"
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "Display"
    .BoundColumn = "ItemID"
End With
With rsCode
    If .State = 1 Then .Close
    temSQL = "SELECT * from tblitem where itemcatogeryID = " & dtcCatogery.BoundText & " order by code"
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
    temSQL = "SELECT * from tblitem order by display"
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "display"
    .BoundColumn = "ItemID"
End With
With rsCode
    If .State = 1 Then .Close
    temSQL = "SELECT * from tblitem order by code"
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcCode
    Set .RowSource = rsCode
    .ListField = "Code"
    .BoundColumn = "ItemID"
End With
End Sub

Private Sub dtcCatogery_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then dtcCatogery.Text = Empty
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
        .Col = 11
        dtcFrequency.BoundText = Val(.Text)
        .Col = 12
        dtcPMessage.BoundText = Val(.Text)
        .Col = 13
        txtDuration.Text = Val(.Text)
        .Col = 14
        dtcDuration.BoundText = Val(.Text)
        bttnDelete_Click
    End With
    dtcItem.SetFocus
End Sub


Private Sub txtCashPaid_Change()
    txtBalance.Text = Format((Val(txtCashPaid.Text) - Val(txtDue.Text)), "0.00")
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

Private Sub txtDuration_Change()
    Call CalculateQuentity
End Sub

Private Sub txtGTotal_Change()
    Call CalculateNetTotal
End Sub

Private Sub txtNTotal_Change()
    txtDue.Text = txtNTotal.Text
    txtCreditDue.Text = txtNTotal.Text
End Sub

Private Sub txtQty_Change()
    Call CalculatePrice
End Sub

Private Sub CalculatePrice()
    txtPrice.Text = Format((Val(txtQty.Text) * Val(txtRate.Text)), "0.00")
    txtItemCost.Text = Val(txtCostRate.Text) * Val(txtQty.Text)
End Sub

Private Sub AddLabels()
If PreferanceLabelCols = 1 Then
    AddSingleLabels
Else
    AddDoubleLabels
End If
End Sub

Private Sub AddDoubleLabels()
    Dim InbetweenY As Long
    Dim LLabelTY As Long
    Dim LLabelBY As Long
    Dim LLabelLX As Long
    Dim LLabelRX As Long
    Dim RLabelTY As Long
    Dim RLabelBY As Long
    Dim RLabelLX As Long
    Dim RLabelRX As Long
    Dim LabelRowSerial As Long
    Dim RLabelRowSerial As Long
    Dim LLabelRowSerial As Long
    Dim IndoorSerial As Long
    Dim LabelSinhalaFontName As String
    Dim LabelSinhalaFontSize As Single
    Dim LabelTamilFontName As String
    Dim LabelTamilFontSize As Single
    Dim LabelEnglishFontName As String
    Dim LabelEnglishFontSize As Single
    
    LabelSinhalaFontName = "kaputadotcom"
    LabelSinhalaFontSize = 12
    LabelEnglishFontName = "Verdana"
    LabelEnglishFontSize = 9
    
    LabelRowSerial = 0
    RLabelRowSerial = 0
    LLabelRowSerial = 0
    
    With Printer
    IndoorSerial = 1
    
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
'   11  Frequency
'   12  PMessage
'   13  txtDuration
'   14  dtcDuration
    
    For i = 1 To GridItem.Rows - 1
    
        If i Mod 2 = 1 Then
            TemString1 = "vrkt" ' varakata
            
            LLabelTY = LabTopY + TextYMargin + LabelRowSerial * InbetweenY
            LLabelLX = LX + (TextXMargin / 3)
            LLabelRX = ((RX - LX) / 2) - (TextXMargin / 3)
            .Font.Name = LabelEnglishFontName
            .Font.Size = LabelEnglishFontSize

            .CurrentY = LLabelTY + LLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = LLabelLX + ((LLabelRX - LLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            
            Printer.Print GridItem.TextMatrix(i, 1) & " " & GridItem.TextMatrix(i, 4)
            
            LLabelRowSerial = LLabelRowSerial + 2
            TemString4 = Dataenvironment1.rscmmdTemDrugs!drug
            .CurrentY = LLabelTY + LLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = LLabelLX + ((LLabelRX - LLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            LLabelRowSerial = LLabelRowSerial + 2
                
            TemString4 = ""
            .CurrentY = LLabelTY + LLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = LLabelLX + ((LLabelRX - LLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            LLabelRowSerial = LLabelRowSerial + 2
                
                If TemString3 <> "" Then
                    .Font.Name = LabelSinhalaFontName
                    .Font.Size = LabelSinhalaFontSize
                     TemString4 = TemString3
                    .CurrentY = LLabelTY + LLabelRowSerial * InbetweenY + TextYMargin
                    .CurrentX = LLabelLX + ((LLabelRX - LLabelLX) / 2) - (.TextWidth(TemString4)) / 2
                    Printer.Print TemString4
                    LLabelRowSerial = LLabelRowSerial + 2
                End If
                
            TemString4 = ""
            .CurrentY = LLabelTY + LLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = LLabelLX + ((LLabelRX - LLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            LLabelRowSerial = LLabelRowSerial + 2
               
                
            .Font.Name = LabelSinhalaFontName
            .Font.Size = LabelSinhalaFontSize
             
             TemString4 = TemString1 & TemString5 & "b#gQn\" & "  "
            .CurrentY = LLabelTY + LLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = LLabelLX + ((LLabelRX - LLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString1 & " ";
            .Font.Name = LabelEnglishFontName
            Printer.Print TemString5 & " ";
            .Font.Name = LabelSinhalaFontName
            Printer.Print "b#gQn\"
            
            
            LLabelRowSerial = LLabelRowSerial + 2
            TemString4 = TemString2
            .CurrentY = LLabelTY + LLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = LLabelLX + ((LLabelRX - LLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            LLabelRowSerial = LLabelRowSerial + 4
            .CurrentY = LLabelTY + LLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = LLabelLX + ((LLabelRX - LLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            LLabelRowSerial = LLabelRowSerial + 2
            LLabelBY = .CurrentY
            Printer.Line (LLabelLX, LLabelTY)-(LLabelRX, LLabelBY), , B
            
        Else
        
            TemString1 = ""
            TemString2 = ""
            TemString3 = ""
            TemString4 = ""
            TemString5 = ""
            RLabelRowSerial = 0
            TemString1 = "vrkt" ' varakata
            Select Case Dataenvironment1.rscmmdTemDrugs!DoseUnitID
            Case 1: TemString1 = TemString1 & " " & "mQlQ gY$m|"
            Case 2: TemString1 = TemString1 & " " & "@@m@kY` gY$m|"
            Case 3: TemString1 = TemString1 & " " & "mQlQ lWtr\"
                    TemString3 = "@b`n @b@hw"
            Case 4: TemString1 = TemString1 & " " & "@pwQ" ' tab
            Case 5: TemString1 = TemString1 & " " & "krl\" 'cap
            Case 6: TemString1 = TemString1 & " " & "@w\ h#qQ"
                    TemString3 = "@b`n @b@hw"
            Case 7: TemString1 = TemString1 & " " & "bQn\qE"
                    TemString3 = "@b`n @b@hw"
            Case 8: TemString1 = TemString1 & " " & "up`n\g"
            Case 9: TemString1 = TemString1 & " " & "ugOr#"
            Case 10: TemString1 = TemString1 & " " & "s@p`\sQtrQ" ' suppository
            Case 11: TemString1 = TemString1 & " " & "en\nw\" ' injections
            Case 12: TemString1 = TemString1 & " " & "v#jyQnl\ @pYsrQ"  'vaginal pressery
            Case 13: TemString1 = TemString1 & " " & "gY$m|"
            Case 14: TemString1 = TemString1 & " " & "@@m@kY` gY$m|"
            Case 15: TemString1 = TemString1 & " " & "kOp\pQ"
            Case 16: TemString1 = TemString1 & " " & "vQk` kn @pwQ"
            Case 17: TemString1 = TemString1 & " " & "qQykrn @pwQ"
            End Select
                        
            If Dataenvironment1.rscmmdTemDrugs!dose = 0.25 Then
                TemString5 = Chr(188)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 0.5 Then
                TemString5 = Chr(189)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 0.75 Then
                TemString5 = Chr(190)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 0.4 Then
                TemString5 = "2/5"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 0.6 Then
                TemString5 = "3/5"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 0.4 And Dataenvironment1.rscmmdTemDrugs!dose > 0.3 Then
                TemString5 = "1/3"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 0.7 And Dataenvironment1.rscmmdTemDrugs!dose > 0.6 Then
                TemString5 = "2/3"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.25 Then
                TemString5 = "1 " & Chr(188)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.5 Then
                TemString5 = "1 " & Chr(189)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.75 Then
                TemString5 = "1 " & Chr(190)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 0.4 And Dataenvironment1.rscmmdTemDrugs!dose > 0.3 Then
                TemString5 = "1/3"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 0.7 And Dataenvironment1.rscmmdTemDrugs!dose > 0.6 Then
                TemString5 = "2/3"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.25 Then
                TemString5 = "1" & Chr(188)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.5 Then
                TemString5 = "1" & Chr(189)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.75 Then
                TemString5 = "1" & Chr(190)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.4 Then
                TemString5 = "1  2/5"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.6 Then
                TemString5 = "1  3/5"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 1.4 And Dataenvironment1.rscmmdTemDrugs!dose > 1.3 Then
                TemString5 = "1  1/3"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 1.7 And Dataenvironment1.rscmmdTemDrugs!dose > 1.6 Then
                TemString5 = "1  2/3"
            Else
                TemString5 = Dataenvironment1.rscmmdTemDrugs!dose
            End If
'            temString1 = temString1 & DataEnvironment1.rscmmdTemDrugs!dose & " " & "b#gQn\"
            TemString2 = ""
            Select Case Dataenvironment1.rscmmdTemDrugs!FrequencyID
            Case 1: TemString2 = "qvst @qvrk\"
            Case 2: TemString2 = "qvst wOn\vrk"
            Case 3: TemString2 = "qvst ek\ vrk\"
            Case 4: TemString2 = "p#y atkt vrk\"
            Case 5: TemString2 = "p#y hykt vrk\"
            Case 6: TemString2 = "p#y phkt vrk"
            Case 7: TemString2 = "p#ykt vrk\"
            Case 8: TemString2 = "p#y @qkkt vrk"
            Case 9: TemString2 = "p#y wOnkt vrk"
            Case 10: TemString2 = "u@q\t ek\ vrk\"
            Case 11: TemString2 = "rHt ek\ vrk\"
            Case 12: TemString2 = "svst ek\ vrk"
            Case 13: TemString2 = "av;& vEn@h`w"
            Case 14: TemString2 = "q#n\"
            Case 15: TemString2 = "av;& vEn@h`w"
            Case 16: TemString2 = "qvst @qvrk\"
            Case 17: TemString2 = "qvst wOn\vrk"
            Case 18: TemString2 = "swQ @qkkt vrk\"
            Case 19: TemString2 = "swQykt vrk\"
            Case 20: TemString2 = "m`sykt vrk\"
            Case 21: TemString2 = "swQ hwrkt vrk"
            Case 22: TemString2 = "s$m sqEq`km"
            Case 23: TemString2 = "s$m aghr#v`q`vkm"
            Case 24: TemString2 = "s$m bq`q`vkm"
            Case 25: TemString2 = "s$m bYhs\pwQn\q`vkm"
            Case 26: TemString2 = "s$m sQkOr`q`vkm"
            Case 27: TemString2 = "s$m @snsErq`vkm"
            Case 28: TemString2 = "s$m irQqvkm"
            End Select
            TemString2 = TemString2 & " " & "gn\n"
        RLabelTY = LabTopY + TextYMargin + LabelRowSerial * InbetweenY
        RLabelLX = ((RX - LX) / 2) + (TextXMargin / 3)
        RLabelRX = RX - (TextXMargin / 3)
            .Font.Name = LabelEnglishFontName
            .Font.Size = LabelEnglishFontSize
            TemString4 = PresentPatientFirstName
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            RLabelRowSerial = RLabelRowSerial + 2
            TemString4 = Dataenvironment1.rscmmdTemDrugs!drug
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            
            TemString4 = ""
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            RLabelRowSerial = RLabelRowSerial + 2
                    
            RLabelRowSerial = RLabelRowSerial + 2
                If TemString3 <> "" Then
                    .Font.Name = LabelSinhalaFontName
                    .Font.Size = LabelSinhalaFontSize
                     TemString4 = TemString3
                    .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
                    .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
                    Printer.Print TemString4
                    RLabelRowSerial = RLabelRowSerial + 2
                End If
                
            TemString4 = ""
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            RLabelRowSerial = RLabelRowSerial + 2
                
            .Font.Name = LabelSinhalaFontName
            .Font.Size = LabelSinhalaFontSize
             
             
             
             TemString4 = TemString1 & TemString5 & "b#gQn\" & "  "
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString1 & " ";
            .Font.Name = LabelEnglishFontName
            Printer.Print TemString5 & " ";
            .Font.Name = LabelSinhalaFontName
            Printer.Print "b#gQn\"
            RLabelRowSerial = RLabelRowSerial + 2
            TemString4 = TemString2
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            RLabelRowSerial = RLabelRowSerial + 4
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            RLabelRowSerial = RLabelRowSerial + 2
            RLabelBY = .CurrentY
            Printer.Line (RLabelLX, RLabelTY)-(RLabelRX, RLabelBY), , B
            
            If LLabelRowSerial > RLabelRowSerial Then
                LabelRowSerial = LabelRowSerial + LLabelRowSerial + 1
            Else
                LabelRowSerial = RLabelRowSerial + LabelRowSerial + 1
            End If
        
        End If
        IndoorSerial = IndoorSerial + 1
    
    Next i
    
.CurrentY = .CurrentY + InbetweenY
'If PreferancePrintingLines = True Then Printer.Line (RX, .CurrentY)-(LX, .CurrentY)
.CurrentY = .CurrentY + InbetweenY
.CurrentY = .CurrentY + InbetweenY

End With

InbetweenY = InbetweenY / 5


End Sub



Private Sub AddSingleLabels()
InbetweenY = InbetweenY * 5
Dim LLabelTY As Long
Dim LLabelBY As Long
Dim LLabelLX As Long
Dim LLabelRX As Long
Dim RLabelTY As Long
Dim RLabelBY As Long
Dim RLabelLX As Long
Dim RLabelRX As Long
Dim LabelRowSerial As Long
Dim RLabelRowSerial As Long
Dim LLabelRowSerial As Long

Dim LabelSinhalaFontName As String
Dim LabelSinhalaFontSize As Single
Dim LabelTamilFontName As String
Dim LabelTamilFontSize As Single
Dim LabelEnglishFontName As String
Dim LabelEnglishFontSize As Single

LabelSinhalaFontName = "kaputadotcom"
LabelSinhalaFontSize = 12
LabelEnglishFontName = "Verdana"
LabelEnglishFontSize = 9

LabelRowSerial = 0
RLabelRowSerial = 0
LLabelRowSerial = 0

With Printer
.CurrentY = LabTopY
Dataenvironment1.rscmmdTemDrugs.MoveFirst
IndoorSerial = 1

While Not Dataenvironment1.rscmmdTemDrugs.EOF

If Dataenvironment1.rscmmdTemDrugs!own = True Then

            TemString1 = ""
            TemString2 = ""
            TemString3 = ""
            TemString4 = ""
            TemString5 = ""
            RLabelRowSerial = 0
            TemString1 = "vrkt" ' varakata
            Select Case Dataenvironment1.rscmmdTemDrugs!DoseUnitID
            Case 1: TemString1 = TemString1 & " " & "mQlQ gY$m|"
            Case 2: TemString1 = TemString1 & " " & "@@m@kY` gY$m|"
            Case 3: TemString1 = TemString1 & " " & "mQlQ lWtr\"
                    TemString3 = "@b`n @b@hw"
            Case 4: TemString1 = TemString1 & " " & "@pwQ" ' tab
            Case 5: TemString1 = TemString1 & " " & "krl\" 'cap
            Case 6: TemString1 = TemString1 & " " & "@w\ h#qQ"
                    TemString3 = "@b`n @b@hw"
            Case 7: TemString1 = TemString1 & " " & "bQn\qE"
                    TemString3 = "@b`n @b@hw"
            Case 8: TemString1 = TemString1 & " " & "up`n\g"
            Case 9: TemString1 = TemString1 & " " & "ugOr#"
            Case 10: TemString1 = TemString1 & " " & "s@p`\sQtrQ" ' suppository
            Case 11: TemString1 = TemString1 & " " & "en\nw\" ' injections
            Case 12: TemString1 = TemString1 & " " & "v#jyQnl\ @pYsrQ"  'vaginal pressery
            Case 13: TemString1 = TemString1 & " " & "gY$m|"
            Case 14: TemString1 = TemString1 & " " & "@@m@kY` gY$m|"
            Case 15: TemString1 = TemString1 & " " & "kOp\pQ"
            Case 16: TemString1 = TemString1 & " " & "vQk` kn @pwQ"
            Case 17: TemString1 = TemString1 & " " & "qQykrn @pwQ"
            End Select
                        
            If Dataenvironment1.rscmmdTemDrugs!dose = 0.25 Then
                TemString5 = Chr(188)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 0.5 Then
                TemString5 = Chr(189)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 0.75 Then
                TemString5 = Chr(190)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 0.4 Then
                TemString5 = "2/5"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 0.6 Then
                TemString5 = "3/5"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 0.4 And Dataenvironment1.rscmmdTemDrugs!dose > 0.3 Then
                TemString5 = "1/3"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 0.7 And Dataenvironment1.rscmmdTemDrugs!dose > 0.6 Then
                TemString5 = "2/3"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.25 Then
                TemString5 = "1 " & Chr(188)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.5 Then
                TemString5 = "1 " & Chr(189)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.75 Then
                TemString5 = "1 " & Chr(190)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 0.4 And Dataenvironment1.rscmmdTemDrugs!dose > 0.3 Then
                TemString5 = "1/3"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 0.7 And Dataenvironment1.rscmmdTemDrugs!dose > 0.6 Then
                TemString5 = "2/3"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.25 Then
                TemString5 = "1" & Chr(188)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.5 Then
                TemString5 = "1" & Chr(189)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.75 Then
                TemString5 = "1" & Chr(190)
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.4 Then
                TemString5 = "1  2/5"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose = 1.6 Then
                TemString5 = "1  3/5"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 1.4 And Dataenvironment1.rscmmdTemDrugs!dose > 1.3 Then
                TemString5 = "1  1/3"
            ElseIf Dataenvironment1.rscmmdTemDrugs!dose < 1.7 And Dataenvironment1.rscmmdTemDrugs!dose > 1.6 Then
                TemString5 = "1  2/3"
            Else
                TemString5 = Dataenvironment1.rscmmdTemDrugs!dose
            End If
'            temString1 = temString1 & DataEnvironment1.rscmmdTemDrugs!dose & " " & "b#gQn\"
            TemString2 = ""
            Select Case Dataenvironment1.rscmmdTemDrugs!FrequencyID
            Case 1: TemString2 = "qvst @qvrk\"
            Case 2: TemString2 = "qvst wOn\vrk"
            Case 3: TemString2 = "qvst ek\ vrk\"
            Case 4: TemString2 = "p#y atkt vrk\"
            Case 5: TemString2 = "p#y hykt vrk\"
            Case 6: TemString2 = "p#y phkt vrk"
            Case 7: TemString2 = "p#ykt vrk\"
            Case 8: TemString2 = "p#y @qkkt vrk"
            Case 9: TemString2 = "p#y wOnkt vrk"
            Case 10: TemString2 = "u@q\t ek\ vrk\"
            Case 11: TemString2 = "rHt ek\ vrk\"
            Case 12: TemString2 = "svst ek\ vrk"
            Case 13: TemString2 = "av;& vEn@h`w"
            Case 14: TemString2 = "q#n\"
            Case 15: TemString2 = "av;& vEn@h`w"
            Case 16: TemString2 = "qvst @qvrk\"
            Case 17: TemString2 = "qvst wOn\vrk"
            Case 18: TemString2 = "swQ @qkkt vrk\"
            Case 19: TemString2 = "swQykt vrk\"
            Case 20: TemString2 = "m`sykt vrk\"
            Case 21: TemString2 = "swQ hwrkt vrk"
            Case 22: TemString2 = "s$m sqEq`km"
            Case 23: TemString2 = "s$m aghr#v`q`vkm"
            Case 24: TemString2 = "s$m bq`q`vkm"
            Case 25: TemString2 = "s$m bYhs\pwQn\q`vkm"
            Case 26: TemString2 = "s$m sQkOr`q`vkm"
            Case 27: TemString2 = "s$m @snsErq`vkm"
            Case 28: TemString2 = "s$m irQqvkm"
            End Select
            TemString2 = TemString2 & " " & "gn\n"
        RLabelTY = LabTopY + TextYMargin + LabelRowSerial * InbetweenY
        RLabelLX = LX + TextXMargin
        RLabelRX = RX - TextXMargin
            .Font.Name = LabelEnglishFontName
            .Font.Size = LabelEnglishFontSize
            TemString4 = PresentPatientFirstName
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            RLabelRowSerial = RLabelRowSerial + 2
            TemString4 = Dataenvironment1.rscmmdTemDrugs!drug
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            
            'TemString4 = ""
            '.CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            '.CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            'Printer.Print TemString4
            'RLabelRowSerial = RLabelRowSerial + 2
                    
            RLabelRowSerial = RLabelRowSerial + 2
                If TemString3 <> "" Then
                    .Font.Name = LabelSinhalaFontName
                    .Font.Size = LabelSinhalaFontSize
                     TemString4 = TemString3
                    .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
                    .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
                    Printer.Print TemString4
                    RLabelRowSerial = RLabelRowSerial + 2
                End If
                
            TemString4 = ""
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            RLabelRowSerial = RLabelRowSerial + 2
                
            .Font.Name = LabelSinhalaFontName
            .Font.Size = LabelSinhalaFontSize
             
             
             
             TemString4 = TemString1 & TemString5 & "b#gQn\" & "  "
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString1 & " ";
            .Font.Name = LabelEnglishFontName
            Printer.Print TemString5 & " ";
            .Font.Name = LabelSinhalaFontName
            Printer.Print "b#gQn\"
            RLabelRowSerial = RLabelRowSerial + 2
            TemString4 = TemString2
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            Printer.Print TemString4
            RLabelRowSerial = RLabelRowSerial + 4
            .CurrentY = RLabelTY + RLabelRowSerial * InbetweenY + TextYMargin
            .CurrentX = RLabelLX + ((RLabelRX - RLabelLX) / 2) - (.TextWidth(TemString4)) / 2
            RLabelRowSerial = RLabelRowSerial + 2
            RLabelBY = .CurrentY
            Printer.Line (RLabelLX, RLabelTY)-(RLabelRX, RLabelBY), , B
            
            If LLabelRowSerial > RLabelRowSerial Then
                LabelRowSerial = LabelRowSerial + LLabelRowSerial + 1
            Else
                LabelRowSerial = RLabelRowSerial + LabelRowSerial + 1
            End If
        IndoorSerial = IndoorSerial + 1
End If
Dataenvironment1.rscmmdTemDrugs.MoveNext
Wend
.CurrentY = .CurrentY + InbetweenY
'If PreferancePrintingLines = True Then Printer.Line (RX, .CurrentY)-(LX, .CurrentY)
.CurrentY = .CurrentY + InbetweenY
.CurrentY = .CurrentY + InbetweenY

End With

InbetweenY = InbetweenY / 5


End Sub

