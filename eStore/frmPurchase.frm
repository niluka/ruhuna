VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase"
   ClientHeight    =   11010
   ClientLeft      =   45
   ClientTop       =   435
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   11880
      TabIndex        =   127
      Top             =   6240
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.OptionButton optPUnits 
      Caption         =   "In Packs"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.OptionButton optIUnits 
      Caption         =   "In Issue Units"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtSPrice 
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtFQty 
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtBatch 
      Height          =   375
      Left            =   12720
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtQty 
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtPurchaseValue 
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtPPrice 
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   4575
      Left            =   11880
      TabIndex        =   83
      Top             =   1560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Prices"
      TabPicture(0)   =   "frmPurchase.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(3)=   "lblNetTotal"
      Tab(0).Control(4)=   "lblGrossTotal"
      Tab(0).Control(5)=   "Label23"
      Tab(0).Control(6)=   "frameCash"
      Tab(0).Control(7)=   "frameCredit"
      Tab(0).Control(8)=   "frameCheque"
      Tab(0).Control(9)=   "frameCreditCard"
      Tab(0).Control(10)=   "dtcPayment"
      Tab(0).Control(11)=   "txtDiscount"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Other"
      TabPicture(1)   =   "frmPurchase.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label22"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label21"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label24"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "dtcSupplier"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "dtcStaff"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "dtcChecked"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtInvoice"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.TextBox txtInvoice 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   1080
         TabIndex        =   30
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   -73800
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   960
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo dtcPayment 
         Height          =   360
         Left            =   -73800
         TabIndex        =   16
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcChecked 
         Height          =   360
         Left            =   1080
         TabIndex        =   32
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcStaff 
         Height          =   360
         Left            =   1080
         TabIndex        =   31
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcSupplier 
         Height          =   360
         Left            =   1080
         TabIndex        =   29
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Frame frameCreditCard 
         Caption         =   "Credit Card"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   114
         Top             =   2280
         Width           =   3015
         Begin VB.TextBox txtCreditCardNo 
            Height          =   375
            Left            =   1080
            TabIndex        =   19
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txtCreditCode 
            Height          =   375
            Left            =   1080
            TabIndex        =   20
            Top             =   1680
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo dtcCardBank 
            Height          =   360
            Left            =   1080
            TabIndex        =   18
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcCreditCard 
            Height          =   360
            Left            =   1080
            TabIndex        =   17
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label50 
            Caption         =   "No"
            Height          =   255
            Left            =   120
            TabIndex        =   118
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label49 
            Caption         =   "Card"
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label48 
            Caption         =   "Bank"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label47 
            Caption         =   "Code"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   1680
            Width           =   1575
         End
      End
      Begin VB.Frame frameCheque 
         Caption         =   "Cheque"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   109
         Top             =   2280
         Width           =   3015
         Begin VB.TextBox txtChequeNo 
            Height          =   375
            Left            =   1080
            TabIndex        =   23
            Top             =   1200
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker dtpChequeDate 
            Height          =   375
            Left            =   1080
            TabIndex        =   24
            Top             =   1680
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   69074947
            CurrentDate     =   39551
         End
         Begin MSDataListLib.DataCombo dtcBranch 
            Height          =   360
            Left            =   1080
            TabIndex        =   22
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcBank 
            Height          =   360
            Left            =   1080
            TabIndex        =   21
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label46 
            Caption         =   "Branch"
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label45 
            Caption         =   "Bank"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label44 
            Caption         =   "No"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label43 
            Caption         =   "Date"
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   1680
            Width           =   1575
         End
      End
      Begin VB.Frame frameCredit 
         Caption         =   "Credit"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   107
         Top             =   2280
         Width           =   3015
         Begin VB.TextBox txtCreditDue 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1080
            TabIndex        =   25
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label42 
            Caption         =   "Due"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame frameCash 
         Caption         =   "Cash"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   103
         Top             =   2280
         Width           =   3015
         Begin VB.TextBox txtDue 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1080
            TabIndex        =   26
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtCashPaid 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1080
            TabIndex        =   27
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1080
            TabIndex        =   28
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label41 
            Caption         =   "Due"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label40 
            Caption         =   "Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label39 
            Caption         =   "Change"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.Label Label24 
         Caption         =   "Supplier"
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Received by"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Checked by"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Invoice No."
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Payment"
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblGrossTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   -73800
         TabIndex        =   87
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   -73800
         TabIndex        =   85
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Gross Total"
         Height          =   255
         Left            =   -74880
         TabIndex        =   88
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Net Total"
         Height          =   255
         Left            =   -74880
         TabIndex        =   86
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Discount"
         Height          =   255
         Left            =   -74880
         TabIndex        =   84
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.TextBox txtDataEntry 
      Height          =   375
      Left            =   2760
      TabIndex        =   82
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin btButtonEx.ButtonEx bttnReceive 
      Height          =   375
      Left            =   11880
      TabIndex        =   33
      Top             =   6600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Purchase"
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
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   4335
      Left            =   120
      TabIndex        =   35
      Top             =   2640
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7646
      _Version        =   393216
      WordWrap        =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   36
      Top             =   7080
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "frmPurchase.frx":0038
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtIStore"
      Tab(0).Control(1)=   "txtDisplay"
      Tab(0).Control(2)=   "txtVTM"
      Tab(0).Control(3)=   "txtVMP"
      Tab(0).Control(4)=   "txtAMP"
      Tab(0).Control(5)=   "txtVMPP"
      Tab(0).Control(6)=   "txtAMPP"
      Tab(0).Control(7)=   "Label14"
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(9)=   "Label9"
      Tab(0).Control(10)=   "Label10"
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(12)=   "Label12"
      Tab(0).Control(13)=   "Label13"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Stocks"
      TabPicture(1)   =   "frmPurchase.frx":0054
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridTotal"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Usage"
      TabPicture(2)   =   "frmPurchase.frx":0070
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dtpUFrom"
      Tab(2).Control(1)=   "GridUsage"
      Tab(2).Control(2)=   "dtpUTo"
      Tab(2).Control(3)=   "Label3"
      Tab(2).Control(4)=   "Label7"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Ordering"
      TabPicture(3)   =   "frmPurchase.frx":008C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dtpOFrom"
      Tab(3).Control(1)=   "GridOrdering"
      Tab(3).Control(2)=   "dtpOTo"
      Tab(3).Control(3)=   "Label8"
      Tab(3).Control(4)=   "Label15"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Prices"
      TabPicture(4)   =   "frmPurchase.frx":00A8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label19"
      Tab(4).Control(1)=   "Label18"
      Tab(4).Control(2)=   "Label17"
      Tab(4).Control(3)=   "Label16"
      Tab(4).Control(4)=   "dtpPTo"
      Tab(4).Control(5)=   "dtpPFrom"
      Tab(4).Control(6)=   "GridSPrice"
      Tab(4).Control(7)=   "GridPPrice"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Distributor"
      TabPicture(5)   =   "frmPurchase.frx":00C4
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "lblDistributor"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label20"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label25"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label26"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label27"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "lblBalance"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "lblTelNo"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "lblAddress"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Label31"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "lblFax"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "lblCity"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "Label33"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).ControlCount=   12
      Begin VB.TextBox txtIStore 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   3360
         Width           =   6615
      End
      Begin VB.TextBox txtDisplay 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   480
         Width           =   6615
      End
      Begin VB.TextBox txtVTM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   960
         Width           =   6615
      End
      Begin VB.TextBox txtVMP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1440
         Width           =   6615
      End
      Begin VB.TextBox txtAMP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   1920
         Width           =   6615
      End
      Begin VB.TextBox txtVMPP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2400
         Width           =   6615
      End
      Begin VB.TextBox txtAMPP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2880
         Width           =   6615
      End
      Begin MSFlexGridLib.MSFlexGrid GridPPrice 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   37
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3836
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpUFrom 
         Height          =   375
         Left            =   -74040
         TabIndex        =   38
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   69074947
         CurrentDate     =   39540
      End
      Begin MSFlexGridLib.MSFlexGrid GridUsage 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   39
         Top             =   840
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   4260
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridTotal 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   40
         Top             =   360
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5106
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpUTo 
         Height          =   375
         Left            =   -71040
         TabIndex        =   48
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   69074947
         CurrentDate     =   39540
      End
      Begin MSComCtl2.DTPicker dtpOFrom 
         Height          =   375
         Left            =   -74040
         TabIndex        =   49
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   69074947
         CurrentDate     =   39540
      End
      Begin MSFlexGridLib.MSFlexGrid GridOrdering 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   50
         Top             =   840
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpOTo 
         Height          =   375
         Left            =   -71040
         TabIndex        =   51
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   69074947
         CurrentDate     =   39540
      End
      Begin MSFlexGridLib.MSFlexGrid GridSPrice 
         Height          =   2175
         Left            =   -66840
         TabIndex        =   52
         Top             =   1080
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3836
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpPFrom 
         Height          =   375
         Left            =   -74160
         TabIndex        =   53
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   69074947
         CurrentDate     =   39540
      End
      Begin MSComCtl2.DTPicker dtpPTo 
         Height          =   375
         Left            =   -71160
         TabIndex        =   54
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   69074947
         CurrentDate     =   39540
      End
      Begin VB.Label Label14 
         Caption         =   "Store :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   81
         Top             =   3360
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "Display Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   80
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Virtual Therapeutic Moiety:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   79
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "Virtual Medicinal Product:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   78
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Actual Medicinal Product:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   77
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Virtual Medicinal Product Pack :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   76
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label Label13 
         Caption         =   "Actual Medicinal Product Pack :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   75
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "From :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   74
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "To :"
         Height          =   255
         Left            =   -71520
         TabIndex        =   73
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "To :"
         Height          =   255
         Left            =   -71520
         TabIndex        =   72
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "From :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   71
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label16 
         Caption         =   "Sales Prices"
         Height          =   255
         Left            =   -66840
         TabIndex        =   70
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label17 
         Caption         =   "Purchase Prices"
         Height          =   255
         Left            =   -74880
         TabIndex        =   69
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label18 
         Caption         =   "To :"
         Height          =   255
         Left            =   -71640
         TabIndex        =   68
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label19 
         Caption         =   "From :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   67
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label33 
         Caption         =   "City"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblCity 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   65
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label lblFax 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   64
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label Label31 
         Caption         =   "Fax No"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   1680
         TabIndex        =   62
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label lblTelNo 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   61
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   1680
         TabIndex        =   60
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Label Label27 
         Caption         =   "Balance"
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "Tel No"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Address"
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Distributor"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblDistributor 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   55
         Top             =   480
         Width           =   3375
      End
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   375
      Left            =   13560
      TabIndex        =   34
      Top             =   6600
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSDataListLib.DataCombo dtcCatogery 
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   120
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
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCode 
      Height          =   360
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDOM 
      Height          =   375
      Left            =   12720
      TabIndex        =   11
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM yyyy"
      Format          =   69074947
      CurrentDate     =   39545
   End
   Begin MSComCtl2.DTPicker dtpDOE 
      Height          =   375
      Left            =   12720
      TabIndex        =   12
      Top             =   1080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM yyyy"
      Format          =   69074947
      CurrentDate     =   39545
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   10320
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   375
      Left            =   10320
      TabIndex        =   14
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin VB.Label Label57 
      Caption         =   "Quentity Unit"
      Height          =   375
      Left            =   120
      TabIndex        =   126
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label56 
      Caption         =   "Sale Price"
      Height          =   495
      Left            =   5280
      TabIndex        =   125
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label55 
      Height          =   375
      Left            =   8640
      TabIndex        =   124
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblFQtyUnit 
      Height          =   375
      Left            =   8640
      TabIndex        =   123
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label54 
      Caption         =   "Free Quantity"
      Height          =   375
      Left            =   5280
      TabIndex        =   122
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblSPriceUnit 
      Height          =   375
      Left            =   8640
      TabIndex        =   121
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblPPriceUnit 
      Height          =   375
      Left            =   8640
      TabIndex        =   120
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblQtyUnit 
      Height          =   375
      Left            =   8640
      TabIndex        =   119
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label38 
      Caption         =   "Batch"
      Height          =   375
      Left            =   10800
      TabIndex        =   102
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label37 
      Caption         =   "Date of Manufacture"
      Height          =   375
      Left            =   10800
      TabIndex        =   101
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label36 
      Caption         =   "Date of Expiary"
      Height          =   375
      Left            =   10800
      TabIndex        =   100
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label35 
      Caption         =   "Code"
      Height          =   375
      Left            =   120
      TabIndex        =   99
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "&Item"
      Height          =   375
      Left            =   120
      TabIndex        =   98
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label32 
      Caption         =   "Quantity"
      Height          =   375
      Left            =   5280
      TabIndex        =   97
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label30 
      Caption         =   "Purchase Value"
      Height          =   375
      Left            =   5280
      TabIndex        =   96
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label29 
      Caption         =   "Purchase Price"
      Height          =   375
      Left            =   5280
      TabIndex        =   95
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label28 
      Caption         =   "Catogery"
      Height          =   375
      Left            =   120
      TabIndex        =   94
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    
    Dim CsetPrinter As New cSetDfltPrinter
    
    Dim TemOrderBillID As Long
    Dim TemDistributorId As Long
    Dim TemDistributorOrderID As Long
    Dim EditingData As Boolean
    Dim TemContent(22) As String
    Dim CurrentRow As Integer
    Dim TemCellContent As String
    Dim temRefillBillID As Long
    
    Dim NewItem As New Item
    
    Dim rsStaff As New ADODB.Recordset
    Dim rsSPrice As New ADODB.Recordset
    Dim rsPPrice As New ADODB.Recordset
    Dim rsCC As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCategory As New ADODB.Recordset
    Dim rsCode As New ADODB.Recordset
    Dim rsBanks As New ADODB.Recordset
    Dim rsCreditCards As New ADODB.Recordset
    Dim rsCities As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsDistributor As New ADODB.Recordset
    
    Dim rsTemOrder As New ADODB.Recordset
    Dim rsTemPrice As New ADODB.Recordset
    Dim rsTemDistributor As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    Dim rsTemOrderBill As New ADODB.Recordset
    Dim rsTemDistributorOrder As New ADODB.Recordset
    Dim rsTemRefill As New ADODB.Recordset
    Dim rsTemRefillBill As New ADODB.Recordset
    Dim rsTemCash As New ADODB.Recordset
    Dim rsTemCredit As New ADODB.Recordset
    Dim rsTemCheque As New ADODB.Recordset
    
Private Sub bttnDelete_Click()
    If GridItem.Rows <= 1 Then Exit Sub
    If GridItem.Rows = 2 Then
        FormatGrid
    Else
        GridItem.RemoveItem (GridItem.Row)
    End If
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
    temSQL = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by display"
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "Display"
    .BoundColumn = "ItemID"
End With
With rsCode
    If .State = 1 Then .Close
    temSQL = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by code"
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
    If KeyCode = vbKeyEscape Then
        dtcCatogery.Text = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcItem.SetFocus
    End If
End Sub


Private Sub dtcCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtQty.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtcCode.Text = Empty
    End If
End Sub

Private Sub dtcItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcCode.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtcItem.Text = Empty
    End If
End Sub

Private Sub dtcPayment_Click(Area As Integer)
    Select Case dtcPayment.Text
        Case "Cash":
            frameCash.Visible = True
            frameCheque.Visible = False
            frameCredit.Visible = False
            frameCreditCard.Visible = False
        Case "Credit":
            frameCash.Visible = False
            frameCheque.Visible = False
            frameCredit.Visible = True
            frameCreditCard.Visible = False
        Case "Cheque":
            frameCash.Visible = False
            frameCheque.Visible = True
            frameCredit.Visible = False
            frameCreditCard.Visible = False
        Case Else
            frameCash.Visible = False
            frameCheque.Visible = False
            frameCredit.Visible = False
            frameCreditCard.Visible = False
    End Select
End Sub

Private Sub dtcSupplier_Click(Area As Integer)
    If IsNumeric(dtcSupplier.BoundText) = False Then Exit Sub
    TemDistributorId = dtcSupplier.BoundText
    DistributorDetails (TemDistributorId)
End Sub





Private Sub dtpDOE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        bttnAdd_Click
    End If
End Sub

Private Sub dtpDOM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpDOE.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    Call SetValues
    GridItem.RowHeight(0) = GridItem.RowHeight(0) * 3
End Sub

Private Sub SetValues()
    dtpDOE.Value = Date
    dtpDOM.Value = Date
    dtpDOE.MinDate = LastDateOfMonth(Date)
    optPUnits.Value = True
    dtcStaff.BoundText = UserID
    dtcChecked.BoundText = UserID
    dtcStaff.Locked = True
    frameCash.Visible = False
    frameCheque.Visible = False
    frameCredit.Visible = False
    frameCreditCard.Visible = False
    dtpOFrom.Value = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
    dtpOTo.Value = Date
    dtpPFrom.Value = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
    dtpPTo.Value = Date
    dtpUFrom.Value = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
    dtpUTo.Value = Date
End Sub

Private Sub FillCombos()
    With rsStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblstaff order by listedname"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcStaff
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With dtcChecked
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With rsDistributor
        If .State = 1 Then .Close
        temSQL = "SELECT tblDistrubutor.* From tblDistrubutor ORDER BY tblDistrubutor.DistributorName"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcSupplier
        Set .RowSource = rsDistributor
        .ListField = "DistributorName"
        .BoundColumn = "DistributorID"
    End With
    With rsCC
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblpaymentMethod " & _
                    "ORDER BY PaymentMethod"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcPayment
        Set .RowSource = rsCC
        .ListField = "PaymentMethod"
        .BoundColumn = "PaymentMethodID"
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
    With rsItemCategory
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblItemCategory order by ItemCategory"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCatogery
        Set .RowSource = rsItemCategory
        .ListField = "ItemCategory"
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
End Sub
    
Private Sub FormatGrid()
    EditingData = False
    With GridItem
        .Cols = 22
        .Rows = 1
        .Row = 0
        .Col = 0
        .FixedCols = 0
        
'        .RowHeight(0) = .RowHeight(0) * 3
        
        Dim i As Integer
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            Select Case i
                Case 0:     .Text = "No"
                            .ColWidth(i) = 400
                Case 1:     .Text = "Item"
                            .ColWidth(i) = 3600
                Case 5:     .Text = "Purchased"
                            .ColWidth(i) = 900
                Case 6:     .Text = "Unit"
                            .ColWidth(i) = 900
                Case 7:     .Text = "Free"
                            .ColWidth(i) = 900
                Case 8:     .Text = "Unit"
                            .ColWidth(i) = 900
                Case 9:     .Text = "Batch"
                            .ColWidth(i) = 900
                Case 10:     .Text = "Pruchase Price Per Pack"
                            .ColWidth(i) = 900
                Case 11:     .Text = "Slaes Price Per Unit Sale"
                            .ColWidth(i) = 900
                Case 18:    .ColWidth(i) = 1200
                            .Text = "Total Pruchase Value"
                Case Else:  .ColWidth(i) = 1
            End Select
        Next i
    
    End With
    '   0   No
    '   1   Item
    '   2   ItemID
    '   3   PackUnitID
    '   4   IssueUnitID
    '   5   PurchaseQuentity
    '   6   PUnit
    '   7   FreeQuentity
    '   8   PUnit
    '   9   Batch
    '   10  Purchase Price
    '   11  Sales Price
    '   12  Sales Margin
    '   13
    '   14
    '   15  IPurchased
    '   16  IFreePurchased
    '   17  IUnitsPerPack
    '   18  Display Price
    '   19  Actual Price
    '   20  DOM
    '   21  DOE
    EditingData = True
End Sub

 Private Sub bttnCancel_Click()
    Unload Me
End Sub
   
Private Sub bttnAdd_Click()
    If CanAdd = False Then Exit Sub
    EditingData = False
    With GridItem
        .Rows = .Rows + 1
        .Row = .Rows - 1
        
        .Col = 0
        .CellAlignment = 7
        .Text = .Row
        
        .Col = 1
        .CellAlignment = 1
        .Text = NewItem.Display
        
        .Col = 2
        .Text = NewItem.ID
        
        .Col = 3
        .Text = NewItem.PUnitID
        
        .Col = 4
        .Text = NewItem.IUnitID
        
        .Col = 5
        .CellAlignment = 7
        .Text = txtQty.Text
        
        .Col = 6
        .CellAlignment = 1
        If optIUnits.Value = True Then
            .Text = NewItem.IUnit
        Else
            .Text = NewItem.PUnit
        End If
        
        .Col = 7
        .CellAlignment = 7
        .Text = txtFQty.Text
        
        .Col = 8
        .CellAlignment = 1
        If optIUnits.Value = True Then
            .Text = NewItem.IUnit
        Else
            .Text = NewItem.PUnit
        End If
        
        .Col = 9
        .CellAlignment = 7
        .Text = txtBatch.Text
        
        .Col = 10
        .CellAlignment = 7
        .Text = Format(Val(txtPPrice.Text), "0.00")
        
        .Col = 11
        .CellAlignment = 7
        .Text = Format((Val(txtSPrice.Text)), "0.00")
        
        .Col = 12
        .CellAlignment = 7
        .Text = NewItem.SalesMargin
        
        .Col = 13
        .Text = Empty
        
        .Col = 14
        .Text = Empty
        
        .Col = 15
        If optIUnits.Value = True Then
            .Text = Val(txtQty.Text)
        Else
            .Text = Val(txtQty.Text) * NewItem.IssueUnitsPerPack
        End If
        
        .Col = 16
        If optIUnits.Value = True Then
            .Text = Val(txtFQty.Text)
        Else
            .Text = Val(txtFQty.Text) * NewItem.IssueUnitsPerPack
        End If
        
        .Col = 17
        .Text = NewItem.IssueUnitsPerPack
        
        .Col = 18
        If optIUnits.Value = True Then
            .Text = Format((Val(txtQty.Text) * NewItem.IssueUnitsPerPack * Val(txtPPrice.Text)), "#,##0.00")
        Else
            .Text = Format((Val(txtQty.Text) * Val(txtPPrice.Text)), "#,##0.00")
        End If
        
        .Col = 19
        If optIUnits.Value = True Then
            .Text = Val(txtQty.Text) * NewItem.IssueUnitsPerPack * Val(txtPPrice.Text)
        Else
            .Text = Val(txtQty.Text) * Val(txtPPrice.Text)
        End If
        
        .Col = 20
        .CellAlignment = 4
        .Text = LastDateOfMonth(dtpDOM.Value)
        
        .Col = 21
        .CellAlignment = 7
        .Text = LastDateOfMonth(dtpDOE.Value)
    End With
    Call ClearAddValues
    Call ClearItemDetails
    Call ClearGrids
    Call CalculateTotal
    dtcItem.SetFocus
    EditingData = True
End Sub
    

Private Sub ClearAddValues()
    txtQty.Text = Empty
    txtPPrice.Text = Empty
    txtSPrice.Text = Empty
    txtFQty.Text = Empty
    txtPurchaseValue.Text = Empty
    dtcItem.Text = Empty
    dtcCatogery.Text = Empty
    dtcCode.Text = Empty
    txtBatch.Text = Empty
End Sub

Private Sub ClearItemDetails()
    txtVMP.Text = Empty
    txtVMPP.Text = Empty
    txtVTM.Text = Empty
    txtAMP.Text = Empty
    txtAMPP.Text = Empty
    txtDisplay.Text = Empty
End Sub

Private Sub ClearGrids()
    With GridOrdering
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridPPrice
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridSPrice
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridTotal
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
    With GridUsage
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = .Width
    End With
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
        If dtpDOE.Value = Date Then
            tr = MsgBox("You have not entered a Date of Expiary", vbCritical, "Expiary Date")
            dtpDOE.SetFocus
            Exit Function
        End If
        If Trim(txtBatch.Text) = Empty Then
            tr = MsgBox("You have not entered a Batch number", vbCritical, "Expiary Date")
            txtBatch.SetFocus
            Exit Function
        End If
        If Val(txtPPrice.Text) = 0 Then
            tr = MsgBox("You have not entered the purchase price", vbCritical, "Purchase Price")
            txtPPrice.SetFocus
            Exit Function
        End If
        If Val(txtSPrice.Text) = 0 Then
            tr = MsgBox("You have not entered the sale price", vbCritical, "Purchase Price")
            txtSPrice.SetFocus
            Exit Function
        End If
        If Val(txtPPrice.Text) >= Val(txtSPrice.Text) * NewItem.IssueUnitsPerPack Then
            tr = MsgBox("You can't sell items at a rate below the purchase rate", vbCritical, "Adjust Sale Price")
            txtSPrice.SetFocus
            Exit Function
        End If
    CanAdd = True
End Function
    
Private Sub dtcItem_Change()
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    dtcCode.BoundText = dtcItem.BoundText
    NewItem.ID = dtcItem.BoundText
    Call FillLabels
    Call GetItemDetails(NewItem.ID)
    Call FillStocks(dtcItem.BoundText)
    Call FillPrice(dtcItem.BoundText)
    Call GetItemDetails(dtcItem.BoundText)
End Sub
    
Private Sub FillLabels()
    If optIUnits.Value = True Then
        lblQtyUnit.Caption = NewItem.IUnit
        lblFQtyUnit.Caption = NewItem.IUnit

    Else
        lblQtyUnit.Caption = NewItem.PUnit
        lblFQtyUnit.Caption = NewItem.PUnit

    End If
    lblPPriceUnit.Caption = "Per " & NewItem.PUnit
    lblSPriceUnit.Caption = "Per " & NewItem.IUnit

End Sub


Private Sub GridItem_DblClick()
    With GridItem
        If IsNumeric(.TextMatrix(.Row, 2)) = False Then Exit Sub
        dtcItem.BoundText = .TextMatrix(.Row, 2)
        If optIUnits.Value = True Then
            txtQty.Text = .TextMatrix(.Row, 15)
            txtFQty.Text = .TextMatrix(.Row, 16)
        Else
            txtQty.Text = .TextMatrix(.Row, 15) / NewItem.IssueUnitsPerPack
            txtFQty.Text = .TextMatrix(.Row, 16) / NewItem.IssueUnitsPerPack
        End If
        txtPPrice.Text = .TextMatrix(.Row, 10)
        txtSPrice.Text = .TextMatrix(.Row, 11)
        txtBatch.Text = .TextMatrix(.Row, 9)
        dtpDOM.Value = .TextMatrix(.Row, 20)
        dtpDOE.Value = .TextMatrix(.Row, 21)
    End With
    bttnDelete_Click
End Sub

Private Sub lblNetTotal_Change()
    txtCreditDue.Text = lblNetTotal.Caption
    txtDue.Text = lblNetTotal.Caption
End Sub

Private Sub optIUnits_Click()
    Call FillLabels
End Sub

Private Sub optPUnits_Click()
    Call FillLabels
End Sub


Private Sub txtBatch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpDOM.SetFocus
    End If
End Sub

Private Sub txtCashPaid_Change()
    Call CalculateBalance
End Sub

Private Sub txtDue_Change()
    Call CalculateBalance
End Sub

Private Sub CalculateBalance()
    txtBalance.Text = Format((Val(txtCashPaid.Text) - Val(txtDue.Text)), "0.00")
End Sub

Private Sub txtFQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPPrice.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtFQty.Text = Empty
    End If
End Sub

Private Sub txtPPrice_Change()
    Call CalculatePurchaseValue
    Call CalculateSalePrice
End Sub

Private Sub txtPPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtSPrice.SetFocus
    End If
End Sub

Private Sub txtQty_Change()
    Call CalculatePurchaseValue
End Sub
    
    
Private Sub CalculatePurchaseValue()
    If optIUnits.Value = True Then
        txtPurchaseValue.Text = Format(((Val(txtQty.Text) / NewItem.IssueUnitsPerPack) * Val(txtPPrice.Text)), "0.00")
    Else
        txtPurchaseValue.Text = Format(((Val(txtQty.Text)) * Val(txtPPrice.Text)), "0.00")
    End If
End Sub
    
Private Sub CalculateSalePrice()
    If NewItem.ID <> 0 Then
        txtSPrice.Text = Format((((Val(txtPPrice.Text) * (NewItem.SalesMargin + 100)) / 100) / NewItem.IssueUnitsPerPack), "0.00")
    End If
End Sub
    
    
Private Function CanReceive() As Boolean
    Dim i As Integer
    Dim tr As Integer
    CanReceive = False
    
    If GridItem.Rows <= 1 Then
        tr = MsgBox("There are no items to sell", vbCritical, "No Items")
        dtcItem.SetFocus
        Exit Function
    End If
    
    If IsNumeric(dtcPayment.BoundText) = False Then
        tr = MsgBox("You have not selected the payment method", vbCritical, "No Items")
        SSTab2.Tab = 0
        dtcPayment.SetFocus
        Exit Function
    End If
    
    If IsNumeric(dtcSupplier.BoundText) = False Then
        tr = MsgBox("You have not selected the supplier", vbCritical, "No Supplier")
        SSTab2.Tab = 0
        dtcSupplier.SetFocus
        Exit Function
    End If
    
    If dtcPayment.Text = "Cash" Then
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
        
    ElseIf dtcPayment.Text = "Credit" = True Then
    
    ElseIf dtcPayment.Text = "Cheque" Then
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
    Else
        tr = MsgBox("You have not selected a Valid Payment Method", vbCritical, "Payment Method?")
        SSTab2.Tab = 1
        dtcPayment.SetFocus
        Exit Function
    End If
    
    If IsNumeric(dtcStaff.BoundText) = False Then
        tr = MsgBox("You have not selected the user", vbCritical, "Issued by?")
        SSTab2.Tab = 0
        dtcStaff.SetFocus
        Exit Function
    End If
    
    If IsNumeric(dtcChecked.BoundText) = False Then
        tr = MsgBox("You have not selected the name of the checked staff member", vbCritical, "Checked by?")
        SSTab2.Tab = 0
        dtcChecked.SetFocus
        Exit Function
    End If
    
    CanReceive = True
End Function
    
Private Sub bttnReceive_Click()
    If CanReceive = False Then Exit Sub
    Dim tr As Integer
    Dim i As Integer
    Dim DiscountPercent As Double
    
    With rsTemRefillBill
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblRefillBill"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !DistributorID = dtcSupplier.BoundText
        !StoreID = UserStoreID
        !StaffID = UserID
        If IsNumeric(dtcChecked.BoundText) = True Then
            !CheckedStaffID = dtcChecked.BoundText
        End If
        !Price = Val(lblGrossTotal.Caption)
        !DIscount = Val(txtDiscount.Text)
        DiscountPercent = (Val(txtDiscount.Text) / Val(lblGrossTotal.Caption)) * 100
        !DiscountPercent = DiscountPercent
        !NetPrice = Val(lblNetTotal.Caption)
        !Date = Date
        !Time = Time
        !PaymentMethodID = dtcPayment.BoundText
        !PaymentMethod = dtcPayment.Text
        If dtcPayment.Text = "Credit" Then
            !FullyPaid = False
        Else
            !FullyPaid = True
        End If
        !purchase = True
        !Autorequest = False
        !ManualRequest = False
        .Update
        temRefillBillID = !RefillBillID
        If dtcPayment.Text = "Cash" Then
            !IssuedCashID = IssueCash(temRefillBillID)
        ElseIf dtcPayment.Text = "Credit" Then
            !IssuedCreditID = IssueCredit(temRefillBillID)
        ElseIf dtcPayment.Text = "Cheque" Then
            !IssuedChequeID = IssueCheque(temRefillBillID)
        End If
        .Update
        .Close
    End With
    
    With GridItem
        For i = 1 To .Rows - 1
            If rsTemRefill.State = 1 Then rsTemRefill.Close
            temSQL = "SELECT * FROM tblRefill"
            rsTemRefill.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            rsTemRefill.AddNew
            rsTemRefill!ItemID = Val(.TextMatrix(i, 2))
            rsTemRefill!StoreID = UserStoreID
            rsTemRefill!Date = Date
            rsTemRefill!Time = Time
            rsTemRefill!StaffID = UserID
            rsTemRefill!DistributorID = dtcSupplier.BoundText
            rsTemRefill!Price = Val(.TextMatrix(i, 19))
            rsTemRefill!DiscountPercent = DiscountPercent
            rsTemRefill!NetPrice = (Val(.TextMatrix(i, 19))) - (Val(.TextMatrix(i, 19)) * DiscountPercent / 100)
            rsTemRefill!RefillBillID = temRefillBillID
            rsTemRefill!purchase = True
            rsTemRefill!Autorequest = False
            rsTemRefill!ManualRequest = False
            rsTemRefill!Amount = Val(.TextMatrix(i, 15))
            rsTemRefill!FreeAmount = Val(.TextMatrix(i, 16))
            rsTemRefill!CheckedStaffID = dtcChecked.BoundText
            
            Dim ThisBatch As Long
            ThisBatch = BatchExist(.TextMatrix(i, 9), Val(.TextMatrix(i, 2)))
            If ThisBatch <> 0 Then
                rsTemRefill!BatchID = ThisBatch
                If AddToStock(ThisBatch, UserStoreID, Val(.TextMatrix(i, 15)) + Val(.TextMatrix(i, 15))) = False Then
                    MsgBox "Error"
                    Exit For
                End If
            Else
                ThisBatch = AddBatch(.TextMatrix(i, 9), Val(.TextMatrix(i, 2)), .TextMatrix(i, 20), .TextMatrix(i, 21))
                rsTemRefill!BatchID = ThisBatch
                
                If AddToStock(ThisBatch, UserStoreID, Val(.TextMatrix(i, 15)) + Val(.TextMatrix(i, 16))) = False Then
                    MsgBox "Error"
                    Exit For
                End If
            End If
            rsTemRefill.Update
            rsTemRefill.Close
            
            If rsSPrice.State = 1 Then rsSPrice.Close
            temSQL = "SELECT tblSalePrice.ItemID, tblSalePrice.SPrice, tblSalePrice.SetDate, tblSalePrice.SetTime, tblSalePrice.StaffID FROM tblSalePrice "
            rsSPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            rsSPrice.AddNew
            rsSPrice!ItemID = Val(.TextMatrix(i, 2))
            rsSPrice!SPrice = Val(.TextMatrix(i, 11))
            rsSPrice!setdate = Date
            rsSPrice!SetTime = Time
            rsSPrice!StaffID = UserID
            rsSPrice.Update
            rsSPrice.Close
            
            If rsSPrice.State = 1 Then rsSPrice.Close
            temSQL = "SELECT * FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
            rsSPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If rsSPrice.RecordCount < 1 Then
                rsSPrice.AddNew
                rsSPrice!ItemID = Val(.TextMatrix(i, 2))
                rsSPrice!SPrice = Val(.TextMatrix(i, 11))
                rsSPrice!setdate = Date
                rsSPrice!SetTime = Time
                rsSPrice!StaffID = UserID
                rsSPrice.Update
            ElseIf rsSPrice.RecordCount = 1 Then
                rsSPrice!SPrice = Val(.TextMatrix(i, 11))
                rsSPrice!setdate = Date
                rsSPrice!SetTime = Time
                rsSPrice!StaffID = UserID
                rsSPrice.Update
            Else
                If rsSPrice.State = 1 Then rsSPrice.Close
                temSQL = "Delete * FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
                rsSPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                If rsSPrice.State = 1 Then rsSPrice.Close
                temSQL = "SELECT * FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
                rsSPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                rsSPrice.AddNew
                rsSPrice!ItemID = Val(.TextMatrix(i, 2))
                rsSPrice!SPrice = Val(.TextMatrix(i, 11))
                rsSPrice!setdate = Date
                rsSPrice!SetTime = Time
                rsSPrice!StaffID = UserID
                rsSPrice.Update
            End If
            rsSPrice.Close
            
            If rsPPrice.State = 1 Then rsPPrice.Close
            temSQL = "SELECT tblPurchasePrice.ItemID, tblPurchasePrice.PPrice, tblPurchasePrice.SetDate, tblPurchasePrice.SetTime, tblPurchasePrice.StaffID FROM tblPurchasePrice"
            rsPPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            rsPPrice.AddNew
            rsPPrice!ItemID = Val(.TextMatrix(i, 2))
            rsPPrice!PPrice = Val(.TextMatrix(i, 10)) / NewItem.IssueUnitsPerPack
            rsPPrice!setdate = Date
            rsPPrice!SetTime = Time
            rsPPrice!StaffID = UserID
            rsPPrice.Update
            rsPPrice.Close
            
            If rsPPrice.State = 1 Then rsPPrice.Close
            temSQL = "SELECT * FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2)) & " Order by SetDate Desc, SetTime Desc"
            rsPPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If rsPPrice.RecordCount < 1 Then
                rsPPrice.AddNew
                rsPPrice!ItemID = Val(.TextMatrix(i, 2))
                rsPPrice!PPrice = Val(.TextMatrix(i, 10)) / NewItem.IssueUnitsPerPack
                rsPPrice!setdate = Date
                rsPPrice!SetTime = Time
                rsPPrice!StaffID = UserID
                rsPPrice.Update
            ElseIf rsPPrice.RecordCount = 1 Then
                rsPPrice!PPrice = Val(.TextMatrix(i, 10)) / NewItem.IssueUnitsPerPack
                rsPPrice!setdate = Date
                rsPPrice!SetTime = Time
                rsPPrice!StaffID = UserID
                rsPPrice.Update
            Else
                If rsPPrice.State = 1 Then rsPPrice.Close
                temSQL = "Delete * FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2)) & " Order by SetDate Desc, SetTime Desc"
                rsPPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                If rsPPrice.State = 1 Then rsPPrice.Close
                temSQL = "SELECT * FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2)) & " Order by SetDate Desc, SetTime Desc"
                rsPPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                rsPPrice.AddNew
                rsPPrice!ItemID = Val(.TextMatrix(i, 2))
                rsPPrice!PPrice = Val(.TextMatrix(i, 10)) / NewItem.IssueUnitsPerPack
                rsPPrice!setdate = Date
                rsPPrice!SetTime = Time
                rsPPrice!StaffID = UserID
                rsPPrice.Update
            End If
            rsPPrice.Close
           
        Next
    End With
    If chkPrint.Value = 1 Then PrintPurchase
    
    tr = MsgBox("The Goods Received and added to stocks successfully", vbInformation, "Success")
    Call FormatGrid
    Call ClearSettleValues
    dtcItem.SetFocus
End Sub



Private Sub PrintPurchase()
    Dim RetVal As Integer
    Dim TemResponce     As Integer
     With Dataenvironment1.rscmmdGoodReceive
         If .State = 1 Then .Close
         .Source = "SELECT tblItem.ItemID, tblItem.Display, tblItem.AMPP, [tblRefill].[Amount]/[tblItem].[IssueUnitsPerPack] AS AmountInPackUnit, [tblRefill].[FreeAmount]/[tblItem].[IssueUnitsPerPack] AS FreeAmountInPackUnit, tblRefill.Amount, tblItem.IssueUnitsPerPack, tblPackUnit.PackUnit, tblIssueUnit.IssueUnit, tblItemCategory.SalesMargin, tblRefill.FreeAmount, tblRefill.* " & _
                     " FROM tblRefill LEFT JOIN (((tblItem LEFT JOIN tblIssueUnit ON tblItem.IssueUnitID = tblIssueUnit.IssueUnitID) LEFT JOIN tblPackUnit ON tblItem.PackUnitID = tblPackUnit.PackUnitID) LEFT JOIN tblItemCategory ON tblItem.ItemCategoryID = tblItemCategory.ItemCategoryID) ON tblRefill.ItemID = tblItem.ItemID " & _
                     " WHERE (((tblRefill.RefillBillID)= " & temRefillBillID & ") AND ((tblRefill.Amount) > 0))"
         .Open
         If .RecordCount > 0 Then
        CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
            With dtrPurchase
                Set .DataSource = Dataenvironment1.rscmmdGoodReceive
                .Sections("Section4").Controls("lblName").Caption = HospitalName
                .Sections("Section4").Controls("lblContact").Caption = HospitalAddress
                .Sections("Section4").Controls("lblTopic").Caption = "Local Purchase Note"
                .Sections("Section4").Controls("lblSUbtopic").Caption = Empty
                .Sections("Section4").Controls("lblTo").Caption = lblDistributor.Caption
                .Sections("Section4").Controls("lblAddress").Caption = lblAddress.Caption
                .Sections("Section4").Controls("lblTel").Caption = lblTelNo.Caption
                .Sections("Section4").Controls("lblFax").Caption = lblFax.Caption
                .Sections("Section4").Controls("lblDate").Caption = Format(Date, LongDateFormat)
                .Sections("Section4").Controls("lblRefillID").Caption = temRefillBillID
                .Sections("Section5").Controls("lblPayee").Caption = lblDistributor.Caption
                .Sections("Section5").Controls("lblTotalAmount").Caption = lblGrossTotal.Caption
                .Sections("Section5").Controls("lblDiscount").Caption = txtDiscount.Text
                .Sections("Section5").Controls("lblNetTotal").Caption = lblNetTotal.Caption
                .Sections("Section5").Controls("lblCheckedBy").Caption = dtcChecked.Text
                .Sections("Section5").Controls("lblreceivedBy").Caption = dtcStaff.Text
                RetVal = SelectForm(ReportPaperName, Me.hwnd)
                If RetVal = FORM_SELECTED Then
                    .Show
                Else
                    TemResponce = MsgBox("An Error in the report printer", vbCritical, "Printer Error")
                    Exit Sub
                End If
            End With
         End If
    End With
End Sub


Private Sub ClearSettleValues()
    txtAMP.Text = Empty
    txtAMPP.Text = Empty
    txtBalance.Text = Empty
    txtBatch.Text = Empty
    txtCashPaid.Text = Empty
    txtChequeNo.Text = Empty
    txtCreditCardNo.Text = Empty
    txtCreditCode.Text = Empty
    txtCreditDue.Text = Empty
    txtDataEntry.Text = Empty
    txtDiscount.Text = Empty
    txtDisplay.Text = Empty
    txtDue.Text = Empty
    txtFQty.Text = Empty
    txtInvoice.Text = Empty
    txtIStore.Text = Empty
    txtPPrice.Text = Empty
    txtPurchaseValue.Text = Empty
    txtQty.Text = Empty
    txtVMP.Text = Empty
    txtVMPP.Text = Empty
    txtVTM.Text = Empty
    dtcSupplier.Text = Empty
    dtcBank.Text = Empty
    dtcBranch.Text = Empty
    dtcCardBank.Text = Empty
    dtcCatogery.Text = Empty
    dtcCode.Text = Empty
    dtcCreditCard.Text = Empty
    dtcItem.Text = Empty
    dtcPayment.Text = Empty
    dtpChequeDate.Value = Date
    dtpOFrom.Value = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
    dtpOTo.Value = Date
    dtpPFrom.Value = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
    dtpPTo.Value = Date
    dtpUFrom.Value = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
    dtpUTo.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim tr As Integer
    If GridItem.Rows > 1 Then
        tr = MsgBox("There are items to be received. Are You sure you want to exit?", vbYesNo + vbQuestion, "Exit?")
        If tr = vbNo Then Cancel = True: Exit Sub
    End If
End Sub

Private Function IssueCredit(RefillBillID As Long) As Long
    With rsTemCredit
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblIssuedCredit"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IssuedSTaffID = dtcStaff.BoundText
        !IssuedDate = Date
        !IssuedTime = Time
        !Price = Val(lblNetTotal.Caption)
        !StoreID = UserStoreID
        !RefillBillID = RefillBillID
        .Update
        IssueCredit = !IssuedCreditID
        .Close
    End With
End Function


Private Function IssueCheque(RefillBillID As Long) As Long
    With rsTemCheque
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblIssuedCheque"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IssuedSTaffID = dtcStaff.BoundText
        !IssuedDate = Date
        !IssuedTime = Time
        !bankID = Val(dtcBank.BoundText)
        If IsNumeric(dtcBranch.BoundText) = True Then
            !BranchID = dtcBranch.BoundText
        End If
        !ChequeDate = dtpChequeDate.Value
        !ChequeNo = txtChequeNo.Text
        !Price = Val(lblNetTotal.Caption)
        !StoreID = UserStoreID
        !RefillBillID = RefillBillID
        .Update
        IssueCheque = !IssuedChequeID
        .Close
    End With
End Function


Private Function IssueCash(RefillBillID As Long) As Long
    With rsTemCash
        If .State = 1 Then .Close
        temSQL = "SELECT tblIssuedCash.* FROM tblIssuedCash"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IssuedSTaffID = dtcStaff.BoundText
        !IssuedDate = Date
        !IssuedTime = Time
        !Price = Val(lblNetTotal.Caption)
        !RefillBillID = RefillBillID
        !StoreID = UserStoreID
        .Update
        IssueCash = !IssuedCashID
        .Close
    End With
End Function

Private Sub CalculateTotal()
    Dim i As Integer
    Dim GrossTotal As Double
    Dim NetTotal As Double
    With GridItem
        For i = 1 To GridItem.Rows - 1
            GrossTotal = GrossTotal + Val(.TextMatrix(i, 19))
        Next
        lblGrossTotal.Caption = Format(GrossTotal, "####.00")
        NetTotal = GrossTotal - Val(txtDiscount.Text)
        lblNetTotal.Caption = Format(NetTotal, "####.00")
    End With
End Sub

Private Sub FillUsage(ByVal ItemID As Long)
    '0 Store
    '1 Sale
    '2 Consum
    '3 Discard
    '4 Adjustments
    '5 Total
    Dim StoreConsumption As Double
    Dim StoreSale As Double
    Dim StoreAdjustment As Double
    Dim StoreDiscard As Double
    Dim StoreUsage As Double
    Dim TotalConsumption As Double
    Dim TotalSale As Double
    Dim TotalAdjustment As Double
    Dim TotalDiscard As Double
    Dim TotalUsage As Double
    Dim TemStore As String
    
    With GridUsage
        .Cols = 6
        .Rows = 1
        .FixedCols = 0
        
        .ColWidth(0) = 3000
        
        .ColWidth(1) = (.Width - (.ColWidth(0) + 100)) / 5
        .ColWidth(2) = .ColWidth(1)
        .ColWidth(3) = .ColWidth(1)
        .ColWidth(4) = .ColWidth(1)
        .ColWidth(5) = .ColWidth(1)
        
        Dim i As Long
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            Select Case i
                Case 0: .Text = "Store"
                Case 1: .Text = "Sale"
                Case 2: .Text = "Consumption"
                Case 3: .Text = "Discard"
                Case 4: .Text = "Adjustment"
                Case 5: .Text = "Total"
            End Select
        Next i
        
    End With

    With rsTemStore
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblStore order by store"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
        
            While .EOF = False
                
                TemStore = !Store
                
                StoreUsage = 0
                
                StoreConsumption = CalculateConsumption(NewItem.ID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreConsumption
                TotalConsumption = TotalConsumption + StoreConsumption
                
                StoreSale = CalculateSale(NewItem.ID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreSale
                TotalSale = TotalSale + StoreSale
                
                StoreDiscard = CalculateDiscard(NewItem.ID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreDiscard
                TotalDiscard = TotalDiscard + StoreDiscard
                
                StoreAdjustment = CalculateAdjustment(NewItem.ID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreAdjustment
                TotalAdjustment = TotalAdjustment + StoreAdjustment
                
                With GridUsage
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = 0
                    .CellAlignment = 1
                    .Text = TemStore
                    .Col = 1
                    .Text = StoreSale & " " & NewItem.IUnit
                    .Col = 2
                    .Text = StoreConsumption & " " & NewItem.IUnit
                    .Col = 3
                    .Text = StoreDiscard & " " & NewItem.IUnit
                    .Col = 4
                    .Text = StoreAdjustment & " " & NewItem.IUnit
                    .Col = 5
                    .Text = StoreUsage & " " & NewItem.IUnit
                End With
                .MoveNext
            Wend
            With GridUsage
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
            .CellAlignment = 1
            .Text = "Total"
            .Col = 1
            .Text = TotalSale & " " & NewItem.IUnit
            .Col = 2
            .Text = TotalConsumption & " " & NewItem.IUnit
            .Col = 3
            .Text = TotalDiscard & " " & NewItem.IUnit
            .Col = 4
            .Text = TotalAdjustment & " " & NewItem.IUnit
            TotalUsage = TotalConsumption + TotalSale + TotalDiscard + TotalAdjustment
            .Col = 5
            .Text = TotalUsage & " " & NewItem.IUnit
            End With
        End If
        .Close
    End With
    
End Sub


Private Sub FillOrdering(ByVal ItemID As Long)
    With GridOrdering
        .Rows = 1
        .Cols = 8
        .FixedCols = 0
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Requested On"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Approved On"

        .Col = 2
        .CellAlignment = 4
        .Text = "Received On"
        
        .Col = 3
        .CellAlignment = 4
        .Text = "Requested Amount"
        
        .Col = 4
        .CellAlignment = 4
        .Text = "Approved Amount"
        
        .Col = 5
        .CellAlignment = 4
        .Text = "Received Amount"
        
        .Col = 6
        .CellAlignment = 4
        .Text = "Requested Distributor"
        
        .Col = 7
        .CellAlignment = 4
        .Text = "Approved Distributor"
        
        Dim i As Integer
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = (.Width - 100) / 8
        Next i
        
    End With
    With rsTemOrder
        If .State = 1 Then .Close
        temSQL = "SELECT tblOrder.RequestDate, tblOrder.ApprovedDate, tblOrder.ReceivedDate, tblOrder.RequestAmount, tblOrder.ApprovedAmount, tblOrder.ReceivedAmount, tblRDistrubutor.DistributorName, tblADistrubutor.DistributorName FROM (tblDistrubutor AS tblRDistrubutor RIGHT JOIN tblOrder ON tblRDistrubutor.DistributorID = tblOrder.ApprovedDistributorID) LEFT JOIN tblDistrubutor AS tblADistrubutor ON tblOrder.RequestDistributorID = tblADistrubutor.DistributorID WHERE (((tblOrder.ItemID)=" & ItemID & ") AND ((tblOrder.RequestDate) Between #" & Format(dtpOFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpOTo.Value, "dd MMMM yyyy") & "#)) ORDER BY tblOrder.RequestDate"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 1 Then
            While .EOF = False
                GridOrdering.Rows = GridOrdering.Rows + 1
                GridOrdering.Row = GridOrdering.Rows - 1
                GridOrdering.Col = 0
                GridOrdering.CellAlignment = 1
                GridOrdering.Text = Format(!requestdate, ShortDateFormat)
                GridOrdering.Col = 1
                GridOrdering.CellAlignment = 1
                If Not IsNull(!Approveddate) Then
                    GridOrdering.Text = Format(!Approveddate, ShortDateFormat)
                Else
                    GridOrdering.Text = "Not Approved"
                End If
                GridOrdering.Col = 2
                GridOrdering.CellAlignment = 1
                If Not IsNull(!ReceivedDate) Then
                    GridOrdering.Text = Format(!ReceivedDate, ShortDateFormat)
                Else
                    GridOrdering.Text = "Not Received"
                End If
                GridOrdering.Col = 3
                GridOrdering.CellAlignment = 7
                If Not IsNull(!RequestAmount) Then
                    GridOrdering.Text = !RequestAmount & " " & NewItem.IUnit
                Else
                    GridOrdering.Text = "Not Requested"
                End If
                GridOrdering.Col = 4
                GridOrdering.CellAlignment = 7
                If Not IsNull(!ApprovedAmount) Then
                    GridOrdering.Text = !ApprovedAmount & " " & NewItem.IUnit
                Else
                    GridOrdering.Text = "Not Approved"
                End If
                GridOrdering.Col = 5
                GridOrdering.CellAlignment = 7
                If Not IsNull(!ReceivedAmount) Then
                    GridOrdering.Text = !ReceivedAmount & " " & NewItem.IUnit
                Else
                    GridOrdering.Text = "Not Received"
                End If
                GridOrdering.Col = 6
                GridOrdering.CellAlignment = 7
                If Not IsNull(.Fields("tblRDistrubutor.DistributorName").Value) Then
                    GridOrdering.Text = .Fields("tblRDistrubutor.DistributorName").Value
                Else
                    GridOrdering.Text = "Not Requested"
                End If
                GridOrdering.Col = 7
                GridOrdering.CellAlignment = 7
                If Not IsNull(.Fields("tblADistrubutor.DistributorName").Value) Then
                    GridOrdering.Text = .Fields("tblADistrubutor.DistributorName").Value
                Else
                    GridOrdering.Text = "Not Approved"
                End If
                .MoveNext
            Wend
        End If
    End With
End Sub

Private Sub FillPrice(ByVal ItemID As Long)
With GridPPrice
    .Cols = 2
    .Rows = 1
    .FixedCols = 0
    
    .Row = 0
    
    .Col = 0
    .CellAlignment = 4
    .Text = "Starting Date"
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Purchase Price per " & NewItem.PUnit
    
    .ColWidth(0) = (.Width - 100) / 2
    .ColWidth(1) = (.Width - 100) / 2
    
End With

With rsTemPrice
    If .State = 1 Then .Close
    temSQL = "SELECT tblPurchasePrice.SetDate, tblPurchasePrice.PPrice FROM tblPurchasePrice WHERE (((tblPurchasePrice.ItemID)=" & ItemID & ") AND ((tblPurchasePrice.SetDate) Between #" & Format(dtpPFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpPTo.Value, "dd MMMM yyyy") & "#)) ORDER BY tblPurchasePrice.SetDate DESC"
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        While .EOF = False
            With GridPPrice
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .CellAlignment = 1
                .Text = Format(rsTemPrice!setdate, LongDateFormat)
                .Col = 1
                .CellAlignment = 7
                .Text = Format(rsTemPrice!PPrice * NewItem.IssueUnitsPerPack, "#,#00.00")
            End With
            .MoveNext
        Wend
    End If
End With


With GridSPrice
    .Cols = 2
    .Rows = 1
    .FixedCols = 0
    
    .Row = 0
    
    .Col = 0
    .CellAlignment = 4
    .Text = "Starting Date"
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Sales Price per " & NewItem.IUnit
    
    .ColWidth(0) = (.Width - 100) / 2
    .ColWidth(1) = (.Width - 100) / 2
    
End With

With rsTemPrice
    If .State = 1 Then .Close
    temSQL = "SELECT tblSalePrice.SetDate, tblSalePrice.SPrice FROM tblSalePrice WHERE (((tblSalePrice.ItemID)=" & ItemID & ") AND ((tblSalePrice.SetDate) Between #" & Format(dtpPFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpPTo.Value, "dd MMMM yyyy") & "#)) ORDER BY tblSalePrice.SetDate DESC"
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        While .EOF = False
            With GridSPrice
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .CellAlignment = 1
                .Text = Format(rsTemPrice!setdate, LongDateFormat)
                .Col = 1
                .CellAlignment = 7
                .Text = Format(rsTemPrice!SPrice, "#,#00.00")
            End With
            .MoveNext
        Wend
    End If
End With


End Sub

Private Sub DistributorDetails(ByVal DistributorID As Long)
    With rsTemDistributor
        If .State = 1 Then .Close
        temSQL = "SELECT tblDistrubutor.*, tblCity.City FROM tblCity RIGHT JOIN tblDistrubutor ON tblCity.CityId = tblDistrubutor.DistributorCityID Where DistributorId = " & DistributorID & ""
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!DistributorName) Then lblDistributor.Caption = !DistributorName
        If Not IsNull(!Balance) Then lblBalance.Caption = Format(!Balance, "#0.00")
        If Not IsNull(!distributorTelephone) Then lblTelNo.Caption = !distributorTelephone
        If Not IsNull(!distributorFax) Then lblFax.Caption = !distributorFax
        If Not IsNull(!distributorAddress) Then lblAddress.Caption = !distributorAddress
        If Not IsNull(!City) Then lblCity.Caption = !City
        If .State = 1 Then .Close
    End With
End Sub

Private Sub FillStocks(ByVal ItemID As Long)
    With GridTotal
        .Visible = False
    
        .Clear
        .Cols = 4
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
        .Text = "Department"
        
        .ColWidth(1) = 1600
        .ColWidth(2) = 1600
        .ColWidth(3) = 1600
        .ColWidth(0) = .Width - (.ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 100)
        
    End With
    With rsTemStore
        If .State = 1 Then .Close
        temSQL = "SELECT tblBatch.Batch, tblBatch.DOE, tblBatchStock.Stock, tblStore.Store, tblBatch.ItemID " & _
                    " FROM tblStore RIGHT JOIN (tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) ON tblStore.StoreID = tblBatchStock.StoreID " & _
                    " WHERE tblBatch.ItemID=" & ItemID & " AND tblBatchStock.Stock > 0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                GridTotal.Rows = GridTotal.Rows + 1
                GridTotal.Row = GridTotal.Rows - 1
                GridTotal.Col = 0
                GridTotal.CellAlignment = 1
                GridTotal.Text = !Batch
                GridTotal.Col = 1
                GridTotal.CellAlignment = 7
                If Not IsNull(!Stock) Then
                    GridTotal.Text = !Stock
                Else
                    GridTotal.Text = 0
                End If
                GridTotal.Col = 2
                GridTotal.CellAlignment = 1
                GridTotal.Text = Format(!DOE, ShortDateFormat)
                GridTotal.Col = 3
                GridTotal.CellAlignment = 1
                If Not IsNull(!Store) Then
                    GridTotal.Text = !Store
                Else
                    GridTotal.Text = Empty
                End If
                .MoveNext
            Wend
        End If
        GridTotal.Visible = True
        .Close
    End With
    
End Sub


Private Sub GetItemDetails(ItemID As Long)
    NewItem.ID = ItemID
    txtAMP.Text = NewItem.AMP
    txtAMPP.Text = NewItem.AMPP
    txtVMP.Text = NewItem.VMP
    txtVMPP.Text = NewItem.VMPP
    txtVTM.Text = NewItem.Generic
    txtDisplay.Text = NewItem.Display
End Sub

Private Sub lblGrossTotal_Change()
    lblNetTotal.Caption = Format((Val(lblGrossTotal.Caption) - Val(txtDiscount.Text)), "#0.00")
End Sub


Private Sub txtDiscount_Change()
    lblNetTotal.Caption = Format((Val(lblGrossTotal.Caption) - Val(txtDiscount.Text)), "#0.00")
End Sub

Private Sub txtDiscount_LostFocus()
    txtDiscount.Text = Format(txtDiscount.Text, "#0.00")
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtFQty.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtQty.Text = Empty
    End If
End Sub


Private Sub txtSPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtBatch.SetFocus
    End If
End Sub
