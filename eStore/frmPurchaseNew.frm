VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchaseNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase"
   ClientHeight    =   9255
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15240
   Begin VB.ComboBox cmbItem 
      Height          =   1380
      Left            =   960
      Style           =   1  'Simple Combo
      TabIndex        =   88
      Top             =   120
      Width           =   3975
   End
   Begin VB.Frame frameCash 
      Caption         =   "Cash"
      Height          =   2175
      Left            =   12120
      TabIndex        =   81
      Top             =   3480
      Width           =   3015
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1080
         TabIndex        =   84
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtCashPaid 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1080
         TabIndex        =   83
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtDue 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1080
         TabIndex        =   82
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label39 
         Caption         =   "Change"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label40 
         Caption         =   "Paid"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label41 
         Caption         =   "Due"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frameCredit 
      Caption         =   "Credit"
      Height          =   2175
      Left            =   12120
      TabIndex        =   78
      Top             =   3480
      Width           =   3015
      Begin VB.TextBox txtCreditDue 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1080
         TabIndex        =   79
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label42 
         Caption         =   "Due"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame frameCheque 
      Caption         =   "Cheque"
      Height          =   2175
      Left            =   12120
      TabIndex        =   69
      Top             =   3480
      Width           =   3015
      Begin VB.TextBox txtChequeNo 
         Height          =   375
         Left            =   1080
         TabIndex        =   70
         Top             =   1200
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpChequeDate 
         Height          =   375
         Left            =   1080
         TabIndex        =   71
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   76152835
         CurrentDate     =   39551
      End
      Begin MSDataListLib.DataCombo dtcBranch 
         Height          =   360
         Left            =   1080
         TabIndex        =   72
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
         TabIndex        =   73
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label43 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label44 
         Caption         =   "No"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label45 
         Caption         =   "Bank"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label46 
         Caption         =   "Branch"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame frameCreditCard 
      Caption         =   "Credit Card"
      Height          =   2175
      Left            =   12120
      TabIndex        =   54
      Top             =   3420
      Width           =   3015
      Begin VB.TextBox txtCreditCode 
         Height          =   375
         Left            =   1080
         TabIndex        =   56
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtCreditCardNo 
         Height          =   375
         Left            =   1080
         TabIndex        =   55
         Top             =   1200
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo dtcCardBank 
         Height          =   360
         Left            =   1080
         TabIndex        =   57
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
         TabIndex        =   58
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label47 
         Caption         =   "Code"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label48 
         Caption         =   "Bank"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label49 
         Caption         =   "Card"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label50 
         Caption         =   "No"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   13200
      TabIndex        =   52
      Top             =   2100
      Width           =   1935
   End
   Begin VB.TextBox txtInvoice 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   12120
      TabIndex        =   44
      Top             =   6600
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   13080
      TabIndex        =   42
      Top             =   8160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   76152835
      CurrentDate     =   39691
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   1935
      Left            =   5040
      TabIndex        =   41
      Top             =   120
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "By Issue Units"
      TabPicture(0)   =   "frmPurchaseNew.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label54"
      Tab(0).Control(1)=   "Label32"
      Tab(0).Control(2)=   "Label29"
      Tab(0).Control(3)=   "txtFQty"
      Tab(0).Control(4)=   "txtQty"
      Tab(0).Control(5)=   "txtPPrice"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "By Pack Units"
      TabPicture(1)   =   "frmPurchaseNew.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label52"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label53"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label58"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtFPQty"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtPQty"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtPPPrice"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtPPPrice 
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtPQty 
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFPQty 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtPPrice 
         Height          =   375
         Left            =   -73200
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtQty 
         Height          =   375
         Left            =   -73200
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFQty 
         Height          =   375
         Left            =   -73200
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label58 
         Caption         =   "Purchase Price"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label53 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label52 
         Caption         =   "Free Quantity"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label29 
         Caption         =   "Purchase Price"
         Height          =   375
         Left            =   -74760
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label32 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label54 
         Caption         =   "Free Quantity"
         Height          =   375
         Left            =   -74760
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.TextBox txtLastPPrice 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   36
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtLastSalePrice 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   255
      Left            =   12120
      TabIndex        =   34
      Top             =   8280
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.TextBox txtSPrice 
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtBatch 
      Height          =   375
      Left            =   12600
      TabIndex        =   18
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtPurchaseValue 
      Height          =   375
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtDataEntry 
      Height          =   375
      Left            =   8520
      TabIndex        =   28
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin btButtonEx.ButtonEx bttnReceive 
      Height          =   375
      Left            =   11880
      TabIndex        =   26
      Top             =   8640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
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
      Height          =   5535
      Left            =   120
      TabIndex        =   25
      Top             =   3120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9763
      _Version        =   393216
      WordWrap        =   -1  'True
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   375
      Left            =   13560
      TabIndex        =   27
      Top             =   8640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
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
   Begin MSComCtl2.DTPicker dtpDOM 
      Height          =   375
      Left            =   12600
      TabIndex        =   20
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "MMMM yyyy"
      Format          =   76152835
      CurrentDate     =   39545
   End
   Begin MSComCtl2.DTPicker dtpDOE 
      Height          =   375
      Left            =   12600
      TabIndex        =   22
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "MMMM yyyy"
      Format          =   76152835
      CurrentDate     =   39545
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   8880
      TabIndex        =   23
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
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
      Left            =   120
      TabIndex        =   24
      Top             =   8760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
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
   Begin MSDataListLib.DataCombo dtcChecked 
      Height          =   360
      Left            =   12120
      TabIndex        =   45
      Top             =   7740
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcStaff 
      Height          =   360
      Left            =   12120
      TabIndex        =   46
      Top             =   7200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcSupplier 
      Height          =   360
      Left            =   12120
      TabIndex        =   47
      Top             =   6000
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcPayment 
      Height          =   360
      Left            =   13200
      TabIndex        =   53
      Top             =   3060
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblItem 
      Caption         =   "Selected Item Name"
      Height          =   375
      Left            =   120
      TabIndex        =   89
      Top             =   1920
      Width           =   4815
   End
   Begin VB.Label Label5 
      Caption         =   "Discount"
      Height          =   255
      Left            =   12120
      TabIndex        =   68
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label lblNetTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   13200
      TabIndex        =   65
      Top             =   2580
      Width           =   1935
   End
   Begin VB.Label lblGrossTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   13200
      TabIndex        =   64
      Top             =   1620
      Width           =   1935
   End
   Begin VB.Label Label23 
      Caption         =   "Payment"
      Height          =   255
      Left            =   12120
      TabIndex        =   63
      Top             =   3060
      Width           =   1455
   End
   Begin VB.Label Label22 
      Caption         =   "Invoice No."
      Height          =   255
      Left            =   12120
      TabIndex        =   51
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "Checked by"
      Height          =   255
      Left            =   12120
      TabIndex        =   50
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Received by"
      Height          =   255
      Left            =   12120
      TabIndex        =   49
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label24 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   12120
      TabIndex        =   48
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblCategory 
      Caption         =   "Selected Item Category"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label57 
      Caption         =   "Last Purchase Price"
      Height          =   255
      Left            =   2400
      TabIndex        =   40
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblLastPPIUnit 
      Height          =   375
      Left            =   3960
      TabIndex        =   39
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblLastSPIUnit 
      Height          =   375
      Left            =   1320
      TabIndex        =   38
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label51 
      Caption         =   "Last Sale Price"
      Height          =   495
      Left            =   120
      TabIndex        =   37
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label56 
      Caption         =   "Sale Price"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label55 
      Height          =   375
      Left            =   9000
      TabIndex        =   33
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblFQtyUnit 
      Height          =   375
      Left            =   9000
      TabIndex        =   32
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblSPriceUnit 
      Height          =   375
      Left            =   9000
      TabIndex        =   31
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblPPriceUnit 
      Height          =   375
      Left            =   9000
      TabIndex        =   30
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblQtyUnit 
      Height          =   375
      Left            =   9000
      TabIndex        =   29
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label38 
      Caption         =   "Batch"
      Height          =   375
      Left            =   10800
      TabIndex        =   17
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label37 
      Caption         =   "Date of Manufacture"
      Height          =   375
      Left            =   10800
      TabIndex        =   19
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label36 
      Caption         =   "Date of Expiary"
      Height          =   375
      Left            =   10800
      TabIndex        =   21
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label34 
      Caption         =   "&Item"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label30 
      Caption         =   "Purchase Value"
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Gross Total"
      Height          =   255
      Left            =   12120
      TabIndex        =   66
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Net Total"
      Height          =   255
      Left            =   12120
      TabIndex        =   67
      Top             =   2580
      Width           =   1215
   End
End
Attribute VB_Name = "frmPurchaseNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    
    Dim billDistributor As New Distributor
    
    
    Dim itemCat As New ItemCategory
    Dim itemItem As New Item1
    
    Dim allItem As Collection
    Dim catItem As Collection
    
    
    
    Dim CsetPrinter As New cSetDfltPrinter
    
    Dim TemOrderBillID As Long
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
    Dim rsDI As New ADODB.Recordset
 
Private Sub bttnDelete_Click()
    If GridItem.Rows <= 1 Then Exit Sub
    If GridItem.Rows = 2 Then
        FormatGrid1
    Else
        GridItem.RemoveItem (GridItem.Row)
    End If
End Sub

Private Sub itemChanged()
    
    Dim temId As Long
    If cmbItem.ListIndex < 0 Then Exit Sub
    temId = Val(cmbItem.ItemData(cmbItem.ListIndex))
    If itemItem.ItemID = temId Then Exit Sub
    NewItem.ID = temId
    lblItem.Caption = NewItem.Display
    lblCategory.Caption = NewItem.Category
    Call FillLabels
    Call GetLastPrices(temId)
    
End Sub


Private Sub cmbItem_Change()
    itemChanged
End Sub

Private Sub cmbItem_Click()
    itemChanged
End Sub

Private Sub cmbItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If SSTab3.Tab = 0 Then
            txtQty.SetFocus
        Else
            txtPQty.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbItem.Text = Empty
    Else
    
    End If
End Sub

Private Sub cmbItem_LostFocus()
    itemChanged
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

Private Sub dtcPayment_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
        dtcPayment.Text = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcSupplier.SetFocus
    End If
End Sub

Private Sub dtcSupplier_Click(Area As Integer)
    If IsNumeric(dtcSupplier.BoundText) = False Then Exit Sub
    billDistributor.DistributorID = Val(dtcSupplier.BoundText)
End Sub

Private Sub dtcSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtcSupplier.Text = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtInvoice.SetFocus
    End If
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        bttnReceive_Click
    End If
End Sub


Private Sub dtpDOE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        bttnAdd_Click
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid1
    Call SetValues
    GridItem.RowHeight(0) = GridItem.RowHeight(0) * 3
    SSTab3.Tab = 0
    dtpDate.Value = Date
End Sub

Private Sub SetValues()
    dtpDOE.Value = Date
    dtpDOM.Value = Date
    dtpDOE.MinDate = LastDateOfMonth(Date)
    dtcStaff.BoundText = UserID
    dtcChecked.BoundText = UserID
    dtcStaff.Locked = True
    frameCash.Visible = False
    frameCheque.Visible = False
    frameCredit.Visible = False
    frameCreditCard.Visible = False
    
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
'    With rsItem
'        If .State = 1 Then .Close
'        temSQL = "SELECT * from tblitem order by display"
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With dtcItem
'        Set .RowSource = rsItem
'        .ListField = "display"
'        .BoundColumn = "ItemID"
'    End With
'    With rsItemCategory
'        If .State = 1 Then .Close
'        temSQL = "SELECT * from tblItemCategory order by categoryCode"
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With dtcCatogery
'        Set .RowSource = rsItemCategory
'        .ListField = "CategoryCode"
'        .BoundColumn = "ItemCategoryID"
'    End With
'    With rsCode
'        If .State = 1 Then .Close
'        temSQL = "SELECT * from tblitem order by code"
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With dtcCode
'        Set .RowSource = rsCode
'        .ListField = "code"
'        .BoundColumn = "ItemID"
'    End With
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
    

    Dim i As Integer
    i = 0
    With rsItem
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblitem order by Code"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            cmbItem.AddItem !Code
            cmbItem.ItemData(i) = !ItemID
            i = i + 1
            .MoveNext
        Wend
        .Close
        temSQL = "SELECT * from tblitem order by display"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            cmbItem.AddItem !Display
            cmbItem.ItemData(i) = !ItemID
            i = i + 1
            .MoveNext
        Wend
    End With
    
    
    
End Sub
    
Private Sub FormatGrid1()
    EditingData = False
    With GridItem
        .Cols = 24
        .Rows = 1
        .Row = 0
        .col = 0
        .FixedCols = 0
        
'        .RowHeight(0) = .RowHeight(0) * 3
        
        Dim i As Integer
        
        For i = 0 To .Cols - 1
            .col = i
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
                Case 10:     .Text = "Pruchase Price Per Unit"
                            .ColWidth(i) = 900
                Case 11:     .Text = "Slaes Price Per Unit"
                            .ColWidth(i) = 900
                Case 13:     .Text = "Purchase Price Per Pack"
                            .ColWidth(i) = 900
                Case 18:    .ColWidth(i) = 1200
                            .Text = "Total Pruchase Value"
                Case 21:    .ColWidth(i) = 1200
                            .Text = "DOE"
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
    '   6   IUnit
    '   7   FreeQuentity
    '   8   IUnit
    '   9   Batch
    '   10  Purchase Price Per Unit
    '   11  Sales Price Per Unit
    '   12  Sales Margin
    '   13  Purchaes Price Per Pack
    '   14
    '   15  IPurchased
    '   16  IFreePurchased
    '   17  IUnitsPerPack
    '   18  Display Price
    '   19  Actual Price
    '   20  DOM
    '   21  DOE
    '   22  Last Sale Price
    '   23  Last Purchase Price
    
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
        
        .col = 0
        .CellAlignment = 7
        .Text = .Row
        
        .col = 1
        .CellAlignment = 1
        .Text = NewItem.Display
        
        .col = 2
        .Text = NewItem.ID
        
        .col = 3
        .Text = NewItem.PUnitID
        
        .col = 4
        .Text = NewItem.IUnitID
        
        .col = 5
        .CellAlignment = 7
        .Text = txtQty.Text
        
        .col = 6
        .CellAlignment = 1
        .Text = NewItem.IUnit
        
        .col = 7
        .CellAlignment = 7
        .Text = txtFQty.Text
        
        .col = 8
        .CellAlignment = 1
        .Text = NewItem.IUnit
        
        .col = 9
        .CellAlignment = 7
        .Text = txtBatch.Text
        
        .col = 10
        .CellAlignment = 7
        .Text = Format(Val(txtPPrice.Text), "0.00")
        
        .col = 11
        .CellAlignment = 7
        .Text = Format((Val(txtSPrice.Text)), "0.00")
        
        .col = 12
        .CellAlignment = 7
        .Text = NewItem.SalesMargin
        
        .col = 13
        .Text = Format((Val(txtPPrice.Text) * NewItem.IssueUnitsPerPack), "0.00")
        
        .col = 14
        .Text = Empty
        
        .col = 15
        .Text = Val(txtQty.Text)
        
        .col = 16
        .Text = Val(txtFQty.Text)
       
        .col = 17
        .Text = NewItem.IssueUnitsPerPack
        
        .col = 18
        .Text = Format((Val(txtQty.Text) * Val(txtPPrice.Text)), "#,##0.00")
        
        .col = 19
        .Text = Val(txtQty.Text) * Val(txtPPrice.Text)
        
        .col = 20
        .CellAlignment = 4
        .Text = LastDateOfMonth(dtpDOM.Value)
        
        .col = 21
        .CellAlignment = 7
        .Text = Format(LastDateOfMonth(dtpDOE.Value), "dd MMM yyyy")
        
        .col = 22
        .Text = Val(txtLastSalePrice.Text)
        
        .col = 23
        .Text = Val(txtLastPPrice.Text)
        
    End With
    Call ClearAddValues
    Call CalculateTotal
    cmbItem.SetFocus
    EditingData = True
End Sub
    

Private Sub ClearAddValues()
    txtQty.Text = Empty
    txtPPrice.Text = Empty
    txtSPrice.Text = Empty
    txtFQty.Text = Empty
    txtPurchaseValue.Text = Empty
    cmbItem.Text = Empty
    
    txtBatch.Text = Empty
    txtLastPPrice.Text = Empty
    txtLastSalePrice.Text = Empty
    txtPQty.Text = Empty
    txtFPQty.Text = Empty
    txtPPPrice.Text = Empty
    
    lblQtyUnit.Caption = Empty
    lblFQtyUnit.Caption = Empty
    lblLastPPIUnit.Caption = Empty
    lblLastSPIUnit.Caption = Empty
    lblPPriceUnit.Caption = Empty
    lblSPriceUnit.Caption = Empty
    
End Sub




Private Function CanAdd() As Boolean
    CanAdd = False
    Dim tr As Integer
        If IsNumeric(Val(cmbItem.ItemData(cmbItem.ListIndex))) = False Then
            tr = MsgBox("You have not entered the item to add", vbCritical, "Item?")
            cmbItem.SetFocus
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
        If Val(txtPPrice.Text) >= Val(txtSPrice.Text) Then
            tr = MsgBox("You can't sell items at a rate below the purchase rate", vbCritical, "Adjust Sale Price")
            txtSPrice.SetFocus
            Exit Function
        End If
    CanAdd = True
End Function
    
Private Sub dtcItem_Change()
'    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
'    dtcCode.BoundText = dtcItem.BoundText
'    NewItem.ID = dtcItem.BoundText
'    Call FillLabels
'    Call FormatGrids
'    Call GetLastPrices(dtcItem.BoundText)
'    Call GetItemDetails(NewItem.ID)
'    Call FillStocks(dtcItem.BoundText)
'    Call FillPurchase(dtcItem.BoundText)
'    Call FillUsage(dtcItem.BoundText)
'    Call FillPrice(dtcItem.BoundText)
'    Call GetItemDetails(dtcItem.BoundText)
'    Call FillOrdering(dtcItem.BoundText)
'    Dim rsDI As New ADODB.Recordset
'    With rsDI
'        If .State = 1 Then .Close
'        temSql = "SELECT tblItemDistributor.DistributorID FROM tblItemDistributor WHERE (((tblItemDistributor.ItemID)=" & Val(dtcItem.BoundText) & "))"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            DistributorDetails (!DistributorID)
'        End If
'        .Close
'    End With
End Sub
    
Private Sub FillLabels()
    lblQtyUnit.Caption = NewItem.IUnit
    lblFQtyUnit.Caption = NewItem.IUnit
    lblLastPPIUnit.Caption = NewItem.IUnit
    lblLastSPIUnit.Caption = NewItem.IUnit
    lblPPriceUnit.Caption = "Per " & NewItem.IUnit
    lblSPriceUnit.Caption = "Per " & NewItem.IUnit
End Sub

Private Sub GetLastPrices(ItemID As Long)
    txtLastPPrice.Text = Empty
    txtLastSalePrice.Text = Empty
    With rsTemPrice
        If .State = 1 Then .Close
        temSQL = "SELECT tblCurrentSalePrice.SPrice FROM tblCurrentSalePrice WHERE tblCurrentSalePrice.ItemID=" & ItemID & " Order By SetDate Desc, SetTime DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtLastSalePrice.Text = Format(!SPrice, "0.00")
        End If
    End With
    With rsTemPrice
        If .State = 1 Then .Close
        temSQL = "SELECT tblCurrentPurchasePrice.PPrice FROM tblCurrentPurchasePrice WHERE tblCurrentPurchasePrice.ItemID=" & ItemID & " Order By SetDate Desc, SetTime DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
           txtLastPPrice.Text = Format(rsTemPrice!PPrice, "##00.00")
        End If
    End With
End Sub

'Private Sub GridItem_DblClick()
'    With GridItem
'        If IsNumeric(.TextMatrix(.Row, 2)) = False Then Exit Sub
'        dtcItem.BoundText = .TextMatrix(.Row, 2)
'        txtQty.Text = .TextMatrix(.Row, 15)
'        txtFQty.Text = .TextMatrix(.Row, 16)
'        txtPPrice.Text = .TextMatrix(.Row, 10)
'        txtSPrice.Text = .TextMatrix(.Row, 11)
'        txtBatch.Text = .TextMatrix(.Row, 9)
'        dtpDOM.Value = .TextMatrix(.Row, 20)
'        dtpDOE.Value = .TextMatrix(.Row, 21)
'    End With
'    bttnDelete_Click
'End Sub

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
        dtpDOE.SetFocus
    End If
End Sub

Private Sub txtBatch_LostFocus()
    txtBatch.Text = UCase(txtBatch.Text)
End Sub

Private Sub txtCashPaid_Change()
    Call CalculateBalance
End Sub

Private Sub txtDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyEscape Then
        txtDiscount.Text = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcPayment.SetFocus
    End If
End Sub

Private Sub txtDue_Change()
    Call CalculateBalance
End Sub

Private Sub CalculateBalance()
    txtBalance.Text = Format((Val(txtCashPaid.Text) - Val(txtDue.Text)), "0.00")
End Sub

Private Sub txtFPQty_Change()
    If SSTab3.Tab = 1 Then
        txtFQty.Text = Val(txtFPQty.Text) * NewItem.IssueUnitsPerPack
    End If
End Sub

Private Sub txtFPQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPPPrice.SetFocus
    End If
End Sub

Private Sub txtFQty_Change()
    If SSTab3.Tab = 0 Then
        If NewItem.IssueUnitsPerPack <> 0 Then txtFPQty.Text = Val(txtFQty.Text) / NewItem.IssueUnitsPerPack
    End If
End Sub

Private Sub txtFQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPPrice.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtFQty.Text = Empty
    End If
End Sub



Private Sub txtInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtInvoice.Text = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpDate.SetFocus
    End If
End Sub

Private Sub txtPPPrice_Change()
    If SSTab3.Tab = 1 Then
        If NewItem.IssueUnitsPerPack <> 0 Then txtPPrice.Text = Val(txtPPPrice.Text) / NewItem.IssueUnitsPerPack
    End If
End Sub

Private Sub txtPPPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtSPrice.SetFocus
    End If
End Sub

Private Sub txtPPrice_Change()
    Call CalculatePurchaseValue
    Call CalculateSalePrice
    If SSTab3.Tab = 0 Then
        txtPPPrice.Text = Val(txtPPrice.Text) * NewItem.IssueUnitsPerPack
    End If
End Sub

Private Sub txtPPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtSPrice.SetFocus
    End If
End Sub

Private Sub txtPQty_Change()
    If SSTab3.Tab = 1 Then
        txtQty.Text = Val(txtPQty.Text) * NewItem.IssueUnitsPerPack
    End If
End Sub

Private Sub txtPQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtFPQty.SetFocus
    End If
End Sub

Private Sub txtQty_Change()
    If SSTab3.Tab = 0 Then
        If NewItem.IssueUnitsPerPack <> 0 Then txtPQty.Text = Val(txtQty.Text) / NewItem.IssueUnitsPerPack
    End If
    Call CalculatePurchaseValue
End Sub
    
    
Private Sub CalculatePurchaseValue()
    txtPurchaseValue.Text = Format(((Val(txtQty.Text)) * Val(txtPPrice.Text)), "0.00")
End Sub
    
Private Sub CalculateSalePrice()
    If NewItem.ID <> 0 Then
        txtSPrice.Text = Format((((Val(txtPPrice.Text) * (NewItem.SalesMargin + 100)) / 100)), "0.00")
    End If
End Sub
    
    
Private Function CanReceive() As Boolean
    Dim i As Integer
    Dim tr As Integer
    CanReceive = False
    
    If GridItem.Rows <= 1 Then
        tr = MsgBox("There are no items to sell", vbCritical, "No Items")
        cmbItem.SetFocus
        Exit Function
    End If
    
    If txtInvoice.Text = Empty Then
        tr = MsgBox("Please enter an Invoice Nuumber")
        Exit Function
    End If
    
    If IsNumeric(dtcPayment.BoundText) = False Then
        tr = MsgBox("You have not selected the payment method", vbCritical, "No Items")
        dtcPayment.SetFocus
        Exit Function
    End If
    
    If IsNumeric(dtcSupplier.BoundText) = False Then
        tr = MsgBox("You have not selected the supplier", vbCritical, "No Supplier")
        dtcSupplier.SetFocus
        Exit Function
    End If
    
    If dtcPayment.Text = "Cash" Then
        If IsNumeric(txtCashPaid.Text) = False Then
            tr = MsgBox("You have not entered a valied cash amount", vbCritical, "Cash?")
            txtCashPaid.SetFocus
            Exit Function
        End If
'        If Val(txtCashPaid.Text) < Val(txtDue.Text) Then
'            tr = MsgBox("The amount you pay is not sufficient", vbCritical, "Not sufficient cash")
'            SSTab2.Tab = 1
'            txtCashPaid.SetFocus
'            Exit Function
'        End If
        
    ElseIf dtcPayment.Text = "Credit" = True Then
    
    ElseIf dtcPayment.Text = "Cheque" Then
        If IsNumeric(dtcBank.BoundText) = False Then
            tr = MsgBox("You have not selected a Bank", vbCritical, "Bank?")
            dtcBank.SetFocus
            Exit Function
        End If
        If Trim(txtChequeNo.Text) = "" Then
            tr = MsgBox("You have not entered the cheque number", vbCritical, "Cheque Number?")
            txtChequeNo.SetFocus
            Exit Function
        End If
    Else
        tr = MsgBox("You have not selected a Valid Payment Method", vbCritical, "Payment Method?")
        dtcPayment.SetFocus
        Exit Function
    End If
    
    If IsNumeric(dtcStaff.BoundText) = False Then
        tr = MsgBox("You have not selected the user", vbCritical, "Issued by?")
        dtcStaff.SetFocus
        Exit Function
    End If
    
    If IsNumeric(dtcChecked.BoundText) = False Then
        tr = MsgBox("You have not selected the name of the checked staff member", vbCritical, "Checked by?")
        dtcChecked.SetFocus
        Exit Function
    End If
    
    CanReceive = True
End Function
    
Private Function NoSameInnvoice(InvoiceNo As String, DistributorID As Long) As Boolean
    Dim rsNTem As New ADODB.Recordset
    With rsNTem
        If .State = 1 Then .Close
        temSQL = "SELECT     DistributorID, InvoiceNo FROM         dbo.tblRefillBill WHERE     (DistributorID = " & DistributorID & ") AND (InvoiceNo = N'" & InvoiceNo & "')"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            NoSameInnvoice = False
        Else
            NoSameInnvoice = True
        End If
        .Close
    End With
End Function


Private Sub bttnReceive_Click()
    
    Dim i As Integer
    
    If CanReceive = False Then Exit Sub
    If NoSameInnvoice(txtInvoice.Text, Val(dtcSupplier.BoundText)) = False Then
        i = MsgBox("This invoice no " & txtInvoice.Text & " from " & dtcSupplier.Text & " is already entered. Do you want to enter it again?", vbYesNo)
        If i = vbNo Then Exit Sub
    End If
    Dim tr As Integer
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
        !Discount = Val(txtDiscount.Text)
        DiscountPercent = (Val(txtDiscount.Text) / Val(lblGrossTotal.Caption)) * 100
        !DiscountPercent = DiscountPercent
        !NetPrice = Val(lblNetTotal.Caption)
        !Date = Date
        !Time = Now
        !PaymentMethodID = dtcPayment.BoundText
        !PaymentMethod = dtcPayment.Text
        !InvoiceNo = txtInvoice.Text
        !InvoiceDate = dtpDate.Value
        If dtcPayment.Text = "Credit" Then
            !FullyPaid = False
        Else
            !FullyPaid = True
        End If
        !purchase = True
        !Autorequest = False
        !ManualRequest = False
        .Update
        temSQL = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        temRefillBillID = !NewID
        .Close
        Dim CashID As Long
        Dim CreditID As Long
        Dim ChequeID As Long
        
        If dtcPayment.Text = "Cash" Then
            CashID = IssueCash(temRefillBillID)
        ElseIf dtcPayment.Text = "Credit" Then
            CreditID = IssueCredit(temRefillBillID)
        ElseIf dtcPayment.Text = "Cheque" Then
            ChequeID = IssueCheque(temRefillBillID)
        End If
        If .State = 1 Then .Close
        temSQL = "Select * from tblRefillBill Where RefillBillID = " & temRefillBillID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If dtcPayment.Text = "Cash" Then
                !IssuedCashID = CashID
            ElseIf dtcPayment.Text = "Credit" Then
                !IssuedCreditID = CreditID
            ElseIf dtcPayment.Text = "Cheque" Then
                !IssuedChequeID = ChequeID
            End If
            .Update
        End If
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
            rsTemRefill!Time = Now
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
                If AddToStock(ThisBatch, UserStoreID, Val(.TextMatrix(i, 15)) + Val(.TextMatrix(i, 16))) = False Then
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
            rsTemRefill!DOM = CDate(.TextMatrix(i, 20))
            rsTemRefill!DOE = CDate(.TextMatrix(i, 21))
            rsTemRefill!PackPPrice = Val(.TextMatrix(i, 13))
            rsTemRefill!SPrice = Val(.TextMatrix(i, 11))
            rsTemRefill!PPrice = Val(.TextMatrix(i, 10))
            rsTemRefill!LastSPrice = Val(.TextMatrix(i, 22))
            rsTemRefill!LastPPrice = Val(.TextMatrix(i, 23))
            
            rsTemRefill.Update
            rsTemRefill.Close
            
            If rsSPrice.State = 1 Then rsSPrice.Close
            temSQL = "SELECT tblSalePrice.ItemID, tblSalePrice.SPrice, tblSalePrice.SetDate, tblSalePrice.SetTime, tblSalePrice.StaffID FROM tblSalePrice "
            rsSPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            rsSPrice.AddNew
            rsSPrice!ItemID = Val(.TextMatrix(i, 2))
            rsSPrice!SPrice = Val(.TextMatrix(i, 11))
            rsSPrice!setdate = Date
            rsSPrice!SetTime = Now
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
                rsSPrice!SetTime = Now
                rsSPrice!StaffID = UserID
                rsSPrice.Update
            ElseIf rsSPrice.RecordCount = 1 Then
                rsSPrice!SPrice = Val(.TextMatrix(i, 11))
                rsSPrice!setdate = Date
                rsSPrice!SetTime = Now
                rsSPrice!StaffID = UserID
                rsSPrice.Update
            Else
                If rsSPrice.State = 1 Then rsSPrice.Close
                temSQL = "Delete FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
                rsSPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                If rsSPrice.State = 1 Then rsSPrice.Close
                temSQL = "SELECT * FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
                rsSPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                rsSPrice.AddNew
                rsSPrice!ItemID = Val(.TextMatrix(i, 2))
                rsSPrice!SPrice = Val(.TextMatrix(i, 11))
                rsSPrice!setdate = Date
                rsSPrice!SetTime = Now
                rsSPrice!StaffID = UserID
                rsSPrice.Update
            End If
            rsSPrice.Close
            
            If rsPPrice.State = 1 Then rsPPrice.Close
            temSQL = "SELECT tblPurchasePrice.ItemID, tblPurchasePrice.PPrice, tblPurchasePrice.SetDate, tblPurchasePrice.SetTime, tblPurchasePrice.StaffID FROM tblPurchasePrice"
            rsPPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            rsPPrice.AddNew
            rsPPrice!ItemID = Val(.TextMatrix(i, 2))
            rsPPrice!PPrice = Val(.TextMatrix(i, 10))
            rsPPrice!setdate = Date
            rsPPrice!SetTime = Now
            rsPPrice!StaffID = UserID
            rsPPrice.Update
            rsPPrice.Close
            
            If rsPPrice.State = 1 Then rsPPrice.Close
            temSQL = "SELECT * FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2)) & " Order by SetDate Desc, SetTime Desc"
            rsPPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If rsPPrice.RecordCount < 1 Then
                rsPPrice.AddNew
                rsPPrice!ItemID = Val(.TextMatrix(i, 2))
                rsPPrice!PPrice = Val(.TextMatrix(i, 10))
                rsPPrice!setdate = Date
                rsPPrice!SetTime = Now
                rsPPrice!StaffID = UserID
                rsPPrice.Update
            ElseIf rsPPrice.RecordCount = 1 Then
                rsPPrice!PPrice = Val(.TextMatrix(i, 10))
                rsPPrice!setdate = Date
                rsPPrice!SetTime = Now
                rsPPrice!StaffID = UserID
                rsPPrice.Update
            Else
                If rsPPrice.State = 1 Then rsPPrice.Close
                temSQL = "Delete FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2)) & " Order by SetDate Desc, SetTime Desc"
                rsPPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                If rsPPrice.State = 1 Then rsPPrice.Close
                temSQL = "SELECT * FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2)) & " Order by SetDate Desc, SetTime Desc"
                rsPPrice.Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                rsPPrice.AddNew
                rsPPrice!ItemID = Val(.TextMatrix(i, 2))
                rsPPrice!PPrice = Val(.TextMatrix(i, 10))
                rsPPrice!setdate = Date
                rsPPrice!SetTime = Now
                rsPPrice!StaffID = UserID
                rsPPrice.Update
            End If
            rsPPrice.Close
           
        Next
    End With
    If chkPrint.Value = 1 Then PrintPurchase
    
    tr = MsgBox("The Goods Received and added to stocks successfully" & vbNewLine & "Bill ID - " & temRefillBillID, vbInformation, "Success")
    Call FormatGrid1
    Call ClearSettleValues
    cmbItem.SetFocus
End Sub



Private Sub PrintPurchase()
    Dim RetVal As Integer
    Dim TemResponce     As Integer
     With Dataenvironment1.rscmmdGoodReceive
         If .State = 1 Then .Close
         .Source = "SELECT tblItem.Display, tblRefill.DOE, tblRefill.Amount, tblRefill.FreeAmount, tblRefill.PPrice, tblRefill.Price, tblRefill.SPrice, tblRefill.LastPPrice " & _
                     " FROM tblRefill LEFT JOIN tblItem ON tblRefill.ItemID = tblItem.ItemID " & _
                     " WHERE (((tblRefill.RefillBillID)= " & temRefillBillID & ") AND ((tblRefill.Amount) > 0))"
         .Open
         If .RecordCount > 0 Then
        
        
        CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
            
        
        Dim MyPrinter As Printer
        
        For Each MyPrinter In VB.Printers
            If MyPrinter.DeviceName = ReportPaperName Then
                Set Printer = MyPrinter
            End If
        Next
            
            With dtrPurchase
                Set .DataSource = Dataenvironment1.rscmmdGoodReceive
                .Sections("Section4").Controls("lblName").Caption = HospitalName
                .Sections("Section4").Controls("lblContact").Caption = HospitalAddress
                .Sections("Section4").Controls("lblTopic").Caption = "Good Receive Note"
                .Sections("Section4").Controls("lblSUbtopic").Caption = Empty
                .Sections("Section4").Controls("lblTo").Caption = billDistributor.DistributorName
                .Sections("Section4").Controls("lblAddress").Caption = billDistributor.DistributorAddress
                .Sections("Section4").Controls("lblTel").Caption = billDistributor.DistributorTelephone
                .Sections("Section4").Controls("lblFax").Caption = billDistributor.DistributorFax
                .Sections("Section4").Controls("lblDate").Caption = Format(Date, LongDateFormat)
                .Sections("Section4").Controls("lblRefillID").Caption = temRefillBillID
                .Sections("Section4").Controls("lblInvoiceDate").Caption = Format(dtpDate.Value, LongDateFormat)
                .Sections("Section4").Controls("lblInvoiceNo").Caption = txtInvoice.Text
                .Sections("Section5").Controls("lblPayee").Caption = billDistributor.DistributorName
                .Sections("Section5").Controls("lblTotalAmount").Caption = lblGrossTotal.Caption
                .Sections("Section5").Controls("lblDiscount").Caption = txtDiscount.Text
                .Sections("Section5").Controls("lblNetTotal").Caption = lblNetTotal.Caption
                .Sections("Section5").Controls("lblOperatedBy").Caption = UserName
                '
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

    
    txtBalance.Text = Empty
    txtBatch.Text = Empty
    txtCashPaid.Text = Empty
    txtChequeNo.Text = Empty
    txtCreditCardNo.Text = Empty
    txtCreditCode.Text = Empty
    txtCreditDue.Text = Empty
    txtDataEntry.Text = Empty
    txtDiscount.Text = Empty
    
    txtDue.Text = Empty
    txtFQty.Text = Empty
    txtInvoice.Text = Empty
    
    txtPPrice.Text = Empty
    txtPurchaseValue.Text = Empty
    txtQty.Text = Empty
    
    
    
    
    dtcSupplier.Text = Empty
    dtcBank.Text = Empty
    dtcBranch.Text = Empty
    dtcCardBank.Text = Empty
    
   
    dtcCreditCard.Text = Empty
    cmbItem.Text = Empty
    dtcPayment.Text = Empty
    
    
    dtpChequeDate.Value = Date
    
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
        !IssuedTime = Now
        !Price = Val(lblNetTotal.Caption)
        !StoreID = UserStoreID
        !RefillBillID = RefillBillID
        .Update
        .Close
        temSQL = "SELECT @@IDENTITY AS NewID"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        IssueCredit = !NewID
        .Close
        Set rsTemCredit = Nothing
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
        !IssuedTime = Now
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
        .Close
        temSQL = "SELECT @@IDENTITY AS NewID"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        IssueCheque = !NewID
        .Close
        Set rsTemCredit = Nothing
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
        !IssuedTime = Now
        !Price = Val(lblNetTotal.Caption)
        !RefillBillID = RefillBillID
        !StoreID = UserStoreID
        .Update
        .Close
        temSQL = "SELECT @@IDENTITY AS NewID"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        IssueCash = !NewID
        .Close
        Set rsTemCredit = Nothing
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


Private Sub txtDiscount_Change()
    lblNetTotal.Caption = Format((Val(lblGrossTotal.Caption) - Val(txtDiscount.Text)), "0.00")
End Sub

Private Sub txtDiscount_LostFocus()
    txtDiscount.Text = Format(txtDiscount.Text, "0.00")
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


