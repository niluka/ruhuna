VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmUserShiftSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Shift Summery"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUserShiftSummery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   8325
   Begin MSDataListLib.DataCombo DataComboUser 
      Bindings        =   "FrmUserShiftSummery.frx":0442
      Height          =   360
      Left            =   1680
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      ListField       =   "StaffName"
      BoundColumn     =   "Staff_ID"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin VB.Frame FrameShiftSummary 
      Height          =   4455
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   7815
      Begin btButtonEx.ButtonEx bttnUserSummaryPrint 
         Height          =   375
         Left            =   5880
         TabIndex        =   34
         Top             =   3720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Details Summary"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnCashFromCreditChanneling 
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "P. Creid &Patients  Cash"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnPrintAgentCash 
         Height          =   375
         Left            =   5880
         TabIndex        =   5
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print &Agent Payments"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnPrintSummary 
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         Top             =   3120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Summery Print"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnCashIncome 
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print &Cash Income"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnCashRefunds 
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print Cash &Refunds"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnDoctorPayments 
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         Top             =   2520
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print &Doctor Payments"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCashFromCreditChanneling 
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
         TabIndex        =   31
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Cettling for Credit Patients "
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label lblAgentCashPayments 
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
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblNetCashCollection 
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
         Left            =   2640
         TabIndex        =   24
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblDoctorPayment 
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
         Left            =   4200
         TabIndex        =   23
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblRefund 
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
         Left            =   4200
         TabIndex        =   22
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblCashFromChanneling 
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
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Cash Collection"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   5760
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   5760
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Expenses"
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Income"
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Refunds / Cancellations "
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash From Channeling"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Cash Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   2295
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   9975
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Toda&y"
      TabPicture(0)   =   "FrmUserShiftSummery.frx":0461
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblToday"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Selected Day"
      TabPicture(1)   =   "FrmUserShiftSummery.frx":047D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).Control(1)=   "Label11"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "S&elected Period"
      TabPicture(2)   =   "FrmUserShiftSummery.frx":0499
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DTPicker2"
      Tab(2).Control(1)=   "DTPicker3"
      Tab(2).Control(2)=   "Label18"
      Tab(2).Control(3)=   "Label17"
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -72360
         TabIndex        =   1
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   56426499
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -74040
         TabIndex        =   2
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   56426499
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   -70560
         TabIndex        =   3
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   56426499
         CurrentDate     =   39442
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "&To"
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
         Left            =   -71040
         TabIndex        =   27
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "&From"
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
         Left            =   -74640
         TabIndex        =   26
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Selec&ted Date"
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
         Left            =   -74040
         TabIndex        =   25
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblToday 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Today :"
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
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   3975
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   6720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cl&ose"
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
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   33
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label lblUser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   32
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User :"
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
      Left            =   600
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmUserShiftSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TemNetChannelingincome As Double
Dim CashFromCahnneling As Double
Dim TemAgentCash As Double
Dim TemCashFromCreditCahnneling As Double
Dim temCashRefund As Double
Dim TemDoctorPayment As Double

Dim A
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

'GridCellBackColor = 5853695
'GridCellForeColor = 658120


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

bttnCashIncome.BackColor = BttnBackColour
bttnCashIncome.ForeColor = BttnForeColour

bttnPrintAgentCash.BackColor = BttnBackColour
bttnPrintAgentCash.ForeColor = BttnForeColour

bttnCashFromCreditChanneling.BackColor = BttnBackColour
bttnCashFromCreditChanneling.ForeColor = BttnForeColour

bttnCashRefunds.BackColor = BttnBackColour
bttnCashRefunds.ForeColor = BttnForeColour

bttnDoctorPayments.BackColor = BttnBackColour
bttnDoctorPayments.ForeColor = BttnForeColour

bttnPrintSummary.BackColor = BttnBackColour
bttnPrintSummary.ForeColor = BttnForeColour

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour

bttnUserSummaryPrint.BackColor = BttnBackColour
bttnUserSummaryPrint.ForeColor = BttnForeColour
'
'bttnDelete.BackColor = BttnBackColour
'bttnDelete.ForeColor = BttnForeColour


FrameShiftSummary.BackColor = FrameBackColour
FrameShiftSummary.ForeColor = FrameForeColour




FrmUserShiftSummery.BackColor = FrameBackColour
FrmUserShiftSummery.ForeColor = FrameForeColour

'FrameOfficial.BackColor = FrameBackColour
'FrameOfficial.ForeColor = FrameForeColour
'
'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour

'chkBypassOrder.BackColor = LblBackColour
'chkBypassOrder.ForeColor = LblForeColour
'
'ChkCalculateTime.BackColor = LblBackColour
'ChkCalculateTime.ForeColor = LblForeColour
'
'chkFullDayLeave.BackColor = LblBackColour
'chkFullDayLeave.ForeColor = LblForeColour
'
'
'DataComboDoctor.BackColor = TxtBackColour
'DataComboDoctor.ForeColor = TxtForeColour

'DataComboPaymenyMethod.BackColor = TxtBackColour
'DataComboPaymenyMethod.ForeColor = TxtForeColour
'
'DataComboSex.BackColor = TxtBackColour
'DataComboSex.ForeColor = TxtForeColour
'
'DataComboSpeciality.BackColor = TxtBackColour
'DataComboSpeciality.ForeColor = TxtForeColour
'
'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour




'Grid1.BackColor = GridBackColor
'Grid1.ForeColor = GridForeColor
'
'Grid1.BackColorBkg = GridBackColorBkg
'Grid1.BackColorFixed = GridBackColorFixed
'Grid1.BackColorSel = GridBackColorSel
'
'Grid1.ForeColor = GridForeColor
'Grid1.ForeColorFixed = GridForeColorFixed
'Grid1.ForeColorSel = GridForeColorSel

'grid1.ForeColor = Grid



'Label1.BackColor = LblBackColour
'Label1.ForeColor = LblForeColour
'
'Label10.BackColor = LblBackColour
'Label10.ForeColor = LblForeColour
'Label11.BackColor = LblBackColour
'Label11.ForeColor = LblForeColour
'Label12.BackColor = LblBackColour
'Label12.ForeColor = LblForeColour
'Label13.BackColor = LblBackColour
'Label13.ForeColor = LblForeColour
'Label14.BackColor = LblBackColour
'Label14.ForeColor = LblForeColour
'Label15.BackColor = LblBackColour
'Label15.ForeColor = LblForeColour
'Label16.BackColor = LblBackColour
'Label16.ForeColor = LblForeColour
'Label2.BackColor = LblBackColour
'Label2.ForeColor = LblForeColour
'Label18.BackColor = LblBackColour
'Label18.ForeColor = LblForeColour
'Label3.BackColor = LblBackColour
'Label3.ForeColor = LblForeColour
'Label20.BackColor = LblBackColour
'Label20.ForeColor = LblForeColour
'Label21.BackColor = LblBackColour
'Label21.ForeColor = LblForeColour
'Label4.BackColor = LblBackColour
'Label4.ForeColor = LblForeColour
'Label23.BackColor = LblBackColour
'Label23.ForeColor = LblForeColour
'Label24.BackColor = LblBackColour
'Label24.ForeColor = LblForeColour
'Label25.BackColor = LblBackColour
'Label25.ForeColor = LblForeColour
'Label26.BackColor = LblBackColour
'Label26.ForeColor = LblForeColour
'Label27.BackColor = LblBackColour
'Label27.ForeColor = LblForeColour
'Label4.BackColor = LblBackColour
'Label4.ForeColor = LblForeColour
'Label5.BackColor = LblBackColour
'Label5.ForeColor = LblForeColour
'Label6.BackColor = LblBackColour
'Label6.ForeColor = LblForeColour
'Label7.BackColor = LblBackColour
'Label7.ForeColor = LblForeColour

'Label8.BackColor = LblBackColour
'Label8.ForeColor = LblForeColour
'Label9.BackColor = LblBackColour
'Label9.ForeColor = LblForeColour

'lblOfficialEmail.BackColor = LblBackColour
'lblOfficialEmail.ForeColor = LblForeColour

'lblOfficialWebsite.BackColor = LblBackColour
'lblOfficialWebsite.ForeColor = LblForeColour


'txtAccount.BackColor = TxtBackColour
'txtAccount.ForeColor = TxtForeColour
'
'txtBankBranch.BackColor = TxtBackColour
'txtBankBranch.ForeColor = TxtForeColour
'
'txtComments.BackColor = TxtBackColour
'txtComments.ForeColor = TxtForeColour
'txtCredit.BackColor = TxtBackColour
'txtCredit.ForeColor = TxtForeColour
'txtDesignation.BackColor = TxtBackColour
'txtDesignation.ForeColor = TxtForeColour
'txtListedName.BackColor = TxtBackColour
'txtListedName.ForeColor = TxtForeColour
'txtName.BackColor = TxtBackColour
'txtName.ForeColor = TxtForeColour
'txtOfficialAddress.BackColor = TxtBackColour
'txtOfficialAddress.ForeColor = TxtForeColour
'txtOfficialEMail.BackColor = TxtBackColour
'txtOfficialEMail.ForeColor = TxtForeColour
'txtOfficialFax.BackColor = TxtBackColour
'txtOfficialFax.ForeColor = TxtForeColour
'txtOfficialTel.BackColor = TxtBackColour
'txtOfficialTel.ForeColor = TxtForeColour
'txtOfficialWebsite.BackColor = TxtBackColour
'txtOfficialWebsite.ForeColor = TxtForeColour
'
'txtPrivateAddress.BackColor = TxtBackColour
'txtPrivateAddress.ForeColor = TxtForeColour
'txtPrivateEmail.BackColor = TxtBackColour
'txtPrivateEmail.ForeColor = TxtForeColour
'txtPrivateFax.BackColor = TxtBackColour
'txtPrivateFax.ForeColor = TxtForeColour
'txtPrivateMobile.BackColor = TxtBackColour
'txtPrivateMobile.ForeColor = TxtForeColour
'txtPrivateTel.BackColor = TxtBackColour
'txtPrivateTel.ForeColor = TxtForeColour
'
'
'txtQualifications.BackColor = TxtBackColour
'txtQualifications.ForeColor = TxtForeColour
'txtRegistation.BackColor = TxtBackColour
'txtRegistation.ForeColor = TxtForeColour
'txtSearch.BackColor = TxtBackColour
'txtSearch.ForeColor = TxtForeColour
'txtTel.ForeColor = TxtForeColour
'txtTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour







End Sub
Private Sub bttnCashFromCreditChanneling_Click()

With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.settlecashDate = #" & Date & "#)  order by tblPatientFacility.patientfacility_ID ")
    ElseIf SSTab1.Tab = 1 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.settlecashDate = #" & DTPicker1.Value & "#) order by tblPatientFacility.patientfacility_ID")
    ElseIf SSTab1.Tab = 2 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.settlecashDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#) order by tblPatientFacility.patientfacility_ID ")
    End If
    
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "No Transaction"): Exit Sub
    
    With dtrCredittBookingsPayment
        Set .DataSource = DataEnvironment1.rssqlTem10
        
        .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        .Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    
        If SSTab1.Tab = 0 Then
        .Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, "dd/MM/YYYY")
        .Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, "dd/MM/YYYY")
        ElseIf SSTab1.Tab = 1 Then
        .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
        .Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
        ElseIf SSTab1.Tab = 2 Then
        .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
        .Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
        End If
        
        .Show
    End With
    
    

End With



End Sub

Private Sub bttnCashRefunds_Click()
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
With DataEnvironment1.rssqlCashireRepost

    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = True) and ( (tblPatientFacility.Cancelled = True)or (tblPatientFacility.Refund = true ))and (tblPatientFacility.RepayDate = #" & Date & "#)")
    ElseIf SSTab1.Tab = 1 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = True) and ( (tblPatientFacility.Cancelled = True)or (tblPatientFacility.Refund = true ))and (tblPatientFacility.RepayDate = #" & DTPicker1.Value & "#)")
    ElseIf SSTab1.Tab = 2 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = True) and ( (tblPatientFacility.Cancelled = True)or (tblPatientFacility.Refund = true ))and (tblPatientFacility.RepayDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)")
    End If
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "No Transaction"): Exit Sub
    
    Set DataReportCashRefunds.DataSource = DataEnvironment1.rssqlCashireRepost
        If HospitalDetails = True Then
            DataReportCashRefunds.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
            DataReportCashRefunds.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
            DataReportCashRefunds.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
            DataReportCashRefunds.Sections("section3").Controls.Item("lblAds").Caption = LongAd
        Else
            DataReportCashRefunds.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
            DataReportCashRefunds.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
            DataReportCashRefunds.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
            DataReportCashRefunds.Sections("section3").Controls.Item("lblAds").Caption = LongAd
        End If

    If SSTab1.Tab = 0 Then
    DataReportCashRefunds.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, "dd/MM/YYYY")
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, "dd/MM/YYYY")
    ElseIf SSTab1.Tab = 1 Then
    DataReportCashRefunds.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
    DataReportCashRefunds.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
    End If
    
DataReportCashRefunds.Show

End With

End Sub

Private Sub bttnCashIncome_Click()
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
With DataEnvironment1.rssqlCashireRepost

    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = #" & Date & "#)  order by tblPatientFacility.patientfacility_ID ")
    ElseIf SSTab1.Tab = 1 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = #" & DTPicker1.Value & "#) order by tblPatientFacility.patientfacility_ID")
    ElseIf SSTab1.Tab = 2 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#) order by tblPatientFacility.patientfacility_ID ")
    End If
    
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "No Transaction"): Exit Sub
  
    Set DataReportCashIncome.DataSource = DataEnvironment1.rssqlCashireRepost
    
    DataReportCashIncome.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportCashIncome.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportCashIncome.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text

    If SSTab1.Tab = 0 Then
    DataReportCashIncome.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, "dd/MM/YYYY")
    DataReportCashIncome.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, "dd/MM/YYYY")
    ElseIf SSTab1.Tab = 1 Then
    DataReportCashIncome.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
    DataReportCashIncome.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
    DataReportCashIncome.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
    DataReportCashIncome.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
    End If
    
DataReportCashIncome.Show

End With

End Sub


Private Sub bttnDoctorPayments_Click()
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
With DataEnvironment1.rssqlTem9
    If .State = 1 Then .Close
   
    If SSTab1.Tab = 0 Then
        .Open ("Select tblstaffpayment.*,tbldoctor.* From tblstaffpayment Left Join tbldoctor On tblstaffpayment.Staff_ID = tbldoctor.Doctor_ID Where  (tblstaffpayment.User_ID = " & Val(DataComboUser.BoundText) & " and (tblstaffpayment.PaidDate = #" & Date & "# )and (tblstaffpayment.IsADoctor = True))Order by StaffPayment_ID")
    ElseIf SSTab1.Tab = 1 Then
        .Open ("Select tblstaffpayment.*,tbldoctor.* From tblstaffpayment Left Join tbldoctor On tblstaffpayment.Staff_ID = tbldoctor.Doctor_ID Where (tblstaffpayment.User_ID = " & Val(DataComboUser.BoundText) & " and (tblstaffpayment.PaidDate = #" & DTPicker1.Value & "#)and (tblstaffpayment.IsADoctor = True))Order by StaffPayment_ID")
    ElseIf SSTab1.Tab = 2 Then
        .Open ("Select tblstaffpayment.*,tbldoctor.* From tblstaffpayment Left Join tbldoctor On tblstaffpayment.Staff_ID = tbldoctor.Doctor_ID Where (tblstaffpayment.User_ID = " & Val(DataComboUser.BoundText) & " and (tblstaffpayment.PaidDate between #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)and (tblstaffpayment.IsADoctor = True))Order by StaffPayment_ID")
    End If
    
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "No Transaction"): Exit Sub
    
    Set DataReportDoctorPayment.DataSource = DataEnvironment1.rssqlTem9
If HospitalDetails = True Then
    DataReportDoctorPayment.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportDoctorPayment.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    DataReportDoctorPayment.Sections("section3").Controls.Item("lblAds").Caption = LongAd
Else
    DataReportDoctorPayment.Sections("Section4").Controls.Item("RptName").Caption = Empty
    DataReportDoctorPayment.Sections("Section4").Controls.Item("RptAddress").Caption = Empty
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    DataReportDoctorPayment.Sections("section3").Controls.Item("lblAds").Caption = LongAd
End If
    If SSTab1.Tab = 0 Then
    DataReportDoctorPayment.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, "dd/MM/YYYY")
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, "dd/MM/YYYY")
    ElseIf SSTab1.Tab = 1 Then
    DataReportDoctorPayment.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
    DataReportDoctorPayment.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
    End If
    
    DataReportDoctorPayment.Show

End With

End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub



Private Sub bttnPrintAgentCash_Click()

If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
With DataEnvironment1.rssqlAgentPayment1

    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate = #" & Date & "#)   ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    ElseIf SSTab1.Tab = 1 Then
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate = #" & DTPicker1.Value & "#) ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    ElseIf SSTab1.Tab = 2 Then
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#) ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    End If
    
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "No Transaction"): Exit Sub

    
    Set dtrAgentCashReceive.DataSource = DataEnvironment1.rssqlAgentPayment1
    
    
    dtrAgentCashReceive.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    dtrAgentCashReceive.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    dtrAgentCashReceive.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text

    If SSTab1.Tab = 0 Then
    dtrAgentCashReceive.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, "dd/MM/YYYY")
    dtrAgentCashReceive.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, "dd/MM/YYYY")
    ElseIf SSTab1.Tab = 1 Then
    dtrAgentCashReceive.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
    dtrAgentCashReceive.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
    dtrAgentCashReceive.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
    dtrAgentCashReceive.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
    End If
    
    dtrAgentCashReceive.Show

End With

End Sub

Private Sub ClearValues()

lblCashFromChanneling.Caption = "0.00"
lblAgentCashPayments.Caption = "0.00"
lblRefund.Caption = "0.00"
lblNetCashCollection.Caption = "0.00"
lblCashFromCreditChanneling = "0.00"
lblDoctorPayment.Caption = "0.00"
End Sub

Private Sub bttnPrintSummary_Click()

With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    .Open "Select * From tblTem"
    Set dtrShiftEndCash.DataSource = DataEnvironment1.rssqlTemSu1
End With

With dtrShiftEndCash
    .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    .Sections("Section4").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    .Sections("Section2").Controls.Item("rptLCashChannaling").Caption = Format(lblCashFromChanneling, "#0.00")
    .Sections("Section2").Controls.Item("rptAgentCashreceive").Caption = Format(lblAgentCashPayments, "#0.00")
    .Sections("Section2").Controls.Item("rptlCashFromCreditChaneling").Caption = Format(lblCashFromCreditChanneling, "#0.00")
    .Sections("Section2").Controls.Item("rptlNetCashChanelling").Caption = Format(lblNetPatientcash, "#0.00")
    .Sections("Section2").Controls.Item("rptTotalCashReceive").Caption = Format(TemTotalCash, "#0.00")
    .Sections("Section2").Controls.Item("rptlCancelRefund").Caption = Format(lblRefund, "#0.00")
    .Sections("Section2").Controls.Item("rptDoctorPayment").Caption = Format(lblDoctorPayment, "#0.00")
    .Sections("Section2").Controls.Item("rptTotalPayment").Caption = Format(TemTotalPayment, "#0.00")
    .Sections("Section2").Controls.Item("rptlNetCashChanelling").Caption = Format(lblNetCashCollection, "#0.00")
    
        If SSTab1.Tab = 0 Then
        .Sections("Section4").Controls.Item("rptFromdate").Caption = Format(Date, "dd/MM/YYYY")
        .Sections("Section4").Controls.Item("RptToDate").Caption = Format(Date, "dd/MM/YYYY")
        ElseIf SSTab1.Tab = 1 Then
        .Sections("Section4").Controls.Item("rptFromdate").Caption = DTPicker1.Value
        .Sections("Section4").Controls.Item("RptToDate").Caption = DTPicker1.Value
        ElseIf SSTab1.Tab = 2 Then
        .Sections("Section4").Controls.Item("rptFromdate").Caption = DTPicker2.Value
        .Sections("Section4").Controls.Item("RptToDate").Caption = DTPicker3.Value
        End If
    
    .Show
End With
        

End Sub

Private Sub bttnUserSummaryPrint_Click()
Const preSHAPE = "SHAPE {"
Const SQL = "SELECT tblPatientFacility.*, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode, tblPatientMainDetails.FirstName, tblDoctor.DoctorName FROM tblInstitutions RIGHT JOIN (tblPatientMainDetails RIGHT JOIN (tblDoctor RIGHT JOIN tblPatientFacility ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID"
Const postSHAPE = "ORDER BY tblPatientFacility.PaymentMode DESC , tblPatientFacility.PatientFacility_ID}  AS sqlPDAF COMPUTE sqlPDAF, SUM(sqlPDAF.'PersonalDue') AS DocFee, SUM(sqlPDAF.'InstitutionDue') AS HosFee, ANY(sqlPDAF.'PaymentMode') AS PaymentMethodName, SUM(sqlPDAF.'TotalDue') AS TotFee BY 'PaymentMethod_Id'"
' SHAPE {SELECT tblPatientFacility.*, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode, tblPatientMainDetails.FirstName, tblDoctor.DoctorName FROM tblInstitutions RIGHT JOIN (tblPatientMainDetails RIGHT JOIN (tblDoctor RIGHT JOIN tblPatientFacility ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID ORDER BY tblPatientFacility.PaymentMode DESC , tblPatientFacility.PatientFacility_ID}  AS sqlPDAF COMPUTE sqlPDAF, SUM(sqlPDAF.'PersonalDue') AS DocFee, SUM(sqlPDAF.'InstitutionDue') AS HosFee, ANY(sqlPDAF.'PaymentMode') AS PaymentMethodName, SUM(sqlPDAF.'TotalDue') AS TotFee BY 'PaymentMethod_Id'

With DataEnvironment1

    If .rssqlPDAF_Grouping.State = 1 Then DataEnvironment1.rssqlPDAF_Grouping.Close
    
   
    Select Case SSTab1.Tab
    
    Case 0
    .Commands!sqlPDAF_Grouping.CommandText = preSHAPE & SQL & " Where (BookingDate = #" & Date & "#)and (User_ID = " & Val(DataComboUser.BoundText) & ") " & postSHAPE
    .sqlPDAF_Grouping
    
    Case 1
    .Commands!sqlPDAF_Grouping.CommandText = preSHAPE & SQL & " Where (BookingDate =#" & DTPicker1.Value & "#)and (User_ID = " & Val(DataComboUser.BoundText) & ")" & postSHAPE
    .sqlPDAF_Grouping
    
    Case 2
    .Commands!sqlPDAF_Grouping.CommandText = preSHAPE & SQL & " Where (BookingDate Between #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#) and (User_ID = " & Val(DataComboUser.BoundText) & ")" & postSHAPE
    .sqlPDAF_Grouping
    
    End Select
    
    If DataEnvironment1.rssqlPDAF_Grouping.RecordCount = 0 Then A = MsgBox("No Transaction to view", vbInformation + vbOKOnly, "No Transactions"): Exit Sub
    
    With dtrShifUserSummary
    
'    Set .DataSource = DataEnvironment1.rssqlPDAF_Grouping
    
    .Sections("ReportHeader").Controls.Item("RptName").Caption = InstitutionName
    .Sections("ReportHeader").Controls.Item("RptAddress").Caption = InstitutionAddress
    .Sections("ReportHeader").Controls.Item("rptLHeding3").Caption = "Cashier Shift End Summary"

    .Sections("ReportFooter").Controls.Item("lblCashirerName").Caption = "Cashier Name :  " & UserName

    If SSTab1.Tab = 0 Then
    .Sections("PageHeader").Controls.Item("rptDate").Caption = "On   " & Format(Date, "dd/MM/YYYY")
    ElseIf SSTab1.Tab = 1 Then
    .Sections("PageHeader").Controls.Item("rptDate").Caption = "On   " & DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
    .Sections("PageHeader").Controls.Item("rptDate").Caption = "Date From   " & DTPicker2.Value & "   To   " & DTPicker3.Value
    End If
    
    .Show
    End With
    
End With

End Sub

Private Sub DataComboUser_Change()
lblUser.Caption = DataComboUser.Text
Call CalculateValues
End Sub

Private Sub DTPicker1_Change()
Call CalculateValues
End Sub

Private Sub DTPicker2_Change()
If (DTPicker2.Value) > (DTPicker3.Value) Then
    Dim TemDate1 As Date
    TemDate1 = DTPicker2.Value
    DTPicker2.Value = DTPicker3.Value
    DTPicker3.Value = TemDate1
End If
Call CalculateValues
End Sub

Private Sub DTPicker3_Change()
Call CalculateValues
End Sub

Private Sub Form_Load()
On Error GoTo Error_Han

With DataEnvironment1.rssqlTem11
    DataComboUser.RowMember = ""
    DataComboUser.ListField = ""
    If .State = 1 Then .Close
    .Source = "Select* From tblStaff Order by StaffName "
    .Open
    If .RecordCount = 0 Then Exit Sub

    
    DataComboUser.RowMember = "sqlTem11"
    DataComboUser.ListField = "StaffName"
   
    If .State = 1 Then .Close

End With

SSTab1.Tab = 0
lblDate.Caption = Format(Date, "dd mmmm yyyy (dddd)")

DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
SSTab1.Tab = 0
DataComboUser.BoundText = UserID
Call CalculateValues
Call Setcolours

Exit Sub

Error_Han:
A = MsgBox(Err.Number & vbNewLine & Err.Description, vbInformation + vbOKOnly, "Loding Error")
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
'Call CalculateValues
End Sub

Private Sub CalculateValues()
Call ClearValues

If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub


Call ChannelingCashIncome

Call AgentCashReceive

Call CashReceiveFromCreditBooking

Call CashRefund

Call DoctorPayment

Call CalculateTotals


Exit Sub
End Sub

Private Sub DoctorPayment()
TemDoctorPayment = 0
lblDoctorPayment.Caption = Format(TemDoctorPayment, "#0.00")

With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    Select Case SSTab1.Tab
    Case 0
    .Source = "Select tblStaffPayment.* From tblStaffPayment Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (PaidDate = #" & Date & "# )and (IsADoctor = True)Order by StaffPayment_ID"
    .Open
    Case 1
     .Source = "Select tblStaffPayment.* From tblStaffPayment Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (PaidDate = #" & DTPicker1.Value & "# )and (IsADoctor = True)Order by StaffPayment_ID"
    .Open
    Case 2
     .Source = "Select tblStaffPayment.* From tblStaffPayment Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (PaidDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)and (IsADoctor = True)Order by StaffPayment_ID"
    .Open
    End Select
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
     TemDoctorPayment = TemDoctorPayment + !PaidAmount
     .MoveNext
    Loop
   
    If .State = 1 Then .Close
End With

lblDoctorPayment.Caption = Format(TemDoctorPayment, "#0.00")

End Sub

Private Sub CashReceiveFromCreditBooking()

TemCashFromCreditCahnneling = 0
lblCashFromCreditChanneling.Caption = Format(TemCashFromCreditCahnneling, "#0.00")

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    If SSTab1.Tab = 0 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.settlecashDate = #" & Date & "#)  order by tblPatientFacility.patientfacility_ID ")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 1 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.settlecashDate = #" & DTPicker1.Value & "#) order by tblPatientFacility.patientfacility_ID")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 2 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.settlecashDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#) order by tblPatientFacility.patientfacility_ID ")
    If .RecordCount = 0 Then Exit Sub
    End If
    
    While .EOF = False
        TemCashFromCreditCahnneling = Val(TemCashFromCreditCahnneling) + Val(!personalfee) + Val(!InstitutionFee) + Val(!otherfee)
        .MoveNext
    Wend
    
    

End With
lblCashFromCreditChanneling.Caption = Format(TemCashFromCreditCahnneling, "#0.00")

End Sub

Private Sub AgentCashReceive()
TemAgentCash = 0
lblAgentCashPayments.Caption = Format(TemAgentCash, "###0.00")

If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
With DataEnvironment1.rssqlAgentPayment1

    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate = #" & Date & "#)   ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 1 Then
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate = #" & DTPicker1.Value & "#) ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 2 Then
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#) ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    If .RecordCount = 0 Then Exit Sub
    End If
    
    Do While .EOF = False
    TemAgentCash = Val(TemAgentCash) + Val(!Cash)
    .MoveNext
    Loop
    
End With
lblAgentCashPayments.Caption = Format(TemAgentCash, "###0.00")
End Sub

Private Sub ChannelingCashIncome()

If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
CashFromCahnneling = 0
lblCashFromChanneling.Caption = Format(CashFromCahnneling, "#0.00")

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    If SSTab1.Tab = 0 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = #" & Date & "#)  order by tblPatientFacility.patientfacility_ID ")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 1 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = #" & DTPicker1.Value & "#) order by tblPatientFacility.patientfacility_ID")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 2 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#) order by tblPatientFacility.patientfacility_ID ")
    If .RecordCount = 0 Then Exit Sub
    End If
    
    While .EOF = False
        CashFromCahnneling = CashFromCahnneling + !personalfee + !InstitutionFee + !otherfee
        .MoveNext
    Wend
End With

lblCashFromChanneling.Caption = Format(CashFromCahnneling, "#0.00")

End Sub

Private Sub CashRefund()
temCashRefund = 0
lblRefund.Caption = Format(temCashRefund, "0.00")


With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    If SSTab1.Tab = 0 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = True) and ( (tblPatientFacility.Cancelled = True)or (tblPatientFacility.Refund = true ))and (tblPatientFacility.RepayDate = #" & Date & "#)")
    ElseIf SSTab1.Tab = 1 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = True) and ( (tblPatientFacility.Cancelled = True)or (tblPatientFacility.Refund = true ))and (tblPatientFacility.RepayDate = #" & DTPicker1.Value & "#)")
    ElseIf SSTab1.Tab = 2 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = True) and ( (tblPatientFacility.Cancelled = True)or (tblPatientFacility.Refund = true ))and (tblPatientFacility.RepayDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)")
    End If
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
           If IsNull(!InstitutionRefund) = False Then temCashRefund = temCashRefund + (!InstitutionRefund)
           If IsNull(!personalrefund) = False Then temCashRefund = temCashRefund + (!personalrefund)
        .MoveNext
    Loop
   
    If .State = 1 Then .Close
End With

lblRefund.Caption = Format(temCashRefund, "0.00")

End Sub

Private Sub CalculateTotals()
TemTotalCash = 0
TemTotalPayment = 0
TemNetChannelingincome = 0
lblNetCashCollection.Caption = Format(TemNetChannelingincome, "#0.00")

TemTotalCash = (CashFromCahnneling + TemAgentCash + TemCashFromCreditCahnneling)
TemTotalPayment = (temCashRefund + TemDoctorPayment)

TemNetChannelingincome = Val(TemTotalCash) - Val(TemTotalPayment)

lblNetCashCollection.Caption = Format(TemNetChannelingincome, "#0.00")

End Sub
