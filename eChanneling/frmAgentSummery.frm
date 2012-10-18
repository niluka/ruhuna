VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmAgentSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selected Agent Summery"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmAgentSummery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8925
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   5520
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSDataListLib.DataCombo dtcAgentCode 
      Bindings        =   "frmAgentSummery.frx":0442
      Height          =   360
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "InstitutionCode"
      BoundColumn     =   "Institution_ID"
      Text            =   ""
      Object.DataMember      =   "SqlTemAgent2"
   End
   Begin MSDataListLib.DataCombo dtcAgentName 
      Bindings        =   "frmAgentSummery.frx":0461
      Height          =   360
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "InstitutionName"
      BoundColumn     =   "Institution_ID"
      Text            =   ""
      Object.DataMember      =   "sqlTemAgent"
   End
   Begin VB.Frame FrameShiftSummary 
      Height          =   3135
      Left            =   600
      TabIndex        =   9
      Top             =   2040
      Width           =   7815
      Begin btButtonEx.ButtonEx bttnCashRefundAgent 
         Height          =   375
         Left            =   5880
         TabIndex        =   25
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print Cash &Refund"
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
      Begin btButtonEx.ButtonEx bttnPrintAgentBooking 
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print &Agent Booking"
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
      Begin btButtonEx.ButtonEx bttnPrintCashReceive 
         Height          =   375
         Left            =   5880
         TabIndex        =   6
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print &Cash Receive"
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
      Begin VB.Label lblAgentCashrefund 
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
         Left            =   2760
         TabIndex        =   27
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Refund To Agent"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Booking Value"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receive"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   4680
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   4440
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCashReceive 
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
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblAgenBooking 
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
         Left            =   2760
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label lblBalance 
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
         Left            =   2760
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   7435
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Today"
      TabPicture(0)   =   "frmAgentSummery.frx":0480
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblToday"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected &Day"
      TabPicture(1)   =   "frmAgentSummery.frx":049C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Selected &Period"
      TabPicture(2)   =   "frmAgentSummery.frx":04B8
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label18"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label17"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DTPicker3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DTPicker2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -72360
         TabIndex        =   3
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58654723
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58654723
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58654723
         CurrentDate     =   39442
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
         Left            =   -73920
         TabIndex        =   21
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "&Selected Date"
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
         TabIndex        =   20
         Top             =   480
         Width           =   1575
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
         Left            =   360
         TabIndex        =   19
         Top             =   480
         Width           =   735
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
         Left            =   3960
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent &Name"
      Height          =   375
      Left            =   3360
      TabIndex        =   24
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   3840
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Cod&e"
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmAgentSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TemCashTotal As Double
Dim TemAgentBookingPatients As Double
Dim TemRefund As Double
Dim A
Dim CSetPrinter As New cSetDfltPrinter

Private Sub bttnClose_Click()
Unload Me
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

bttnPrintAgentBooking.BackColor = BttnBackColour
bttnPrintAgentBooking.ForeColor = BttnForeColour

bttnPrintCashReceive.BackColor = BttnBackColour
bttnPrintCashReceive.ForeColor = BttnForeColour

bttnCashRefundAgent.BackColor = BttnBackColour
bttnCashRefundAgent.ForeColor = BttnForeColour

bttnPrintAgentBooking.BackColor = BttnBackColour
bttnPrintAgentBooking.ForeColor = BttnForeColour

'bttnDoctorPayments.BackColor = BttnBackColour
'bttnDoctorPayments.ForeColor = BttnForeColour

'bttnPrintSummary.BackColor = BttnBackColour
'bttnPrintSummary.ForeColor = BttnForeColour

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour

'bttnChange.BackColor = BttnBackColour
'bttnChange.ForeColor = BttnForeColour
'
'bttnDelete.BackColor = BttnBackColour
'bttnDelete.ForeColor = BttnForeColour


FrameShiftSummary.BackColor = FrameBackColour
FrameShiftSummary.ForeColor = FrameForeColour




frmAgentSummery.BackColor = FrameBackColour
frmAgentSummery.ForeColor = FrameForeColour

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

Private Sub FindAgentCashReceive()
TemCashTotal = 0
lblCashReceive.Caption = "0.00"

With DataEnvironment1.rssqlTem1
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "Select tblAgentCashSettle.* fROM tblAgentCashSettle Where (Institution_ID = " & dtcAgentName.BoundText & ")  and (SettledDate = '" & Date & "')"
    .Open
    
    Case 1
    .Source = "Select tblAgentCashSettle.* fROM tblAgentCashSettle Where (Institution_ID = " & dtcAgentName.BoundText & ")  and (SettledDate = '" & DTPicker1 & "')"
    .Open
    
    Case 2
    .Source = "Select tblAgentCashSettle.* fROM tblAgentCashSettle Where (Institution_ID = " & dtcAgentName.BoundText & ")  and (SettledDate between '" & DTPicker2 & "' and '" & DTPicker3 & "')"
    .Open
    
    End Select
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    
    TemCashTotal = Val(TemCashTotal) + Val(!Cash)
    .MoveNext
    Loop
End With

lblCashReceive.Caption = Format(TemCashTotal, "0.00")
End Sub
Private Sub ClearValues()
lblCashReceive.Caption = "0.00"
lblAgenBooking.Caption = "0.00"
lblBalance.Caption = "0.00"
lblAgentCashrefund = "0.00"
End Sub
Private Sub bttnPrintAgentBooking_Click()
If IsNumeric(dtcAgentName.BoundText) = False Then Exit Sub
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
   
    Select Case SSTab1.Tab
    
'SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblAgentBooking.AgentRefNo, tblDoctor.DoctorListedName FROM tblDoctor INNER JOIN ((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) INNER JOIN tblAgentBooking ON tblPatientFacility.PatientFacility_ID = tblAgentBooking.PatientFacility_ID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID
    Case 0
    .Source = "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblAgentBooking.AgentRefNo, tblDoctor.DoctorListedName FROM tblDoctor INNER JOIN ((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) INNER JOIN tblAgentBooking ON tblPatientFacility.PatientFacility_ID = tblAgentBooking.PatientFacility_ID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID Where (tblPatientFacility.Agent_ID = " & dtcAgentName.BoundText & ")  and ( tblPatientFacility.BookingDate = '" & Date & "') and (tblPatientFacility.PaymentMode ='Agent') ORDER BY tblPatientFacility.PatientFacility_ID"
    .Open
    
    Case 1
    .Source = "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblAgentBooking.AgentRefNo, tblDoctor.DoctorListedName FROM tblDoctor INNER JOIN ((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) INNER JOIN tblAgentBooking ON tblPatientFacility.PatientFacility_ID = tblAgentBooking.PatientFacility_ID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID Where (tblPatientFacility.Agent_ID = " & dtcAgentName.BoundText & ")  and ( tblPatientFacility.BookingDate = '" & DTPicker1.Value & "') and (tblPatientFacility.PaymentMode ='Agent') ORDER BY tblPatientFacility.PatientFacility_ID"
    .Open
    
    Case 2
    .Source = "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblAgentBooking.AgentRefNo, tblDoctor.DoctorListedName FROM tblDoctor INNER JOIN ((tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) INNER JOIN tblAgentBooking ON tblPatientFacility.PatientFacility_ID = tblAgentBooking.PatientFacility_ID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID Where (tblPatientFacility.Agent_ID = " & dtcAgentName.BoundText & ")  and ( tblPatientFacility.BookingDate between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and (PaymentMode ='Agent') ORDER BY tblPatientFacility.PatientFacility_ID"

'    .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (Agent_ID = " & dtcAgentName.BoundText & ")  and ( BookingDate between '" & DTPicker2 & "' and '" & DTPicker3 & "') and (PaymentMode ='Agent')Order by PatientFacility_ID "
    .Open
    
    End Select
    
    If .RecordCount = 0 Then A = MsgBox("No Agent Booking to view", vbCritical + vbOKOnly, "No Data"): Exit Sub

    With dtrAgentBookings2
    
        If HospitalDetails = True Then
        .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        Else
        .Sections("Section4").Controls.Item("RptName").Caption = Empty
        .Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        End If

         Select Case SSTab1.Tab
         
         Case 0
         .Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
         .Sections("Section2").Controls.Item("rptTodate").Caption = Format(Date, DefaultLongDate)
         Case 1
         .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
         .Sections("Section2").Controls.Item("rptTodate").Caption = DTPicker1.Value
         
         Case 2
         .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
         .Sections("Section2").Controls.Item("rptTodate").Caption = DTPicker3.Value
        
         End Select
         
        .Sections("Section2").Controls.Item("rptAgentName").Caption = dtcAgentName.Text
        .Sections("Section2").Controls.Item("rptlAgentcode").Caption = dtcAgentCode.Text
        
        .Sections("Section5").Controls.Item("lblNoofchanneling").Caption = " No Of Channeling   :  " & DataEnvironment1.rssqlTem10.RecordCount
        
        Set .DataSource = DataEnvironment1.rssqlTem10
        .Show
    End With



End With

End Sub



Private Sub bttnPrintCashReceive_Click()
If IsNumeric(dtcAgentName.BoundText) = False Then Exit Sub
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1.rssqlTem3
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "Select tblAgentCashSettle.*, tblInstitutions.* fROM tblAgentCashSettle Left Join tblInstitutions On tblAgentCashSettle.Institution_Id = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.Institution_ID = " & dtcAgentName.BoundText & ")  and (tblAgentCashSettle.SettledDate = '" & Date & "') ORDER BY SettledDate "
    .Open
    
    Case 1
    .Source = "Select tblAgentCashSettle.*, tblInstitutions.* fROM tblAgentCashSettle Left Join tblInstitutions On tblAgentCashSettle.Institution_Id = tblInstitutions.Institution_ID Where (tblAgentCashSettle.Institution_ID = " & dtcAgentName.BoundText & ")  and (tblAgentCashSettle.SettledDate = '" & DTPicker1 & "')  ORDER BY SettledDate "
    .Open
    
    Case 2
    .Source = "SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID Where (tblAgentCashSettle.Institution_ID = " & dtcAgentName.BoundText & ")  and (tblAgentCashSettle.SettledDate between '" & DTPicker2 & "' and '" & DTPicker3 & "')  ORDER BY SettledDate "
    .Open
    
    End Select
    
    If .RecordCount = 0 Then A = MsgBox("No Cash receive to view", vbCritical + vbOKOnly, "No Data"): Exit Sub
    
    With dtrAgentCashReceive2
        If HospitalDetails = True Then

        .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        Else
        .Sections("Section4").Controls.Item("RptName").Caption = Empty
        .Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        End If
         
         Select Case SSTab1.Tab
         
         Case 0
         .Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
         .Sections("Section2").Controls.Item("rptTodate").Caption = Format(Date, DefaultLongDate)
         Case 1
         .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
         .Sections("Section2").Controls.Item("rptTodate").Caption = DTPicker1.Value
         
         Case 2
         .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
         .Sections("Section2").Controls.Item("rptTodate").Caption = DTPicker3.Value
        
         End Select
         
        .Sections("Section2").Controls.Item("rptlAgentName").Caption = dtcAgentName.Text
        .Sections("Section2").Controls.Item("rptlAgentcode").Caption = dtcAgentCode.Text
    
        
        Set .DataSource = DataEnvironment1.rssqlTem3
        .Show
    
    End With

End With

End Sub

Private Sub dtcAgentCode_Change()
dtcAgentName.BoundText = dtcAgentCode.BoundText
End Sub

Private Sub dtcAgentCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then dtcAgentName.SetFocus

End Sub

Private Sub dtcAgentName_Change()
If IsNumeric(dtcAgentName.BoundText) = False Then Exit Sub
dtcAgentCode.BoundText = dtcAgentName.BoundText
CalculateValues
End Sub

Private Sub FindAgentBookingPatients()
TemAgentBookingPatients = 0
lblAgenBooking.Caption = "0.00"

With DataEnvironment1.rssqlTem2
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (Agent_ID = " & dtcAgentName.BoundText & ")  and ( BookingDate = '" & Date & "') and (PaymentMode ='Agent')"
    .Open
    
    Case 1
    .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (Agent_ID = " & dtcAgentName.BoundText & ")  and ( BookingDate = '" & DTPicker1 & "')and (PaymentMode ='Agent')"
    .Open
    
    Case 2
    .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (Agent_ID = " & dtcAgentName.BoundText & ")  and ( BookingDate between '" & DTPicker2 & "' and '" & DTPicker3 & "') and (PaymentMode ='Agent')"
    .Open
    
    End Select
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemAgentBookingPatients = Val(TemAgentBookingPatients) + Val(!totalfee)
    
    .MoveNext
    Loop
    
End With

lblAgenBooking.Caption = Format(TemAgentBookingPatients, "0.00")
End Sub

Private Sub FindAgentRefund()
TemRefund = 0

With DataEnvironment1.rssqlTem12

    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (Agent_ID = " & dtcAgentName.BoundText & ") and (RefundToAgent = 1) and ( RepayDate = '" & Date & "') and (PaymentMode ='Agent') Order by PatientFacility_ID "
     .Open
    Case 1
     .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (Agent_ID = " & dtcAgentName.BoundText & ") and (RefundToAgent = 1) and ( RepayDate = '" & DTPicker1 & "')and (PaymentMode ='Agent')Order by PatientFacility_ID "
     .Open
    Case 2
     .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (Agent_ID = " & dtcAgentName.BoundText & ") and (RefundToAgent = 1) and ( RepayDate between '" & DTPicker2 & "' and '" & DTPicker3 & "') and (PaymentMode ='Agent')Order by PatientFacility_ID "
     .Open
    End Select
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemRefund = Val(TemRefund) + Val(!totalrefund)
    .MoveNext
    Loop
    
    If .State = 1 Then .Close
 
End With
lblAgentCashrefund.Caption = Format(TemRefund, "0.00")
End Sub

Private Sub bttnCashRefundAgent_Click()
If IsNumeric(dtcAgentName.BoundText) = False Then Exit Sub
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

TemRefund = 0

With DataEnvironment1.rssqlTem12

    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode FROM (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (Agent_ID = " & dtcAgentName.BoundText & ") and (RefundToAgent = 1) and ( RepayDate = '" & Date & "') and (PaymentMode ='Agent') Order by PatientFacility_ID "
     .Open
    Case 1
     .Source = "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode FROM (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (Agent_ID = " & dtcAgentName.BoundText & ") and (RefundToAgent = 1) and ( RepayDate = '" & DTPicker1 & "')and (PaymentMode ='Agent')Order by PatientFacility_ID "
     .Open
    Case 2
     .Source = "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode FROM (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (Agent_ID = " & dtcAgentName.BoundText & ") and (RefundToAgent = 1) and ( RepayDate between '" & DTPicker2 & "' and '" & DTPicker3 & "') and (PaymentMode ='Agent')Order by PatientFacility_ID "
     .Open
    End Select
    
    If .RecordCount = 0 Then A = MsgBox("No Agent Booking to view", vbCritical + vbOKOnly, "No Data"): Exit Sub
        
        With DataReportAgentRefund
    If HospitalDetails = True Then
        .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        .Sections("section3").Controls.Item("lblAds").Caption = LongAd
    Else
        .Sections("Section4").Controls.Item("RptName").Caption = Empty
        .Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        .Sections("section3").Controls.Item("lblAds").Caption = LongAd
    End If
         Select Case SSTab1.Tab
         
         Case 0
         .Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
         .Sections("Section2").Controls.Item("rptTodate").Caption = Format(Date, DefaultLongDate)
         Case 1
         .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
         .Sections("Section2").Controls.Item("rptTodate").Caption = DTPicker1.Value
         
         Case 2
         .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
         .Sections("Section2").Controls.Item("rptTodate").Caption = DTPicker3.Value
        
         End Select
         
        .Sections("Section2").Controls.Item("rptAgentName").Caption = dtcAgentName.Text
        .Sections("Section2").Controls.Item("rptAgentCode").Caption = dtcAgentCode.Text
        
        Set .DataSource = DataEnvironment1.rssqlTem12
        
        .Show
        End With

 
End With

End Sub
'Private Sub FindInstutionBalanc()
'lblBalance.Caption = "0.00"
'
'With DataEnvironment1.rssqlTem2
'
'    If .State = 1 Then .Close
'
'     .Source = "Select tblInstitutions.* fROM tblInstitutions Where (Institution_ID = " & dtcAgentName.BoundText & ") "
'     .Open
'     If .RecordCount = 0 Then Exit Sub
'
'     lblBalance.Caption = Format(!InstitutionCredit, "0.00")
'
'    If .State = 1 Then .Close
'
'End With
'
'End Sub

Private Sub FindInstutionBalanc()
    Dim TemInstutionBal As Double
    Dim TemDate As Date
    
    If SSTab1.Tab = 0 Then
        TemDate = Date
    ElseIf SSTab1.Tab = 1 Then
        TemDate = DTPicker1
    ElseIf SSTab1.Tab = 2 Then
        TemDate = DTPicker3
    End If
    TemInstutionBal = 0
    lblBalance.Caption = "0.00"
    With DataEnvironment1.rssqlTem2
        If .State = 1 Then .Close
        .Source = "SELECT sum(tblInstitutionBalance.EBalance) as InsBal From tblInstitutionBalance where tblInstitutionBalance.Date = '" & Format(TemDate, "dd MMMM yyyy") & "' And Institution_ID = " & dtcAgentName.BoundText & " "
        .Open
        If .RecordCount = 0 Then Exit Sub
        If IsNull(!InsBal) = True Then Exit Sub
        TemInstutionBal = Val(!InsBal)
        lblBalance.Caption = Format(TemInstutionBal, "0.00")
    End With
End Sub

Private Sub DTPicker1_Change()
    If IsNumeric(dtcAgentName.BoundText) = False Then Exit Sub
    CalculateValues
End Sub

Private Sub DTPicker2_Change()
If IsNumeric(dtcAgentName.BoundText) = False Then Exit Sub
CalculateValues
End Sub

Private Sub CalculateValues()
ClearValues
FindAgentCashReceive
FindAgentBookingPatients
FindInstutionBalanc
FindAgentRefund
End Sub

Private Sub DTPicker3_Change()
If IsNumeric(dtcAgentName.BoundText) = False Then Exit Sub
CalculateValues
End Sub

Private Sub Form_Load()
If SetPrinter = False Then Unload Me: Exit Sub
DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
SSTab1.Tab = 0
    Me.Top = (Screen.Height / 2) - (Me.Height)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Call Setcolours
If UserAuthority <> AuthorityOwner Then
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
End If
End Sub

Private Function SetPrinter() As Boolean
SetPrinter = False
Dim MyPrinter As Printer

For Each MyPrinter In Printers
    If MyPrinter.DeviceName = ReportPrinterName Then
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


Private Sub SSTab1_Click(PreviousTab As Integer)

If IsNumeric(dtcAgentName.BoundText) = False Then Exit Sub

Select Case SSTab1.Tab

Case 0
CalculateValues

End Select

End Sub

