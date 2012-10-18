VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmInstitutionPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Institution Payment"
   ClientHeight    =   8580
   ClientLeft      =   1155
   ClientTop       =   2115
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInstitutionPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   9360
   Begin VB.Frame Frame2 
      Height          =   8055
      Left            =   480
      TabIndex        =   12
      Top             =   360
      Width           =   8415
      Begin MSDataListLib.DataCombo dtcAgentName 
         Height          =   360
         Left            =   3720
         TabIndex        =   44
         Top             =   840
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtReceiptno 
         Height          =   375
         Left            =   360
         MaxLength       =   50
         TabIndex        =   43
         Top             =   7440
         Width           =   3015
      End
      Begin VB.ComboBox cmbAgentCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3720
         TabIndex        =   41
         Top             =   360
         Width           =   4455
      End
      Begin VB.Frame FramePaymentMethod 
         Caption         =   "Pa&yment Method"
         Height          =   1695
         Left            =   5880
         TabIndex        =   13
         Top             =   5400
         Width           =   2295
         Begin VB.OptionButton OptionCreditCard 
            Caption         =   "Cred&it Card"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton OptionCheque 
            Caption         =   "Che&que"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton OptionCash 
            Caption         =   "&Cash"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker dtpPaymentdate 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   4800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   39423
      End
      Begin btButtonEx.ButtonEx bttnClose 
         Height          =   375
         Left            =   6240
         TabIndex        =   11
         Top             =   7440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Close"
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
      Begin btButtonEx.ButtonEx bttnUpdate 
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   7440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Update"
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
      Begin btButtonEx.ButtonEx bttnSerchBill 
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Serch Agent "
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
      Begin VB.Frame framInstitution 
         Height          =   2415
         Left            =   360
         TabIndex        =   36
         Top             =   1320
         Width           =   7815
         Begin MSFlexGridLib.MSFlexGrid grid1 
            Height          =   1815
            Left            =   240
            TabIndex        =   1
            Top             =   360
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   3201
            _Version        =   393216
            ScrollBars      =   2
            SelectionMode   =   1
            AllowUserResizing=   2
         End
      End
      Begin VB.Frame FrameCash 
         Caption         =   "Cash"
         Height          =   1695
         Left            =   360
         TabIndex        =   14
         Top             =   5400
         Width           =   5415
         Begin VB.TextBox txtCashPayment 
            Height          =   375
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash      :                                      Rs."
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label lblCashBalance 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rs. 0.00"
            Height          =   375
            Left            =   3600
            TabIndex        =   17
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Change   :"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame FrameCheque 
         Caption         =   "C&heque"
         Height          =   1695
         Left            =   360
         TabIndex        =   28
         Top             =   5400
         Width           =   5415
         Begin VB.TextBox txtChequeNo 
            Height          =   360
            Left            =   840
            MaxLength       =   20
            TabIndex        =   5
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtChequeAmount 
            Height          =   375
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   4
            Top             =   720
            Width           =   3015
         End
         Begin MSDataListLib.DataCombo DataComboBank 
            Bindings        =   "frmInstitutionPayment.frx":0442
            Height          =   360
            Left            =   2280
            TabIndex        =   3
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   635
            _Version        =   393216
            ListField       =   "BankName"
            BoundColumn     =   "Bank_ID"
            Text            =   ""
            Object.DataMember      =   "sqlBank"
         End
         Begin MSComCtl2.DTPicker DTPickerChequeDate 
            Height          =   375
            Left            =   3600
            TabIndex        =   6
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   20709379
            CurrentDate     =   39414
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "&No. :"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "&Amount :"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "&Date :"
            Height          =   255
            Left            =   3000
            TabIndex        =   30
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "&Bank      :"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame FrameCreditCard 
         Caption         =   "Credit Card"
         Height          =   1695
         Left            =   360
         TabIndex        =   19
         Top             =   5400
         Width           =   5415
         Begin VB.OptionButton OptionABC 
            Caption         =   "ABC"
            Height          =   255
            Left            =   3720
            TabIndex        =   25
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtAuthorizationCode 
            Height          =   375
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   24
            Top             =   1200
            Width           =   2895
         End
         Begin VB.OptionButton OptionAmEx 
            Caption         =   "AmEX"
            Height          =   255
            Left            =   2640
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptionMaster 
            Caption         =   "MASTER"
            Height          =   255
            Left            =   1320
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptionVISA 
            Caption         =   "VISA"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtAmount 
            Height          =   375
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   20
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount                :"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "Authorization Code:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1200
            Width           =   2055
         End
      End
      Begin MSDataListLib.DataCombo dtcAgentCode 
         Height          =   360
         Left            =   5520
         TabIndex        =   40
         Top             =   4920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt No"
         Height          =   255
         Left            =   360
         TabIndex        =   46
         Top             =   7200
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Name"
         Height          =   375
         Left            =   2280
         TabIndex        =   45
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Receipt No"
         Height          =   375
         Left            =   360
         TabIndex        =   42
         Top             =   6480
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent C&ode"
         Height          =   375
         Left            =   2280
         TabIndex        =   39
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblInstitutionName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3360
         TabIndex        =   38
         Top             =   4320
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Name"
         Height          =   255
         Left            =   3360
         TabIndex        =   37
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Payment Date"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Balance"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lblInstitutionBalance 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   4320
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmInstitutionPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim i As Integer
    Dim rcount As Integer
    Dim TemInstitutionID As Integer
    Dim PaymentSuccess As Boolean

Private Sub SetColour()

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

'bttnClose.BackColor = BttnBackColour
'bttnClose.ForeColor = BttnForeColour
'
'bttnAdd.BackColor = BttnBackColour
'bttnAdd.ForeColor = BttnForeColour
'
'bttnAddLeave.BackColor = BttnBackColour
'bttnAddLeave.ForeColor = BttnForeColour
'
'bttnCancel.BackColor = BttnBackColour
'bttnCancel.ForeColor = BttnForeColour
'
'bttnChange.BackColor = BttnBackColour
'bttnChange.ForeColor = BttnForeColour
'
'bttnClose.BackColor = BttnBackColour
'bttnClose.ForeColor = BttnForeColour
'
'bttnDelete.BackColor = BttnBackColour
'bttnDelete.ForeColor = BttnForeColour
'
'bttnEdit.BackColor = BttnBackColour
'bttnEdit.ForeColor = BttnForeColour

'ButtonEx3.BackColor = BttnBackColour
'ButtonEx3.ForeColor = BttnForeColour

bttnSerchBill.BackColor = BttnBackColour
bttnSerchBill.ForeColor = BttnForeColour

bttnUpdate.BackColor = BttnBackColour
bttnUpdate.ForeColor = BttnForeColour

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour


OptionABC.BackColor = FrmBackColour
OptionABC.ForeColor = FrmForeColour

OptionAmEx.BackColor = FrmBackColour
OptionAmEx.ForeColor = FrmForeColour

OptionCash.BackColor = FrmBackColour
OptionCash.ForeColor = FrmForeColour

OptionVISA.BackColor = FrmBackColour
OptionVISA.ForeColor = FrmForeColour

OptionMaster.BackColor = FrmBackColour
OptionMaster.ForeColor = FrmForeColour

OptionCheque.BackColor = FrmBackColour
OptionCheque.ForeColor = FrmForeColour

OptionCreditCard.BackColor = FrmBackColour
OptionCreditCard.ForeColor = FrmForeColour

OptionCreditCard.BackColor = FrmBackColour
OptionCreditCard.ForeColor = FrmForeColour


FramePaymentMethod.BackColor = FrmBackColour
FramePaymentMethod.ForeColor = FrmForeColour

FrameCreditCard.BackColor = FrameBackColour
FrameCreditCard.ForeColor = FrameForeColour

FrameCash.BackColor = FrameBackColour
FrameCash.ForeColor = FrameForeColour



Frame2.BackColor = FrameBackColour
Frame2.ForeColor = FrameForeColour

framInstitution.BackColor = FrameBackColour
framInstitution.ForeColor = FrameForeColour

FrameCheque.BackColor = FrameBackColour
FrameCheque.ForeColor = FrameForeColour

frmInstitutionPayment.BackColor = FrameBackColour
frmInstitutionPayment.ForeColor = FrameForeColour
'FrameCredit.BackColor = FrameBackColour
'FrameCredit.ForeColor = FrameForeColour
'FrameCreditCard.BackColor = FrameBackColour
'FrameCreditCard.ForeColor = FrameForeColour
'FramePatient.BackColor = FrameBackColour
'FramePatient.ForeColor = FrameForeColour
'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour
'
'FramePaymentMethod.BackColor = FrameBackColour
'FramePaymentMethod.ForeColor = FrameForeColour
'frameSearchPatient.BackColor = FrameBackColour
'frameSearchPatient.ForeColor = FrameForeColour




'chkEvening.BackColor = LblBackColour
'chkEvening.ForeColor = LblForeColour
'
'chkFridayEvening.BackColor = LblBackColour
'chkFridayEvening.ForeColor = LblForeColour
'
'chkMondayEvening.BackColor = LblBackColour
'chkMondayEvening.ForeColor = LblForeColour
'
'chkMondayFullLeave.BackColor = LblBackColour
'chkMondayFullLeave.ForeColor = LblForeColour
'
'chkMondayMorning.BackColor = LblBackColour
'chkMondayMorning.ForeColor = LblForeColour
'
'chkMoning.BackColor = LblBackColour
'chkMoning.ForeColor = LblForeColour
'
'chkSaturdayEvening.BackColor = LblBackColour
'chkSaturdayEvening.ForeColor = LblForeColour
'
'chkSaturdayFullLeave.BackColor = LblBackColour
'chkSaturdayFullLeave.ForeColor = LblForeColour
'
'chkSaturdayMorning.BackColor = LblBackColour
'chkSaturdayMorning.ForeColor = LblForeColour
'
'chkSundayEvening.BackColor = LblBackColour
'chkSundayEvening.ForeColor = LblForeColour
'
'chkSundayFullLeave.BackColor = LblBackColour
'chkSundayFullLeave.ForeColor = LblForeColour
'
'chkSundayEvening.BackColor = LblBackColour
'chkSundayEvening.ForeColor = LblForeColour
'
'chkSundayMorning.BackColor = LblBackColour
'chkSundayMorning.ForeColor = LblForeColour
'
'chkThursdayEvening.BackColor = LblBackColour
'chkThursdayEvening.ForeColor = LblForeColour
'
'chkThursdayFullLeave.BackColor = LblBackColour
'chkThursdayFullLeave.ForeColor = LblForeColour
'
'chkThursdayEvening.BackColor = LblBackColour
'chkThursdayEvening.ForeColor = LblForeColour
'
'chkThursdayMorning.BackColor = LblBackColour
'chkThursdayMorning.ForeColor = LblForeColour
'
'chkTuesdayEvening.BackColor = LblBackColour
'chkTuesdayEvening.ForeColor = LblForeColour
'
'chkTuesdayFullLeave.BackColor = LblBackColour
'chkTuesdayFullLeave.ForeColor = LblForeColour
'
'chkTuesdayEvening.BackColor = LblBackColour
'chkTuesdayEvening.ForeColor = LblForeColour
'
'chkTuesdayMorning.BackColor = LblBackColour
'chkTuesdayMorning.ForeColor = LblForeColour
'
'chkWednesdayEvening.BackColor = LblBackColour
'chkWednesdayEvening.ForeColor = LblForeColour
'
'chkWednesdayFullLeave.BackColor = LblBackColour
'chkWednesdayFullLeave.ForeColor = LblForeColour
'
'chkWednesdayEvening.BackColor = LblBackColour
'chkWednesdayEvening.ForeColor = LblForeColour
'
'chkWednesdayMorning.BackColor = LblBackColour
'chkWednesdayMorning.ForeColor = LblForeColour

DataComboBank.BackColor = TxtBackColour
DataComboBank.ForeColor = TxtForeColour

'DataComboDoctorStaff.BackColor = TxtBackColour
'DataComboDoctorStaff.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboDoctorStaff.BackColor = TxtBackColour
'DataComboDoctorStaff.ForeColor = TxtForeColour
'
'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour
'
'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour




Grid1.BackColor = GridBackColor
Grid1.ForeColor = GridForeColor

Grid1.BackColorBkg = GridBackColorBkg
Grid1.BackColorFixed = GridBackColorFixed
Grid1.BackColorSel = GridBackColorSel

Grid1.ForeColor = GridForeColor
Grid1.ForeColorFixed = GridForeColorFixed
Grid1.ForeColorSel = GridForeColorSel

Grid1.ForeColor = GridForeColor




Label1.BackColor = LblBackColour
Label1.ForeColor = LblForeColour

'lblDoctorStaff.BackColor = LblBackColour
'lblDoctorStaff.ForeColor = LblForeColour
'lblInstitutionFee.BackColor = LblBackColour
'lblInstitutionFee.ForeColor = LblForeColour
'Lbl.BackColor = LblBackColour
'LblCommentsLX.ForeColor = LblForeColour
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
'
'Label8.BackColor = LblBackColour
'Label8.ForeColor = LblForeColour
'Label9.BackColor = LblBackColour
'Label9.ForeColor = LblForeColour
'
'lblAmount.BackColor = LblBackColour
'lblAmount.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour
'
'lblCashPaid.BackColor = LblBackColour
'lblCashPaid.ForeColor = LblForeColour
'
'lblChequeAmount.BackColor = LblBackColour
'lblChequeAmount.ForeColor = LblForeColour
'
'lblThisTimeCredit.BackColor = LblBackColour
'lblThisTimeCredit.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour

'chkFridayFullLeave.BackColor = FrameBackColour
'chkFridayFullLeave.ForeColor = FrameForeColour
'
'chkFridayMorning.BackColor = FrameBackColour
'chkFridayMorning.ForeColor = FrameForeColour
'
'chkFullDayLeave.BackColor = FrameBackColour
'chkFullDayLeave.ForeColor = FrameForeColour

'chkHLine2.BackColor = FrameBackColour
'chkHLine2.ForeColor = FrameForeColour
'
'chkHLine3.BackColor = FrameBackColour
'chkHLine3.ForeColor = FrameForeColour
''
'chkHLine4.BackColor = FrameBackColour
'chkHLine4.ForeColor = FrameForeColour

'chk.BackColor = FrameBackColour
'chkParaethesia.ForeColor = FrameForeColour
'
'chkMuscleWeak.BackColor = FrameBackColour
'chkMuscleWeak.ForeColor = FrameForeColour
'
'chkSleep.BackColor = FrameBackColour
'chkSleep.ForeColor = FrameForeColour
'
'chkVisual.BackColor = FrameBackColour
'chkVisual.ForeColor = FrameForeColour
''
'chkSmell.BackColor = FrameBackColour
'chkSmell.ForeColor = FrameForeColour
'
'chkTaste.BackColor = FrameBackColour
'chkTaste.ForeColor = FrameForeColour
''
'chkSpeech.BackColor = FrameBackColour
'chkSpeech.ForeColor = FrameForeColour
'
'chkPsychiatric.BackColor = FrameBackColour
'chkPsychiatric.ForeColor = FrameForeColour
'
'chkThinHair.BackColor = FrameBackColour
'chkThinHair.ForeColor = FrameForeColour
'
'chkHoarseVoice.BackColor = FrameBackColour
'chkHoarseVoice.ForeColor = FrameForeColour
'
'chkUrgency.BackColor = FrameBackColour
'chkUrgency.ForeColor = FrameForeColour
'
'chkUrinaryFrequency.BackColor = FrameBackColour
'chkUrinaryFrequency.ForeColor = FrameForeColour
'
'chkUrgeIncontinence.BackColor = FrameBackColour
'chkUrgeIncontinence.ForeColor = FrameForeColour
'
'txt.BackColor = TxtBackColour
'txtAddress.ForeColor = TxtForeColour
'
'txtAge.BackColor = TxtBackColour
'txtAge.ForeColor = TxtForeColour
'
'txtAgentBalance.BackColor = TxtBackColour
'txtAgentBalance.ForeColor = TxtForeColour
'txtAuthorizationCode.BackColor = TxtBackColour
'txtAuthorizationCode.ForeColor = TxtForeColour
'txtCashDue.BackColor = TxtBackColour
'txtCashDue.ForeColor = TxtForeColour
'txtChequeNo.BackColor = TxtBackColour
'txtChequeNo.ForeColor = TxtForeColour
'txtDiscount.BackColor = TxtBackColour
'txtDiscount.ForeColor = TxtForeColour
'txtEmail.BackColor = TxtBackColour
'txtEmail.ForeColor = TxtForeColour
'txtFax.BackColor = TxtBackColour
'txtFax.ForeColor = TxtForeColour
'txtFirstName.BackColor = TxtBackColour
'txtFirstName.ForeColor = TxtForeColour
'txtGrossTotal.BackColor = TxtBackColour
'txtGrossTotal.ForeColor = TxtForeColour
'txtNetTotal.BackColor = TxtBackColour
'txtNetTotal.ForeColor = TxtForeColour
'
'txtNIC.BackColor = TxtBackColour
'txtNIC.ForeColor = TxtForeColour
'txtNotes.BackColor = TxtBackColour
'txtNotes.ForeColor = TxtForeColour
'txtOtherName.BackColor = TxtBackColour
'txtOtherName.ForeColor = TxtForeColour
'txtPaidForCredit.BackColor = TxtBackColour
'txtPaidForCredit.ForeColor = TxtForeColour
'txtSearchFirstName.BackColor = TxtBackColour
'txtSearchFirstName.ForeColor = TxtForeColour
'
'
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'
'Text2.BackColor = TxtBackColour
'Text2.ForeColor = TxtForeColour
'
'Text3.BackColor = TxtBackColour
'Text3.ForeColor = TxtForeColour
'
'Text4.BackColor = TxtBackColour
'Text4.ForeColor = TxtForeColour
'
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'
'
'
'OptionAgent.BackColor = TxtBackColour
'OptionAgent.ForeColor = TxtForeColour
'
'OptionCash.BackColor = TxtBackColour
'OptionCash.ForeColor = TxtForeColour
'
'OptionCheque.BackColor = TxtBackColour
'OptionCheque.ForeColor = TxtForeColour
'OptionDoNotPrint.BackColor = TxtBackColour
'OptionDoNotPrint.ForeColor = TxtForeColour
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCreditCard.BackColor = TxtBackColour
'OptionCreditCard.ForeColor = TxtForeColour
'
'OptionMaster.BackColor = TxtBackColour
'OptionMaster.ForeColor = TxtForeColour
'
'OptionPrintOne.BackColor = TxtBackColour
'OptionPrintOne.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'
'
'txtSearchID.BackColor = TxtBackColour
'txtSearchID.ForeColor = TxtForeColour
'txtSearchSurname.BackColor = TxtBackColour
'txtSearchSurname.ForeColor = TxtForeColour
'txtSurname.BackColor = TxtBackColour
'txtSurname.ForeColor = TxtForeColour
'txtTelephone.BackColor = TxtBackColour
'txtTelephone.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour





End Sub
Private Sub FormatGrid()
With Grid1
    .Clear
    .Cols = 5
    .ColWidth(0) = 500
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 1
    .ColWidth(4) = 1000
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + 350)
    .Row = 0
    .Col = 0
    .Text = "ID"
    .Col = 1
    .Text = "Institution Name"
    .CellAlignment = 1
    .Col = 2
    .Text = "Code"
    .CellAlignment = 7
    .Col = 3
    .Text = ""
    .Col = 4
    .Text = "Balance"
    .CellAlignment = 7
    .Rows = 2
End With
End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnSerchBill_Click()
Call FindDueBills
End Sub

Private Sub checkPaymentMode()
Dim A As Integer
Dim B As Integer
Dim C As Integer


PaymentSuccess = False

If txtReceiptNo.Text = "" Then A = MsgBox("Enter Receipt No", vbCritical + vbOKOnly, "Error"): Exit Sub

If TemInstitutionID = 0 Then A = MsgBox("Click Institution Name From Grid", vbCritical + vbOKOnly, "Error"): Exit Sub

If OptionCash = False And OptionCheque = False And OptionCreditCard.Value = False Then A = MsgBox("Select Payment Mode", vbCritical + vbOKOnly, "Error"): Exit Sub

If OptionCash = True Then
If txtCashPayment.Text = "" Or Val(txtCashPayment.Text) <= 0 Then A = MsgBox("Enter Cash Amount", vbCritical + vbOKOnly, "Error"): Exit Sub

ElseIf OptionCheque = True Then

If DataComboBank.Text = "" Then A = MsgBox("Select Bank", vbCritical + vbOKOnly, "Error"): Exit Sub
If txtChequeAmount.Text = "" Or Val(txtChequeAmount.Text) <= 0 Then B = MsgBox("Enter Cheque Amount", vbCritical + vbOKOnly, "Error"): Exit Sub
If txtChequeNo.Text = "" Then C = MsgBox("Enter Cheque No", vbCritical + vbOKOnly, "Error"): Exit Sub

ElseIf OptionCreditCard = True Then

If OptionVISA.Value = False And OptionABC.Value = False And OptionMaster.Value = False And OptionAmEx = False Then A = MsgBox("Select Credit Card Name", vbCritical + vbOKOnly, "Error"): Exit Sub
If txtAmount.Text = "" Or Val(txtAmount.Text) <= 0 Then B = MsgBox("Enter Payment Amount", vbCritical + vbOKOnly, "Error"): Exit Sub
If txtAuthorizationCode = "" Then A = MsgBox("Enter Credit Card Authorization Code", vbCritical + vbOKOnly, "Error"): Exit Sub

End If
PaymentSuccess = True
End Sub

Private Sub bttnUpdate_Click()
    Call checkPaymentMode
    If PaymentSuccess = False Then Exit Sub
    If DuplicateNo = True Then
        MsgBox "This receipt number is already entered"
        txtReceiptNo.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    Call PaymentUpdate
    Call UpdateInstitutionbill
    Call ClearVales
End Sub
'
Private Sub PaymentUpdate()

With DataEnvironment1.rssqlAgentCashSettle
If .State = 1 Then .Close

.Source = "Select tblAgentCashSettle.* From tblAgentCashSettle"
If .State = 0 Then .Open

.AddNew
!Institution_Id = TemInstitutionID
!SettledDate = dtpPaymentdate.Value
!ReceiptNo = txtReceiptNo.Text
If OptionCash.Value = True Then !SettleMethod = "Cash": !Cash = Val(txtCashPayment.Text)
If OptionCheque.Value = True Then !SettleMethod = "Cheque": !Cash = Val(txtChequeAmount.Text): !ChequeAmount = Val(txtChequeAmount): !Bank_ID = DataComboBank.BoundText: !ChequeNo = txtChequeNo.Text: !ChequeDate = DTPickerChequeDate.Value
If OptionCreditCard.Value = True Then
    If OptionCreditCard.Value = True Then !SettleMethod = "CreditCard": !Cash = Val(txtAmount.Text): !CreditCardAmount = Val(txtAmount.Text): !AuthorizationCode = txtAuthorizationCode.Text
    If OptionVISA.Value = True Then !CreditCard = "Visa"
    If OptionMaster.Value = True Then !CreditCard = "Master"
    If OptionAmEx.Value = True Then !CreditCard = "American Experss"
    If OptionABC.Value = True Then !CreditCard = "ABC"
End If
!user_ID = UserID
'!Branch
'CreditCardNo
'ExpiaryDate
.Update

If .State = 1 Then .Close

End With


End Sub
'
Private Sub UpdateInstitutionbill()


With DataEnvironment1.rssqlTem14

    If .State = 1 Then .Close
    .Source = "Select tblInstitutions.* From tblInstitutions Where ( Institution_ID = " & TemInstitutionID & " ) "
    If .State = 0 Then .Open
    
    If .RecordCount = 0 Then Exit Sub

    If OptionCash.Value = True Then
    
        !InstitutionCredit = Val(!InstitutionCredit) + Val(txtCashPayment)
        .Update
    End If

    If OptionCheque.Value = True Then
    
        !InstitutionCredit = Val(!InstitutionCredit) + Val(txtChequeAmount)
        .Update
    End If

    If OptionCreditCard.Value = True Then
    
        !InstitutionCredit = Val(!InstitutionCredit) + Val(txtAmount.Text)
        .Update
    End If

    If .State = 1 Then .Close

End With


End Sub
'
Private Sub ClearVales()

OptionCash.Value = True
txtCashPayment.Text = ""
lblCashBalance.Caption = "0.00"
OptionCheque.Value = False
txtChequeAmount.Text = ""
DataComboBank.Text = ""
txtChequeNo.Text = ""
OptionCreditCard = False
txtAmount.Text = ""
txtAuthorizationCode.Text = ""
lblInstitutionBalance.Caption = ""
lblInstitutionName.Caption = ""
txtReceiptNo.Text = ""
OptionVISA = False
OptionMaster = False
OptionAmEx = False
OptionABC = False
Call FormatGrid
End Sub

Private Sub cmbAgentCode_Change()
dtcAgentCode.Text = cmbAgentCode.Text
If IsNumeric(dtcAgentCode.BoundText) = False Then Exit Sub
FindInstitutionBalance2
dtcAgentName.BoundText = dtcAgentCode.BoundText
End Sub

Private Sub FindInstitutionBalance2()

lblInstitutionBalance.Caption = "0.00"
If IsNumeric(dtcAgentCode.BoundText) = False Then Exit Sub

With DataEnvironment1.rssqlPatientBill

    If .State = 1 Then .Close
    .Source = "Select tblInstitutions.* From tblInstitutions Where (Institution_ID = " & dtcAgentCode.BoundText & " )" 'and InstitutionCredit  <> 0) "
    If .State = 0 Then .Open

    If .RecordCount = 0 Then Exit Sub
    TemInstitutionID = Val(dtcAgentCode.BoundText)
    
    lblInstitutionBalance.Caption = Format(!InstitutionCredit, "0.00")
    lblInstitutionName.Caption = !InstitutionName
    
    txtChequeAmount.Text = Format(!InstitutionCredit, "0.00")
    txtAmount.Text = Format(!InstitutionCredit, "0.00") 'Credit card
    txtChequeAmount.Text = Format(!InstitutionCredit, "0.00")
    
    i = 1
    rcount = 1
    Call FormatGrid
    
        Do While .EOF = False
        rcount = rcount + 1
        Grid1.Rows = rcount
        Grid1.TextMatrix(i, 0) = i
        Grid1.TextMatrix(i, 1) = !InstitutionName
        Grid1.CellAlignment = 1
        Grid1.TextMatrix(i, 2) = !InstitutionCode
        Grid1.CellAlignment = 7
        Grid1.TextMatrix(i, 3) = !Institution_Id
        Grid1.TextMatrix(i, 4) = Format(!InstitutionCredit, "0.00")
        Grid1.CellAlignment = 7
        
        i = i + 1
        .MoveNext
        Loop
    
    If .State = 1 Then .Close
   
End With

End Sub

Private Sub fillgrid2()

'    I = 1
'    rcount = 1
'
'    Do While .EOF = False
'    rcount = rcount + 1
'    grid1.Rows = rcount
'    grid1.TextMatrix(I, 0) = I
'    grid1.TextMatrix(I, 1) = !InstitutionName
'    grid1.TextMatrix(I, 2) = Format(!InstitutionCredit, "0.00")
'    grid1.TextMatrix(I, 3) = !Institution_ID
'    I = I + 1
'    .MoveNext
'    Loop


End Sub

Private Sub cmbAgentCode_KeyPress(KeyAscii As Integer)
dtcAgentCode.Text = cmbAgentCode.Text
If IsNumeric(dtcAgentCode.BoundText) = False Then Exit Sub
FindInstitutionBalance2
dtcAgentName.BoundText = dtcAgentCode.BoundText

End Sub

Private Sub dtcAgentCode_Change()
If IsNumeric(dtcAgentCode.BoundText) = False Then Exit Sub
dtcAgentName.BoundText = dtcAgentCode.BoundText
cmbAgentCode.Text = dtcAgentCode.Text

End Sub

Private Sub cmbAgentCode_Click()
dtcAgentCode.Text = cmbAgentCode.Text
If IsNumeric(dtcAgentCode.BoundText) = False Then Exit Sub
FindInstitutionBalance2
dtcAgentName.BoundText = dtcAgentCode.BoundText
End Sub

Private Sub dtcAgentCode_Click(Area As Integer)
If IsNumeric(dtcAgentCode.BoundText) = False Then Exit Sub
dtcAgentName.BoundText = dtcAgentCode.BoundText
cmbAgentCode.Text = dtcAgentCode.Text
End Sub

Private Sub dtcAgentName_Click(Area As Integer)
If IsNumeric(dtcAgentName.BoundText) = False Then Exit Sub
dtcAgentCode.BoundText = dtcAgentName.BoundText
FindInstitutionBalance2
End Sub

Private Sub Form_Load()
Call FormatGrid
Call SetColour
dtpPaymentdate = Date
DTPickerChequeDate = Date
OptionCash.Value = True
With DataEnvironment1.rssqlTem1

    If .State = 1 Then .Close
    .Source = "Select tblInstitutions.* From tblInstitutions Order by InstitutionName " ' Where " '(InstitutionCredit  <> 0) "
    If .State = 0 Then .Open

    If .RecordCount = 0 Then Exit Sub
    
        Do While .EOF = False
        cmbAgentCode.AddItem !InstitutionCode
        .MoveNext
        Loop
        
    If .State = 1 Then .Close
End With

With DataEnvironment1.rssqlTemAgent123

    If .State = 1 Then .Close
    
    .Source = "Select tblInstitutions.* From tblInstitutions Order by InstitutionName " ' Where " '(InstitutionCredit  <> 0) "
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    
    Set dtcAgentCode.RowSource = DataEnvironment1.rssqlTemAgent123
    dtcAgentCode.BoundColumn = "Institution_ID"
    dtcAgentCode.ListField = "InstitutionCode"
    
    Set dtcAgentName.RowSource = DataEnvironment1.rssqlTemAgent123
    dtcAgentName.BoundColumn = "Institution_ID"
    dtcAgentName.ListField = "InstitutionName"
   
End With

End Sub
'
Private Sub FindDueBills()

With DataEnvironment1.rssqlPatientBill

    If .State = 1 Then .Close
    .Source = "Select tblInstitutions.* From tblInstitutions Order by InstitutionName " ' Where " '(InstitutionCredit  <> 0) "
    If .State = 0 Then .Open

    If .RecordCount = 0 Then Exit Sub
    i = 1
    rcount = 1

    Do While .EOF = False
    rcount = rcount + 1
    Grid1.Rows = rcount
    Grid1.TextMatrix(i, 0) = i
    Grid1.TextMatrix(i, 1) = !InstitutionName
    Grid1.CellAlignment = 1
    Grid1.TextMatrix(i, 2) = !InstitutionCode
    Grid1.CellAlignment = 7
    Grid1.TextMatrix(i, 3) = !Institution_Id
    Grid1.TextMatrix(i, 4) = Format(!InstitutionCredit, "0.00")
     Grid1.CellAlignment = 7
    i = i + 1
    .MoveNext
    Loop

    If .State = 1 Then .Close

End With


End Sub

Private Sub Grid1_Click()
Call FindInstitutionBalance
Grid1.Col = 0
Grid1.ColSel = Grid1.Cols - 1
End Sub

Private Sub FindInstitutionBalance()
Grid1.Col = 3

lblInstitutionBalance.Caption = "0.00"

If IsNumeric(Grid1.Text) = False Then Exit Sub

With DataEnvironment1.rssqlTem16

    If .State = 1 Then .Close
    .Source = "Select tblInstitutions.* From tblInstitutions Where (Institution_ID= " & Grid1.Text & " )" 'and InstitutionCredit  <> 0) "
    .Open

    If .RecordCount = 0 Then Exit Sub
    TemInstitutionID = Val(Grid1.Text)
    dtcAgentCode.BoundText = TemInstitutionID
    lblInstitutionBalance.Caption = Format(!InstitutionCredit, "0.00")
    lblInstitutionName.Caption = !InstitutionName
    
    txtChequeAmount.Text = Format(!InstitutionCredit, "0.00")
    txtAmount.Text = Format(!InstitutionCredit, "0.00") 'Credit card
    txtChequeAmount.Text = Format(!InstitutionCredit, "0.00")
    
    If .State = 1 Then .Close
   
End With

End Sub

Private Function DuplicateNo() As Boolean
    DuplicateNo = True
    With DataEnvironment1.rssqlTem16
        If .State = 1 Then .Close
        .Source = "Select tblAgentCashSettle.* From tblAgentCashSettle Where (ReceiptNo = '" & txtReceiptNo.Text & "')"
        .Open
        If .RecordCount = 0 Then
            DuplicateNo = False
        Else
            DuplicateNo = True
        End If
        If .State = 1 Then .Close
    End With
End Function

Private Sub grid1_KeyPress(KeyAscii As Integer)
If IsNumeric(Grid1.Text) = False Then Exit Sub
Call FindInstitutionBalance
End Sub

Private Sub OptionCash_Click()
FrameCreditCard.Visible = False
FrameCheque.Visible = False
FrameCash.Visible = True
End Sub

Private Sub OptionCheque_Click()
FrameCreditCard.Visible = False
FrameCash.Visible = False
FrameCheque.Visible = True
End Sub

Private Sub OptionCreditCard_Click()
FrameCheque.Visible = False
FrameCash.Visible = False
FrameCreditCard.Visible = True
End Sub
'
Private Sub txtCashPayment_Change()
lblCashBalance.Caption = Val(lblInstitutionBalance) + Val(txtCashPayment)
End Sub

Private Sub txtCashPayment_KeyPress(KeyAscii As Integer)
If Val(txtCashPayment.Text) > 0 And KeyAscii = 13 Then bttnUpdate_Click
End Sub
