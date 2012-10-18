VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paitient Payment"
   ClientHeight    =   8430
   ClientLeft      =   4740
   ClientTop       =   1995
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   9090
   Begin VB.Frame Frame1 
      Height          =   7935
      Left            =   360
      TabIndex        =   14
      Top             =   240
      Width           =   8415
      Begin MSDataListLib.DataCombo dtcPatient 
         Height          =   360
         Left            =   2880
         TabIndex        =   41
         Top             =   600
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpPaymentdate 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   4680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61931521
         CurrentDate     =   39423
      End
      Begin btButtonEx.ButtonEx ButtonEx3 
         Height          =   495
         Left            =   6240
         TabIndex        =   13
         Top             =   7080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
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
         Height          =   495
         Left            =   2640
         TabIndex        =   12
         Top             =   7080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
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
         Height          =   495
         Left            =   480
         TabIndex        =   0
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "&Serch Credit Patients"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame FramePaymentMethod 
         Caption         =   "Payment Method"
         Height          =   1695
         Left            =   5880
         TabIndex        =   15
         Top             =   5160
         Width           =   2295
         Begin VB.OptionButton OptionCash 
            Caption         =   "&Cash"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton OptionCheque 
            Caption         =   "Che&que"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton OptionCreditCard 
            Caption         =   "Credi&t Card"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   1695
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   2175
         Left            =   480
         TabIndex        =   17
         Top             =   1200
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3836
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   2
      End
      Begin VB.Frame FrameCash 
         Caption         =   "Cash"
         Height          =   1695
         Left            =   240
         TabIndex        =   22
         Top             =   5160
         Width           =   5415
         Begin VB.TextBox txtCashPayment 
            Height          =   375
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   23
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Change   :"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblCashBalance 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rs. 0.00"
            Height          =   375
            Left            =   3600
            TabIndex        =   24
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Cash      :                                      Rs."
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Frame FrameCreditCard 
         Caption         =   "Credit Card"
         Height          =   1695
         Left            =   240
         TabIndex        =   35
         Top             =   5160
         Width           =   5415
         Begin VB.TextBox txtAmount 
            Height          =   375
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   10
            Top             =   720
            Width           =   2895
         End
         Begin VB.OptionButton OptionVISA 
            Caption         =   "&VISA"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptionMaster 
            Caption         =   "&MASTER"
            Height          =   255
            Left            =   1320
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptionAmEx 
            Caption         =   "&AmEX"
            Height          =   255
            Left            =   2640
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtAuthorizationCode 
            Height          =   375
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   11
            Top             =   1200
            Width           =   2895
         End
         Begin VB.OptionButton OptionABC 
            Caption         =   "A&BC"
            Height          =   255
            Left            =   3720
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "Authori&zation Code:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Amo&unt                :"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   2895
         End
      End
      Begin VB.Frame FrameCheque 
         Caption         =   "Cheque"
         Height          =   1695
         Left            =   240
         TabIndex        =   27
         Top             =   5160
         Width           =   5415
         Begin VB.TextBox txtChequeAmount 
            Height          =   375
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   39
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtChequeNo 
            Height          =   360
            Left            =   840
            MaxLength       =   20
            TabIndex        =   28
            Top             =   1200
            Width           =   1695
         End
         Begin MSDataListLib.DataCombo DataComboBank 
            Bindings        =   "frmPayment.frx":0442
            Height          =   360
            Left            =   2280
            TabIndex        =   29
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
            TabIndex        =   30
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   61931523
            CurrentDate     =   39414
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank      :"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "Date :"
            Height          =   255
            Left            =   2880
            TabIndex        =   32
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount :"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "No. :"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.Label lblPatientName 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   4560
         TabIndex        =   43
         Top             =   4200
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   42
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment &Date"
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Bill Balance"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label lblDuebalance 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2640
         TabIndex        =   20
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Date"
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Patient Bill ID"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame framCoustomerbill 
      Height          =   2775
      Left            =   600
      TabIndex        =   38
      Top             =   1080
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid grid2 
         Height          =   2175
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   3836
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   2
      End
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim rcount As Integer
Dim TemPatientName As String
Dim TemTotalDuebalance As Double
Dim PaymentSuccess As Boolean
Dim TemPaitientBillID As Long
Dim TemBalance As Double
Dim A, B, C, D, E, F

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

ButtonEx3.BackColor = BttnBackColour
ButtonEx3.ForeColor = BttnForeColour

bttnSerchBill.BackColor = BttnBackColour
bttnSerchBill.ForeColor = BttnForeColour

bttnUpdate.BackColor = BttnBackColour
bttnUpdate.ForeColor = BttnForeColour


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

FrameCash.BackColor = FrameBackColour
FrameCash.ForeColor = FrameForeColour

Frame1.BackColor = FrmBackColour
Frame1.ForeColor = FrmForeColour

FrameCheque.BackColor = FrameBackColour
FrameCheque.ForeColor = FrameForeColour


FrameCreditCard.BackColor = FrameBackColour
FrameCreditCard.ForeColor = FrameForeColour

FramePaymentMethod.BackColor = FrameBackColour
FramePaymentMethod.ForeColor = FrameForeColour

framCoustomerbill.BackColor = FrameBackColour
framCoustomerbill.ForeColor = FrameForeColour

frmPayment.BackColor = FrameBackColour
frmPayment.ForeColor = FrameForeColour

'FrameCheque.BackColor = FrameBackColour
'FrameCheque.ForeColor = FrameForeColour
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
'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour



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
    .Cols = 6
    .ColWidth(0) = 500
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 1300
    .ColWidth(4) = 1
    .ColWidth(5) = 1
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + 350)
    .Row = 0
    .Col = 0
    .Text = "Invoice No"
    .Col = 1
    .Text = "Patient Name"
    .Col = 2
    .Text = "Balance"
    .CellAlignment = 6
    .Col = 3
    .Text = "Date"
    .Rows = 2
End With
End Sub

Private Sub FormatGrid2()
'Dim MyVoice As New SpeechLib.SpVoice
' MyVoice.Speak "Good Afternoon, Mr. Sudesh Pathirana,  I wish you a very successful day to you."
'With grid2
'    .Clear
'    .Cols = 6
'    .ColWidth(0) = 500
'    .ColWidth(1) = 1500
'    .ColWidth(2) = 1500
'    .ColWidth(3) = 1300
'    .ColWidth(4) = 1
'    .ColWidth(5) = 1
'    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + 350)
'    .Row = 0
'    .Col = 0
'    .Text = "Invoice No"
'    .Col = 1
'    .Text = "Patient Name"
'    .Col = 2
'    .Text = "Balance"
'    .CellAlignment = 6
'    .Col = 3
'    .Text = "Date"
'    .Rows = 2
'End With

End Sub


Private Sub bttnSerchBill_Click()
framCoustomerbill.Visible = False
Grid1.Visible = True
Call FormatGrid
Call FindDueBills
End Sub

Private Sub checkPaymentMode()
PaymentSuccess = False

If dtcPatient.Text = "" Then A = MsgBox("Select Patient Name", vbCritical + vbOKOnly, "Error"): Exit Sub

If OptionCash = False And OptionCheque = False And OptionCreditCard.Value = False Then A = MsgBox("Select Payment Mode", vbCritical + vbOKOnly, "Error"): Exit Sub

If OptionCash = True Then
If txtCashPayment.Text = "" Then A = MsgBox("Enter Cash Amount", vbCritical + vbOKOnly, "Error"): Exit Sub

ElseIf OptionCheque = True Then

If DataComboBank.Text = "" Then A = MsgBox("Select Bank", vbCritical + vbOKOnly, "Error"): Exit Sub
If txtChequeAmount.Text = "" Then B = MsgBox("Enter Cheque Amount", vbCritical + vbOKOnly, "Error"): Exit Sub
If txtChequeNo.Text = "" Then C = MsgBox("Enter Cheque No", vbCritical + vbOKOnly, "Error"): Exit Sub

ElseIf OptionCreditCard = True Then

If OptionVISA.Value = False And OptionABC.Value = False And OptionMaster.Value = False And OptionAmEx = False Then A = MsgBox("Select Credit Card Name", vbCritical + vbOKOnly, "Error"): Exit Sub
If txtAmount.Text = "" Then B = MsgBox("Enter Payment Amount", vbCritical + vbOKOnly, "Error"): Exit Sub
If txtAuthorizationCode = "" Then A = MsgBox("Enter Credit Card Authorization Code", vbCritical + vbOKOnly, "Error"): Exit Sub

End If
PaymentSuccess = True
End Sub


Private Sub bttnUpdate_Click()
Call checkPaymentMode
If PaymentSuccess = False Then Exit Sub
Call PaymentUpdate
Call UpdatePatientbill
Call PatientFacilityUpdate
Call ClearVales
Call FindDueBills
OptionCash.Value = True

End Sub

Private Sub PatientFacilityUpdate()
Dim TemDoctorval As Double
Dim TemInstituioncal As Double
Grid1.Col = 4
With DataEnvironment1.rssqlTem7

    If .State = 1 Then .Close
    .Source = "Select * From tblPatientFacility Where (PatientBill_ID =" & Grid1.Text & ") and (PatientID =" & Val(dtcPatient.BoundText) & " )"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    
    !CreditSettleUser_ID = UserID
    TemDoctorval = !PersonalFeeToPay
    TemInstituioncal = !InstitutionFeeToPay
    !personalfee = TemDoctorval
    !personaldue = TemDoctorval
    !PersonalFeeToPay = 0
    !institutionfee = TemInstituioncal
    !institutiondue = TemInstituioncal
    !InstitutionFeeToPay = 0
    
        If OptionCash.Value = True Then
            !totalfee = Val(txtCashPayment.Text)
            !TotalDue = Val(txtCashPayment.Text)
        ElseIf OptionCheque.Value = True Then
            !totalfee = txtChequeAmount.Text
            !TotalDue = txtChequeAmount.Text
        ElseIf OptionCreditCard.Value = True Then
            !totalfee = txtAmount.Text
            !TotalDue = txtAmount.Text
        End If
    
    !fullypaid = 1
    !fullypaidnull = 1
    !SettleCashDate = dtpPaymentdate.Value
    .Update
    If .State = 1 Then .Close

End With

End Sub

Private Sub PaymentUpdate()

With DataEnvironment1.rssqlPatientCashSettle
If .State = 1 Then .Close

.Source = "Select tblPatientCashSettle.* From tblPatientCashSettle"

If .State = 0 Then .Open
.AddNew
!patient_ID = Val(dtcPatient.BoundText)
!SettledDate = dtpPaymentdate
If OptionCash = True Then !SettleMethod = "Cash": !Cash = Val(txtCashPayment.Text)
If OptionCheque = True Then !SettleMethod = "Cheque": !ChequeAmount = Val(txtChequeAmount): !Bank_ID = DataComboBank.BoundText: !ChequeNo = txtChequeNo.Text: !ChequeDate = DTPickerChequeDate.Value
If OptionCreditCard = True Then !SettleMethod = "CreditCard": !CreditCardAmount = Val(txtAmount.Text): AuthorizationCode = txtAuthorizationCode
If OptionVISA = True Then !CreditCard = "Visa"
If OptionMaster = True Then !CreditCard = "Master"
If OptionAmEx = True Then !CreditCard = "American Experss"
If OptionABC = True Then !CreditCard = "ABC"
'!Branch
'CreditCardNo
'ExpiaryDate
!user_ID = UserID
.Update

If .State = 1 Then .Close

End With


End Sub

Private Sub UpdatePatientbill() ' update Paitent bill
Grid1.Col = 4
Dim TemSelectedbillvalue As Double
TemSelectedbillvalue = 0
TemBalance = 0

With DataEnvironment1.rssqlTem1
    If .State = 1 Then .Close
    .Source = "Select tblPatientBill.* From tblPatientBill Where ( PatientBill_ID = " & Grid1.Text & ") and (Patient_ID  = " & Val(dtcPatient.BoundText) & " ) "
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    
    If OptionCash.Value = True Then
        If Val(!Credit) >= Val(txtCashPayment.Text) Then
        !Credit = Val(!Credit) - Val(txtCashPayment)
        .Update
        Else
        TemBalance = Val(txtCashPayment.Text) - Val(!Credit)
        TemSelectedbillvalue = Val(txtCashPayment.Text) - Val(TemBalance)
        !Credit = Val(!Credit) - Val(TemSelectedbillvalue)
'        !Cash = !Cash + Val(TemSelectedbillvalue)
        .Update
'            If Val(TemBalance) > 0 Then
'            Call UpdatePatientBalance
'            End If
        End If
    End If
    
    If OptionCheque.Value = True Then
    If Val(!Credit) >= Val(txtChequeAmount.Text) Then
        !Credit = Val(!Credit) - Val(txtChequeAmount)
        .Update
        Else
        TemBalance = Val(txtCashPayment.Text) - Val(!Credit)
        TemSelectedbillvalue = Val(txtCashPayment.Text) - Val(TemBalance)
        !Credit = Val(!Credit) - Val(TemSelectedbillvalue)
'        !Cash = !Cash + Val(TemSelectedbillvalue)
        .Update
       
'            If Val(TemBalance) > 0 Then
'            Call UpdatePatientBalance
'            End If
        End If
    End If
    
    If OptionCreditCard.Value = True Then
    If Val(!Credit) >= Val(txtAmount.Text) Then
        !Credit = Val(!Credit) - Val(txtAmount)
        .Update
        Else
        TemBalance = Val(txtCashPayment.Text) - Val(!Credit)
        TemSelectedbillvalue = Val(txtCashPayment.Text) - Val(TemBalance)
        !Credit = Val(!Credit) - Val(TemSelectedbillvalue)
'        !Cash = !Cash + Val(TemSelectedbillvalue)
        .Update
'            If Val(TemBalance) > 0 Then
'            Call UpdatePatientBalance
'            End If
        End If
    End If
    
    If .State = 1 Then .Close

End With
Call UpdatePatient

End Sub

Private Sub ClearVales()

OptionCash.Value = False
txtCashPayment.Text = ""
lblCashBalance.Caption = "0.00"
OptionCheque.Value = False
txtChequeAmount.Text = ""
DataComboBank.Text = ""
txtChequeNo.Text = ""
OptionCreditCard = False
txtAmount.Text = ""
txtAuthorizationCode.Text = ""
OptionVISA = False
OptionMaster = False
OptionAmEx = False
OptionABC = False
lblPatientName.Caption = ""
lblDate.Caption = ""
lblDuebalance.Caption = "0.00"
Call FormatGrid
End Sub

Private Sub UpdatePatient() ' Patient  Update

With DataEnvironment1.rssqlTem2
If .State = 1 Then .Close
.Source = "Select* From tblPatientMainDetails Where (Patient_ID = " & Val(dtcPatient.BoundText) & ")"
If .State = 0 Then .Open
If .RecordCount = 0 Then Exit Sub
If OptionCash.Value = True Then !Credit = Val(!Credit) + Val(txtCashPayment)
If OptionCreditCard.Value = True Then !Credit = Val(!Credit) + Val(txtAmount)
If OptionCheque.Value = True Then !Credit = Val(!Credit) + Val(txtChequeAmount)
.Update

If .State = 1 Then .Close

End With
End Sub

Private Sub UpdatePatientBalance() 'patient Balance update

With DataEnvironment1.rssqlTem3
    If .State = 1 Then .Close
    .Source = "Select* From tblPatientMainDetails Where (Patient_ID = " & Val(dtcPatient.BoundText) & ")"
    If .State = 0 Then .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    !Credit = Val(!Credit) - Val(TemBalance)
    .Update
    
    If .State = 1 Then .Close

End With



End Sub

Private Sub ButtonEx3_Click()
Unload Me
End Sub

Private Sub cmbPatients_Click()
dtcPatient.Text = cmbPatients.Text
If IsNumeric(dtcPatient.BoundText) = False Then Exit Sub
End Sub

Private Sub dtcPatient_Change()
'grid1.Visible = False
'framCoustomerbill.Visible = True
'Call FindPatientBalance
'Call FindPatientDueBills
End Sub

Private Sub Form_Load()
'dtcPatient.RowSource = ""
'dtcPatient.ListField = ""
'dtcPatient.BoundColumn = ""

'
With DataEnvironment1.rssqlTem12
    If .State = 1 Then .Close
    .Open "Select tblPatientFacility.*, tblPatientMainDetails.* FROM tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.fullypaid = 0) and (tblPatientFacility.PaymentMode = 'Credit') and (tblPatientFacility.ResultSuccess = True )Order by tblPatientMainDetails.FirstName"
'    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    
'    Do While .EOF = False
'    If Not (!firstname) = "" Then cmbPatients.AddItem !firstname
'
'    .MoveNext
'    Loop
'    .MoveFirst
    
    Set dtcPatient.RowSource = DataEnvironment1.rssqlTem12
    dtcPatient.BoundColumn = "Patient_ID"
    dtcPatient.ListField = "PatientFacility_ID"

'    If .State = 1 Then .Close

End With


'With DataEnvironment1.rssqlTem12
'
'    If .State = 1 Then .Close
'    .Source = "Select tblPatientBill.*, tblPatientMainDetails.* FROM tblPatientBill Left Join tblPatientMainDetails On tblPatientBill.Patient_Id = tblPatientMainDetails.Patient_ID Where (tblPatientBill.BillSuccess = True) and (tblPatientBill.Credit > 0) Order by tblPatientMainDetails.FirstName"
''    .Open
'
''    If .RecordCount = 0 Then Exit Sub
'
'End With
OptionCash.Value = True
Call SetColour
Call FormatGrid
Call FormatGrid2
dtpPaymentdate = Date
End Sub


Private Sub FindDueBills()
'lblNumberofbills.Caption = ""
With DataEnvironment1.rssqlTem4

    If .State = 1 Then .Close
    .Source = "Select tblPatientBill.*, tblPatientMainDetails.* FROM tblPatientBill Left Join tblPatientMainDetails On tblPatientBill.Patient_Id = tblPatientMainDetails.Patient_ID Where (tblPatientBill.BillSuccess = True) and (tblPatientBill.PaymentMethod = 'Credit')and (tblPatientBill.Credit > 0) and (tblPatientBill.Cash <> tblPatientBill.NetTotal )Order by tblPatientMainDetails.FirstName"
    If .State = 0 Then .Open
    
    If .RecordCount = 0 Then Exit Sub
    i = 1
    rcount = 1
    
    Do While .EOF = False
    rcount = rcount + 1
    Grid1.Rows = rcount
    Grid1.TextMatrix(i, 0) = i
    If Not DataEnvironment1.rssqlTem4.Fields(28) = "" Then Grid1.TextMatrix(i, 1) = DataEnvironment1.rssqlTem4.Fields(28)
    Grid1.TextMatrix(i, 2) = Format(DataEnvironment1.rssqlTem4.Fields(9), "0.00")
    Grid1.TextMatrix(i, 3) = !Date
    Grid1.TextMatrix(i, 4) = !PatientBill_ID
    Grid1.TextMatrix(i, 5) = DataEnvironment1.rssqlTem4.Fields(1)

    i = i + 1
    .MoveNext
    Loop
    
    If .State = 1 Then .Close
    
End With


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

With DataEnvironment1
If .rssqlPatientMain.State = 1 Then .rssqlPatientMain.Close
If .rssqlTem12.State = 1 Then .rssqlTem12.Close

End With
End Sub





Private Sub Grid1_Click()
Grid1.Col = 5
dtcPatient.BoundText = Grid1.Text
Call FindSelectedPatient3
'Call FindTotalDueBalance2
Grid1.Col = 0
Grid1.ColSel = Grid1.Cols - 1
End Sub

Private Sub FindPatientBalance()
lblPaitentBalance.Caption = "0.00"
lblDuebalance.Caption = "0.00"
lblDate.Caption = ""

With DataEnvironment1.rssqlTem6

    If .State = 1 Then .Close
    .Source = "Select tblPatientMainDetails.* From tblPatientMainDetails Where ( Patient_ID = " & Val(dtcPatient.BoundText) & ") "
    If .State = 0 Then .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    lblPaitentBalance.Caption = Format(!Credit, "0.00")
    
    
    If .State = 1 Then .Close

End With
End Sub

Private Sub FindSelectedPatient2()
'If IsNumeric(dtcPatient.BoundText) = False Then Exit Sub
' lblDuebalance.Caption = "0.00"
'
'With DataEnvironment1.rssqlTem7
'
'    If .State = 1 Then .Close
'    .Source = "Select tblPatientMainDetails.* From tblPatientMainDetails Where ( Patient_ID = " & Val(dtcPatient.BoundText) & ") "
'    If .State = 0 Then .Open
'
'    If .RecordCount = 0 Then Exit Sub
''    lblDate.Caption = !Date
'    lblDuebalance.Caption = Format(!Credit, "0.00")
'    FindTemPatitntBillId
'    If .State = 1 Then .Close
'
'End With
End Sub

Private Sub FindTemPatitntBillId()
grid2.Col = 4
TemPaitientBillID = Val(grid2.Text)

End Sub
Private Sub FindSelectedPatient3()
Grid1.Col = 4
 lblDuebalance.Caption = "0.00"
 
With DataEnvironment1.rssqlTem7

    If .State = 1 Then .Close
    .Source = "Select tblPatientBill.* From tblPatientBill Where (PatientBill_ID = " & Val(Grid1.Text) & ") "

'    .Source = "Select tblPatientMainDetails.* From tblPatientMainDetails Where ( Patient_ID = " & Val(Grid1.Text) & ") "
    If .State = 0 Then .Open

    
    If .RecordCount = 0 Then Exit Sub
    lblDate.Caption = !Date
    lblDuebalance = Format(!Credit, "0.00")
    Grid1.Col = 1
    lblPatientName.Caption = "M/S  " & Grid1.Text
'    FindTemPatitntBillId2
    If .State = 1 Then .Close

End With
End Sub

Private Sub FindTemPatitntBillId2()
grid2.Col = 4
TemPaitientBillID = Val(Grid1.Text)

End Sub

Private Sub FindTotalDueBalance()
'TemTotalDuebalance = Empty
'lblTotalAmount.Caption = "0.00"
'txtAmount.Text = "0.00"
'txtChequeAmount = "0.00"
'grid2.Col = 5
'
'With DataEnvironment1.rssqlPatientBill
'
'    If .State = 1 Then .Close
''    .Source = "Select tblPatientBill.* From tblPatientBill Where ( Patient_ID = " & Val(dtcPatient.BoundText) & " and BillSuccess = True) and ( PaymentMethod = 'Credit')"
''    If .State = 0 Then .Open
'    .Source = "Select tblPatientBill.* From tblPatientBill Where (PatientBill_ID = " & Val(TemPaitientBillID) & ") and (BillSuccess = True) and (PaymentMethod = 'Credit')and (Credit > 0 ) "
'    If .State = 0 Then .Open
'
'    If .RecordCount = 0 Then Exit Sub
'    lblDate.Caption = !Date
'
'    Do While .EOF = False
'
'        TemTotalDuebalance = Val(TemTotalDuebalance) + Val(!Credit)
'
'        .MoveNext
'    Loop
'     lblTotalAmount.Caption = Format(TemTotalDuebalance, "0.00")
'    txtAmount.Text = Format(TemTotalDuebalance, "0.00")
'    txtChequeAmount = Format(TemTotalDuebalance, "0.00")
'
'    If .State = 1 Then .Close
'
'End With
End Sub

Private Sub FindTotalDueBalance2()
'Grid1.Col = 4
'If IsNumeric(Grid1.Text) = False Then Exit Sub
'
'TemTotalDuebalance = Empty
'lblTotalAmount.Caption = "0.00"
'txtAmount.Text = "0.00"
'txtChequeAmount = "0.00"
'
'
'With DataEnvironment1.rssqlPatientBill
'
'    If .State = 1 Then .Close
''    .Source = "Select tblPatientBill.* From tblPatientBill Where ( Patient_ID = " & Val(grid1.Text) & " and BillSuccess = True) and ( PaymentMethod = 'Credit')"
''    If .State = 0 Then .Open
'    .Source = "Select tblPatientBill.* From tblPatientBill Where (PatientBill_ID = " & Val(Grid1.Text) & ") and (BillSuccess = True) and (PaymentMethod = 'Credit')and (Credit > 0 ) "
'    If .State = 0 Then .Open
'
'
'    If .RecordCount = 0 Then Exit Sub
'    lblDate.Caption = !Date
'    Do While .EOF = False
'
'        TemTotalDuebalance = Val(TemTotalDuebalance) + Val(!Credit)
'
'
'        .MoveNext
'    Loop
'
'    lblTotalAmount.Caption = Format(TemTotalDuebalance, "0.00")
'    txtAmount.Text = Format(TemTotalDuebalance, "0.00")
'    txtChequeAmount = Format(TemTotalDuebalance, "0.00")
'
'    If .State = 1 Then .Close
'
'End With
End Sub


Private Sub FindPatientDueBills()
'lblNumberofbills.Caption = "0"
'TemTotalDuebalance = Empty
'lblTotalAmount.Caption = "0.00"
'txtAmount.Text = "0.00"
'txtChequeAmount = "0.00"
'
'With DataEnvironment1.rssqlPatientBill
'
'    If .State = 1 Then .Close
'    .Source = "Select tblPatientBill.* From tblPatientBill Where (Patient_ID = " & Val(dtcPatient.BoundText) & ") and (BillSuccess = True) and (PaymentMethod = 'Credit')and (Credit > 0 ) "
'    If .State = 0 Then .Open
'    If .RecordCount = 0 Then Exit Sub
'    Call FormatGrid2
'    If .RecordCount = 0 Then Exit Sub
'
'    I = 1
'    rcount = 1
'
'    Do While .EOF = False
'    rcount = rcount + 1
'    grid2.Rows = rcount
'    grid2.TextMatrix(I, 0) = I
'    grid2.TextMatrix(I, 1) = dtcPatient.Text
'    grid2.TextMatrix(I, 2) = Format(!Credit, "0.00")
'    grid2.TextMatrix(I, 3) = !Date
'    Grid1.TextMatrix(I, 4) = !PatientBill_ID
'    Grid1.TextMatrix(I, 5) = !Patient_ID
'
'    I = I + 1
'    .MoveNext
'    Loop
'
'    lblNumberofbills.Caption = Val(rcount) - 1
'
'If .State = 1 Then .Close
'End With


End Sub

Private Sub Grid2_Click()
Call FindSelectedPatient2
Call FindTotalDueBalance

Grid1.Col = 0
Grid1.ColSel = Grid1.Cols - 1
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


Private Sub txtAmount_Change()
lblCashBalance.Caption = Val(lblDuebalance) - Val(txtAmount)
If Val(lblDuebalance.Caption) <> Val(txtAmount) Then
bttnUpdate.Enabled = False
Else
bttnUpdate.Enabled = True

End If
End Sub

Private Sub txtCashPayment_Change()
lblCashBalance.Caption = Val(lblDuebalance) - Val(txtCashPayment)

If Val(lblDuebalance.Caption) <> Val(txtCashPayment) Then
bttnUpdate.Enabled = False
Else
bttnUpdate.Enabled = True
End If

End Sub


Private Sub txtChequeAmount_Change()
lblCashBalance.Caption = Val(lblDuebalance) - Val(txtChequeAmount)

If Val(lblDuebalance.Caption) <> Val(txtChequeAmount) Then
bttnUpdate.Enabled = False
Else
bttnUpdate.Enabled = True
End If
End Sub
