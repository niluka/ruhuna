VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form FrmInstitutions1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Institution"
   ClientHeight    =   8625
   ClientLeft      =   2385
   ClientTop       =   2085
   ClientWidth     =   10680
   Icon            =   "FrmInstitutions1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   8625
   ScaleWidth      =   10680
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
   Begin VB.Frame framInstitution 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   4680
      TabIndex        =   20
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton OptionCreditAgent 
         Caption         =   "Credit Agent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   36
         Top             =   6840
         Width           =   2295
      End
      Begin VB.OptionButton OptionCashAgent 
         Caption         =   "Cash Agent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   35
         Top             =   6600
         Width           =   2295
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   33
         Top             =   720
         Width           =   3495
      End
      Begin VB.CheckBox chkBlackListed 
         Caption         =   "Black Listed"
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
         Left            =   2040
         TabIndex        =   15
         Top             =   6120
         Width           =   3495
      End
      Begin VB.TextBox txtMaxCredit 
         Height          =   360
         Left            =   2040
         TabIndex        =   13
         Top             =   5280
         Width           =   3495
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtTel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   6
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtFax 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtComment 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   9
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtAccount 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   4800
         Width           =   3495
      End
      Begin VB.TextBox txtCredit 
         Height          =   360
         Left            =   2040
         TabIndex        =   14
         Top             =   5760
         Width           =   3495
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   8
         Top             =   2760
         Width           =   3495
      End
      Begin VB.CheckBox CheckAgent 
         Caption         =   "Is an agent"
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
         Left            =   2040
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DataComboBank 
         Bindings        =   "FrmInstitutions1.frx":0442
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   4440
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "BankName"
         BoundColumn     =   "Bank_ID"
         Text            =   ""
         Object.DataMember      =   "sqlBank"
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
      Begin MSDataListLib.DataCombo DataComboPaymenyMethod 
         Bindings        =   "FrmInstitutions1.frx":0461
         Height          =   360
         Left            =   2040
         TabIndex        =   10
         Top             =   3960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "PaymentMethod"
         BoundColumn     =   "PaymentMethod_ID"
         Text            =   ""
         Object.DataMember      =   "sqlPaymentMethod"
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
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution Code"
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
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Limit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution Name"
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
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution Address"
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
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution Tel:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution Fax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution Comments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Method"
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
         Left            =   120
         TabIndex        =   25
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution  Bank"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution Account"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Credit / Debit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "E - Mail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1575
      End
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   495
      Left            =   9000
      TabIndex        =   18
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Ca&ncel"
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9000
      TabIndex        =   19
      Top             =   8040
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
   Begin btButtonEx.ButtonEx bttnChange 
      Height          =   495
      Left            =   5520
      TabIndex        =   16
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&hange"
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
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   495
      Left            =   5520
      TabIndex        =   17
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Sa&ve"
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
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Edit"
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
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   6135
      Left            =   120
      TabIndex        =   31
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   10821
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
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
End
Attribute VB_Name = "FrmInstitutions1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemInstitutionID As Long
Dim FromGrid As Boolean
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

bttnAdd.BackColor = BttnBackColour
bttnAdd.ForeColor = BttnForeColour

bttnCancel.BackColor = BttnBackColour
bttnCancel.ForeColor = BttnForeColour

bttnChange.BackColor = BttnBackColour
bttnChange.ForeColor = BttnForeColour

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour

bttnEdit.BackColor = BttnBackColour
bttnEdit.ForeColor = BttnForeColour

bttnSave.BackColor = BttnBackColour
bttnSave.ForeColor = BttnForeColour

'frmStaff.BackColor = FrmBackColour
'frmStaff.ForeColor = FrmForeColour


CheckAgent.BackColor = LblBackColour
CheckAgent.ForeColor = LblForeColour

chkBlackListed.BackColor = LblBackColour
chkBlackListed.ForeColor = LblForeColour


DataComboBank.BackColor = TxtBackColour
DataComboBank.ForeColor = TxtForeColour

DataComboPaymenyMethod.BackColor = TxtBackColour
DataComboPaymenyMethod.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboSex.ForeColor = TxtForeColour

'DataComboSpeciality.BackColor = TxtBackColour
'DataComboSpeciality.ForeColor = TxtForeColour

'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
DataComboBank.BackColor = TxtBackColour
DataComboBank.ForeColor = TxtForeColour
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
''Label21.BackColor = LblBackColour
''Label21.ForeColor = LblForeColour
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
'lblOfficialEmail.BackColor = LblBackColour
'lblOfficialEmail.ForeColor = LblForeColour
'
'lblOfficialWebsite.BackColor = LblBackColour
'lblOfficialWebsite.ForeColor = LblForeColour


txtAccount.BackColor = TxtBackColour
txtAccount.ForeColor = TxtForeColour

'txtBankBranch.BackColor = TxtBackColour
'txtBankBranch.ForeColor = TxtForeColour

txtComment.BackColor = TxtBackColour
txtComment.ForeColor = TxtForeColour
txtCredit.BackColor = TxtBackColour
txtCredit.ForeColor = TxtForeColour
txtMaxCredit.BackColor = TxtBackColour
txtMaxCredit.ForeColor = TxtForeColour
'txtListedName.BackColor = TxtBackColour
'txtListedName.ForeColor = TxtForeColour
txtName.BackColor = TxtBackColour
txtName.ForeColor = TxtForeColour
txtAddress.BackColor = TxtBackColour
txtAddress.ForeColor = TxtForeColour
txtEmail.BackColor = TxtBackColour
txtEmail.ForeColor = TxtForeColour
txtFax.BackColor = TxtBackColour
txtFax.ForeColor = TxtForeColour
txtTel.BackColor = TxtBackColour
txtTel.ForeColor = TxtForeColour
'txtOfficialWebsite.BackColor = TxtBackColour
'txtOfficialWebsite.ForeColor = TxtForeColour

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

'txtUserName.BackColor = TxtBackColour
'txtUserName.ForeColor = TxtForeColour
'txtPassword.BackColor = TxtBackColour
'txtPassword.ForeColor = TxtForeColour
'txtReenterPassword.BackColor = TxtBackColour
'txtReenterPassword.ForeColor = TxtForeColour
'txtOfficialFax.BackColor = TxtBackColour
'txtOfficialFax.ForeColor = TxtForeColour
'txtOfficialTel.BackColor = TxtBackColour
'txtOfficialTel.ForeColor = TxtForeColour
'txtOfficialWebsite.BackColor = TxtBackColour
'txtOfficialWebsite.ForeColor = TxtForeColour

framInstitution.BackColor = FrmBackColour
framInstitution.ForeColor = FrmForeColour

FrmInstitutions1.BackColor = FrmBackColour
FrmInstitutions1.ForeColor = FrmForeColour


'txtQualifications.BackColor = TxtBackColour
'txtQualifications.ForeColor = TxtForeColour
'txtRegistation.BackColor = TxtBackColour
'txtRegistation.ForeColor = TxtForeColour
txtSearch.BackColor = TxtBackColour
txtSearch.ForeColor = TxtForeColour
End Sub

Private Sub bttnAdd_Click()
    Call AfterAdd
    Call ClearValues
End Sub

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub bttnChange_Click()
    Dim TemResponce  As Integer
    If Trim(txtName.Text) = "" Then
        TemResponce = MsgBox("Please enter the name of the institution", vbCritical + vbOKOnly, "No Name")
        txtName.SetFocus
        Exit Sub
    End If
    Call EditData
    Call ClearValues
    Call FormatGrid
    Call FillGrid
    Call BeforeAddEdit
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    Call AfterEdit
End Sub

Private Sub bttnSave_Click()
    Dim TemResponce  As Integer
    If Trim(txtName.Text) = "" Then
        TemResponce = MsgBox("Please enter the name of the institution", vbCritical + vbOKOnly, "No Name")
        txtName.SetFocus
        Exit Sub
    End If
    Call SaveData
    Call FormatGrid
    Call FillGrid
    Call ClearValues
    Call AfterAdd
End Sub

Private Sub Form_Load()
    If AgentCashOnly = True Then
        OptionCashAgent.Visible = False
        OptionCreditAgent.Visible = False
    Else
        OptionCashAgent.Visible = True
        OptionCreditAgent.Visible = True
    End If
    Call FormatGrid
    Call FillGrid
    Call BeforeAddEdit
    Call ClearValues
    Call Setcolours
End Sub

Private Sub BeforeAddEdit()
    bttnEdit.Enabled = True
    bttnAdd.Enabled = True
    
    bttnSave.Visible = False
    bttnChange.Visible = False
    bttnCancel.Visible = False
    
    framInstitution.Enabled = False
    Grid1.Enabled = True
End Sub

Private Sub AfterAdd()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
    
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    
    framInstitution.Enabled = True
    Grid1.Enabled = False
End Sub
Private Sub AfterEdit()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
    
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    
    framInstitution.Enabled = True
    Grid1.Enabled = True
End Sub

Private Sub SaveData()
    'On Error GoTo ErrorHandler
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblInstitutions.* FROM tblInstitutions ORDER BY InstitutionName"
        If .State = 0 Then .Open
        .AddNew
        !InstitutionName = Trim(txtName.Text)
        !InstitutionAddress = Trim(txtAddress.Text)
        !InstitutionTelephone = Trim(txtTel.Text)
        !InstitutionFax = Trim(txtFax.Text)
        !InstitutionEmail = Trim(txtEmail.Text)
        !InstitutionComments = Trim(txtComment.Text)
        If IsNumeric(DataComboPaymenyMethod.BoundText) Then !InstitutionPaymentMethod_ID = DataComboPaymenyMethod.BoundText
        If IsNumeric(DataComboBank.BoundText) Then !InstitutionBank_ID = DataComboBank.BoundText
        !InstitutionAccount = Trim(txtAccount.Text)
        If CheckAgent.Value = 1 Then
            !InstitutionIsAnAgent = 1
        Else
            !InstitutionIsAnAgent = 0
        End If
        If chkBlackListed.Value = 1 Then
            !InstitutionBlackListed = 1
        Else
            !InstitutionBlackListed = 0
        End If
        !InstitutionMaxCredit = Val(txtMaxCredit.Text)
        !InstitutionCode = txtCode.Text
        If AgentCashOnly = False Then
            If OptionCashAgent.Value = True Then
                !Cashagent = 1
            Else
                !Cashagent = 0
            End If
        End If
        .Update
        .Close
    Exit Sub
    
    
ErrorHandler:
     MsgBox Err.Description
    .CancelUpdate
    End With
    
End Sub


Private Sub EditData()
    'On Error GoTo ErrorHandler
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblInstitutions.* FROM tblInstitutions where Institution_ID = " & TemInstitutionID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        !InstitutionName = Trim(txtName.Text)
        !InstitutionAddress = Trim(txtAddress.Text)
        !InstitutionTelephone = Trim(txtTel.Text)
        !InstitutionFax = Trim(txtFax.Text)
        !InstitutionEmail = Trim(txtEmail.Text)
        !InstitutionComments = Trim(txtComment.Text)
        If IsNumeric(DataComboPaymenyMethod.BoundText) Then !InstitutionPaymentMethod_ID = DataComboPaymenyMethod.BoundText
        If IsNumeric(DataComboBank.BoundText) Then !InstitutionBank_ID = DataComboBank.BoundText
        !InstitutionAccount = Trim(txtAccount.Text)
        If CheckAgent.Value = 1 Then
            !InstitutionIsAnAgent = 1
        Else
            !InstitutionIsAnAgent = 0
        End If
        If chkBlackListed.Value = 1 Then
            !InstitutionBlackListed = 1
        Else
            !InstitutionBlackListed = 0
        End If
        !InstitutionMaxCredit = Val(txtMaxCredit.Text)
        !InstitutionCode = txtCode.Text
        If AgentCashOnly = False Then
            If OptionCashAgent.Value = True Then
                !Cashagent = 1
            Else
                !Cashagent = 0
            End If
        End If
        .Update
        .Close
    Exit Sub
ErrorHandler:
     MsgBox Err.Description
    .CancelUpdate
    End With
    
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    txtAddress.Text = Empty
    txtTel.Text = Empty
    txtFax.Text = Empty
    DataComboPaymenyMethod.Text = Empty
    DataComboBank.Text = Empty
    txtAccount.Text = Empty
    txtComment.Text = Empty
    txtAccount.Text = Empty
    txtCredit.Text = Empty
    CheckAgent.Value = 0
    txtMaxCredit.Text = Empty
    chkBlackListed.Value = 0
    txtCode.Text = Empty
    OptionCashAgent.Value = True
End Sub


Private Sub GetData()
    Call ClearValues
    If Grid1.Row < 1 Then Exit Sub
    Grid1.Col = 2
    If IsNumeric(Grid1.Text) = False Then Exit Sub
    TemInstitutionID = Val(Grid1.Text)
    With DataEnvironment1.rssqlTem7
        If .State = 1 Then .Close
        .Source = "SELECT tblInstitutions.* FROM tblInstitutions where Institution_ID = " & TemInstitutionID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!InstitutionName) Then
            txtName.Text = !InstitutionName
        End If
        If Not IsNull(!InstitutionAddress) Then
            txtAddress.Text = !InstitutionAddress
        End If
        If Not IsNull(!InstitutionTelephone) Then
            txtTel.Text = !InstitutionTelephone
        End If
        If Not IsNull(!InstitutionFax) Then
            txtFax.Text = !InstitutionFax
        End If
        If Not IsNull(!InstitutionEmail) Then
            txtEmail.Text = !InstitutionEmail
        End If
        If Not IsNull(!InstitutionComments) Then
            txtComment.Text = !InstitutionComments
        End If
        If Not IsNull(!InstitutionPaymentMethod_ID) Then
            DataComboPaymenyMethod.BoundText = !InstitutionPaymentMethod_ID
        End If
        If Not IsNull(!InstitutionBank_ID) Then
            DataComboBank.BoundText = !InstitutionBank_ID
        End If
        If Not IsNull(!InstitutionAccount) Then
            txtAccount.Text = !InstitutionAccount
        End If
        If Not IsNull(!InstitutionCredit) Then
            txtCredit.Text = !InstitutionCredit
        End If
        If !InstitutionIsAnAgent = True Then CheckAgent.Value = 1
        If !InstitutionBlackListed = True Then chkBlackListed.Value = 1
        If Not IsNull(!InstitutionMaxCredit) Then txtMaxCredit.Text = !InstitutionMaxCredit
        If Not IsNull(!InstitutionCode) Then
            txtCode.Text = !InstitutionCode
        End If
        If AgentCashOnly = False Then
            If !Cashagent = 1 Then
                OptionCashAgent.Value = True
            Else
                OptionCreditAgent.Value = True
            End If
        End If
        .Close
End With

End Sub

Private Sub FormatGrid()
    Dim BorderMargin As Long
    BorderMargin = 100
    With Grid1
        .Clear
        .Cols = 3
        .Rows = 1
        .ColWidth(0) = 600
        .ColWidth(2) = 1
        .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + BorderMargin)
        .Row = 0
        .Col = 0
        .Text = "NO"
        .CellAlignment = 6
        .Col = 1
        .Text = "Institution Name"
        .Col = 2
        .Text = "ID"
        .CellAlignment = 6
    End With
End Sub


Private Sub FillGrid()
    Dim NowROw As Long
    With DataEnvironment1.rssqlTem6
    If .State = 1 Then .Close
    .Source = "SELECT tblInstitutions.* FROM tblInstitutions ORDER BY InstitutionName"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
        Do While .EOF = False
            If Not IsNull(!InstitutionName) Then
            NowROw = NowROw + 1
            Grid1.Rows = NowROw + 1
            Grid1.Row = NowROw
            Grid1.Col = 0
            Grid1.CellAlignment = 7
            Grid1.Text = NowROw
            Grid1.Col = 1
            Grid1.CellAlignment = 1
            Grid1.Text = !InstitutionName
            Grid1.Col = 2
            Grid1.CellAlignment = 7
            Grid1.Text = !Institution_Id
            End If
        .MoveNext
        Loop
    End With
End Sub



Private Sub Grid1_Click()
    If Grid1.Rows < 1 Then Exit Sub
    Grid1.Col = 2
    If Not IsNumeric(Grid1.Text) Then Exit Sub
    Call GetData
    Grid1.Col = 0
    Grid1.ColSel = Grid1.Cols - 1
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then Grid1_Click
End Sub

Private Sub txtSearch_Change()
    
' **************************************

    If FromGrid = True Then Exit Sub
    Dim TemFRows As Long
    Dim TemNowRow As Long
    Dim TemArray As Long
    Dim SearchSuccess As Boolean
    Dim TemLength As Single
    TemFRows = Grid1.Rows
    Grid1.Col = 1
    SearchSuccess = False
    If Len(txtSearch.Text) = 0 Then GoTo MeasureSuccess
    For TemArray = 1 To (TemFRows - 1)
        Grid1.Row = TemArray
        If Len(txtSearch.Text) > Len(Grid1.Text) Then
            GoTo FinishLoop
        Else
            TemLength = Len(txtSearch.Text)
        End If
        If UCase(Left((Grid1.Text), TemLength)) = UCase(txtSearch.Text) Then
            SearchSuccess = True
            Exit For
        Else
            SearchSuccess = False
        End If
FinishLoop:
    Next
    
MeasureSuccess:
    
    If SearchSuccess = True Then
        Grid1.TopRow = TemArray
        Grid1.Row = TemArray
        Grid1.Col = 0
        Grid1.ColSel = (Grid1.Cols - 1)
        bttnEdit.Enabled = True
        bttnAdd.Enabled = False
        Grid1.Col = 2
        TemInstitutionID = Grid1.Text
        Call GetData
        Grid1.Col = 0
        Grid1.ColSel = Grid1.Cols - 1
    Else
        Grid1.TopRow = 1
        Grid1.Row = 0
        Grid1.Col = 0
        Grid1.ColSel = 0
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
    End If
'**************************************
End Sub


