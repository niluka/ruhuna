VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCatogerisedIncome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catogerioused Income"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatogerisedIncome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   7095
   Begin VB.Frame frameIncome 
      Caption         =   "Income"
      Height          =   2055
      Left            =   240
      TabIndex        =   31
      Top             =   4440
      Width           =   6495
      Begin VB.Label lbl1 
         Caption         =   "Personal Income"
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
         TabIndex        =   43
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblSelectedPeriodPersonalIncome 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3600
         TabIndex        =   42
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblSelectedPeriodInstitutionIncome 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3600
         TabIndex        =   41
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Institution Income"
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
         TabIndex        =   40
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblSelectedPeriodOtherIncome 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3600
         TabIndex        =   39
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label12 
         Caption         =   "Other Income"
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
         TabIndex        =   38
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblSelectedPeriodTotalIncome 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3600
         TabIndex        =   37
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label14 
         Caption         =   "Total Income"
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
         TabIndex        =   36
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Rs."
         Height          =   255
         Left            =   2760
         TabIndex        =   35
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Rs."
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Rs."
         Height          =   255
         Left            =   2760
         TabIndex        =   33
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Rs."
         Height          =   255
         Left            =   2760
         TabIndex        =   32
         Top             =   1560
         Width           =   495
      End
   End
   Begin VB.Frame FramePaymentMethod 
      Caption         =   "Payment Method"
      Height          =   1935
      Left            =   2760
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
      Begin VB.OptionButton OptionAll 
         Caption         =   "&All"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton OptionCash 
         Caption         =   "&Cash"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptionCredit 
         Caption         =   "C&redit"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptionCheque 
         Caption         =   "C&heque"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton OptionCreditCard 
         Caption         =   "Cr&edit Card"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton OptionAgent 
         Caption         =   "A&gent"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.Frame FrameSecession 
      Caption         =   "Secession"
      Height          =   1935
      Left            =   240
      TabIndex        =   21
      Top             =   1080
      Width           =   2415
      Begin VB.OptionButton OptionNotRelevent 
         Caption         =   "&Not Relevent"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton OptionBoth 
         Caption         =   "&Both"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton OptionEvening 
         Caption         =   "&Evening Secession"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton OptionMorning 
         Caption         =   "&Morning Secession"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   5640
      TabIndex        =   19
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Close"
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
      Height          =   3495
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Today"
      TabPicture(0)   =   "frmCatogerisedIncome.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DTPickerToday"
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(3)=   "Label1"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Selected Day"
      TabPicture(1)   =   "frmCatogerisedIncome.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPickerSelectedDay"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label3"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Selected Period"
      TabPicture(2)   =   "frmCatogerisedIncome.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DTPickerFrom"
      Tab(2).Control(1)=   "DTPickerTo"
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(3)=   "Label5"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Graphs"
      TabPicture(3)   =   "frmCatogerisedIncome.frx":0496
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "frameGraphs"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame frameGraphs 
         Caption         =   "Graphs"
         Height          =   3015
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   6495
         Begin VB.OptionButton OptionPieChart 
            Caption         =   "Pie Chart"
            Height          =   255
            Left            =   4320
            TabIndex        =   52
            Top             =   1440
            Width           =   1215
         End
         Begin VB.OptionButton OptionLineChart 
            Caption         =   "Line Chart"
            Height          =   255
            Left            =   4320
            TabIndex        =   51
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton OptionBarChart 
            Caption         =   "Bar Char&t"
            Height          =   375
            Left            =   4320
            TabIndex        =   17
            Top             =   480
            Width           =   1935
         End
         Begin btButtonEx.ButtonEx bttnThisYear 
            Height          =   375
            Left            =   840
            TabIndex        =   47
            Top             =   1440
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "This Year"
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
         Begin btButtonEx.ButtonEx bttnThisMonth 
            Height          =   375
            Left            =   840
            TabIndex        =   16
            Top             =   480
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "Thi&s Month"
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
         Begin btButtonEx.ButtonEx bttnLastMonth 
            Height          =   375
            Left            =   840
            TabIndex        =   48
            Top             =   960
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "Last Month"
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
         Begin btButtonEx.ButtonEx BttnLastYear 
            Height          =   375
            Left            =   840
            TabIndex        =   49
            Top             =   1920
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "Last Year"
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
         Begin btButtonEx.ButtonEx bttnLastFiveYears 
            Height          =   375
            Left            =   840
            TabIndex        =   50
            Top             =   2400
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "Last Five Years"
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
      Begin MSComCtl2.DTPicker DTPickerToday 
         Height          =   375
         Left            =   -72120
         TabIndex        =   12
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55050243
         CurrentDate     =   39421
      End
      Begin MSComCtl2.DTPicker DTPickerSelectedDay 
         Height          =   375
         Left            =   -72120
         TabIndex        =   13
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55050243
         CurrentDate     =   39421
      End
      Begin MSComCtl2.DTPicker DTPickerFrom 
         Height          =   375
         Left            =   -74160
         TabIndex        =   14
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55050243
         CurrentDate     =   39421
      End
      Begin MSComCtl2.DTPicker DTPickerTo 
         Height          =   375
         Left            =   -71640
         TabIndex        =   15
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55050243
         CurrentDate     =   39421
      End
      Begin VB.Label Label15 
         Caption         =   "Today"
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
         Left            =   -73200
         TabIndex        =   45
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Date"
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
         Left            =   -73200
         TabIndex        =   44
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "To"
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
         Left            =   -72240
         TabIndex        =   30
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "From"
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
         Left            =   -74760
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Income"
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
         Left            =   -71520
         TabIndex        =   28
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Income"
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
         Left            =   -73800
         TabIndex        =   27
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Income"
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
         Left            =   -71520
         TabIndex        =   26
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Income"
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
         Left            =   -73680
         TabIndex        =   25
         Top             =   2040
         Width           =   1215
      End
   End
   Begin MSDataListLib.DataCombo DataComboFacility 
      Bindings        =   "frmCatogerisedIncome.frx":04B2
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "HospitalFacility"
      BoundColumn     =   "HospitalFacility_ID"
      Text            =   ""
      Object.DataMember      =   "sqlHospitalFacility"
   End
   Begin MSDataListLib.DataCombo DataComboDoctorStaff 
      Bindings        =   "frmCatogerisedIncome.frx":04D1
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   635
      _Version        =   393216
      ListField       =   "DoctorListedName"
      BoundColumn     =   "Doctor_ID"
      Text            =   ""
      Object.DataMember      =   "sqlDoctor"
   End
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   4320
      TabIndex        =   18
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Print"
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
   Begin VB.Label lblFacility 
      BackStyle       =   0  'Transparent
      Caption         =   "&Facility :"
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
      Left            =   240
      TabIndex        =   23
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblDoctorStaff 
      BackStyle       =   0  'Transparent
      Caption         =   "&Doctor :"
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
      Left            =   240
      TabIndex        =   22
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmCatogerisedIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemPatientID As Long
    Dim TemCatogery  As Integer
    Dim TemDailyMaximum As Integer
    Dim TemStaffFacilityID As Long
    Dim TemHospitalFacilityID As Long
    Dim TemstaffID As Long
    Dim TemPatientFacilityID As Long
    Dim TemBillID As Long
    Dim TemSecession  As Integer
    Dim TemPersonalIncome As Double
    Dim TemInstitutionIncome As Double
    Dim TemOtherIncome As Double
    Dim TemTotalIncome As Double
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

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour
'
bttnLastFiveYears.BackColor = BttnBackColour
bttnLastFiveYears.ForeColor = BttnForeColour
'
BttnLastYear.BackColor = BttnBackColour
BttnLastYear.ForeColor = BttnForeColour

bttnLastMonth.BackColor = BttnBackColour
bttnLastMonth.ForeColor = BttnForeColour

bttnPrint.BackColor = BttnBackColour
bttnPrint.ForeColor = BttnForeColour
'
bttnThisMonth.BackColor = BttnBackColour
bttnThisMonth.ForeColor = BttnForeColour

bttnThisYear.BackColor = BttnBackColour
bttnThisYear.ForeColor = BttnForeColour
'
'bttnPrint.BackColor = BttnBackColour
'bttnPrint.ForeColor = BttnForeColour

'OptionPrintSeperately.BackColor = TxtBackColour
'OptionPrintSeperately.ForeColor = TxtForeColour
'
'bttnRemove.BackColor = BttnBackColour
'bttnRemove.ForeColor = BttnForeColour


FrameSecession.BackColor = FrmBackColour
FrameSecession.ForeColor = FrmForeColour

FramePaymentMethod.BackColor = FrameBackColour
FramePaymentMethod.ForeColor = FrameForeColour

'FrameAgent.BackColor = FrameBackColour
'FrameAgent.ForeColor = FrameForeColour

frmCatogerisedIncome.BackColor = FrameBackColour
frmCatogerisedIncome.ForeColor = FrameForeColour
'
SSTab1.BackColor = FrameBackColour
SSTab1.ForeColor = FrameForeColour
'
'Shape1.BackColor = FrameBackColour
'Shape1.ForeColor = FrameForeColour
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
''FramePayment.BackColor = FrameBackColour
''FramePayment.ForeColor = FrameForeColour
'
'
'
''chk.BackColor = LblBackColour
''chkCurrentlyChanneling.ForeColor = LblForeColour
'
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboAgent.BackColor = TxtBackColour
'DataComboAgent.ForeColor = TxtForeColour
'
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboDoctorStaff.BackColor = TxtBackColour
'DataComboDoctorStaff.ForeColor = TxtForeColour
'
'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour
'
''DataCombo.BackColor = TxtBackColour
''DataComboBank.ForeColor = TxtForeColour
''
''DataComboBank.BackColor = TxtBackColour
''DataComboBank.ForeColor = TxtForeColour
''DataComboBank.BackColor = TxtBackColour
''DataComboBank.ForeColor = TxtForeColour
'
'
'
'
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
'
''grid1.ForeColor = Grid



'Label1.BackColor = LblBackColour
'Label1.ForeColor = LblForeColour

'lblLRMP.BackColor = LblBackColour
'lblLRMP.ForeColor = LblForeColour
'lblMenstrualFlow.BackColor = LblBackColour
'lblMenstrualFlow.ForeColor = LblForeColour
'LblCommentsLX.BackColor = LblBackColour
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

'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour

'chkCough.BackColor = FrameBackColour
'chkCough.ForeColor = FrameForeColour
'
'chkPleuritic.BackColor = FrameBackColour
'chkPleuritic.ForeColor = FrameForeColour
'
'chkSputum.BackColor = FrameBackColour
'chkSputum.ForeColor = FrameForeColour
''
'chkHaemoptysis.BackColor = FrameBackColour
'chkHaemoptysis.ForeColor = FrameForeColour
''
'chkWheeze.BackColor = FrameBackColour
'chkWheeze.ForeColor = FrameForeColour
'''
'chkDyspnoea.BackColor = FrameBackColour

'chkDyspnoea.ForeColor = FrameForeColour
'
'chkParaethesia.BackColor = FrameBackColour
'chkParaethesia.ForeColor = FrameForeColour
''
'chkMuscleWeak.BackColor = FrameBackColour
'chkMuscleWeak.ForeColor = FrameForeColour
''
'chkSleep.BackColor = FrameBackColour
'chkSleep.ForeColor = FrameForeColour
''
'chkVisual.BackColor = FrameBackColour
'chkVisual.ForeColor = FrameForeColour
'''
'chkSmell.BackColor = FrameBackColour
'chkSmell.ForeColor = FrameForeColour
'
'chkTaste.BackColor = FrameBackColour
'chkTaste.ForeColor = FrameForeColour
'''
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

'chkUrgeIncontinence.BackColor = FrameBackColour
'chkUrgeIncontinence.ForeColor = FrameForeColour

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
'OptionMorning.BackColor = TxtBackColour
'OptionMorning.ForeColor = TxtForeColour
'
OptionEvening.BackColor = FrameBackColour
OptionEvening.ForeColor = TxtForeColour


OptionBoth.BackColor = FrameBackColour
OptionBoth.ForeColor = TxtForeColour

OptionNotRelevent.BackColor = FrameBackColour
OptionNotRelevent.ForeColor = TxtForeColour

OptionCash.BackColor = FrameBackColour
OptionCash.ForeColor = TxtForeColour

'
OptionCredit.BackColor = FrameBackColour
OptionCredit.ForeColor = TxtForeColour


OptionCheque.BackColor = FrameBackColour
OptionCheque.ForeColor = TxtForeColour


OptionCreditCard.BackColor = FrameBackColour
OptionCreditCard.ForeColor = TxtForeColour


OptionAgent.BackColor = FrameBackColour
OptionAgent.ForeColor = TxtForeColour


OptionAll.BackColor = FrameBackColour
OptionAll.ForeColor = TxtForeColour


OptionMorning.BackColor = FrameBackColour
OptionMorning.ForeColor = TxtForeColour
''
''OptionCredit.BackColor = TxtBackColour
''OptionCredit.ForeColor = TxtForeColour
''
''OptionCredit.BackColor = TxtBackColour
''OptionCredit.ForeColor = TxtForeColour
''
''OptionCredit.BackColor = TxtBackColour
''OptionCredit.ForeColor = TxtForeColour
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

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnGraph_Click()
    frmGraph01.Show
End Sub

Private Sub bttnLastFiveYears_Click()

Dim MyWorkBook As Excel.Workbook
Dim MyWorkSheet As Excel.Worksheet
Dim MyChart As Excel.Chart
Dim DayNumber As Integer
Dim TemDate1 As Date
Dim TemDate2 As Date
Dim TemSecession As Integer
Dim TemIncome As IncomeByDates
Dim TemNum As Long


Set MyWorkBook = GetObject(App.Path & "\graph01.xls")
Set MyWorkSheet = MyWorkBook.Worksheets.Item(1)
Set MyChart = MyWorkBook.Charts.Item(1)

TemNum = 1

For DayNumber = (Year(Date) - 5) To Year(Date)

   TemDate1 = DateSerial(DayNumber, 1, 1)
   TemDate2 = DateSerial(DayNumber, 12, 31)
    
    TemSecession = 0
    
    If OptionMorning.Value = True Then TemSecession = MorningSecession
    If OptionEvening.Value = True Then TemSecession = EveningSecession
    
    If IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, DataComboFacility.BoundText, DataComboDoctorStaff.BoundText)
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income by " & DataComboFacility.Text
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    ElseIf IsNumeric(DataComboFacility.BoundText) And Not IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, Val(DataComboFacility.BoundText), Val(DataComboDoctorStaff.BoundText))
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income by " & DataComboFacility.Text
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    
    
    
    ElseIf Not IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, DataComboFacility.BoundText, DataComboDoctorStaff.BoundText)
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income"
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    
    End If

    MyWorkSheet.Cells(1, TemNum + 1) = DayNumber
    MyWorkSheet.Cells(2, TemNum + 1) = TemIncome.PersonalIncome
    MyWorkSheet.Cells(3, TemNum + 1) = TemIncome.InstitutionIncome
    MyWorkSheet.Cells(4, TemNum + 1) = TemIncome.OtherIncome

    TemNum = TemNum + 1
    
Next


If OptionBarChart.Value = True Then MyChart.ChartType = 51
If OptionLineChart.Value = True Then MyChart.ChartType = 65
If OptionPieChart.Value = True Then MyChart.ChartType = xlPie

MyChart.HasTitle = False
MyChart.HasLegend = True
MyChart.Legend.Font.Size = 8


MyChart.SetSourceData MyWorkSheet.Range("a2:g4")

MyChart.Activate

frmGraph01.WindowState = 2
frmGraph01.Show





End Sub

Private Sub bttnLastMonth_Click()

Dim MyWorkBook As Excel.Workbook
Dim MyWorkSheet As Excel.Worksheet
Dim MyChart As Excel.Chart
Dim DayNumber As Integer
Dim TemDate1 As Date
Dim TemDate2 As Date
Dim TemSecession As Integer
Dim TemIncome As IncomeByDates

Set MyWorkBook = GetObject(App.Path & "\graph01.xls")
Set MyWorkSheet = MyWorkBook.Worksheets.Item(1)
Set MyChart = MyWorkBook.Charts.Item(1)


For DayNumber = 1 To Day(Date)

If Month(Date) <> 1 Then
    TemDate1 = DateSerial(Year(Date), (Month(Date) - 1), DayNumber)
    TemDate2 = DateSerial(Year(Date), (Month(Date) - 1), DayNumber)
Else
    TemDate1 = DateSerial(Year(Date) - 1, (12), DayNumber)
    TemDate2 = DateSerial(Year(Date) - 1, (12), DayNumber)
End If
    TemSecession = 0
    
    If OptionMorning.Value = True Then TemSecession = MorningSecession
    If OptionEvening.Value = True Then TemSecession = EveningSecession
    
    If IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, DataComboFacility.BoundText, DataComboDoctorStaff.BoundText)
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income by " & DataComboFacility.Text
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    ElseIf IsNumeric(DataComboFacility.BoundText) And Not IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, Val(DataComboFacility.BoundText), Val(DataComboDoctorStaff.BoundText))
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income by " & DataComboFacility.Text
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    
    
    
    ElseIf Not IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, DataComboFacility.BoundText, DataComboDoctorStaff.BoundText)
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income"
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    
    End If

    MyWorkSheet.Cells(1, DayNumber + 1) = DayNumber
    MyWorkSheet.Cells(2, DayNumber + 1) = TemIncome.PersonalIncome
    MyWorkSheet.Cells(3, DayNumber + 1) = TemIncome.InstitutionIncome
    MyWorkSheet.Cells(4, DayNumber + 1) = TemIncome.OtherIncome

Next

MyChart.Activate

If OptionBarChart.Value = True Then MyChart.ChartType = 51
If OptionLineChart.Value = True Then MyChart.ChartType = 65
If OptionPieChart.Value = True Then MyChart.ChartType = xlPie

MyChart.HasTitle = False
MyChart.HasLegend = True
MyChart.Legend.Font.Size = 8


MyChart.SetSourceData MyWorkSheet.Range("a2:ae4")

frmGraph01.WindowState = 2
frmGraph01.Show


End Sub

Private Sub BttnLastYear_Click()

Dim MyWorkBook As Excel.Workbook
Dim MyWorkSheet As Excel.Worksheet
Dim MyChart As Excel.Chart
Dim DayNumber As Integer
Dim TemDate1 As Date
Dim TemDate2 As Date
Dim TemSecession As Integer
Dim TemIncome As IncomeByDates

Set MyWorkBook = GetObject(App.Path & "\graph01.xls")
Set MyWorkSheet = MyWorkBook.Worksheets.Item(1)
Set MyChart = MyWorkBook.Charts.Item(1)


For DayNumber = 1 To Month(Date)

    TemDate1 = DateSerial(Year(Date) - 1, (DayNumber), 1)
    TemDate2 = DateSerial(Year(Date) - 1, (DayNumber), 31)
    
    TemSecession = 0
    
    If OptionMorning.Value = True Then TemSecession = MorningSecession
    If OptionEvening.Value = True Then TemSecession = EveningSecession
    
    If IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, DataComboFacility.BoundText, DataComboDoctorStaff.BoundText)
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income by " & DataComboFacility.Text
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    ElseIf IsNumeric(DataComboFacility.BoundText) And Not IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, Val(DataComboFacility.BoundText), Val(DataComboDoctorStaff.BoundText))
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income by " & DataComboFacility.Text
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    
    
    
    ElseIf Not IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, DataComboFacility.BoundText, DataComboDoctorStaff.BoundText)
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income"
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    
    End If

    MyWorkSheet.Cells(1, DayNumber + 1) = MonthName(DayNumber)
    MyWorkSheet.Cells(2, DayNumber + 1) = TemIncome.PersonalIncome
    MyWorkSheet.Cells(3, DayNumber + 1) = TemIncome.InstitutionIncome
    MyWorkSheet.Cells(4, DayNumber + 1) = TemIncome.OtherIncome

Next

MyChart.Activate

If OptionBarChart.Value = True Then MyChart.ChartType = 51
If OptionLineChart.Value = True Then MyChart.ChartType = 65
If OptionPieChart.Value = True Then MyChart.ChartType = xlPie

MyChart.HasTitle = False
MyChart.HasLegend = True
MyChart.Legend.Font.Size = 8


MyChart.SetSourceData MyWorkSheet.Range("a1:m4")

frmGraph01.WindowState = 2
frmGraph01.Show

End Sub

Private Sub bttnThisMonth_Click()

Dim MyWorkBook As Excel.Workbook
Dim MyWorkSheet As Excel.Worksheet
Dim MyChart As Excel.Chart
Dim DayNumber As Integer
Dim TemDate1 As Date
Dim TemDate2 As Date
Dim TemSecession As Integer
Dim TemIncome As IncomeByDates

Set MyWorkBook = GetObject(App.Path & "\graph01.xls")
Set MyWorkSheet = MyWorkBook.Worksheets.Item(1)
Set MyChart = MyWorkBook.Charts.Item(1)

For DayNumber = 1 To Day(Date)
    TemDate1 = DateSerial(Year(Date), Month(Date), DayNumber)
    TemDate2 = DateSerial(Year(Date), Month(Date), DayNumber)
    
    TemSecession = 0
    
    If OptionMorning.Value = True Then TemSecession = MorningSecession
    If OptionEvening.Value = True Then TemSecession = EveningSecession
    
    If IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, DataComboFacility.BoundText, DataComboDoctorStaff.BoundText)
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income by " & DataComboFacility.Text
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    ElseIf IsNumeric(DataComboFacility.BoundText) And Not IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, Val(DataComboFacility.BoundText), Val(DataComboDoctorStaff.BoundText))
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income by " & DataComboFacility.Text
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    
    
    
    ElseIf Not IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, DataComboFacility.BoundText, DataComboDoctorStaff.BoundText)
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income"
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    
    End If

    MyWorkSheet.Cells(1, DayNumber + 1) = DayNumber
    MyWorkSheet.Cells(2, DayNumber + 1) = TemIncome.PersonalIncome
    MyWorkSheet.Cells(3, DayNumber + 1) = TemIncome.InstitutionIncome
    MyWorkSheet.Cells(4, DayNumber + 1) = TemIncome.OtherIncome

Next

MyChart.Activate

If OptionBarChart.Value = True Then MyChart.ChartType = 51
If OptionLineChart.Value = True Then MyChart.ChartType = 65
If OptionPieChart.Value = True Then MyChart.ChartType = xlPie

MyChart.HasTitle = False
MyChart.HasLegend = True
MyChart.Legend.Font.Size = 8


MyChart.SetSourceData MyWorkSheet.Range("a2:ae4")

frmGraph01.WindowState = 2
frmGraph01.Show

End Sub

Private Function FacilityStaffIncomeByDate(FromDate As Date, ToDate As Date) As Double
    
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        If .State = 1 Then .Close
        Select Case SSTab1.Tab
            Case 0:
                DTPickerToday.Value = Date
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ")  and (staff_ID = " & TemstaffID & ") and (bookingdate = #" & Date & "#) and (fullypaid = true) "
            Case 1:
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ") and (staff_ID = " & TemstaffID & ") and (bookingdate = #" & DTPickerSelectedDay.Value & "#) and (fullypaid = true) "
            Case 2:
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ") and (staff_ID = " & TemstaffID & ") and (bookingdate between #" & DTPickerFrom.Value & "# and #" & DTPickerTo.Value & "#) and (fullypaid = true) "
        End Select
        If .State = 0 Then .Open
    
        TemTotalIncome = 0
        TemPersonalIncome = 0
        TemInstitutionIncome = 0
        TemOtherIncome = 0
        If .RecordCount <> 0 Then
        .MoveFirst
            While Not .EOF
                If OptionMorning.Value = True Then
                    If !secession = MorningSecession Then
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                    End If
                ElseIf OptionEvening.Value = True Then
                    If !secession = EveningSecession Then
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                    End If
                Else
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                End If
                .MoveNext
            Wend
        End If
        TemTotalIncome = TemPersonalIncome + TemInstitutionIncome + TemOtherIncome
    End With

lblSelectedPeriodPersonalIncome.Caption = Format(TemPersonalIncome, "#0.00")
lblSelectedPeriodInstitutionIncome.Caption = Format(TemInstitutionIncome, "#0.00")
lblSelectedPeriodOtherIncome.Caption = Format(TemOtherIncome, "#0.00")
lblSelectedPeriodTotalIncome.Caption = Format(TemTotalIncome, "#0.00")
    
    
End Function



Private Function FacilityIncomeByDate(FromDate As Date, ToDate As Date) As Double
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        Select Case SSTab1.Tab
            Case 0:
                DTPickerToday.Value = Date
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ") and (bookingdate = #" & Date & "#) and (fullypaid = true) "
            Case 1:
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ") and (bookingdate = #" & DTPickerSelectedDay.Value & "#)and (fullypaid = true) "
            Case 2:
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ") and (bookingdate between #" & DTPickerFrom.Value & "# and #" & DTPickerTo.Value & "#) and (fullypaid = true) "
            Case 3: Exit Function
        End Select
        If .State = 0 Then .Open

        TemTotalIncome = 0
        TemPersonalIncome = 0
        TemInstitutionIncome = 0
        TemOtherIncome = 0
        If .RecordCount <> 0 Then
        .MoveFirst
            While Not .EOF
                If OptionMorning.Value = True Then
                    If !secession = MorningSecession Then
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                    End If
                ElseIf OptionEvening.Value = True Then
                    If !secession = EveningSecession Then
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                    End If
                Else
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                End If
                .MoveNext
            Wend
        End If
        TemTotalIncome = TemPersonalIncome + TemInstitutionIncome + TemOtherIncome
    End With

lblSelectedPeriodPersonalIncome.Caption = Format(TemPersonalIncome, "#0.00")
lblSelectedPeriodInstitutionIncome.Caption = Format(TemInstitutionIncome, "#0.00")
lblSelectedPeriodOtherIncome.Caption = Format(TemOtherIncome, "#0.00")
lblSelectedPeriodTotalIncome.Caption = Format(TemTotalIncome, "#0.00")


End Function




Private Sub bttnThisYear_Click()

Dim MyWorkBook As Excel.Workbook
Dim MyWorkSheet As Excel.Worksheet
Dim MyChart As Excel.Chart
Dim DayNumber As Integer
Dim TemDate1 As Date
Dim TemDate2 As Date
Dim TemSecession As Integer
Dim TemIncome As IncomeByDates

Set MyWorkBook = GetObject(App.Path & "\graph01.xls")
Set MyWorkSheet = MyWorkBook.Worksheets.Item(1)
Set MyChart = MyWorkBook.Charts.Item(1)


For DayNumber = 1 To Month(Date)

    TemDate1 = DateSerial(Year(Date), (DayNumber), 1)
    TemDate2 = DateSerial(Year(Date), (DayNumber), 31)
    
    TemSecession = 0
    
    If OptionMorning.Value = True Then TemSecession = MorningSecession
    If OptionEvening.Value = True Then TemSecession = EveningSecession
    
    If IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, DataComboFacility.BoundText, DataComboDoctorStaff.BoundText)
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income by " & DataComboFacility.Text
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    ElseIf IsNumeric(DataComboFacility.BoundText) And Not IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, Val(DataComboFacility.BoundText), Val(DataComboDoctorStaff.BoundText))
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income by " & DataComboFacility.Text
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    
    
    
    ElseIf Not IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then
        TemIncome = GetIncomeByDates(TemDate1, TemDate2, TemSecession, DataComboFacility.BoundText, DataComboDoctorStaff.BoundText)
        MyWorkSheet.Cells(2, 1) = DataComboDoctorStaff.Text & "'s income"
        MyWorkSheet.Cells(3, 1) = "Institution Income"
        MyWorkSheet.Cells(4, 1) = "Other Oncome"
    
    
    End If

    MyWorkSheet.Cells(1, DayNumber + 1) = MonthName(DayNumber)
    MyWorkSheet.Cells(2, DayNumber + 1) = TemIncome.PersonalIncome
    MyWorkSheet.Cells(3, DayNumber + 1) = TemIncome.InstitutionIncome
    MyWorkSheet.Cells(4, DayNumber + 1) = TemIncome.OtherIncome

Next

MyChart.Activate

If OptionBarChart.Value = True Then MyChart.ChartType = 51
If OptionLineChart.Value = True Then MyChart.ChartType = 65
If OptionPieChart.Value = True Then MyChart.ChartType = xlPie

MyChart.HasTitle = False
MyChart.HasLegend = True
MyChart.Legend.Font.Size = 8


MyChart.SetSourceData MyWorkSheet.Range("a1:m4")

frmGraph01.WindowState = 2
frmGraph01.Show

End Sub

Private Sub DataComboDoctorStaff_Click(Area As Integer)
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    
    If IsNumeric(DataComboFacility.BoundText) = True Then
        With DataEnvironment1.rssqlTem2
            If .State = 1 Then .Close
            .Source = "SELECT tblfacilitystaff.* from tblfacilitystaff where (HospitalFacility_ID = " & DataComboFacility.BoundText & ") and (Staff_ID = " & DataComboDoctorStaff.BoundText & ")"
            If .State = 0 Then .Open
            If .RecordCount = 0 Then Exit Sub
            If Not IsNull(!FacilityStaff_ID) Then TemStaffFacilityID = !FacilityStaff_ID
            If Not IsNull(!staff_ID) Then TemstaffID = !staff_ID
            If !TwoSecessions = True Then
                PrepareForTwoSecessions
            Else
                PrepareForOneSecession
            End If
            Select Case TemCatogery
                Case Doctor:
                    LblDoctorStaff.Caption = "Doctor :"
                Case Staff:
                    LblDoctorStaff.Caption = "Staff Member :"
                Case Investigation:
                    LblDoctorStaff.Caption = "Investigation :"
                Case Other:
            End Select
            .Close
        End With
    Else
        With DataEnvironment1.rssqlDoctor
                If .State = 1 Then .Close
                .Source = "SELECT * from tbldoctor order by doctorlistedname"
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                DataComboDoctorStaff.RowMember = "sqlBookingFacility"
                DataComboDoctorStaff.ListField = "doctorlistedname"
                DataComboDoctorStaff.BoundColumn = "doctor_ID"
                .Close
        End With
    End If
    CalculateIncome
End Sub


Private Sub PrepareForOneSecession()
    OptionMorning.Enabled = False
    OptionEvening.Enabled = False
    OptionBoth.Enabled = False
    OptionNotRelevent.Enabled = True
    OptionNotRelevent.Value = True
End Sub

Private Sub PrepareForTwoSecessions()
    OptionMorning.Enabled = True
    OptionEvening.Enabled = True
    OptionBoth.Enabled = True
    OptionNotRelevent.Enabled = False
    OptionBoth.Value = True
End Sub

Private Sub DataComboFacility_Click(Area As Integer)
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    TemHospitalFacilityID = DataComboFacility.BoundText
    DataComboDoctorStaff.Text = Empty
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = "SELECT tblhospitalfacility.* from tblhospitalfacility where hospitalfacility_ID = " & DataComboFacility.BoundText
        If .State = 0 Then .Open
        TemCatogery = !PersonCatogery
        .Close
    End With
    With DataComboDoctorStaff
        .RowMember = Empty
        .ListField = Empty
        .BoundColumn = Empty
    End With
    With DataEnvironment1.rssqlBookingFacility
        If .State = 1 Then .Close
        Select Case TemCatogery
            Case Doctor:
                .Source = "SELECT tblfacilitystaff.* , tbldoctor.* FROM tblfacilitystaff left join tbldoctor on tblfacilitystaff.staff_ID = tbldoctor.doctor_ID where HospitalFacility_ID = " & DataComboFacility.BoundText & " order by doctorname"
                If .State = 0 Then .Open
                If .RecordCount = 0 Then CalculateIncome: Exit Sub
                DataComboDoctorStaff.RowMember = "sqlBookingFacility"
                DataComboDoctorStaff.ListField = "doctorname"
                DataComboDoctorStaff.BoundColumn = "doctor_ID"
            Case Staff:
                .Source = "SELECT tblfacilitystaff.* , tblstaff.* FROM tblfacilitystaff left join tblstaff on tblfacilitystaff.staff_ID = tblstaff.staff_ID where HospitalFacility_ID = " & DataComboFacility.BoundText & " order by staffname"
                If .State = 0 Then .Open
                If .RecordCount = 0 Then CalculateIncome: Exit Sub
                DataComboDoctorStaff.RowMember = "sqlBookingFacility"
                DataComboDoctorStaff.ListField = "staffname"
                DataComboDoctorStaff.BoundColumn = "tblstaff.Staff_ID"
            Case Investigation:
                .Source = "SELECT tblfacilitystaff.* , tblinvestigations.* FROM tblfacilitystaff left join tblinvestigations on tblfacilitystaff.staff_ID = tblinvestigations.investigation_ID where HospitalFacility_ID = " & DataComboFacility.BoundText & " order by investigation"
                If .State = 0 Then .Open
                If .RecordCount = 0 Then CalculateIncome: Exit Sub
                DataComboDoctorStaff.RowMember = "sqlBookingFacility"
                DataComboDoctorStaff.ListField = "investigation"
                DataComboDoctorStaff.BoundColumn = "investigation_ID"
            Case Other:
        End Select
        .Close
    End With
    Call CalculateIncome
End Sub


Private Sub CalculateIncome()
    If IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then FacilityStaffIncome
    If IsNumeric(DataComboFacility.BoundText) And Not IsNumeric(DataComboDoctorStaff.BoundText) Then FacilityIncome
    If Not IsNumeric(DataComboFacility.BoundText) And IsNumeric(DataComboDoctorStaff.BoundText) Then StaffIncome
End Sub


Private Sub FacilityIncome()
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        Select Case SSTab1.Tab
            Case 0:
                DTPickerToday.Value = Date
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ") and (bookingdate = #" & Date & "#) and (fullypaid = true) "
            Case 1:
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ") and (bookingdate = #" & DTPickerSelectedDay.Value & "#) and (fullypaid = true) "
            Case 2:
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ") and (bookingdate between #" & DTPickerFrom.Value & "# and #" & DTPickerTo.Value & "#) and (fullypaid = true) "
            Case 3: Exit Sub
        End Select
        If .State = 0 Then .Open

        TemTotalIncome = 0
        TemPersonalIncome = 0
        TemInstitutionIncome = 0
        TemOtherIncome = 0
        If .RecordCount <> 0 Then
        .MoveFirst
            While Not .EOF
                If OptionMorning.Value = True Then
                    If !secession = MorningSecession Then
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                    End If
                ElseIf OptionEvening.Value = True Then
                    If !secession = EveningSecession Then
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                    End If
                Else
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                End If
                .MoveNext
            Wend
        End If
        TemTotalIncome = TemPersonalIncome + TemInstitutionIncome + TemOtherIncome
    End With

lblSelectedPeriodPersonalIncome.Caption = Format(TemPersonalIncome, "#0.00")
lblSelectedPeriodInstitutionIncome.Caption = Format(TemInstitutionIncome, "#0.00")
lblSelectedPeriodOtherIncome.Caption = Format(TemOtherIncome, "#0.00")
lblSelectedPeriodTotalIncome.Caption = Format(TemTotalIncome, "#0.00")


End Sub

Private Sub FacilityStaffIncome()
    
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        If .State = 1 Then .Close
        
        Select Case SSTab1.Tab
            Case 0:
                DTPickerToday.Value = Date
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ")  and (staff_ID = " & TemstaffID & ") and (bookingdate = #" & Date & "#) and (fullypaid = true) "
            Case 1:
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ") and (staff_ID = " & TemstaffID & ") and (bookingdate = #" & DTPickerSelectedDay.Value & "#) and (fullypaid = true) "
            Case 2:
                .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & TemHospitalFacilityID & ") and (staff_ID = " & TemstaffID & ") and (bookingdate between #" & DTPickerFrom.Value & "# and #" & DTPickerTo.Value & "#)and (fullypaid = true) "
            Case Else
                Exit Sub
        End Select
        
        If .State = 0 Then .Open
    
        TemTotalIncome = 0
        TemPersonalIncome = 0
        TemInstitutionIncome = 0
        TemOtherIncome = 0
        If .RecordCount <> 0 Then
        .MoveFirst
            While Not .EOF
                If OptionMorning.Value = True Then
                    If !secession = MorningSecession Then
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                    End If
                ElseIf OptionEvening.Value = True Then
                    If !secession = EveningSecession Then
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                    End If
                Else
                    If Not IsNull(!Personalfee) Then
                        TemPersonalIncome = TemPersonalIncome + (!Personalfee)
                    End If
                    If Not IsNull(!personalrefund) Then
                        TemPersonalIncome = TemPersonalIncome - Val(!personalrefund)
                    End If
                    
                    If Not IsNull(!institutionfee) Then
                        TemInstitutionIncome = TemInstitutionIncome + (!institutionfee)
                    End If
                    If Not IsNull(!InstitutionRefund) Then
                        TemInstitutionIncome = TemInstitutionIncome - (!InstitutionRefund)
                    End If
                    If Not IsNull(!otherfee) Then
                        TemOtherIncome = TemOtherIncome + (!otherfee)
                    End If
                    If Not IsNull(!OtherRefund) Then
                        TemOtherIncome = TemOtherIncome - (!OtherRefund)
                    End If
                
                End If
                .MoveNext
            Wend
        End If
        TemTotalIncome = TemPersonalIncome + TemInstitutionIncome + TemOtherIncome
    End With

lblSelectedPeriodPersonalIncome.Caption = Format(TemPersonalIncome, "#0.00")
lblSelectedPeriodInstitutionIncome.Caption = Format(TemInstitutionIncome, "#0.00")
lblSelectedPeriodOtherIncome.Caption = Format(TemOtherIncome, "#0.00")
lblSelectedPeriodTotalIncome.Caption = Format(TemTotalIncome, "#0.00")
    
    
End Sub

Private Sub StaffIncome()

End Sub

Private Sub DTPickerFrom_Change()
Call CalculateIncome
End Sub

Private Sub DTPickerSelectedDay_Change()
Call CalculateIncome
End Sub


Private Sub DTPickerTo_Change()
Call CalculateIncome
End Sub


Private Sub DTPickerToday_Change()
Call CalculateIncome
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
Call SetColour
frameGraphs.Visible = False
frameIncome.Visible = True
End Sub

Private Sub OptionAgent_Click()
Call CalculateIncome
End Sub

Private Sub OptionAll_Click()
Call CalculateIncome
End Sub

Private Sub OptionBoth_Click()
Call CalculateIncome
End Sub

Private Sub OptionCash_Click()
Call CalculateIncome
End Sub

Private Sub OptionCheque_Click()
Call CalculateIncome
End Sub

Private Sub OptionCredit_Click()
Call CalculateIncome
End Sub

Private Sub OptionCreditCard_Click()
Call CalculateIncome
End Sub

Private Sub OptionEvening_Click()
Call CalculateIncome
End Sub

Private Sub OptionMorning_Click()
Call CalculateIncome
End Sub

Private Sub OptionNotRelevent_Click()
Call CalculateIncome
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Tab = 3 Then
    frameIncome.Visible = False
    frameGraphs.Visible = True
Else
    frameIncome.Visible = True
    frameGraphs.Visible = False
End If

Call CalculateIncome

End Sub

