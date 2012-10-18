VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPatientFacilityListManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Lists"
   ClientHeight    =   8010
   ClientLeft      =   375
   ClientTop       =   1755
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatientFacilityManage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin btButtonEx.ButtonEx bttnCloseList 
      Height          =   375
      Left            =   7920
      TabIndex        =   23
      Top             =   7440
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
   Begin VB.Frame frameRepay 
      Caption         =   "Repayment"
      Height          =   4575
      Left            =   600
      TabIndex        =   32
      Top             =   2520
      Width           =   6375
      Begin VB.OptionButton OptionNo 
         Caption         =   "No Prints"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3240
         Width           =   1575
      End
      Begin VB.OptionButton OptionTwo 
         Caption         =   "Two Prints"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   3720
         Width           =   1575
      End
      Begin VB.OptionButton OptionOne 
         Caption         =   "One Print"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   3480
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtStaffRepay 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtInstitutionRepay 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtOtherRepay 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtRepayTotal 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtRepayComments 
         Height          =   1095
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   2880
         Width           =   4215
      End
      Begin btButtonEx.ButtonEx bttnConfirmRepay 
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "R&epay"
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
      Begin btButtonEx.ButtonEx bttnCancelRepay 
         Height          =   375
         Left            =   3480
         TabIndex        =   15
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
      Begin VB.Label Label1 
         Caption         =   "Doctor Fee :"
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
         TabIndex        =   48
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Institution Fee:"
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
         TabIndex        =   47
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Other Fee:"
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
         TabIndex        =   46
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Paid Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   45
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Re-Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblStaffFeePaid 
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
         Left            =   2040
         TabIndex        =   43
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblInstitutionFeePaid 
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
         Left            =   2040
         TabIndex        =   42
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblOtherFeePaid 
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
         Left            =   2040
         TabIndex        =   41
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblTotalPaid 
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
         Left            =   2040
         TabIndex        =   40
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Total"
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
         TabIndex        =   39
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Comments"
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
         TabIndex        =   38
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblPreviousStaffRepay 
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
         Left            =   3480
         TabIndex        =   37
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblPreviousInstitutionRepay 
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
         Left            =   3480
         TabIndex        =   36
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblPreviousOtherRepay 
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
         Left            =   3480
         TabIndex        =   35
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblPreviousTotalRepay 
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
         Left            =   3480
         TabIndex        =   34
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Previous Repays"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FramePatientList 
      Height          =   7215
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   9375
      Begin btButtonEx.ButtonEx bttnStillToComplete 
         Height          =   375
         Left            =   4920
         TabIndex        =   20
         Top             =   5640
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print Still &to complete"
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
      Begin btButtonEx.ButtonEx bttnRepay 
         Height          =   375
         Left            =   7080
         TabIndex        =   17
         Top             =   3120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Refund"
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
      Begin VB.Frame FrameSecession 
         Caption         =   "Secession"
         Height          =   615
         Left            =   480
         TabIndex        =   28
         Top             =   1680
         Width           =   8775
         Begin VB.OptionButton OptionMorning 
            Caption         =   "&Morning Secession"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton OptionEvening 
            Caption         =   "&Evening Secession"
            Height          =   255
            Left            =   3000
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton OptionNoPreferance 
            Caption         =   "No Preferance"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.OptionButton OptionNotRelevent 
            Caption         =   "&Not Relevent"
            Height          =   255
            Left            =   6360
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridList1 
         Height          =   4575
         Left            =   480
         TabIndex        =   25
         Top             =   2400
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   8070
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo DataComboFacility 
         Bindings        =   "frmPatientFacilityManage.frx":0442
         Height          =   360
         Left            =   4440
         TabIndex        =   0
         Top             =   240
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
         Bindings        =   "frmPatientFacilityManage.frx":0461
         Height          =   360
         Left            =   4440
         TabIndex        =   1
         Top             =   720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "DoctorListedName"
         BoundColumn     =   "Doctor_ID"
         Text            =   ""
         Object.DataMember      =   "sqlDoctor"
      End
      Begin MSComCtl2.DTPicker DTPickerAppointment 
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58720259
         CurrentDate     =   39421
      End
      Begin btButtonEx.ButtonEx bttnMarkAsCompleted 
         Height          =   375
         Left            =   7080
         TabIndex        =   6
         Top             =   2640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Cancel Booking"
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
      Begin btButtonEx.ButtonEx bttnRemove 
         Height          =   375
         Left            =   7080
         TabIndex        =   31
         Top             =   3120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Remove"
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
      Begin btButtonEx.ButtonEx bttnPrintAll 
         Height          =   375
         Left            =   7080
         TabIndex        =   18
         Top             =   4560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Print All"
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
      Begin btButtonEx.ButtonEx bttnPrintFullyPaid 
         Height          =   375
         Left            =   7080
         TabIndex        =   19
         Top             =   5040
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print Fully Pa&id"
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
      Begin btButtonEx.ButtonEx bttnPayDoctor 
         Height          =   375
         Left            =   7080
         TabIndex        =   22
         Top             =   5640
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Pa&y Doctor"
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
      Begin btButtonEx.ButtonEx bttnPrintCompleted 
         Height          =   375
         Left            =   4920
         TabIndex        =   21
         Top             =   5160
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print &Completed"
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
      Begin btButtonEx.ButtonEx bttnTelephonePay 
         Height          =   375
         Left            =   7080
         TabIndex        =   49
         Top             =   3720
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Pay For Telephone"
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
         BackStyle       =   0  'Transparent
         Caption         =   "D&ate :"
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
         TabIndex        =   30
         Top             =   1200
         Width           =   5175
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
         Left            =   360
         TabIndex        =   27
         Top             =   720
         Width           =   5175
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
         Left            =   360
         TabIndex        =   26
         Top             =   240
         Width           =   5175
      End
   End
End
Attribute VB_Name = "frmPatientFacilityListManage"
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
    Dim IsACancellation As Boolean
    Dim IsARefund As Boolean
    
    
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

bttnCancelRepay.BackColor = BttnBackColour
bttnCancelRepay.ForeColor = BttnForeColour

bttnCloseList.BackColor = BttnBackColour
bttnCloseList.ForeColor = BttnForeColour

bttnConfirmRepay.BackColor = BttnBackColour
bttnConfirmRepay.ForeColor = BttnForeColour

bttnMarkAsCompleted.BackColor = BttnBackColour
bttnMarkAsCompleted.ForeColor = BttnForeColour

bttnPayDoctor.BackColor = BttnBackColour
bttnPayDoctor.ForeColor = BttnForeColour

bttnPrintAll.BackColor = BttnBackColour
bttnPrintAll.ForeColor = BttnForeColour

bttnPrintCompleted.BackColor = BttnBackColour
bttnPrintCompleted.ForeColor = BttnForeColour

bttnPrintFullyPaid.BackColor = BttnBackColour
bttnPrintFullyPaid.ForeColor = BttnForeColour

bttnRemove.BackColor = BttnBackColour
bttnRemove.ForeColor = BttnForeColour

bttnRepay.BackColor = BttnBackColour
bttnRepay.ForeColor = BttnForeColour

bttnStillToComplete.BackColor = BttnBackColour
bttnStillToComplete.ForeColor = BttnForeColour

bttnPrintFullyPaid.BackColor = BttnBackColour
bttnPrintFullyPaid.ForeColor = BttnForeColour
'
'bttnSearch.BackColor = BttnBackColour
'bttnSearch.ForeColor = BttnForeColour



FramePatientList.BackColor = FrmBackColour
FramePatientList.ForeColor = FrmForeColour

frmPatientFacilityListManage.BackColor = FrameBackColour
frmPatientFacilityListManage.ForeColor = FrameForeColour
'
FrameSecession.BackColor = FrameBackColour
FrameSecession.ForeColor = FrameForeColour
'
'FrameSearchNames.BackColor = FrameBackColour
'FrameSearchNames.ForeColor = FrameForeColour

'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour

'chkCurrentlyChanneling.BackColor = LblBackColour
'chkCurrentlyChanneling.ForeColor = LblForeColour

DataComboDoctorStaff.BackColor = TxtBackColour
DataComboDoctorStaff.ForeColor = TxtForeColour

DataComboFacility.BackColor = TxtBackColour
DataComboFacility.ForeColor = TxtForeColour

'DataComboSex.BackColor = TxtBackColour
'DataComboSex.ForeColor = TxtForeColour

'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour

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



Label1.BackColor = LblBackColour
Label1.ForeColor = LblForeColour

Label10.BackColor = LblBackColour
Label10.ForeColor = LblForeColour
Label11.BackColor = LblBackColour
Label11.ForeColor = LblForeColour
'Label12.BackColor = LblBackColour
'Label12.ForeColor = LblForeColour
Label13.BackColor = LblBackColour
Label13.ForeColor = LblForeColour
'Label14.BackColor = LblBackColour
'Label14.ForeColor = LblForeColour
'Label15.BackColor = LblBackColour
'Label15.ForeColor = LblForeColour
'Label16.BackColor = LblBackColour
'Label16.ForeColor = LblForeColour
Label2.BackColor = LblBackColour
Label2.ForeColor = LblForeColour
'Label18.BackColor = LblBackColour
'Label18.ForeColor = LblForeColour
Label3.BackColor = LblBackColour
Label3.ForeColor = LblForeColour
'Label20.BackColor = LblBackColour
'Label20.ForeColor = LblForeColour
'Label21.BackColor = LblBackColour
'Label21.ForeColor = LblForeColour
Label4.BackColor = LblBackColour
Label4.ForeColor = LblForeColour
'Label23.BackColor = LblBackColour
'Label23.ForeColor = LblForeColour
'Label24.BackColor = LblBackColour
'Label24.ForeColor = LblForeColour
'Label25.BackColor = LblBackColour
'Label25.ForeColor = LblForeColour
'Label17.BackColor = LblBackColour
'Label17.ForeColor = LblForeColour
'Label27.BackColor = LblBackColour
'Label27.ForeColor = LblForeColour
Label4.BackColor = LblBackColour
Label4.ForeColor = LblForeColour
Label5.BackColor = LblBackColour
Label5.ForeColor = LblForeColour
'Label6.BackColor = LblBackColour
'Label6.ForeColor = LblForeColour
'Label7.BackColor = LblBackColour
'Label7.ForeColor = LblForeColour

'Label8.BackColor = LblBackColour
'Label8.ForeColor = LblForeColour
'Label9.BackColor = LblBackColour
'Label9.ForeColor = LblForeColour

Label1.BackColor = LblBackColour
Label1.ForeColor = LblForeColour

'lblOfficialWebsite.BackColor = LblBackColour
'lblOfficialWebsite.ForeColor = LblForeColour


txtInstitutionRepay.BackColor = TxtBackColour
txtInstitutionRepay.ForeColor = TxtForeColour

txtOtherRepay.BackColor = TxtBackColour
txtOtherRepay.ForeColor = TxtForeColour

txtRepayComments.BackColor = TxtBackColour
txtRepayComments.ForeColor = TxtForeColour
txtRepayTotal.BackColor = TxtBackColour
txtRepayTotal.ForeColor = TxtForeColour
txtStaffRepay.BackColor = TxtBackColour
txtStaffRepay.ForeColor = TxtForeColour
'txtNIC.BackColor = TxtBackColour
'txtNIC.ForeColor = TxtForeColour
'txtNotes.BackColor = TxtBackColour
'txtNotes.ForeColor = TxtForeColour
'txtOtherName.BackColor = TxtBackColour
'txtOtherName.ForeColor = TxtForeColour
'txtPhoto.BackColor = TxtBackColour
'txtPhoto.ForeColor = TxtForeColour
'txtSearchFirstName.BackColor = TxtBackColour
'txtSearchFirstName.ForeColor = TxtForeColour
'txtSearchID.BackColor = TxtBackColour
'txtSearchID.ForeColor = TxtForeColour
'txtSearchSurname.BackColor = TxtBackColour
'txtSearchSurname.ForeColor = TxtForeColour
'
'txtSurname.BackColor = TxtBackColour
'txtSurname.ForeColor = TxtForeColour
'txtTelephone.BackColor = TxtBackColour
'txtTelephone.ForeColor = TxtForeColour
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
OptionEvening.ForeColor = FrmForeColour
OptionEvening.BackColor = FrmBackColour

OptionNotRelevent.ForeColor = FrmForeColour
OptionNotRelevent.BackColor = FrmBackColour

OptionMorning.ForeColor = FrmForeColour
OptionMorning.BackColor = FrmBackColour

End Sub

Public Sub FormatPatientFacilityList()
    Dim BorderMargin As Long
    BorderMargin = 150
    With GridList1
        .Clear
        .Rows = 1
        .Row = 0
        .Cols = 16
        
        .ColWidth(0) = 900
        .Col = 0
        .CellAlignment = 4
        .Text = "Serial"
        
        .ColWidth(1) = 3000
        .Col = 1
        .CellAlignment = 4
        .Text = "Patient Name"
    
        .ColWidth(2) = 1200
        .Col = 2
        .CellAlignment = 4
        .Text = "Fully Paid"
    
        .ColWidth(3) = 1200
        .Col = 3
        .CellAlignment = 4
        .Text = "Repay/Cancellation"
        
        .ColWidth(4) = 1800
        .Col = 4
        .CellAlignment = 4
        .Text = "Personal Fee"
        
        .ColWidth(5) = 2000
        .Col = 5
        .CellAlignment = 4
        .Text = "Institution Fee"

        .ColWidth(6) = 2000
        .Col = 6
        .CellAlignment = 4
        .Text = "Other Fee"

        .ColWidth(7) = 2000
        .Col = 7
        .CellAlignment = 4
        .Text = "Personal Refund"
        
        .ColWidth(8) = 2000
        .Col = 8
        .CellAlignment = 4
        .Text = "Institution Refund"

        .ColWidth(9) = 2000
        .Col = 9
        .CellAlignment = 4
        .Text = "Other Refund"

        .ColWidth(10) = 1700
        .Col = 10
        .CellAlignment = 4
        .Text = "Paid to Personal"

        .ColWidth(11) = 1
        .Col = 11
        .CellAlignment = 4
        .Text = "PatientFacilit_ID"
    
        .ColWidth(12) = 1
        .Col = 12
        .Text = "Patient_ID"
        
        .ColWidth(13) = 1
        .Col = 13
        .Text = "HospitalFacility_ID"
        
        .ColWidth(14) = 1
        .Col = 14
        .Text = "FacilityStaff_ID"
        
        .ColWidth(15) = 1
        .Col = 15
        .Text = "PatientBill_ID"
        
        
        
    End With
End Sub

Private Sub PrintAllPaitent()
Dim I As Long

'Printer.PaperSize = vbPRPSA5
Printer.Font = "Bernard MT Condensed"
Printer.Print
Printer.FontSize = 14
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print

'Printer.Print Tab(2); InstitutionName
'Printer.FontSize = 10
'Printer.Print Tab(3); InstitutionAddress
'Printer.Print Tab(3); InstitutionTelephone
Printer.FontName = "Arial Black"
Printer.FontSize = 10
Printer.Print
Printer.Print Tab(2); "All Patient List As at "; Tab(26); DTPickerAppointment.Value
Printer.Print Tab(2); lblFacility.Caption; Tab(15); DataComboFacility.Text
Printer.Print Tab(2); LblDoctorStaff; Tab(15); ":  "; DataComboDoctorStaff.Text
Printer.FontName = ""
Printer.FontSize = 10
Printer.Print
Printer.FontName = "Arial"
Printer.Print Tab(2); "No"; Tab(7); "Patient Name"; Tab(34); "Paid"; Tab(43); "Visit Complete"
Printer.Print
For I = 1 To GridList1.Rows - 1

Printer.Print Tab(2); GridList1.TextMatrix(I, 0); Tab(7); GridList1.TextMatrix(I, 1); Tab(37); GridList1.TextMatrix(I, 2); Tab(49); GridList1.TextMatrix(I, 3)


Next I
Printer.EndDoc

End Sub

Private Sub PrintFullyPaidPaitent()

Dim I As Long

Printer.Font = "Bernard MT Condensed"
Printer.Print
Printer.FontSize = 14
Printer.Print Tab(2); InstitutionName
Printer.FontSize = 10
Printer.Print Tab(3); InstitutionAddress
Printer.Print Tab(3); InstitutionTelephone
Printer.FontName = "Arial Black"
Printer.FontSize = 10
Printer.Print
Printer.Print Tab(2); "Fully Paid Patient List "
Printer.Print Tab(2); "Date   : "; Format(DTPickerAppointment.Value, "dd mmmm yyyy")
Printer.Print Tab(2); lblFacility.Caption; Tab(15); DataComboFacility.Text
Printer.Print Tab(2); LblDoctorStaff; Tab(15); DataComboDoctorStaff.Text
Printer.FontName = ""
Printer.FontSize = 10
Printer.Print
Printer.FontName = "Arial"
Printer.Print Tab(2); "No"; Tab(7); "Patient Name"; Tab(40); "Visit Complete"
Printer.Print

For I = 1 To GridList1.Rows - 1

If GridList1.TextMatrix(I, 2) = "Yes" Then
''Printer.Print Tab(2); GridList1.TextMatrix(I, 0); Tab(7); GridList1.TextMatrix(I, 1); Tab(37); GridList1.TextMatrix(I, 2); Tab(49); GridList1.TextMatrix(I, 3)
Printer.Print Tab(2); GridList1.TextMatrix(I, 0); Tab(7); GridList1.TextMatrix(I, 1); Tab(46); GridList1.TextMatrix(I, 3) ' Tab(49); GridList1.TextMatrix(I, 3)
End If

Next I
Printer.EndDoc
End Sub

Private Sub PrintStilitoCompletePaitent()

Dim I As Long

Printer.Font = "Bernard MT Condensed"
Printer.Print
Printer.FontSize = 14
Printer.Print Tab(2); InstitutionName
Printer.FontSize = 10
Printer.Print Tab(3); InstitutionAddress
Printer.Print Tab(3); InstitutionTelephone
Printer.FontName = "Arial Black"
Printer.FontSize = 10
Printer.Print
Printer.Print Tab(2); "Still to Complete Patient List"
Printer.Print Tab(2); "Date   : "; Format(DTPickerAppointment.Value, "dd mmmm yyyy")
Printer.Print Tab(2); lblFacility.Caption; Tab(15); DataComboFacility.Text
Printer.Print Tab(2); LblDoctorStaff; Tab(15); DataComboDoctorStaff.Text
Printer.FontName = ""
Printer.FontSize = 10
Printer.Print
Printer.FontName = "Arial"
Printer.Print Tab(2); "No"; Tab(7); "Patient Name"; Tab(40); "Paid"
Printer.Print
For I = 1 To GridList1.Rows - 1

If GridList1.TextMatrix(I, 3) = "No" Then
'Printer.Print Tab(2); GridList1.TextMatrix(I, 0); Tab(7); GridList1.TextMatrix(I, 1); Tab(37); GridList1.TextMatrix(I, 2); Tab(49); GridList1.TextMatrix(I, 3)
Printer.Print Tab(2); GridList1.TextMatrix(I, 0); Tab(7); GridList1.TextMatrix(I, 1); Tab(41); GridList1.TextMatrix(I, 2)
End If

Next I
Printer.EndDoc

End Sub

Private Sub PrintVisitCompletedpatient()

Dim I As Long

Printer.Font = "Bernard MT Condensed"
Printer.Print
Printer.FontSize = 14
Printer.Print Tab(2); InstitutionName
Printer.FontSize = 10
Printer.Print Tab(3); InstitutionAddress
Printer.Print Tab(3); InstitutionTelephone
Printer.FontName = "Arial Black"
Printer.FontSize = 10
Printer.Print
Printer.Print Tab(2); "Visit Completed Patient List As at "
Printer.Print Tab(2); "Date   : "; Format(DTPickerAppointment.Value, "dd mmmm yyyy")
Printer.Print Tab(2); lblFacility.Caption; Tab(15); DataComboFacility.Text
Printer.Print Tab(2); LblDoctorStaff.Caption; Tab(15); DataComboDoctorStaff.Text
Printer.FontName = ""
Printer.FontSize = 10
Printer.Print
Printer.FontName = "Arial"
Printer.Print Tab(2); "No"; Tab(7); "Patient Name"; Tab(40); "Paid"
Printer.Print
For I = 1 To GridList1.Rows - 1

If GridList1.TextMatrix(I, 3) = "Yes" Then
'Printer.Print Tab(2); GridList1.TextMatrix(I, 0); Tab(7); GridList1.TextMatrix(I, 1); Tab(37); GridList1.TextMatrix(I, 2); Tab(49); GridList1.TextMatrix(I, 3)
Printer.Print Tab(2); GridList1.TextMatrix(I, 0); Tab(7); GridList1.TextMatrix(I, 1); Tab(41); GridList1.TextMatrix(I, 2)
End If

Next I
Printer.EndDoc





End Sub


Public Sub FillPatientFacilityList()
    Dim NowRow As Long
    Dim TemNum As Long
    If TemDailyMaximum > 1 Then
        GridList1.Rows = TemDailyMaximum + 1
    Else
        GridList1.Rows = 1
    End If
    GridList1.Col = 0
    For TemNum = 1 To TemDailyMaximum
        GridList1.Rows = TemNum + 1
        GridList1.Row = TemNum
        GridList1.Text = TemNum
    Next
    
    If OptionMorning.Value = True Then
        TemSecession = MorningSecession
    ElseIf OptionEvening.Value = True Then
        TemSecession = EveningSecession
    ElseIf OptionNotRelevent.Value = True Then
        TemSecession = NoReleventSecession
    End If
    
    With DataEnvironment1.rssqlTem3
        If .State = 1 Then .Close
        .Source = "select tblpatientfacility.* from tblpatientfacility where (FacilityStaff_ID = " & TemStaffFacilityID & ") and (appointmentdate = #" & DTPickerAppointment.Value & "#) and (secession = " & TemSecession & ") order by dayserial"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        While Not .EOF
            TemNum = !DaySerial
            If TemNum + 1 >= GridList1.Rows Then GridList1.Rows = TemNum + 1
            GridList1.Row = TemNum
            
            GridList1.Col = 0
            GridList1.CellAlignment = 4
            GridList1.Text = !DaySerial
            
            GridList1.Col = 1
            GridList1.CellAlignment = 4
            GridList1.Text = FindPatientByID(Val(!patientid))
        
            GridList1.Col = 2
            GridList1.CellAlignment = 4
            If !fullypaid = True Then
                GridList1.Text = "Yes"
            Else
                GridList1.Text = "No"
            End If
        
            GridList1.Col = 3
            GridList1.CellAlignment = 4
            If !cancelled = True Then
                GridList1.Text = "Cancelled"
            ElseIf !refund = True Then
                GridList1.Text = "Repaied"
            End If
            
            
            GridList1.Col = 4
            GridList1.CellAlignment = 4
            GridList1.Text = !Personalfee
            
            GridList1.Col = 5
            GridList1.CellAlignment = 4
            GridList1.Text = !institutionfee
    
            GridList1.Col = 6
            GridList1.CellAlignment = 4
            GridList1.Text = !otherfee
    
            GridList1.Col = 7
            GridList1.CellAlignment = 4
            If Not IsNull(!personalrefund) Then
                GridList1.Text = !personalrefund
            Else
                GridList1.Text = 0
            End If
            
            GridList1.Col = 8
            GridList1.CellAlignment = 4
            If Not IsNull(!InstitutionRefund) Then
                GridList1.Text = !InstitutionRefund
            Else
                GridList1.Text = 0
            End If
    
            GridList1.Col = 9
            GridList1.CellAlignment = 4
            
            If Not IsNull(!OtherRefund) Then
                GridList1.Text = !OtherRefund
            Else
                GridList1.Text = 0
            End If
            GridList1.Col = 10
            GridList1.CellAlignment = 4
            If !PaidToSTaff = True Then
                GridList1.Text = "Yes"
            Else
                GridList1.Text = "No"
            End If
    
            GridList1.Col = 11
            GridList1.CellAlignment = 4
            GridList1.Text = !patientfacility_ID
            
            GridList1.Col = 12
            GridList1.Text = !patientid
            GridList1.Col = 13
            GridList1.Text = !HospitalFacility_id
            GridList1.Col = 14
            GridList1.Text = !FacilityStaff_ID
            GridList1.Col = 15
            GridList1.Text = !PatientBill_ID
            
            
            .MoveNext
        Wend
        .Close
    GridList1.Col = 0
    For TemNum = 1 To GridList1.Rows - 1
        GridList1.Row = TemNum
        GridList1.Text = TemNum
    Next
    End With
    GridList1.Row = 0
    GridList1.Col = 0
End Sub

Private Sub bttnCancelRepay_Click()
    frameRepay.Visible = False
    frameRepay.Enabled = False
    FramePatientList.Visible = True
    FramePatientList.Enabled = True
    FramePatientList.ZOrder 0
End Sub

Private Sub bttnCloseList_Click()
Unload Me
End Sub

Private Sub bttnConfirmRepay_Click()
    Dim TemResponce  As Integer
    If Val(lblPreviousTotalRepay.Caption) + Val(txtRepayTotal.Text) > Val(lblTotalPaid.Caption) Then
        TemResponce = MsgBox("You can't repay an amount grater than that paid initially by the patient", vbCritical, "Exceeds Payment")
        txtStaffRepay.SetFocus
        Exit Sub
    End If
    With DataEnvironment1.rssqlTem6
        If .State = 1 Then .Close
        .Source = "select * from tblpatientrepay"
        If .State = 0 Then .Open
        .AddNew
        GridList1.Col = 12
        !patient_ID = GridList1.Text
        GridList1.Col = 13
        !HospitalFacility_id = GridList1.Text
        GridList1.Col = 14
        !FacilityStaff_ID = GridList1.Text
        !repayuser_ID = UserID
        !catogery = TemCatogery
        GridList1.Col = 15
        !PatientBill_ID = GridList1.Text
        !RepayDate = Date
        !StaffRepay = Val(txtStaffRepay.Text)
        !InstitutionRepay = Val(txtInstitutionRepay.Text)
        !OtherRepay = Val(txtOtherRepay.Text)
        !TotalRepay = Val(txtRepayTotal.Text)
        If Trim(txtRepayComments.Text) = "" Then
            If IsACancellation = True Then !repaycomments = "Cancellation"
            If IsARefund = True Then !repaycomments = "Refund"
        Else
            !repaycomments = txtRepayComments.Text
        End If
        GridList1.Col = 11
        !patientfacility_ID = Val(GridList1.Text)
        !RepayDate = Date
        .Update
        .Close
        .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(GridList1.Text)
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
            If IsNull(!personalrefund) Then
                !personalrefund = Val(txtStaffRepay.Text)
            Else
                !personalrefund = Val(!personalrefund) + Val(txtStaffRepay.Text)
            End If
            If IsNull(!InstitutionRefund) Then
                !InstitutionRefund = Val(txtInstitutionRepay.Text)
            Else
                !InstitutionRefund = Val(!InstitutionRefund) + Val(txtInstitutionRepay.Text)
            End If
            If IsNull(!OtherRefund) Then
                !OtherRefund = Val(txtOtherRepay.Text)
            Else
                !OtherRefund = Val(!OtherRefund) + Val(txtOtherRepay.Text)
            End If
            If IsNull(!totalRefund) Then
                !totalRefund = Val(txtRepayTotal.Text)
            Else
                !totalRefund = Val(!totalRefund) + Val(txtRepayTotal.Text)
            End If
            If Trim(txtRepayComments.Text) = "" Then
                If IsACancellation = True Then !repaycomments = "Cancellation"
                If IsARefund = True Then !repaycomments = "Refund"
            Else
                !repaycomments = txtRepayComments.Text
            End If
            GridList1.Col = 11
            !RepayDate = Date
            If IsACancellation = True Then !cancelled = True
            If IsARefund = True Then !refund = True
            !repayuser_ID = UserID
            .Update
        .Close
    End With
    
    If OptionOne.Value = True Then
        PrintRepay
    ElseIf OptionTwo.Value = True Then
        PrintRepay
        PrintRepay
    End If

    frameRepay.Visible = False
    frameRepay.Enabled = False
    FramePatientList.Visible = True
    FramePatientList.Enabled = True
    FramePatientList.ZOrder 0
    FormatPatientFacilityList
    FillPatientFacilityList
    
    
End Sub

Private Sub PrintRepay()

Printer.Font = "Bernard MT Condensed"
Printer.Print
Printer.FontSize = 14
Printer.Print Tab(2); InstitutionName
Printer.FontSize = 10
Printer.Print Tab(3); InstitutionAddress
Printer.Print Tab(3); InstitutionTelephone
Printer.FontName = "Arial"
Printer.FontSize = 11
Printer.Print Tab(6); "--------------------------------"
Printer.Print
Printer.Print Tab(6); "        REPAYMENT"
Printer.Print
Printer.Print Tab(6); "--------------------------------"
Printer.Print
GridList1.Col = 1
Printer.FontBold = False
Printer.Print Tab(2); "Patient   :     "; GridList1.Text
Printer.Print
Printer.Print Tab(2); "Refund   :   Rs. "; Format(Val(txtRepayTotal.Text), "#00.00")
Printer.Print
Printer.Print
Printer.Print
Printer.Print Tab(2); "-----------------          ----------------"
Printer.Print Tab(2); "Customer Signature     User Signature"
Printer.Print Tab(2); ""
Printer.EndDoc


End Sub

Private Sub bttnMarkAsCompleted_Click()
    If TemPatientFacilityID = 0 Then Exit Sub
    
    IsARefund = False
    IsACancellation = True
    
'    With DataEnvironment1.rssqlTem
'        If .State = 1 Then .Close
'        .Source = "SELECT tblpatientfacility.* from tblpatientfacility where patientfacility_ID = " & TemPatientFacilityID
'        If .State = 0 Then .Open
'        If .RecordCount = 0 Then Exit Sub
'        !resultsuccess = True
'        .Update
'        .Close
'    End With
    

    Dim TemResponce As Long
    
    If GridList1.Row < 1 Then Exit Sub
        
    GridList1.Col = 2
    If GridList1.Text = "No" Then
            With DataEnvironment1.rssqlTem7
                If .State = 1 Then .Close
                GridList1.Col = 11
                .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(GridList1.Text)
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                    If IsACancellation = True Then !cancelled = True
                    !repayuser_ID = UserID
                    .Update
                    .Close
            End With
        Exit Sub
    End If
    GridList1.Col = 3
    If GridList1.Text = "Cancelled" Then
        TemResponce = MsgBox("This booking has already cancelled", vbCritical, "Already Cancelled")
        Exit Sub
    End If
    GridList1.Col = 3
    If GridList1.Text = "Repaied" Then
        TemResponce = MsgBox("Money for this booking has already repaied", vbCritical, "Already Repaied")
        Exit Sub
    End If
    
    FramePatientList.Enabled = False
    frameRepay.Visible = True
    frameRepay.Enabled = True
    frameRepay.ZOrder 0
    txtStaffRepay.Text = Empty
    txtInstitutionRepay.Text = Empty
    txtOtherRepay.Text = Empty
    
    GridList1.Col = 4
    lblStaffFeePaid.Caption = Format(Val(GridList1.Text), "#0.00")
    
    GridList1.Col = 5
    lblInstitutionFeePaid.Caption = Format(Val(GridList1.Text), "#0.00")
    
    GridList1.Col = 6
    lblOtherFeePaid.Caption = Format(Val(GridList1.Text), "#0.00")
    
    GridList1.Col = 7
    lblPreviousStaffRepay.Caption = Format(Val(GridList1.Text), "#0.00")
    
    GridList1.Col = 8
    lblPreviousInstitutionRepay.Caption = Format(Val(GridList1.Text), "#0.00")
    
    GridList1.Col = 9
    lblPreviousOtherRepay.Caption = Format(Val(GridList1.Text), "#0.00")

'    FormatPatientFacilityList
'    FillPatientFacilityList
End Sub

Private Sub bttnPayDoctor_Click()

Dim DoctorFee As Double
Dim I As Long

Printer.Font = "Bernard MT Condensed"
Printer.Print
Printer.FontSize = 14
Printer.Print Tab(2); InstitutionName
Printer.FontSize = 10
Printer.Print Tab(3); InstitutionAddress
Printer.Print Tab(3); InstitutionTelephone
Printer.FontName = "Arial Black"
Printer.FontSize = 10
Printer.Print
Printer.Print Tab(2); DataComboFacility.Text
Printer.Print Tab(2); LblDoctorStaff.Caption & "   " & DataComboDoctorStaff.Text
Printer.Print Tab(2); "Date   : "; Format(DTPickerAppointment.Value, "dd mmmm yyyy")
'Printer.Print Tab(2); lblFacility.Caption; Tab(15); DataComboFacility.Text
'Printer.Print Tab(2); lblDoctorStaff; Tab(15); DataComboDoctorStaff.Text
Printer.FontName = ""
Printer.FontSize = 10
Printer.Print
Printer.FontName = "Arial"
Printer.Print Tab(2); "No"; Tab(7); "Patient Name"; Tab(40); LblDoctorStaff.Caption & " Fee"
Printer.Print
DoctorFee = 0

For I = 1 To GridList1.Rows - 1
    
    If GridList1.TextMatrix(I, 2) = "Yes" Then
        Printer.Print Tab(2); GridList1.TextMatrix(I, 0); Tab(7); GridList1.TextMatrix(I, 1); Tab(40);
        If GridList1.TextMatrix(1, 10) = "Yes" Then
            Printer.Print "Already Paid"
        Else
            Printer.Print Format((Val(GridList1.TextMatrix(I, 4)) - Val(GridList1.TextMatrix(I, 7))), "#0.00")
            DoctorFee = DoctorFee + (Val(GridList1.TextMatrix(I, 4)) - Val(GridList1.TextMatrix(I, 7)))
        End If
    End If
    
Next I

Printer.Print
Printer.Print "Total " & LblDoctorStaff.Caption & " Fee :  Rs. " & Format(DoctorFee, "#0.00")
Printer.Print
Printer.Print
Printer.Print Tab(2); "----------------------------"; Tab(30); "----------------------------"
Printer.Print Tab(2); UserName; Tab(30); DataComboDoctorStaff.Text

Printer.EndDoc

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblStaffPayment"
    If .State = 0 Then .Open
    .AddNew
    !HospitalFacility_id = TemHospitalFacilityID
    !FacilityStaff_ID = TemStaffFacilityID
    If TemCatogery = Doctor Then
        !isadoctor = True
    Else
        !isadoctor = False
    End If
    !staff_ID = DataComboDoctorStaff.BoundText
    !PaidAmount = DoctorFee
    !paiddate = Date
    !User_ID = UserID
    .Update
    .Close
End With

For I = 1 To GridList1.Rows - 1
    If GridList1.TextMatrix(I, 2) = "Yes" Then
        With DataEnvironment1.rssqlTem
            If .State = 1 Then .Close
            .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(GridList1.TextMatrix(I, 11))
            If .State = 0 Then .Open
            If .RecordCount = 0 Then Exit Sub
            !PaidToSTaff = True
            .Update
            .Close
        End With
    End If
Next

FormatPatientFacilityList
FillPatientFacilityList

End Sub

Private Sub bttnPrintAll_Click()
'    Dim TemNum As Long
'    Select Case TemCatogery
'
'
'    Case Doctor:
'        PrintPatientName (FindDoctorFromID(TemstaffID))
'    Case Staff:
'        PrintPatientName (FindStaffFromID(TemstaffID))
'    Case Investigation:
'        PrintPatientName (FindInvestigationFromID(TemstaffID))
'    End Select
'
'    If PreferancechkPatientAge = True Then Call PrintPatientAge(txtAge.Text)
'    If PreferancechkPatientSex = True Then Call PrintPatientSex(DataComboSex.Text)
'    If PreferancechkPatientID = True Then Call PrintPatientID(TemPatientID)
'    If PreferanceChkLblComments = True Then Call PrintLblComments(PreferanceTxtLblComments)
'
'
'    For TemNum = 1 To GridList1.Rows - 1
'
'    Next

Call PrintAllPaitent
End Sub

Private Sub bttnPrintCompleted_Click()
Call PrintVisitCompletedpatient
End Sub

Private Sub bttnPrintFullyPaid_Click()
Call PrintFullyPaidPaitent
End Sub

Private Sub bttnRepay_Click()
    IsARefund = True
    IsACancellation = False

    Dim TemResponce As Long
    If GridList1.Row < 1 Then Exit Sub
        
    GridList1.Col = 2
    
    If GridList1.Text = "No" Then
        TemResponce = MsgBox("This patient has not completed the payment. Therefore you can't repay", vbCritical, "Not Paid")
        Exit Sub
    End If
    GridList1.Col = 3
    If GridList1.Text = "Cancelled" Then
        TemResponce = MsgBox("This booking has already cancelled", vbCritical, "Already Cancelled")
        Exit Sub
    End If
    GridList1.Col = 3
    If GridList1.Text = "Repaied" Then
        TemResponce = MsgBox("Money for this booking has already repaied", vbCritical, "Already Repaied")
        Exit Sub
    End If
    
    
    
    
    FramePatientList.Enabled = False
    frameRepay.Visible = True
    frameRepay.Enabled = True
    frameRepay.ZOrder 0
    txtStaffRepay.Text = Empty
    txtInstitutionRepay.Text = Empty
    txtOtherRepay.Text = Empty
    
    GridList1.Col = 4
    lblStaffFeePaid.Caption = Format(Val(GridList1.Text), "#0.00")
    
    GridList1.Col = 5
    lblInstitutionFeePaid.Caption = Format(Val(GridList1.Text), "#0.00")
    
    GridList1.Col = 6
    lblOtherFeePaid.Caption = Format(Val(GridList1.Text), "#0.00")
    
    GridList1.Col = 7
    lblPreviousStaffRepay.Caption = Format(Val(GridList1.Text), "#0.00")
    
    GridList1.Col = 8
    lblPreviousInstitutionRepay.Caption = Format(Val(GridList1.Text), "#0.00")
    
    GridList1.Col = 9
    lblPreviousOtherRepay.Caption = Format(Val(GridList1.Text), "#0.00")

End Sub

Private Sub bttnStillToComplete_Click()
Call PrintStilitoCompletePaitent
End Sub


Private Sub bttnTelephonePay_Click()
    Dim TemResponce As Long
    If GridList1.Row < 1 Then Exit Sub
    
    GridList1.Col = 2
   
    If GridList1.Text = "Yes" Then
        TemResponce = MsgBox("This patient has already completed the payment. Therefore you can't repay", vbCritical, "Already Paid")
        Exit Sub
    End If
    GridList1.Col = 3
    If GridList1.Text = "Cancelled" Then
        TemResponce = MsgBox("This booking has already cancelled", vbCritical, "Already Cancelled")
        Exit Sub
    End If
    GridList1.Col = 3
    If GridList1.Text = "Repaied" Then
        TemResponce = MsgBox("Money for this booking has already repaied", vbCritical, "Already Repaied")
        Exit Sub
    End If

With DataEnvironment1.rssqlTem3
    If .State = 1 Then .Close
        GridList1.Col = 11
        .Source = "select tblpatientfacility.* from tblpatientfacility where patientfacility_ID = " & Val(GridList1.Text)
        .Open
    If .RecordCount = 0 Then .Close: Exit Sub
    !fullypaid = True
    .Update
    



    Dim TemRows As Long

        
        Printer.Font = "Bernard MT Condensed"
        Printer.Print
        Printer.FontSize = 14
        Printer.Print 'Tab(2); InstitutionName
        Printer.FontSize = 12
        Printer.Print ' Tab(3); InstitutionAddress
        Printer.Print 'Tab(3); InstitutionTelephone
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        
        Printer.FontName = "Courier"
        Printer.FontSize = 10
        Printer.Print
        
        Dim TemTab1 As Long
        Dim TemTab2 As Long
        Dim TemTab3 As Long
        Dim TemTab4 As Long
        Dim TemTab5 As Long
        Dim TemTab6 As Long
        
        TemTab1 = 2
        TemTab2 = 6
        TemTab3 = 20
        TemTab4 = 25
        TemTab5 = 36
        TemTab6 = 16
        
        Printer.Print Tab(TemTab1); "Patient";
        Printer.Print Tab(TemTab6); " : ";
        GridList1.Col = 1
        Printer.Print Tab(TemTab3); GridList1.Text
        Printer.Print
        Printer.Print Tab(TemTab1); "Consultant";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); DataComboDoctorStaff.Text
        Printer.Print
        Printer.Print Tab(TemTab1); "Appo. Date ";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); Format(DTPickerAppointment.Value, "dd mmmm yyyy")
        
'        Printer.Print Tab(TemTab1); "Appo. Time";
'        Printer.Print Tab(TemTab6); " : ";
'        Printer.Print Tab(TemTab3); TemAppointmentTime
        
        GridList1.Col = 0

        Printer.Print Tab(TemTab1); "Appo. No.";
        Printer.Print Tab(TemTab6); " : ";
        
        GridList1.Col = 14
        
        Printer.Print Tab(TemTab3); GridList1.Text
        
        Printer.Print Tab(TemTab1); "Appo. ID";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); GridList1.Text
        Printer.Print
        
        GridList1.Col = 4
        Printer.Print Tab(TemTab1); "Doctor Fee";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3 + 8 - Len(Format(Val(GridList1.Text), "0.00"))); Format(Val(GridList1.Text), "0.00")
        
        GridList1.Col = 5
        Printer.Print Tab(TemTab1); "Hospital Fee";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3 + 8 - Len(Format(Val(GridList1.Text), "0.00"))); Format(Val(GridList1.Text), "0.00")
        
        
        Printer.Print Tab(TemTab1); "Total Fee";
        Printer.Print Tab(TemTab6); " : ";
        
        Dim TemTemTotalFee As Double
        GridList1.Col = 4
        TemTemTotalFee = Val(GridList1.Text)
        GridList1.Col = 5
        TemTemTotalFee = TemTemTotalFee + Val(GridList1.Text)
        GridList1.Col = 6
        TemTemTotalFee = TemTemTotalFee + Val(GridList1.Text)
        
        Printer.Print Tab(TemTab3 + 8 - Len(Format(TemTemTotalFee, "0.00"))); Format(TemTemTotalFee, "0.00")
        Printer.Print
        Printer.Print Tab(TemTab2); "--------------------"
        Printer.Print Tab(TemTab2); UserName
        Printer.Print Tab(TemTab2); Format(Date, "dd mmmm yyyy")
                
        Printer.EndDoc
        
        .Close
        
        
    End With

FormatPatientFacilityList
FillPatientFacilityList













End Sub


Private Sub PrintTelephonePayment()

End Sub



Private Sub DataComboDoctorStaff_Change()
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
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
    Call FormatPatientFacilityList
    Call FillPatientFacilityList
End Sub

Private Sub DataComboFacility_Change()
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
                If .RecordCount = 0 Then Exit Sub
                DataComboDoctorStaff.RowMember = "sqlBookingFacility"
                DataComboDoctorStaff.ListField = "doctorname"
                DataComboDoctorStaff.BoundColumn = "doctor_ID"
            Case Staff:
                .Source = "SELECT tblfacilitystaff.* , tblstaff.* FROM tblfacilitystaff left join tblstaff on tblfacilitystaff.staff_ID = tblstaff.staff_ID where HospitalFacility_ID = " & DataComboFacility.BoundText & " order by staffname"
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                DataComboDoctorStaff.RowMember = "sqlBookingFacility"
                DataComboDoctorStaff.ListField = "staffname"
                DataComboDoctorStaff.BoundColumn = "tblstaff.Staff_ID"
            Case Investigation:
                .Source = "SELECT tblfacilitystaff.* , tblinvestigations.* FROM tblfacilitystaff left join tblinvestigations on tblfacilitystaff.staff_ID = tblinvestigations.investigation_ID where HospitalFacility_ID = " & DataComboFacility.BoundText & " order by investigation"
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                DataComboDoctorStaff.RowMember = "sqlBookingFacility"
                DataComboDoctorStaff.ListField = "investigation"
                DataComboDoctorStaff.BoundColumn = "investigation_ID"
            Case Other:
        End Select
        .Close
    End With
    FormatPatientFacilityList
End Sub

Private Sub PrepareForTwoSecessions()
    FrameSecession.Enabled = True
    OptionNotRelevent.Value = False
    OptionNotRelevent.Enabled = False
    OptionMorning.Enabled = True
    OptionEvening.Enabled = True
    OptionNoPreferance.Enabled = True
End Sub

Private Sub PrepareForOneSecession()
    FrameSecession.Enabled = False
    OptionMorning.Enabled = False
    OptionEvening.Enabled = False
    OptionNoPreferance.Enabled = False
    OptionNotRelevent.Enabled = True
    OptionNotRelevent.Value = True
End Sub

Private Sub DTPickerAppointment_Change()
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    FormatPatientFacilityList
    FillPatientFacilityList
End Sub

Private Sub Form_Load()
    frameRepay.Visible = False
    frameRepay.Enabled = False
    FramePatientList.Visible = True
    FramePatientList.Enabled = True
    FramePatientList.ZOrder 0
    Call Setcolours
    DTPickerAppointment.Value = Date
    FormatPatientFacilityList
End Sub

Private Sub CalculateRefundTotals()
    txtRepayTotal.Text = Format((Val(txtStaffRepay.Text) + Val(txtInstitutionRepay.Text) + Val(txtOtherRepay.Text)), "#0.00")
    lblTotalPaid.Caption = Format((Val(lblStaffFeePaid.Caption) + Val(lblInstitutionFeePaid.Caption) + Val(lblOtherFeePaid.Caption)), "#0.00")
    lblPreviousTotalRepay.Caption = Format((Val(lblPreviousStaffRepay.Caption) + Val(lblPreviousInstitutionRepay.Caption) + Val(lblPreviousOtherRepay.Caption)), "#0.00")
End Sub

Private Sub GridList1_Click()
    GridList1.Col = 11
    If GridList1.Rows <= 1 Then
        Beep
        bttnMarkAsCompleted.Enabled = False
        bttnRepay.Enabled = False
        bttnPrintAll.Enabled = False
        bttnPrintCompleted.Enabled = False
        bttnStillToComplete.Enabled = False
        bttnPrintFullyPaid.Enabled = False
        Exit Sub
    ElseIf Not IsNumeric(GridList1.Text) Then
        Beep
        bttnMarkAsCompleted.Enabled = False
        bttnRepay.Enabled = False
        bttnPrintAll.Enabled = False
        bttnPrintCompleted.Enabled = False
        bttnStillToComplete.Enabled = False
        bttnPrintFullyPaid.Enabled = False
        Exit Sub
    Else
        TemPatientFacilityID = Val(GridList1.Text)
        GridList1.Col = 0
        GridList1.ColSel = GridList1.Cols - 1
        bttnMarkAsCompleted.Enabled = True
        bttnRepay.Enabled = True
        bttnPrintAll.Enabled = True
        bttnPrintCompleted.Enabled = True
        bttnStillToComplete.Enabled = True
        bttnPrintFullyPaid.Enabled = True
    End If
End Sub

Private Sub lblInstitutionFeePaid_Click()
    Call CalculateRefundTotals
End Sub

Private Sub lblOtherFeePaid_Click()
    Call CalculateRefundTotals
End Sub

Private Sub lblPreviousInstitutionRepay_Click()
    Call CalculateRefundTotals
End Sub

Private Sub lblPreviousOtherRepay_Click()
    Call CalculateRefundTotals
End Sub

Private Sub lblPreviousStaffRepay_Click()
    Call CalculateRefundTotals
End Sub

Private Sub lblPreviousTotalRepay_Click()
    Call CalculateRefundTotals
End Sub

Private Sub lblStaffFeePaid_Click()
    Call CalculateRefundTotals
End Sub

Private Sub lblTotalPaid_Click()
    Call CalculateRefundTotals
End Sub


Private Sub OptionEvening_Click()
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    FormatPatientFacilityList
    FillPatientFacilityList
End Sub

Private Sub OptionMorning_Click()
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    FormatPatientFacilityList
    FillPatientFacilityList
End Sub

Private Sub OptionNotRelevent_Click()
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    FormatPatientFacilityList
    FillPatientFacilityList
End Sub

Private Sub txtInstitutionRepay_Change()
    Call CalculateRefundTotals
End Sub

Private Sub txtOtherRepay_Change()
    Call CalculateRefundTotals
End Sub

Private Sub txtStaffRepay_Change()
    Call CalculateRefundTotals
End Sub
