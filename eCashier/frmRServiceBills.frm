VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRServiceBills 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Roentgents Service Bills"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12900
   FillColor       =   &H0080FFFF&
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
   ScaleHeight     =   8310
   ScaleWidth      =   12900
   Begin VB.TextBox txtDisplayBillID 
      Height          =   375
      Left            =   6720
      TabIndex        =   87
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo cmbHSS 
      Height          =   360
      Left            =   2520
      TabIndex        =   85
      Top             =   6960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.CheckBox chkForeigner 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Foreigner"
      Height          =   375
      Left            =   6720
      TabIndex        =   84
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtSerialNo 
      Height          =   375
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtPaymentMethod 
      Height          =   375
      Left            =   2520
      TabIndex        =   67
      Top             =   6480
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   735
      Left            =   120
      TabIndex        =   77
      Top             =   7440
      Width           =   12615
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   240
         Width           =   4695
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   210
         Width           =   4695
      End
      Begin VB.Label Label29 
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Paper"
         Height          =   255
         Left            =   6120
         TabIndex        =   81
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label30 
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   210
         Width           =   1815
      End
   End
   Begin VB.TextBox txtPtID 
      Height          =   375
      Left            =   6720
      TabIndex        =   74
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkPrint 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Print"
      Height          =   255
      Left            =   10200
      TabIndex        =   61
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtOPDBillID 
      Height          =   375
      Left            =   6720
      TabIndex        =   73
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   6
      Left            =   7920
      TabIndex        =   57
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   5
      Left            =   7920
      TabIndex        =   52
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   4
      Left            =   7920
      TabIndex        =   47
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   3
      Left            =   7920
      TabIndex        =   42
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   2
      Left            =   7920
      TabIndex        =   37
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   32
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   26
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   6
      Left            =   8400
      TabIndex        =   58
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   53
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   4
      Left            =   8400
      TabIndex        =   48
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   43
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   2
      Left            =   8400
      TabIndex        =   38
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   33
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   28
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtCharge 
      Height          =   375
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtProfessionalCharge 
      Height          =   375
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   6
      Left            =   11520
      TabIndex        =   60
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   5
      Left            =   11520
      TabIndex        =   55
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   4
      Left            =   11520
      TabIndex        =   50
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   3
      Left            =   11520
      TabIndex        =   45
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   11520
      TabIndex        =   40
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   11520
      TabIndex        =   35
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   11520
      TabIndex        =   30
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   0
      Left            =   8880
      TabIndex        =   29
      Top             =   3120
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtEditID 
      Height          =   360
      Left            =   11400
      TabIndex        =   71
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtDelID 
      Height          =   360
      Left            =   10920
      TabIndex        =   69
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2520
      TabIndex        =   70
      Top             =   6960
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   75038723
      CurrentDate     =   39956
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11520
      TabIndex        =   63
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   32896
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   32896
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
   Begin MSFlexGridLib.MSFlexGrid gridService 
      Height          =   2895
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      _Version        =   393216
   End
   Begin VB.TextBox txtHospitalCharge 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   2160
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo cmbPatient 
      Height          =   360
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   2040
      TabIndex        =   4
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSC 
      Height          =   360
      Left            =   2040
      TabIndex        =   8
      Top             =   1680
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   32896
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
   Begin btButtonEx.ButtonEx btnUpdate 
      Height          =   495
      Left            =   10200
      TabIndex        =   62
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   32896
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
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   1
      Left            =   8880
      TabIndex        =   34
      Top             =   3600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   2
      Left            =   8880
      TabIndex        =   39
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   3
      Left            =   8880
      TabIndex        =   44
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   4
      Left            =   8880
      TabIndex        =   49
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   5
      Left            =   8880
      TabIndex        =   54
      Top             =   5520
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   6
      Left            =   8880
      TabIndex        =   59
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   375
      Left            =   4800
      TabIndex        =   72
      Top             =   6960
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   75038722
      CurrentDate     =   39956
   End
   Begin MSDataListLib.DataCombo cmbTitle 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2520
      TabIndex        =   65
      Top             =   6000
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSecession 
      Height          =   360
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtComments 
      Height          =   375
      Left            =   2040
      TabIndex        =   25
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker dtpAppDate 
      Height          =   375
      Left            =   10440
      TabIndex        =   17
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   75038723
      CurrentDate     =   39956
   End
   Begin MSComCtl2.DTPicker dtpAppTime 
      Height          =   375
      Left            =   10440
      TabIndex        =   19
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   75038722
      CurrentDate     =   39956
   End
   Begin VB.TextBox txtDuration 
      Height          =   360
      Left            =   10440
      TabIndex        =   82
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   375
      Left            =   12360
      TabIndex        =   83
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   75038722
      CurrentDate     =   39956
   End
   Begin VB.Label lblHSS 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Company"
      Height          =   255
      Left            =   120
      TabIndex        =   86
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label lblSC 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Su&bcategory"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date && Time"
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   6960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Left            =   8280
      TabIndex        =   16
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Serial No"
      Height          =   255
      Left            =   8280
      TabIndex        =   14
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Session"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Co&mments"
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment &Method"
      Height          =   255
      Left            =   120
      TabIndex        =   64
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Total Charge"
      Height          =   255
      Left            =   8280
      TabIndex        =   22
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pro&fessional Charge"
      Height          =   255
      Left            =   8280
      TabIndex        =   20
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   56
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   5
      Left            =   6720
      TabIndex        =   51
      Top             =   5520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   46
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   41
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   36
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   31
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   0
      Left            =   6720
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   2520
      TabIndex        =   76
      Top             =   5640
      Width           =   4095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   120
      TabIndex        =   75
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Hospital Charge"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Cate&gory"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pa&tient"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmRServiceBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsSC As New ADODB.Recordset
    Dim temSQL As String
    Dim rsSPC As New ADODB.Recordset
    Dim rsStaff() As New ADODB.Recordset
    Dim PSCCount As Long
    Dim FirstActi As Boolean
    Dim rsHSS As New ADODB.Recordset

    Dim CsetPrinter As New cSetDfltPrinter
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean

    Dim rsSecession As New ADODB.Recordset

Private Sub btnAdd_Click()
    If IsNumeric(txtOPDBillID.Text) = False Then txtOPDBillID.Text = NewRBillID(dtpDate.Value, dtpTime.Value)
    Dim rsTem As New ADODB.Recordset
       
'    If gridService.Rows >= 2 Then
'        MsgBox "You can't add more than one service for a R Service Bill"
'        Exit Sub
'    End If
   
   If cmbSC.Visible = True And IsNumeric(cmbSC.BoundText) = False Then
        MsgBox "Please select a subcategoty"
        Exit Sub
   End If
   
    Dim n As Integer
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Payment method?"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    If Trim(cmbPatient.Text) = "" Then
        cmbPatient.Text = cmbPaymentMethod.Text & " customer"
    End If
    
    If IsNumeric(cmbCategory.BoundText) = False Then
        MsgBox "Service?"
        cmbCategory.SetFocus
        Exit Sub
    End If
    
        On Error Resume Next

    
    If IsNumeric(txtPtID.Text) = False Then
        With rsTem
            If .State = 1 Then .Close
            temSQL = "Select * from tblPatientMainDetails where PatientID = 0 "
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !FirstName = cmbPatient.Text
            !TitleID = Val(cmbTitle.BoundText)
            .Update
            temSQL = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            txtPtID.Text = !NewID
            .Close
        End With
    End If
    
        On Error Resume Next

    
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtEditID.Text) = True Then
            temSQL = "Select * from tblPatientService where PatientServiceID = " & Val(txtEditID.Text)
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount <= 0 Then
                .AddNew
            End If
        Else
            temSQL = "Select * from tblPatientService  where PatientServiceID = 0 "
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
        End If
        !RBillID = Val(txtOPDBillID.Text)
        !ServiceCategoryID = Val(cmbCategory.BoundText)
        !ServicesubcategoryID = Val(cmbSC.BoundText)
        !Comments = txtComments.Text
        !ServiceDate = dtpDate.Value
        !ServiceTime = dtpTime.Value
        !Charge = Val(txtCharge.Text)
        !ProfessionalCharge = Val(txtProfessionalCharge.Text)
        !HospitalCharge = Val(txtHospitalCharge.Text)
        !UserID = UserID
        !SerialNo = Val(txtSerialNo.Text)
        !SecessionID = Val(cmbSecession.BoundText)
        .Update
        temSQL = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        txtEditID.Text = !NewID
        .Close
    End With
    For n = 0 To lblSpeciality1.UBound
        If lblSpeciality1(n).Visible = True Then
            With rsTem
                If .State = 1 Then .Close
                temSQL = "Select * from tblProfessionalCharges where ServiceProfessionalChargesID = " & Val(txtServiceProfessionalChargesID(n).Text) & " AND PatientServiceID = " & Val(txtEditID.Text)
                .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    
                Else
                    .AddNew
                    !UserID = UserID
                    !ForRBillID = Val(txtOPDBillID.Text)
                    !PatientServiceID = Val(txtEditID.Text)
                    !ServiceProfessionalChargesID = Val(txtServiceProfessionalChargesID(n).Text)
                    !StaffID = Val(cmbStaff1(n).BoundText)
                End If
                !Date = dtpDate.Value
                !Time = dtpTime.Value
                !Fee = Val(txtFee1(n).Text)
                !IsRBill = True
                .Update
            End With
        End If
    Next n
    Call fillGrid
    Call ClearAddValues
    
    cmbCategory.Locked = True
    cmbPaymentMethod.Locked = True
    cmbSecession.Locked = True
    If cmbSC.Visible = True Then
        cmbSC.SetFocus
    Else
        cmbCategory.SetFocus
    End If
End Sub

Private Sub ClearAddValues()
    Dim n As Long
    
'    cmbCategory.Text = Empty
    
    cmbSC.Text = Empty
    txtComments.Text = Empty
    txtProfessionalCharge.Text = Empty
    txtHospitalCharge.Text = Empty
    txtCharge.Text = Empty
    txtEditID.Text = Empty
    txtDelID.Text = Empty
    
    For n = 0 To lblSpeciality1.UBound
        lblSpeciality1(n).Visible = False
        lblSpeciality1(n).Caption = Empty
        cmbStaff1(n).Visible = False
        cmbStaff1(n).Text = Empty
        txtServiceProfessionalChargesID(n).Text = Empty
        txtFee1(n).Visible = False
        txtFee1(n).Text = Empty
        txtSpecialityID(n).Text = Empty
    Next
    
End Sub

Private Sub ClearBillValues()
    cmbCategory.Text = Empty
    cmbPaymentMethod.BoundText = 1
    cmbPatient.Text = Empty
    txtOPDBillID.Text = Empty
    txtPtID.Text = Empty
    lblTotal.Caption = "0.00"
    txtPaymentMethod.Text = Empty
    txtSerialNo.Text = Empty
    cmbTitle.Text = Empty
    chkForeigner.Value = 0
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()

    If Val(txtDelID.Text) = 0 Then
        MsgBox "Please select one to delete"
        Exit Sub
    End If


    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtDelID.Text) = True Then
            temSQL = "Select * from tblPatientService where PatientServiceID = " & Val(txtDelID.Text)
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Deleted = True
                !DeletedUserID = UserID
                !DeletedDate = Date
                !DeletedTime = Now
                .Update
            End If
            .Close
        End If
    End With
    
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblProfessionalCharges where PatientServiceID = " & Val(txtDelID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !Cancelled = True
            !CancelledDate = Date
            !CancelledTime = Now
            !CancelledDateTime = Now
            !CancelledUserID = UserID
            .Update
            .MoveNext
        Wend
        .Close
    End With
    
    
    Call fillGrid
    Call ClearAddValues
    If gridService.Rows < 2 Then
        txtSerialNo.Text = Empty
        cmbPaymentMethod.Locked = False
        cmbCategory.Locked = False
        cmbSecession.Locked = False
        cmbCategory.SetFocus
    Else
        If cmbSC.Visible = True Then
            cmbSC.SetFocus
        Else
            cmbCategory.SetFocus
        End If
    End If
    
End Sub

Private Sub btnUpdate_Click()
        
    If gridService.Rows < 2 Then
        MsgBox "Nothing to add"
        cmbCategory.SetFocus
        Exit Sub
    End If
    If Trim(cmbPatient.Text) = "" Then
        MsgBox "Please select a patient"
        cmbPatient.SetFocus
        Exit Sub
    End If
        
    If cmbPaymentMethod.BoundText = 4 Then
        If Val(cmbHSS.BoundText) = 0 Then
            MsgBox "Please select a Health Scheme Supplier"
            cmbHSS.Visible = True
            cmbHSS.SetFocus
            Exit Sub
        End If
    End If
        
        
    cmbCategory.Locked = False
    cmbPaymentMethod.Locked = False
    cmbSecession.Locked = False
    
    Dim rsTem As New ADODB.Recordset
    Dim DisplayBillID As Long
'
'    With rsTem
'        If .State = 1 Then .Close
'        temSQL = "Select Count(IncomeBillID) as BillCount from tblIncomeBill where Completed = 1 AND IsRBill = 1"
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            If IsNull(!BillCount) = False Then
'                DisplayBillID = !BillCount + 1
'            Else
'                DisplayBillID = 1
'            End If
'        Else
'            DisplayBillID = 1
'        End If
'    End With
'
    DisplayBillID = NewRDisplayBillID(Val(txtOPDBillID.Text))
    
    txtDisplayBillID.Text = DisplayBillID
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeBill where IncomeBillID = " & Val(txtOPDBillID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !DisplayBillID = DisplayBillID
            !Completed = True
            !CompletedDate = Date
            !CompletedTime = Now
            !CompletedUserID = UserID
            !PatientID = Val(txtPtID.Text)
            !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
            !PaymentComments = txtPaymentMethod.Text
            !IsRBill = True
            !GrossTotal = Val(lblTotal.Caption)
            !NetTotal = Val(lblTotal.Caption)
            !StoreID = UserStoreID
            !HSSID = Val(cmbHSS.BoundText)
            .Update
        Else
            MsgBox "Error. Bill NOT added. Please reenter"
            Exit Sub
        End If
            
        
        .Close
    End With
    
    If cmbPaymentMethod.BoundText = 4 Then
        UpdateCompanyBalance Val(cmbHSS.BoundText), Val(lblTotal.Caption), False, True, True, Val(cmbPaymentMethod.BoundText), txtPaymentMethod.Text
    End If
    
    If chkPrint.Value = 1 Then printBill
    
    Call ClearAddValues
    Call ClearBillValues
    Call fillGrid
    cmbTitle.SetFocus
    SendKeys "{M}"
End Sub

Private Sub printBill()
    Dim temBillPoints As MyBillPoints
    
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    
    Dim CenterX As Long
    Dim FieldX As Long
    Dim NoX As Long
    Dim ValueX As Long
    Dim AllLines() As String
    Dim i As Integer
    Dim temY As Long
    Dim n As Long
    
    With MyFOnt
        .Name = DefaultFont.Name
        .Bold = False
        .Italic = False
        .Size = 9
        .Italic = False
        .Underline = False
    End With
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        temBillPoints = PrintThisBill(txtDisplayBillID.Text, cmbPaymentMethod.Text, cmbTitle.Text & " " & cmbPatient.Text, Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "Roentgents Bills", RadiologyName)
        
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        
        temY = temBillPoints.DY
        
        Printer.CurrentY = temY
               
        Printer.Print
        
        If Trim(txtSerialNo.Text) <> "" Then
            temY = Printer.CurrentY
            PrintingText FieldX, temY, ValueX, 0, "Secession : " & cmbSecession.Text, leftAlign, MyFOnt
            temY = Printer.CurrentY
            PrintingText FieldX, temY, ValueX, 0, "Serial : " & txtSerialNo.Text, leftAlign, MyFOnt
        End If
        
            temY = Printer.CurrentY
            PrintingText FieldX, temY, ValueX, 0, "Service : " & cmbCategory.Text, leftAlign, MyFOnt

'        For i = 1 To gridService.Rows - 1
'            temY = Printer.CurrentY
'            n = i
'            temY = Printer.CurrentY
'            PrintingText FieldX, temY, ValueX, 0, "Service : " & gridService.TextMatrix(i, 2), LeftAlign, MyFOnt
'            temY = Printer.CurrentY
'            PrintingText FieldX, temY, ValueX, 0, "Charge : " & gridService.TextMatrix(i, 4), RightAlign, MyFOnt
'
'        Next
        
        For i = 1 To gridService.Rows - 1
            temY = Printer.CurrentY
            n = i
            'PrintingText FieldX, temY, NoX, 0, CStr(n), rightAlign, MyFOnt
            PrintingText FieldX, temY, ValueX, 0, gridService.TextMatrix(i, 2), leftAlign, MyFOnt
            PrintingText FieldX, temY, ValueX, 0, gridService.TextMatrix(i, 4), rightAlign, MyFOnt
        Next
        
        
        Printer.Print
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "---------", rightAlign, MyFOnt
        temY = Printer.CurrentY
        MyFOnt.Bold = False
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText FieldX, temY, ValueX, 0, "Total", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, lblTotal.Caption, rightAlign, MyFOnt
        MyFOnt.Bold = False
        MyFOnt.Bold = False
        MyFOnt.Size = 9
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "---------", rightAlign, MyFOnt
        Printer.Print
        Printer.Print
        
        temY = temBillPoints.CY
        PrintingText FieldX, temY - 720, ValueX, 0, "Cashier :  " & UserFullName, rightAlign, MyFOnt
        
        Printer.EndDoc
        
    End If
End Sub

Private Sub PrintingText(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, PrintText As String, PrintAlignment As TextAlignment, PrintFont As ReportFont)
    
    If PrintAlignment = leftAlign Then
        Printer.CurrentX = X1
    ElseIf PrintAlignment = rightAlign Then
        Printer.CurrentX = X2 - Printer.TextWidth(PrintText)
    ElseIf PrintAlignment = CentreAlign Then
        Printer.CurrentX = (X1 + X2 / 2) - (Printer.TextWidth(PrintText) / 2)
    Else
        Printer.CurrentX = X1
    End If
    If Y1 <> 0 Then Printer.CurrentY = Y1
    Printer.Font.Name = PrintFont.Name
    Printer.Font.Size = PrintFont.Size
    Printer.Font.Italic = PrintFont.Italic
    Printer.Font.Bold = PrintFont.Bold
    Printer.Font.Underline = PrintFont.Underline
    
    Printer.Print PrintText
End Sub

Private Sub chkForeigner_Click()
    If Val(txtOPDBillID.Text) = 0 Then Exit Sub

    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * " & _
                    "FROM tblPatientService " & _
                    "WHERE Deleted = 0 AND RBillID = " & Val(txtOPDBillID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 10 Then Exit Sub
        While .EOF = False
            If chkForeigner.Value = 1 Then
                !Charge = !Charge * 2
                !ProfessionalCharge = !ProfessionalCharge * 2
                !HospitalCharge = !HospitalCharge * 2
            Else
                !Charge = !Charge / 2
                !ProfessionalCharge = !ProfessionalCharge / 2
                !HospitalCharge = !HospitalCharge / 2
            End If
            .Update
            .MoveNext
        Wend
        If .State = 1 Then .Close
        temSQL = "Select * from tblProfessionalCharges where ForRBillID = " & Val(txtOPDBillID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 10 Then Exit Sub
        While .EOF = False
            If chkForeigner.Value = 1 Then
                !Fee = !Fee * 2
            Else
                !Fee = !Fee / 2
            End If
            .Update
            .MoveNext
        Wend
        .Close
    End With
    Call ClearAddValues
    Call FormatGrid
    Call fillGrid
End Sub

Private Sub cmbCategory_Click(Area As Integer)
    If cmbCategory.Locked = True Then
        MsgBox "You can't change the category once Services are added. Only one service category can be added for one Roentgents service bill. If you want to add a new service category, finish this bill and start a new bill or delete the already entered services."
        gridService.SetFocus
    End If
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty

        cmbSecession.SetFocus
        
'        If cmbSC.Visible = True Then
'            cmbSC.SetFocus
'        Else
'            If txtComments.Visible = True Then
'                txtComments.SetFocus
'            Else
'                txtHospitalCharge.SetFocus
'            End If
'        End If
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub cmbPatient_Click(Area As Integer)
'    If IsNumeric(cmbPatient.BoundText) = False Then
'        txtPtID.Text = Empty
'        Exit Sub
'    Else
'        txtPtID.Text = Val(cmbPatient.BoundText)
'    End If
'    Call FillGrid
    
End Sub

Private Sub cmbCategory_Change()
    If IsNumeric(cmbCategory.BoundText) = False Then Exit Sub
    
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblServiceCategory where ServiceCategoryID = " & Val(cmbCategory.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If chkForeigner.Value = 1 Then
                txtHospitalCharge.Text = Format(!Fee * 2, "0.00")
            Else
                txtHospitalCharge.Text = Format(!Fee, "0.00")
            End If
            If !CanChange = True Then
                txtHospitalCharge.Locked = False
            Else
                txtHospitalCharge.Locked = True
            End If
        End If
        .Close
    End With
    
    With rsSC
        If .State = 1 Then .Close
        temSQL = "Select * from tblServiceSubCategory where   Deleted = 0 AND ForR = 1 AND ServiceCategoryID = " & Val(cmbCategory.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbSC.Visible = True
            lblSC.Visible = True
        Else
            cmbSC.Visible = False
            lblSC.Visible = False
        End If
    End With
    With cmbSC
        Set .RowSource = rsSC
        .ListField = "ServiceSubcategory"
        .BoundColumn = "ServiceSubcategoryID"
        .Text = Empty
    End With
    
    With rsSecession
        If .State = 1 Then .Close
        temSQL = "Select * from tblServiceSecession where  ServiceCategoryID = " & Val(cmbCategory.BoundText) & " AND ServiceSubcategoryID = 0 Order by ServiceSecession"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSecession
        Set .RowSource = rsSecession
        .ListField = "ServiceSecession"
        .BoundColumn = "ServiceSecessionID"
        .Text = Empty
    End With
    
    txtFee1_Change (0)
    txtHospitalCharge_Change
    txtProfessionalCharge_Change
    
    cmbSecession_Change
    
End Sub

Private Sub ClearServiceValues()
    Dim n As Integer
    txtComments.Text = Empty
    txtProfessionalCharge.Text = Empty
    txtHospitalCharge.Text = Empty
    txtCharge.Text = Empty
    txtEditID.Text = Empty
    txtDelID.Text = Empty
    For n = 0 To lblSpeciality1.UBound
        lblSpeciality1(n).Visible = False
        lblSpeciality1(n).Caption = Empty
        cmbStaff1(n).Visible = False
        cmbStaff1(n).Text = Empty
        txtServiceProfessionalChargesID(n).Text = Empty
        txtFee1(n).Visible = False
        txtFee1(n).Text = Empty
        txtSpecialityID(n).Text = Empty
    Next

End Sub

Private Sub cmbPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbCategory.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbPatient.Text = Empty
    End If
End Sub

Private Sub cmbPatient_LostFocus()
    cmbPatient.Text = UCase(cmbPatient.Text)
    Dim rsTem As New ADODB.Recordset
    If IsNumeric(cmbPatient.BoundText) = True Then
        txtPtID.Text = cmbPatient.BoundText
    Else
        With rsTem
            If .State = 1 Then .Close
            temSQL = "Select * from tblPatientMainDetails where PatientID = 0 "
            .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !FirstName = cmbPatient.Text
            !TitleID = Val(cmbTitle.BoundText)
            .Update
            temSQL = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            txtPtID.Text = !NewID
            .Close
        End With
    End If
End Sub

Private Sub cmbPaymentMethod_Change()
    If cmbPaymentMethod.BoundText = 1 Then
        txtComments.Visible = False
        txtComments.Text = Empty
    Else
        txtComments.Visible = True
    End If
    If cmbPaymentMethod.BoundText = 4 Then
        cmbHSS.Visible = True
        lblHSS.Visible = True
    Else
        cmbHSS.Visible = False
        lblHSS.Visible = False
    End If

End Sub

Private Sub cmbPaymentMethod_Click(Area As Integer)
    If cmbPaymentMethod.Locked = True Then
        MsgBox "You can't change the Payment Method once Services are added. If you want to add a new service category, finish this bill and start a new bill or delete the already entered services."
        gridService.SetFocus
    End If

End Sub

Private Sub cmbSC_Change()
    Call ClearServiceValues
    
    If IsNumeric(cmbCategory.BoundText) = False Then Exit Sub
    If IsNumeric(cmbSC.BoundText) = False Then Exit Sub
    If cmbSC.Visible = False Then Exit Sub
    
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblServiceSubCategory where ServiceSubCategoryID = " & Val(cmbSC.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If chkForeigner.Value = 1 Then
                txtHospitalCharge.Text = Format(!Fee * 2, "0.00")
            Else
                txtHospitalCharge.Text = Format(!Fee, "0.00")
            End If
            If !CanChange = True Then
                txtHospitalCharge.Locked = False
            Else
                txtHospitalCharge.Locked = True
            End If
        End If
        .Close
    End With
    
    Dim n As Integer
    
    With rsSPC
        If .State = 1 Then .Close
        temSQL = "SELECT Top 7 tblSpeciality.Speciality, tblSpeciality.SpecialityID, tblServiceProfessionalCharges.Fee,  tblServiceProfessionalCharges.StaffID, tblServiceProfessionalCharges.ServiceProfessionalChargesID " & _
                    "FROM tblSpeciality RIGHT JOIN tblServiceProfessionalCharges ON tblSpeciality.SpecialityID = tblServiceProfessionalCharges.SpecialityID " & _
                    "Where (((tblServiceProfessionalCharges.ServiceSubcategoryID) = " & Val(cmbSC.BoundText) & ") AND ((tblServiceProfessionalCharges.Deleted)=0 ))" & _
                    "ORDER BY tblServiceProfessionalCharges.ServiceProfessionalChargesID DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        PSCCount = .RecordCount
        ReDim rsStaff(.RecordCount)
        For n = 0 To PSCCount - 1
            lblSpeciality1(n).Visible = True
            lblSpeciality1(n).Caption = !Speciality
            
            txtServiceProfessionalChargesID(n).Text = !ServiceProfessionalChargesID
            txtSpecialityID(n).Text = !SpecialityID
            
            cmbStaff1(n).Visible = True
            If rsStaff(n).State = 1 Then rsStaff(n).Close
            temSQL = "SELECT tblStaff.Name as TitleStaff, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID Where SpecialityID = " & !SpecialityID & " ORDER BY tblTitle.Title, tblStaff.Name"
            rsStaff(n).Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            Set cmbStaff1(n).RowSource = rsStaff(n)
            cmbStaff1(n).ListField = "TitleStaff"
            cmbStaff1(n).BoundColumn = "StaffID"
            cmbStaff1(n).BoundText = !StaffID
        
            txtFee1(n).Visible = True
            If chkForeigner.Value = 1 Then
                txtFee1(n).Text = Format(!Fee * 2, "0.00")
            Else
                txtFee1(n).Text = Format(!Fee, "0.00")
            End If
            .MoveNext
            
        Next
        If PSCCount = 0 Then
            For n = 0 To lblSpeciality1.UBound
                lblSpeciality1(n).Visible = False
                lblSpeciality1(n).Caption = Empty
                cmbStaff1(n).Visible = False
                cmbStaff1(n).Text = Empty
                txtServiceProfessionalChargesID(n).Text = Empty
                txtFee1(n).Visible = False
                txtFee1(n).Text = Empty
                txtSpecialityID(n).Text = Empty
            Next
        Else
            For n = PSCCount To lblSpeciality1.UBound
                lblSpeciality1(n).Visible = False
                lblSpeciality1(n).Caption = Empty
                cmbStaff1(n).Visible = False
                cmbStaff1(n).Text = Empty
                txtServiceProfessionalChargesID(n).Text = Empty
                txtFee1(n).Visible = False
                txtFee1(n).Text = Empty
                txtSpecialityID(n).Text = Empty
            Next
        End If
    End With
    
'    With rsSecession
'        If .State = 1 Then .Close
'        temSql = "Select * from tblServiceSecession where ServiceSubCategoryID = " & Val(cmbSC.BoundText) & " Order by ServiceSecession"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With cmbSecession
'        Set .RowSource = rsSecession
'        .ListField = "ServiceSecession"
'        .BoundColumn = "ServiceSecessionID"
'        .Text = Empty
'    End With
    
    txtFee1_Change (0)
    txtHospitalCharge_Change
    txtProfessionalCharge_Change
        
    
    cmbSecession_Change
    
End Sub

Private Sub cmbSC_Click(Area As Integer)
    If cmbSC.Locked = True Then
        MsgBox "You can't change the Secession once Services are added. If you want to change the secession, finish this bill and start a new bill or delete the already entered services."
        gridService.SetFocus
    End If
End Sub

Private Sub cmbSC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtHospitalCharge.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbSC.SetFocus
    End If
End Sub

Private Sub cmbSecession_Change()
    Dim rsTem As New ADODB.Recordset
    If cmbSecession.Text = Empty Then Exit Sub
'    If IsNumeric(txtSerialNo.Text) = True Then Exit Sub
    If cmbSecession.Locked = True Then Exit Sub
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select Max(tblPatientService.SerialNo) as MaxSerialNo from tblPatientService Where Deleted = 0 AND ServiceDate = '" & Format(dtpAppDate.Value, "dd MMMM yyyy") & "' AND ServiceCategoryID = " & Val(cmbCategory.BoundText) & " and SECESSIONID  = " & Val(cmbSecession.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If IsNull(!MaxSerialNo) = False Then
            txtSerialNo.Text = !MaxSerialNo + 1
        Else
            txtSerialNo.Text = 1
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblServiceSecession Where ServiceSecessionID = " & Val(cmbSecession.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtDuration.Text = !DurationMinutes
            dtpStart.Value = !StartTime
        Else
            txtDuration.Text = 0
        End If
        .Close
    End With
    dtpAppTime.Value = DateAdd("n", Val(txtDuration.Text) + Val(txtSerialNo.Text) - 1, dtpStart.Value)
End Sub

Private Sub cmbSecession_Click(Area As Integer)
    If cmbSecession.Locked = True Then
        MsgBox "You can't change the Secession once Services are added. If you want to change the secession, finish this bill and start a new bill or delete the already entered services."
        gridService.SetFocus
    End If

End Sub

Private Sub cmbSecession_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If cmbSC.Visible = True Then
            cmbSC.SetFocus
        Else
            If txtComments.Visible = True Then
                txtComments.SetFocus
            Else
                txtHospitalCharge.SetFocus
            End If
        End If
        
'        If txtComments.Visible = True Then
'            txtComments.SetFocus
'        Else
'            txtHospitalCharge.SetFocus
'        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbSecession.Text = Empty
    End If
End Sub

Private Sub cmbStaff1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtFee1(Index).SetFocus
    ElseIf KeyCode = vbKeyEscape Then
'        cmbStaff1(Index).Text = Empty
    End If
End Sub

Private Sub cmbTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPatient.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbTitle.Text = Empty
    End If
End Sub

Private Sub Form_Activate()
    If FirstActi = True Then
        Call GetSettings
        Me.Caption = Me.Caption & RadiologyName
        FirstActi = False
    End If
End Sub

Private Sub Form_Load()
    FirstActi = True
    Call PopulatePrinters
    Call PopulatePapers
    Call GetSettings
    Call FillCombos
    Call ClearServiceValues
    Call ClearAddValues
    Call ClearBillValues
    If IsNumeric(txtOPDBillID.Text) = False Then txtOPDBillID.Text = NewRBillID(dtpDate.Value, dtpTime.Value)
    Call fillGrid
    
End Sub

Private Sub GetSettings(): On Error Resume Next
    GetCommonSettings Me
    dtpDate.Value = Date
    dtpTime.Value = Time
    dtpAppDate.Value = Date
    cmbPaymentMethod.BoundText = 1 ' Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
    chkPrint.Value = GetSetting(App.EXEName, Me.Name, chkPrint.Name, 1)
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCategory, "ServiceCategory", "ServiceCategory", "ForR", True
    Dim BHT As New clsFillCombos
    ''BHT.FillSpecificIDField cmbPatient, "PatientMainDetails", "PatientID", "FirstName", False
    Dim PayM As New clsFillCombos
    PayM.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanReceive", False
    Dim Title As New clsFillCombos
    Title.FillAnyCombo cmbTitle, "Title", False
    With rsHSS
        If .State = 1 Then .Close
        temSQL = "SELECT tblHealthSchemeSuppliers.* FROM tblHealthSchemeSuppliers ORDER BY tblHealthSchemeSuppliers.HealthSchemeSupplierName"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbHSS
        Set .RowSource = rsHSS
        .ListField = "HealthSchemeSupplierName"
        .BoundColumn = "HealthSchemeSupplierID"
    End With
    
End Sub

Private Sub fillGrid()
    Call FormatGrid
    Dim rsTem As New ADODB.Recordset
    Dim TotalCharge As Double
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT tblPatientService.PatientServiceID, tblPatientService.ServiceDate, tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory, tblPatientService.Comments, tblPatientService.Charge, tblServiceSecession.ServiceSecession, tblPatientService.SerialNo " & _
                    "FROM ((tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID) LEFT JOIN tblServiceSecession ON tblPatientService.SecessionID = tblServiceSecession.ServiceSecessionID " & _
                    "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.RBillID)<> 0)  AND ((tblPatientService.RBillID)=" & Val(txtOPDBillID.Text) & ")) " & _
                    "ORDER BY tblPatientService.PatientServiceID"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridService.Rows = gridService.Rows + 1
            gridService.Row = gridService.Rows - 1
            gridService.Col = 0
            gridService.Text = !PatientServiceID
            gridService.Col = 1
            gridService.Text = !ServiceDate
            gridService.Col = 2
            If IsNull(!ServiceSubcategory) = True Then
                gridService.Text = !ServiceCategory
            Else
                gridService.Text = !ServiceSubcategory   '  !ServiceCategory & " - " & !ServiceSubCategory
            End If
            gridService.Col = 3
            gridService.Text = !Comments
            gridService.Col = 4
            gridService.Text = Format(!Charge, "0.00")
            TotalCharge = TotalCharge + !Charge
            
            gridService.Col = 5
            gridService.Text = Format(!ServiceSecession, "")
            gridService.Col = 6
            gridService.Text = Format(!SerialNo, "")
            
            
            .MoveNext
        Wend
    End With
    lblTotal.Caption = Format(TotalCharge, "0.00")
End Sub

Private Sub FormatGrid()
    '   0   ID
    '   1   Date
    '   2   Service
    '   3   Comments
    '   4   Charges
    '   5   Secession
    '   6   Serial
    
    With gridService
        .Cols = 7
        .Rows = 1
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        .ColWidth(0) = 0
        
        
        .Col = 1
        .ColWidth(1) = 0
        .Text = "Date"
        
        .Col = 2
        .ColWidth(2) = 2500
        .Text = "Service"
        
        .Col = 3
        .ColWidth(3) = 2500
        .Text = "Comments "
        
        .Col = 4
        .ColWidth(4) = 1200
        .Text = "Charge"
        
        .Col = 5
        .ColWidth(4) = 1200
        .Text = "Secession"
        
        .Col = 6
        .ColWidth(4) = 1200
        .Text = "Serial"
        
        
        
    End With
    lblTotal.Caption = "0.00"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ClearGrid
    Call SaveSettings
End Sub

Private Sub ClearGrid()
    With gridService
        While .Rows >= 2
            .Row = .Rows - 1
            txtDelID.Text = Val(.TextMatrix(.Row, 0))
            btnDelete_Click
        Wend
    End With
End Sub

Private Sub gridService_Click()
    With gridService
        txtDelID.Text = Val(.TextMatrix(.Row, 0))
        .Col = .Cols - 1
        .ColSel = 0
    End With
End Sub

Private Sub gridService_DblClick()
'    Dim rsTem As New ADODB.Recordset
'    With gridService
'        txtEditID.Text = Val(.TextMatrix(.Row, 0))
'        .Col = .Cols - 1
'        .ColSel = 0
'    End With
'    With rsTem
'        If .State = 1 Then .Close
'        If IsNumeric(txtEditID.Text) = True Then
'            temSql = "Select * from tblPatientService where PatientServiceID = " & Val(txtEditID.Text)
'            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'            If .RecordCount > 0 Then
'                cmbCategory.BoundText = !ServiceCategoryID
'                cmbSC.BoundText = !ServiceSubcategoryID
'                txtComments.Text = !Comments
'                dtpDate.Value = !ServiceDate
'                dtpTime.Value = !ServiceTime
'                txtCharge.Text = Format(!Charge, "0.00")
'                txtHospitalCharge.Text = Format(!HospitalCharge, "0.00")
'                txtProfessionalCharge.Text = Format(!ProfessionalCharge, "0.00")
'            End If
'            .Close
'        End If
'    End With
'    Dim n As Integer
'    For n = 0 To lblSpeciality1.UBound
'        If lblSpeciality1(n).Visible = True Then
'            With rsTem
'                If .State = 1 Then .Close
'                temSql = "Select * from tblProfessionalCharges where ServiceProfessionalChargesID = " & Val(txtServiceProfessionalChargesID(n).Text) & " AND PatientServiceID = " & Val(txtEditID.Text)
'                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                If .RecordCount > 0 Then
'                    cmbStaff1(n).BoundText = !StaffID
'                    If chkForeigner.Value = 1 Then
'                        txtFee1(n).Text = Format(!Fee * 2, "0.00")
'                    Else
'                        txtFee1(n).Text = Format(!Fee, "0.00")
'                    End If
'                End If
'                .Close
'            End With
'        Else
'            txtFee1(n).Text = 0
'        End If
'    Next n
End Sub


Private Sub txtCharge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If cmbStaff1(0).Visible = True Then
            cmbStaff1(0).SetFocus
        Else
            btnAdd_Click
        End If
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtCharge.Text = Empty
    End If
End Sub

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtHospitalCharge.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtComments.Text = Empty
    End If
End Sub

Private Sub txtFee1_Change(Index As Integer)
    Dim n As Long
    Dim temTotal As Double
    For n = 0 To txtFee1.UBound
        temTotal = temTotal + Val(txtFee1(n).Text)
    Next
    txtProfessionalCharge.Text = Format(temTotal, "0.00")
End Sub

Private Sub txtFee1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Index >= txtFee1.UBound Then
            If cmbStaff1(Index + 1).Visible = True Then
                cmbStaff1(Index + 1).SetFocus
            Else
                btnAdd_Click
            End If
        Else
            btnAdd_Click
        End If
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        
    End If
End Sub

Private Sub txtHospitalCharge_Change()
    txtCharge.Text = Format(Val(txtHospitalCharge.Text) + Val(txtProfessionalCharge.Text), "0.00")
End Sub

Private Sub txtHospitalCharge_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtHospitalCharge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtProfessionalCharge.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtHospitalCharge.Text = Empty
    End If
End Sub

Private Sub txtProfessionalCharge_Change()
    txtCharge.Text = Format(Val(txtHospitalCharge.Text) + Val(txtProfessionalCharge.Text), "0.00")
End Sub

Private Sub txtProfessionalCharge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtCharge.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtProfessionalCharge.Text = Empty
    End If
End Sub


Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub PopulatePapers(): On Error Resume Next
    cmbPaper.Clear
    SetPrinter = False
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
'        With FormSize
'            .cx = BillPaperHeight
'            .cy = BillPaperWidth
'        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                cmbPaper.AddItem PtrCtoVbString(.pName)
'                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle): DoEvents
    End If
End Sub

Private Sub cmbPrinter_Change()
    Call PopulatePapers
End Sub

Private Sub cmbPrinter_Click()
    Call PopulatePapers
End Sub
