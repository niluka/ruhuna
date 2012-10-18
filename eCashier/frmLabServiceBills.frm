VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLabServiceBills 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laboratory Bills"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12900
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
   ScaleHeight     =   8985
   ScaleWidth      =   12900
   Begin VB.TextBox txtDisplayBillID 
      Height          =   375
      Left            =   8640
      TabIndex        =   78
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkForeigner 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Foreigner"
      Height          =   255
      Left            =   2040
      TabIndex        =   75
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Height          =   735
      Left            =   120
      TabIndex        =   70
      Top             =   8160
      Width           =   12615
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   210
         Width           =   4695
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   7680
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Paper"
         Height          =   255
         Left            =   6000
         TabIndex        =   73
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtPaymentMethod 
      Height          =   375
      Left            =   9000
      TabIndex        =   61
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox txtPtID 
      Height          =   375
      Left            =   9120
      TabIndex        =   67
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkPrint 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Print"
      Height          =   255
      Left            =   8880
      TabIndex        =   51
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox txtLabBillID 
      Height          =   375
      Left            =   8160
      TabIndex        =   66
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   6
      Left            =   7920
      TabIndex        =   47
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   5
      Left            =   7920
      TabIndex        =   42
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   4
      Left            =   7920
      TabIndex        =   37
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   3
      Left            =   7920
      TabIndex        =   32
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   2
      Left            =   7920
      TabIndex        =   27
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   6
      Left            =   8400
      TabIndex        =   48
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   43
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   4
      Left            =   8400
      TabIndex        =   38
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   33
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   2
      Left            =   8400
      TabIndex        =   28
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   23
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtCharge 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3240
      Width           =   4575
   End
   Begin VB.TextBox txtProfessionalCharge 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2760
      Width           =   4575
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   6
      Left            =   11520
      TabIndex        =   50
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   5
      Left            =   11520
      TabIndex        =   45
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   4
      Left            =   11520
      TabIndex        =   40
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   3
      Left            =   11520
      TabIndex        =   35
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   11520
      TabIndex        =   30
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   11520
      TabIndex        =   25
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   11520
      TabIndex        =   20
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   0
      Left            =   8880
      TabIndex        =   19
      Top             =   4200
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
      Height          =   375
      Left            =   8160
      TabIndex        =   64
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtDelID 
      Height          =   360
      Left            =   8640
      TabIndex        =   62
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   7920
      TabIndex        =   55
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   20578307
      CurrentDate     =   39956
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11520
      TabIndex        =   53
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   16711935
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
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711935
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
      Height          =   3735
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6588
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtHospitalCharge 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2280
      Width           =   4575
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
      Left            =   9000
      TabIndex        =   65
      Top             =   2760
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSC 
      Height          =   1320
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2328
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711935
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
      TabIndex        =   52
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   16711935
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
      TabIndex        =   24
      Top             =   4680
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
      TabIndex        =   29
      Top             =   5160
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
      TabIndex        =   34
      Top             =   5640
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
      TabIndex        =   39
      Top             =   6120
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
      TabIndex        =   44
      Top             =   6600
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
      TabIndex        =   49
      Top             =   7080
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
      Left            =   11160
      TabIndex        =   57
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   20578306
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
      Left            =   9000
      TabIndex        =   59
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtComments 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   4575
   End
   Begin MSDataListLib.DataCombo cmbHSS 
      Height          =   360
      Left            =   9000
      TabIndex        =   76
      Top             =   1080
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblHSS 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Credit Company"
      Height          =   255
      Left            =   6960
      TabIndex        =   77
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Payment Co&mments"
      Height          =   255
      Left            =   6960
      TabIndex        =   60
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Payment &Method"
      Height          =   255
      Left            =   6960
      TabIndex        =   58
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Time"
      Height          =   255
      Left            =   10440
      TabIndex        =   56
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Date"
      Height          =   255
      Left            =   6960
      TabIndex        =   54
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Total Charge"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pro&fessional Charge"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   46
      Top             =   7560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   5
      Left            =   6720
      TabIndex        =   41
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   36
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   31
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   26
      Top             =   5160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   21
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   0
      Left            =   6720
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
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
      Height          =   255
      Left            =   5040
      TabIndex        =   69
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
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
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Hospital Charge"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblSC 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Investigation"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Service Cate&gory"
      Height          =   255
      Left            =   6960
      TabIndex        =   63
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pa&tient"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmLabServiceBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsSC As New ADODB.Recordset
    Dim temSql As String
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

Private Sub chkForeigner_Click()
    If Val(txtLabBillID.Text) = 0 Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT * " & _
                    "FROM tblPatientService " & _
                    "WHERE Deleted = 0 AND LabBillID = " & Val(txtLabBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        
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
        temSql = "Select * from tblProfessionalCharges where ForLabBillID = " & Val(txtLabBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        
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

Private Sub cmbPatient_Change()
'    cmbPatient.Text = UCase(cmbPatient.Text)
'    SendKeys "{end}"
End Sub

Private Sub btnAdd_Click()
    'If IsNumeric(txtLabBillID.Text) = False Then txtLabBillID.Text = NewLabBillID(dtpDate.Value, dtpTime.Value) ' 2010 04 03
    
    If IsNumeric(txtLabBillID.Text) = False Then txtLabBillID.Text = NewLabBillID(Date, Time)
    
    Dim rsTem As New ADODB.Recordset
   
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
    
    If IsNumeric(cmbSC.BoundText) = False Then
        MsgBox "Investigation?"
        cmbSC.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtPtID.Text) = False Then
        With rsTem
            If .State = 1 Then .Close
            temSql = "Select * from tblPatientMainDetails where PatientID = 0 "
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !FirstName = cmbPatient.Text
            !TitleID = Val(cmbTitle.BoundText)
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            txtPtID.Text = !NewID
            .Close
        End With
    End If
    
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtEditID.Text) = True Then
            temSql = "Select * from tblPatientService where PatientServiceID = " & Val(txtEditID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount <= 0 Then
                .AddNew
            End If
        Else
            temSql = "Select * from tblPatientService  where PatientServiceID = 0 "
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
        End If
        !LABBillID = Val(txtLabBillID.Text)
        !ServiceCategoryID = Val(cmbCategory.BoundText)
        !ServicesubcategoryID = Val(cmbSC.BoundText)
        !Comments = txtComments.Text
        !ServiceDate = dtpDate.Value
        !ServiceTime = dtpTime.Value
        !Charge = Val(txtCharge.Text)
        !ProfessionalCharge = Val(txtProfessionalCharge.Text)
        !HospitalCharge = Val(txtHospitalCharge.Text)
        !UserID = UserID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        txtEditID.Text = !NewID
    End With
    For n = 0 To lblSpeciality1.UBound
        If lblSpeciality1(n).Visible = True Then
            With rsTem
                If .State = 1 Then .Close
                temSql = "Select * from tblProfessionalCharges where ServiceProfessionalChargesID = " & Val(txtServiceProfessionalChargesID(n).Text) & " AND PatientServiceID = " & Val(txtEditID.Text)
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    
                Else
                    .AddNew
                    !UserID = UserID
                    !ForLABBILLID = Val(txtLabBillID.Text)
                    !PatientServiceID = Val(txtEditID.Text)
                    !ServiceProfessionalChargesID = Val(txtServiceProfessionalChargesID(n).Text)
                    !StaffID = Val(cmbStaff1(n).BoundText)
                End If
                !Date = dtpDate.Value
                !Time = dtpTime.Value
                !Fee = Val(txtFee1(n).Text)
                !IsLabBill = True
                .Update
            End With
        End If
    Next n
    Call fillGrid
    Call ClearAddValues
    cmbSC.SetFocus
End Sub

Private Sub ClearAddValues()
    Dim n As Long
    'cmbSC.Text = Empty
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
    cmbPatient.Text = Empty
    txtLabBillID.Text = Empty
    txtPtID.Text = Empty
    lblTotal.Caption = "0.00"
    txtPaymentMethod.Text = Empty
    cmbTitle.Text = Empty
    chkForeigner.Value = 0
    cmbPaymentMethod.BoundText = 1
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
            temSql = "Select * from tblPatientService where PatientServiceID = " & Val(txtDelID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
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
        temSql = "Select * from tblProfessionalCharges where PatientServiceID = " & Val(txtDelID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
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
    cmbSC.SetFocus
End Sub

Private Sub btnUpdate_Click()
    If gridService.Rows < 2 Then
        MsgBox "Nothing to update"
        cmbSC.SetFocus
        Exit Sub
    End If

    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Please select a payment method"
        cmbPaymentMethod.SetFocus
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
    
    Dim rsTem As New ADODB.Recordset
    Dim DisplayBillID As Long
    
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "Select count(IncomeBillID) as BillCount from tblIncomeBill where IsLabBill = 1AND Completed = 1 "
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
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
    
    DisplayBillID = NewLabDisplayBillID(Val(txtLabBillID.Text))
    
    txtDisplayBillID.Text = DisplayBillID
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where IncomeBillID = " & Val(txtLabBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Completed = True
            !CompletedDate = Date
            !CompletedTime = Now
            !CompletedUserID = UserID
            !DisplayBillID = DisplayBillID
            !PatientID = Val(txtPtID.Text)
            !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
            !PaymentComments = txtPaymentMethod.Text
            !IsLabBill = True
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
    
    cmbSC_Change
    
    cmbTitle.SetFocus
    cmbTitle.BoundText = 25
    
    
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If cmbSC.Visible = True Then
            cmbSC.SetFocus
        Else
            txtComments.SetFocus
        End If
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
        temSql = "Select * from tblServiceCategory where ServiceCategoryID = " & Val(cmbCategory.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
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
        temSql = "Select * from tblServiceSubCategory where  Deleted = 0 AND ForLab = 1 AND ServiceCategoryID = " & Val(cmbCategory.BoundText) & " Order By ServiceSubCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
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
    End With
    
    
    
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
        cmbSC.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbPatient.Text = Empty
    End If
End Sub

Private Sub cmbPatient_LostFocus()
    cmbPatient.Text = UCase(cmbPatient.Text)
    SendKeys "{end}"

End Sub

Private Sub cmbPaymentMethod_Click(Area As Integer)
    If cmbPaymentMethod.BoundText = 4 Then
        cmbHSS.Visible = True
        lblHSS.Visible = True
    Else
        cmbHSS.Visible = False
        lblHSS.Visible = False
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
        temSql = "Select * from tblServiceSubCategory where ServiceSubCategoryID = " & Val(cmbSC.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
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
        temSql = "SELECT Top 7 tblSpeciality.Speciality, tblSpeciality.SpecialityID, tblServiceProfessionalCharges.Fee,  tblServiceProfessionalCharges.StaffID, tblServiceProfessionalCharges.ServiceProfessionalChargesID " & _
                    "FROM tblSpeciality RIGHT JOIN tblServiceProfessionalCharges ON tblSpeciality.SpecialityID = tblServiceProfessionalCharges.SpecialityID " & _
                    "Where (((tblServiceProfessionalCharges.ServiceSubcategoryID) = " & Val(cmbSC.BoundText) & ") AND ((tblServiceProfessionalCharges.Deleted)=0 ))" & _
                    "ORDER BY tblServiceProfessionalCharges.ServiceProfessionalChargesID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        PSCCount = .RecordCount
        ReDim rsStaff(.RecordCount)
        For n = 0 To PSCCount - 1
            lblSpeciality1(n).Visible = True
            lblSpeciality1(n).Caption = !Speciality
            
            txtServiceProfessionalChargesID(n).Text = !ServiceProfessionalChargesID
            txtSpecialityID(n).Text = !SpecialityID
            
            cmbStaff1(n).Visible = True
            If rsStaff(n).State = 1 Then rsStaff(n).Close
            temSql = "SELECT tblStaff.Name as TitleStaff, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID Where SpecialityID = " & !SpecialityID & " ORDER BY tblTitle.Title, tblStaff.Name"
            rsStaff(n).Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            Set cmbStaff1(n).RowSource = rsStaff(n)
            cmbStaff1(n).ListField = "TitleStaff"
            cmbStaff1(n).BoundColumn = "StaffID"
            cmbStaff1(n).BoundText = !StaffID
        
            txtFee1(n).Visible = True
            txtFee1(n).Text = Format(!Fee, "0.00")
            
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
    
    
End Sub

Private Sub cmbSC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If txtComments.Visible = True Then
            txtComments.SetFocus
        Else
            txtHospitalCharge.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbSC.SetFocus
    End If
End Sub

Private Sub cmbStaff1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtFee1(Index).SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbStaff1(Index).Text = Empty
    End If
End Sub

Private Sub cmbTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPatient.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbTitle.Text = Empty
    End If
End Sub

Private Sub Form_Activate()
    If FirstActi = True Then
        Call GetSettings
        FirstActi = False
    End If
End Sub

Private Sub Form_Load()
    FirstActi = True
    Call FillCombos
    Call PopulatePrinters
    Call ClearServiceValues
    Call ClearAddValues
    Call ClearBillValues
    ' If IsNumeric(txtLabBillID.Text) = False Then txtLabBillID.Text = NewLabBillID(dtpDate.Value, dtpTime.Value) ' 2010 04 03
    
    If IsNumeric(txtLabBillID.Text) = False Then txtLabBillID.Text = NewLabBillID(Date, Time) ' 2010 04 03
    
    Call fillGrid
    
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    dtpTime.Value = Time
    cmbPaymentMethod.BoundText = 1 ' Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
    chkPrint.Value = Val(GetSetting(App.EXEName, Me.Name, chkPrint.Name, 0))
    cmbCategory.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbCategory.Name, 1))
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
    GetCommonSettings Me
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, cmbCategory.Name, cmbCategory.BoundText
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    SaveCommonSettings Me
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCategory, "ServiceCategory", "ServiceCategory", "ForLab", True
    Dim BHT As New clsFillCombos
    ''BHT.FillSpecificIDField cmbPatient, "PatientMainDetails", "PatientID", "FirstName", False
    Dim PayM As New clsFillCombos
    PayM.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanReceive", False
    Dim Title As New clsFillCombos
    Title.FillAnyCombo cmbTitle, "Title", False
    With rsHSS
        If .State = 1 Then .Close
        temSql = "SELECT tblHealthSchemeSuppliers.* FROM tblHealthSchemeSuppliers ORDER BY tblHealthSchemeSuppliers.HealthSchemeSupplierName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
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
        temSql = "SELECT tblPatientService.PatientServiceID, tblPatientService.ServiceDate, tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory, tblPatientService.Comments, tblPatientService.Charge " & _
                    "FROM (tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID " & _
                    "WHERE (((tblPatientService.Deleted)=0)  AND ((tblPatientService.LABBillID)<>0) AND ((tblPatientService.LABBillID)=" & Val(txtLabBillID.Text) & ")) " & _
                    "ORDER BY tblPatientService.PatientServiceID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridService.Rows = gridService.Rows + 1
            gridService.Row = gridService.Rows - 1
            gridService.Col = 0
            gridService.Text = !PatientServiceID
            gridService.Col = 1
            gridService.Text = !ServiceDate
            gridService.Col = 2
            If IsNull(!ServiceSubcategory) = False Then
                gridService.Text = !ServiceSubcategory
            End If
            gridService.Col = 3
            gridService.Text = !Comments
            gridService.Col = 4
            gridService.Text = Format(!Charge, "0.00")
            TotalCharge = TotalCharge + !Charge
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
    With gridService
        .Cols = 5
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
'                    txtFee1(n).Text = Format(!Fee, "0.00")
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
        temBillPoints = PrintThisBill(txtDisplayBillID.Text, cmbPaymentMethod.Text, cmbTitle.Text & " " & cmbPatient.Text, Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "Laboratory Bills", LabName)
        
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        
        temY = temBillPoints.DY
        
        Printer.CurrentY = temY
                
        Printer.Print
        
        
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
        
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, "..........", rightAlign, MyFOnt
        
        temY = temBillPoints.CY
        PrintingText FieldX, temY, ValueX, 0, "Cashier :  " & UserFullName, rightAlign, MyFOnt
        
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
