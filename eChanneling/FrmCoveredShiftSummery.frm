VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmCoveredShiftSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift Summery"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCoveredShiftSummery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   8445
   Begin VB.Frame FrameShiftSummary 
      Height          =   3855
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   7815
      Begin btButtonEx.ButtonEx bttnPrintSummary 
         Height          =   375
         Left            =   5880
         TabIndex        =   2
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
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Cash Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   2295
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Refunds / Cancellations "
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Income"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Expenses"
         Height          =   255
         Left            =   4440
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   5760
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   5760
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Cash Collection"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   2175
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
         TabIndex        =   9
         Top             =   480
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
         TabIndex        =   8
         Top             =   2040
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
         TabIndex        =   7
         Top             =   2520
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
         TabIndex        =   6
         Top             =   3120
         Width           =   1455
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
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Cettling for Credit Patients "
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
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
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
   End
   Begin MSDataListLib.DataCombo DataComboUser 
      Bindings        =   "FrmCoveredShiftSummery.frx":0442
      Height          =   360
      Left            =   1560
      TabIndex        =   0
      Top             =   240
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   240
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   8916
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Toda&y"
      TabPicture(0)   =   "FrmCoveredShiftSummery.frx":0461
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblToday"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Selected Day"
      TabPicture(1)   =   "FrmCoveredShiftSummery.frx":047D
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DTPicker1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "S&elected Period"
      TabPicture(2)   =   "FrmCoveredShiftSummery.frx":0499
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label18"
      Tab(2).Control(1)=   "Label17"
      Tab(2).Control(2)=   "DTPicker3"
      Tab(2).Control(3)=   "DTPicker2"
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   161742851
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -74040
         TabIndex        =   19
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   161742851
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   -70560
         TabIndex        =   20
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   161742851
         CurrentDate     =   39442
      End
      Begin VB.Label lblToday 
         Alignment       =   2  'Center
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
         TabIndex        =   24
         Top             =   600
         Width           =   3975
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
         Left            =   960
         TabIndex        =   23
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label17 
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
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label18 
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
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   6840
      TabIndex        =   25
      Top             =   6360
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
      Left            =   480
      TabIndex        =   28
      Top             =   120
      Width           =   1215
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
      Left            =   1560
      TabIndex        =   27
      Top             =   120
      Width           =   5655
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
      Left            =   1560
      TabIndex        =   26
      Top             =   720
      Width           =   5655
   End
End
Attribute VB_Name = "FrmCoveredShiftSummery"
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
Dim csetPriner As New cSetDfltPrinter

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

'        bttn.BackColor = BttnBackColour
'        bttnCancelRepay.ForeColor = BttnForeColour
'
'        bttnConfirmRepay.BackColor = BttnBackColour
'        bttnConfirmRepay.ForeColor = BttnForeColour
'
'        bttnSearchBookingID.BackColor = BttnBackColour
'        bttnSearchBookingID.ForeColor = BttnForeColour
'
        bttnClose.BackColor = BttnBackColour
        bttnClose.ForeColor = BttnForeColour
'
'        bttnEdit.BackColor = BttnBackColour
'        bttnEdit.ForeColor = BttnForeColour
'
'        bttnSave.BackColor = BttnBackColour
'        bttnSave.ForeColor = BttnForeColour
        
        FrmCoveredShiftSummery.BackColor = FrmBackColour
        FrmCoveredShiftSummery.ForeColor = FrmForeColour
        
        FrameShiftSummary.BackColor = FrmBackColour
        FrameShiftSummary.ForeColor = FrmForeColour
'
'        FrameFacilityDetails.BackColor = FrmBackColour
'        FrameFacilityDetails.ForeColor = FrmForeColour
'
'        frameRepay.BackColor = FrmBackColour
'        frameRepay.ForeColor = FrmForeColour
        
        
'        chkHLine2.BackColor = LblBackColour
'        chkHLine2.ForeColor = LblForeColour
'
'        chkHLine3.BackColor = LblBackColour
'        chkHLine3.ForeColor = LblForeColour
'
'        chkHLine4.BackColor = LblBackColour
'        chkHLine4.ForeColor = LblForeColour


'        OptionTwo.BackColor = FrmBackColour
'        OptionTwo.ForeColor = FrmForeColour
'
'        OptionOne.BackColor = FrmBackColour
'        OptionOne.ForeColor = FrmForeColour
'
'        OptionNo.BackColor = FrmBackColour
'        OptionNo.ForeColor = FrmForeColour
        
        
        
        
        
'        Label1.BackColor = LblBackColour
'        Label1.ForeColor = LblForeColour
'
'        Label10.BackColor = LblBackColour
'        Label10.ForeColor = LblForeColour
'        Label11.BackColor = LblBackColour
'        Label11.ForeColor = LblForeColour
'        Label12.BackColor = LblBackColour
'        Label12.ForeColor = LblForeColour
'        Label13.BackColor = LblBackColour
'        Label13.ForeColor = LblForeColour
'        Label14.BackColor = LblBackColour
'        Label14.ForeColor = LblForeColour
'        Label15.BackColor = LblBackColour
'        Label15.ForeColor = LblForeColour
'        Label2.BackColor = LblBackColour
'        Label2.ForeColor = LblForeColour
'        Label3.BackColor = LblBackColour
'        Label3.ForeColor = LblForeColour
'        Label4.BackColor = LblBackColour
'        Label4.ForeColor = LblForeColour
'        Label4.BackColor = LblBackColour
'        Label4.ForeColor = LblForeColour
'        Label5.BackColor = LblBackColour
'        Label5.ForeColor = LblForeColour
'        Label6.BackColor = LblBackColour
'        Label6.ForeColor = LblForeColour
'        Label7.BackColor = LblBackColour
'        Label7.ForeColor = LblForeColour
'
'        Label8.BackColor = LblBackColour
'        Label8.ForeColor = LblForeColour
'        Label9.BackColor = LblBackColour
'        Label9.ForeColor = LblForeColour
        
    
        
'        txtAccount.BackColor = TxtBackColour
'        txtAccount.ForeColor = TxtForeColour
'
'        txtBankBranch.BackColor = TxtBackColour
'        txtBankBranch.ForeColor = TxtForeColour
'
'        txtComments.BackColor = TxtBackColour
'        txtComments.ForeColor = TxtForeColour
'        txtCredit.BackColor = TxtBackColour
'        txtCredit.ForeColor = TxtForeColour
'        txtDesignation.BackColor = TxtBackColour
'        txtDesignation.ForeColor = TxtForeColour
'        txtListedName.BackColor = TxtBackColour
'        txtListedName.ForeColor = TxtForeColour
'        txtName.BackColor = TxtBackColour
'        txtName.ForeColor = TxtForeColour
'        txtOfficialAddress.BackColor = TxtBackColour
'        txtOfficialAddress.ForeColor = TxtForeColour
'        txtOfficialEMail.BackColor = TxtBackColour
'        txtOfficialEMail.ForeColor = TxtForeColour
'        txtOfficialFax.BackColor = TxtBackColour
'        txtOfficialFax.ForeColor = TxtForeColour
'        txtOfficialTel.BackColor = TxtBackColour
'        txtOfficialTel.ForeColor = TxtForeColour
'        txtOfficialWebsite.BackColor = TxtBackColour
'        txtOfficialWebsite.ForeColor = TxtForeColour
'
'        txtPrivateAddress.BackColor = TxtBackColour
'        txtPrivateAddress.ForeColor = TxtForeColour
'        txtPrivateEmail.BackColor = TxtBackColour
'        txtPrivateEmail.ForeColor = TxtForeColour
'        txtPrivateFax.BackColor = TxtBackColour
'        txtPrivateFax.ForeColor = TxtForeColour
'        txtPrivateMobile.BackColor = TxtBackColour
'        txtPrivateMobile.ForeColor = TxtForeColour
'        txtPrivateTel.BackColor = TxtBackColour
'        txtPrivateTel.ForeColor = TxtForeColour
'
'
'        txtQualifications.BackColor = TxtBackColour
'        txtQualifications.ForeColor = TxtForeColour
'        txtRegistation.BackColor = TxtBackColour
'        txtRegistation.ForeColor = TxtForeColour
'        txtSearch.BackColor = TxtBackColour
'        txtSearch.ForeColor = TxtForeColour
'        'txtTel.ForeColor = TxtForeColour
'        'txtTel.ForeColor = TxtForeColour
'        'txtPrivateTel.ForeColor = TxtForeColour
'        'txtPrivateTel.ForeColor = TxtForeColour
'        'txtPrivateTel.ForeColor = TxtForeColour
'        'txtPrivateTel.ForeColor = TxtForeColour
'        'txtPrivateTel.ForeColor = TxtForeColour
'        'txtPrivateTel.ForeColor = TxtForeColour
'        'txtPrivateTel.ForeColor = TxtForeColour
'        'txtPrivateTel.ForeColor = TxtForeColour

Me.BackColor = FrameBackColour
Me.ForeColor = FrameForeColour





End Sub
Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub ClearValues()
lblCashFromChanneling.Caption = "0.00"
lblAgentCashPayments.Caption = "0.00"
lblRefund.Caption = "0.00"
lblDoctorPayment.Caption = "0.00"
lblNetCashCollection.Caption = "0.00"
lblCashFromCreditChanneling = "0.00"
lblDoctorPayment.Caption = "0.00"
End Sub

Private Sub bttnPrintSummary_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    .Open "Select * From tblTem"
    Set dtrShiftEndCash.DataSource = DataEnvironment1.rssqlTemSu1
End With

With dtrShiftEndCash
    .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    .Sections("Section4").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    .Sections("Section2").Controls.Item("rptLCashChannaling").Caption = Format(lblCashFromChanneling, "0.00")
    .Sections("Section2").Controls.Item("rptAgentCashreceive").Caption = Format(lblAgentCashPayments, "0.00")
    .Sections("Section2").Controls.Item("rptlCashFromCreditChaneling").Caption = Format(lblCashFromCreditChanneling, "0.00")
    .Sections("Section2").Controls.Item("rptlNetCashChanelling").Caption = Format(lblNetPatientcash, "0.00")
    .Sections("Section2").Controls.Item("rptTotalCashReceive").Caption = Format(TemTotalCash, "0.00")
    .Sections("Section2").Controls.Item("rptlCancelRefund").Caption = Format(lblRefund, "0.00")
    .Sections("Section2").Controls.Item("rptDoctorPayment").Caption = Format(lblDoctorPayment, "0.00")
    .Sections("Section2").Controls.Item("rptTotalPayment").Caption = Format(TemTotalPayment, "0.00")
    .Sections("Section2").Controls.Item("rptlNetCashChanelling").Caption = Format(lblNetCashCollection, "0.00")
    
        If SSTab1.Tab = 0 Then
        .Sections("Section4").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        .Sections("Section4").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
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

Private Sub DataComboUser_Change()
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
DataComboUser.ListField = Empty
DataComboUser.RowMember = Empty
With DataEnvironment1.rssqlTem16
    If .State = 1 Then .Close
    .Source = "Select * From tblStaff Order by StaffName "
    .Open
    If .RecordCount = 0 Then Exit Sub
    DataComboUser.RowMember = "SQLTEM16"
    DataComboUser.ListField = "StaffName"
    If .State = 1 Then .Close
End With

DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
SSTab1.Tab = 0
lblUser.Caption = UserName
lblDate.Caption = Format(Date, DefaultLongDate)
DataComboUser.BoundText = UserID
Call Setcolours
CalculateValues
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

With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    Select Case SSTab1.Tab
    Case 0
    .Source = "Select tblStaffPayment.* From tblStaffPayment Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (PaidDate = '" & Date & "' )"
    .Open
    Case 1
     .Source = "Select tblStaffPayment.* From tblStaffPayment Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (PaidDate = '" & DTPicker1.Value & "' )"
    .Open
    Case 2
     .Source = "Select tblStaffPayment.* From tblStaffPayment Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (PaidDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "')"
    .Open
    End Select
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
        If !paiddate = Date Then
           TemDoctorPayment = TemDoctorPayment + !PaidAmount
        ElseIf !staffpayment_ID Mod IncomeDeflation = 0 Then
               TemDoctorPayment = TemDoctorPayment + !PaidAmount
        End If
     
     .MoveNext
    Loop
   
    If .State = 1 Then .Close
End With

lblDoctorPayment.Caption = Format(TemDoctorPayment, "0.00")

End Sub

Private Sub CashReceiveFromCreditBooking()


TemCashFromCreditCahnneling = 0

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.settlecashDate = '" & Date & "')  order by tblPatientFacility.patientfacility_ID ")
    ElseIf SSTab1.Tab = 1 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.settlecashDate = '" & DTPicker1.Value & "') order by tblPatientFacility.patientfacility_ID")
    ElseIf SSTab1.Tab = 2 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.settlecashDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') order by tblPatientFacility.patientfacility_ID ")
    End If
    
    While .EOF = False
        If !BookingDate = Date Then
        TemCashFromCreditCahnneling = Val(TemCashFromCreditCahnneling) + Val(!PersonalFee) + Val(!InstitutionFee) + Val(!otherfee)
        ElseIf patientfacility_ID Mod IncomeDeflation = 0 Then
        TemCashFromCreditCahnneling = Val(TemCashFromCreditCahnneling) + Val(!PersonalFee) + Val(!InstitutionFee) + Val(!otherfee)
        End If
        
        .MoveNext
    Wend

End With
lblCashFromCreditChanneling.Caption = Format(TemCashFromCreditCahnneling, "0.00")

End Sub

Private Sub AgentCashReceive()
TemAgentCash = 0

If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
With DataEnvironment1.rssqlAgentPayment1

    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate = '" & Date & "')   ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 1 Then
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate = '" & DTPicker1.Value & "') ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 2 Then
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    If .RecordCount = 0 Then Exit Sub
    End If
    
    Do While .EOF = False
        If !SettledDate = Date Then
            TemAgentCash = Val(TemAgentCash) + Val(!Cash)
        ElseIf !AgentCashSettle_ID Mod IncomeDeflation = 0 Then
            TemAgentCash = Val(TemAgentCash) + Val(!Cash)
        End If
    
    .MoveNext
    Loop
    
End With
lblAgentCashPayments.Caption = Format(TemAgentCash, "#,##0.00")
End Sub

Private Sub ChannelingCashIncome()

If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
CashFromCahnneling = 0

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    If SSTab1.Tab = 0 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = '" & Date & "')  order by tblPatientFacility.patientfacility_ID ")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 1 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = '" & DTPicker1.Value & "') order by tblPatientFacility.patientfacility_ID")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 2 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') order by tblPatientFacility.patientfacility_ID ")
    If .RecordCount = 0 Then Exit Sub
    End If
    
    While .EOF = False
        If !BookingDate = Date Then
        CashFromCahnneling = CashFromCahnneling + !PersonalFee + !InstitutionFee + !otherfee
        
        ElseIf !patientfacility_ID Mod IncomeDeflation = 0 Then
        CashFromCahnneling = CashFromCahnneling + !PersonalFee + !InstitutionFee + !otherfee
        
        End If
        
        .MoveNext
    
    Wend
End With

lblCashFromChanneling.Caption = Format(CashFromCahnneling, "0.00")

End Sub

Private Sub CashRefund()
temCashRefund = 0


With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = 1) and ( (tblPatientFacility.cancelled = 1)or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate = '" & Date & "')")
    ElseIf SSTab1.Tab = 1 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = 1) and ( (tblPatientFacility.cancelled = 1)or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate = '" & DTPicker1.Value & "')")
    ElseIf SSTab1.Tab = 2 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = 1) and ( (tblPatientFacility.cancelled = 1)or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "')")
    End If
 
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
     
        If !repaydate = Date Then
           If IsNull(!institutionrefund) = False Then temCashRefund = temCashRefund + (!institutionrefund)
           If IsNull(!Personalrefund) = False Then temCashRefund = temCashRefund + (!Personalrefund)
        ElseIf patientfacility_ID Mod IncomeDeflation = 0 Then
           If IsNull(!institutionrefund) = False Then temCashRefund = temCashRefund + (!institutionrefund)
           If IsNull(!Personalrefund) = False Then temCashRefund = temCashRefund + (!Personalrefund)
        End If
     
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

TemTotalCash = (CashFromCahnneling + TemAgentCash + TemCashFromCreditCahnneling)
TemTotalPayment = (temCashRefund + TemDoctorPayment)

TemNetChannelingincome = Val(TemTotalCash) - Val(TemTotalPayment)

lblNetCashCollection.Caption = Format(TemNetChannelingincome, "0.00")

End Sub
