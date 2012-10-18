VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmChannelingSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Channeling Summery"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmChannelingSummery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6765
   Begin VB.Frame FrameShiftSummary 
      Height          =   4935
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   6015
      Begin btButtonEx.ButtonEx bttnPrintCash 
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   3120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Detail Print Patient"
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
      Begin btButtonEx.ButtonEx bttnPrintAgentDetail 
         Height          =   375
         Left            =   4200
         TabIndex        =   18
         Top             =   3720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Detail Print Agent"
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
         Left            =   4200
         TabIndex        =   26
         Top             =   4320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Summery Print"
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
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Other Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label12 
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
         TabIndex        =   29
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctors Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label1 
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
         TabIndex        =   27
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblNetPatientcash 
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
         TabIndex        =   16
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblCancel 
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
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblNetChannelingincome 
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
         TabIndex        =   14
         Top             =   4320
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
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblNetAgentbookingCash 
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
         TabIndex        =   12
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblChannelingincome 
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
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.Line Line4 
         X1              =   360
         X2              =   5880
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   5760
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Channeling Income"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Value of Agent Bookings"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Cash Collection"
         Height          =   255
         Left            =   240
         TabIndex        =   8
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
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancellation Settling"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Expenses"
         Height          =   255
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Income"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Refund Settling"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash From Channeling"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   10821
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Today"
      TabPicture(0)   =   "FrmChannelingSummery.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblToday"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected Day"
      TabPicture(1)   =   "FrmChannelingSummery.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).Control(1)=   "Label11"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Selected Period"
      TabPicture(2)   =   "FrmChannelingSummery.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DTPicker2"
      Tab(2).Control(1)=   "DTPicker3"
      Tab(2).Control(2)=   "Label18"
      Tab(2).Control(3)=   "Label17"
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -72360
         TabIndex        =   19
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58064899
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -74040
         TabIndex        =   21
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58064899
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   -70920
         TabIndex        =   23
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   58064899
         CurrentDate     =   39442
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
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
         Left            =   -71400
         TabIndex        =   24
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
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
         Left            =   -74640
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Date"
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
         Top             =   600
         Width           =   1575
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
         TabIndex        =   1
         Top             =   600
         Width           =   3975
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   5040
      TabIndex        =   25
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
End
Attribute VB_Name = "FrmChannelingSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TemTotalfee As Double
Dim TemRefundRepayment As Double
Dim TemCancellationRepayment As Double
Dim TemNetTotal1 As Double
Dim TemNetTotal2 As Double
Dim TemNetChannelingincome As Double
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

        bttnCancelRepay.BackColor = BttnBackColour
        bttnCancelRepay.ForeColor = BttnForeColour
        
        bttnConfirmRepay.BackColor = BttnBackColour
        bttnConfirmRepay.ForeColor = BttnForeColour
        
        bttnSearchBookingID.BackColor = BttnBackColour
        bttnSearchBookingID.ForeColor = BttnForeColour
'
'        bttnClose.BackColor = BttnBackColour
'        bttnClose.ForeColor = BttnForeColour
'
'        bttnEdit.BackColor = BttnBackColour
'        bttnEdit.ForeColor = BttnForeColour
'
'        bttnSave.BackColor = BttnBackColour
'        bttnSave.ForeColor = BttnForeColour
        
        FrameShiftSummary.BackColor = FrmBackColour
        FrameShiftSummary.ForeColor = FrmForeColour
        
'        FrameSearch.BackColor = FrmBackColour
'        FrameSearch.ForeColor = FrmForeColour
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


        OptionTwo.BackColor = FrmBackColour
        OptionTwo.ForeColor = FrmForeColour
        
        OptionOne.BackColor = FrmBackColour
        OptionOne.ForeColor = FrmForeColour
        
        OptionNo.BackColor = FrmBackColour
        OptionNo.ForeColor = FrmForeColour
        
        
        
        
        
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







End Sub
Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnPrintAgentDetail_Click()

If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub

With DataEnvironment1.rssqlAgentReport
    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
    .Open ("SELECT tblPatientFacility.*, tblInstitutions.* FROM tblPatientFacility LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (repayUser_ID = " & DataComboUser.BoundText & ") and (PaymentMode = 'Agent')and (BookingDate = #" & Date & "#)")
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "Error User"): Exit Sub
    ElseIf SSTab1.Tab = 1 Then
    .Open ("SELECT tblPatientFacility.*, tblInstitutions.* FROM tblPatientFacility LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (repayUser_ID = " & DataComboUser.BoundText & ") and (PaymentMode = 'Agent')and (BookingDate = #" & DTPicker1.Value & "#)")
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "Error User"): Exit Sub
    ElseIf SSTab1.Tab = 2 Then
    .Open ("SELECT tblPatientFacility.*, tblInstitutions.* FROM tblPatientFacility LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (repayUser_ID = " & DataComboUser.BoundText & ") and (PaymentMode = 'Agent')and (BookingDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)")
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "Error User"): Exit Sub
    End If
    
    Set dtrAgentSummary.DataSource = DataEnvironment1.rssqlAgentReport
    
    dtrAgentSummary.Sections("ReportHeader").Controls.Item("RptName").Caption = InstitutionName
    dtrAgentSummary.Sections("ReportHeader").Controls.Item("RptAddress").Caption = InstitutionAddress
    dtrAgentSummary.Sections("PageHeader").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    If SSTab1.Tab = 0 Then
    dtrAgentSummary.Sections("PageHeader").Controls.Item("rptFromdate").Caption = Format(Date, "dd/MM/YYY")
    dtrAgentSummary.Sections("PageHeader").Controls.Item("RptToDate").Caption = Format(Date, "dd/MM/YYY")
    ElseIf SSTab1.Tab = 1 Then
    dtrAgentSummary.Sections("PageHeader").Controls.Item("rptFromdate").Caption = DTPicker1.Value
    dtrAgentSummary.Sections("PageHeader").Controls.Item("RptToDate").Caption = DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
    dtrAgentSummary.Sections("PageHeader").Controls.Item("rptFromdate").Caption = DTPicker2.Value
    dtrAgentSummary.Sections("PageHeader").Controls.Item("RptToDate").Caption = DTPicker3.Value
    End If
    
    dtrAgentSummary.Show
    
End With
End Sub

Private Sub bttnPrintCash_Click()
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
With DataEnvironment1.rssqlCashireRepost

    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (repayUser_ID = " & DataComboUser.BoundText & ") and (PaymentMode = 'Cash')and (BookingDate = #" & Date & "#)")
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "Error User"): Exit Sub
    ElseIf SSTab1.Tab = 1 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (repayUser_ID = " & DataComboUser.BoundText & ") and (PaymentMode = 'Cash')and (BookingDate = #" & DTPicker1.Value & "#)")
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "Error User"): Exit Sub
    ElseIf SSTab1.Tab = 2 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (repayUser_ID = " & DataComboUser.BoundText & ") and (PaymentMode = 'Cash')and (BookingDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)")
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "Error User"): Exit Sub
    End If
    
    Set dtrPatientSummary.DataSource = DataEnvironment1.rssqlCashireRepost
    
    dtrPatientSummary.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    dtrPatientSummary.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    dtrPatientSummary.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text

    If SSTab1.Tab = 0 Then
    dtrPatientSummary.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, "dd/MM/YYY")
    dtrPatientSummary.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, "dd/MM/YYY")
    ElseIf SSTab1.Tab = 1 Then
    dtrPatientSummary.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
    dtrPatientSummary.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
    dtrPatientSummary.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
    dtrPatientSummary.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
    End If
    
dtrPatientSummary.Show

End With

End Sub


Private Sub FindTodayChanelingIncome() 'today
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
TemCancellationRepayment = 0
TemTotalfee = 0
TemRefundRepayment = 0

With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & Date & "#)and (PaymentMode = 'Cash')"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
     
     TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
     If !refund = True Then
        If IsNull(!InstitutionRefund) = False Then TemRefundRepayment = TemRefundRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemRefundRepayment = TemRefundRepayment + (!personalrefund)
     End If
     If !cancelled = True Then
        If IsNull(!InstitutionRefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!personalrefund)
     End If
     
     .MoveNext
    Loop
   
    If .State = 1 Then .Close
End With
lblChannelingincome.Caption = Format(TemTotalfee, "#0.00")
lblRefund.Caption = Format(TemRefundRepayment, "#0.00")
lblCancel.Caption = Format(TemCancellationRepayment, "#0.00")
End Sub

Private Sub FindSelectedDayChanelingIncome() 'Selected ady
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
TemCancellationRepayment = 0
TemTotalfee = 0
TemRefundRepayment = 0

With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & DTPicker1.Value & "#)and (PaymentMode = 'Cash')"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
     
     TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
     If !refund = True Then
        If IsNull(!InstitutionRefund) = False Then TemRefundRepayment = TemRefundRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemRefundRepayment = TemRefundRepayment + (!personalrefund)
     End If
     If !cancelled = True Then
        If IsNull(!InstitutionRefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!personalrefund)
     End If
     .MoveNext
    Loop
   
    If .State = 1 Then .Close
End With
lblChannelingincome.Caption = Format(TemTotalfee, "#0.00")
lblRefund.Caption = Format(TemRefundRepayment, "#0.00")
lblCancel.Caption = Format(TemCancellationRepayment, "#0.00")
End Sub

Private Sub ClearValues()
lblChannelingincome.Caption = "0.00"
lblRefund.Caption = "0.00"
lblCancel.Caption = "0.00"
lblNetPatientcash.Caption = "0.00"
lblNetAgentbookingCash.Caption = "0.00"
lblNetChannelingincome.Caption = "0.00"
End Sub
Private Sub FindSelectedPeriodChanelingIncome() 'Selected Peried
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
TemCancellationRepayment = 0
TemTotalfee = 0
TemRefundRepayment = 0

With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)and (PaymentMode = 'Cash')"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
     
     TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
     If !refund = True Then
        If IsNull(!InstitutionRefund) = False Then TemRefundRepayment = TemRefundRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemRefundRepayment = TemRefundRepayment + (!personalrefund)
     End If
     If !cancelled = True Then
        If IsNull(!InstitutionRefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!personalrefund)
     End If
     
     .MoveNext
    Loop
   
    If .State = 1 Then .Close
End With
lblChannelingincome.Caption = Format(TemTotalfee, "#0.00")
lblRefund.Caption = Format(TemRefundRepayment, "#0.00")
lblCancel.Caption = Format(TemCancellationRepayment, "#0.00")
End Sub
Private Sub FindDirectPatientIncome() 'Today
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
TemCancellationRepayment = 0
TemTotalfee = 0
TemRefundRepayment = 0
TemNetTotal1 = 0

With DataEnvironment1.rssqlTemSu2
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & Date & "#)and (PaymentMode = 'Cash')"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
     TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
     If !refund = True Then
        If IsNull(!InstitutionRefund) = False Then TemRefundRepayment = TemRefundRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemRefundRepayment = TemRefundRepayment + (!personalrefund)
     End If
     If !cancelled = True Then
        If IsNull(!InstitutionRefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!personalrefund)
     End If
     .MoveNext
     Loop
   
    If .State = 1 Then .Close
End With

TemNetTotal1 = Val(TemTotalfee) - (Val(TemCancellationRepayment) + Val(TemRefundRepayment))
lblNetPatientcash.Caption = Format(TemNetTotal1, "#0.00")

End Sub

Private Sub FindDirectPatientIncomeSelectedDay() 'Selected Day
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
TemCancellationRepayment = 0
TemTotalfee = 0
TemRefundRepayment = 0
TemNetTotal1 = 0

With DataEnvironment1.rssqlTemSu2
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & DTPicker1.Value & "#)and (PaymentMode = 'Cash')"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
     TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
     If !refund = True Then
        If IsNull(!InstitutionRefund) = False Then TemRefundRepayment = TemRefundRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemRefundRepayment = TemRefundRepayment + (!personalrefund)
     End If
     If !cancelled = True Then
        If IsNull(!InstitutionRefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!personalrefund)
     End If
     .MoveNext
     Loop
   
    If .State = 1 Then .Close
End With

TemNetTotal1 = Val(TemTotalfee) - (Val(TemCancellationRepayment) + Val(TemRefundRepayment))
lblNetPatientcash.Caption = Format(TemNetTotal1, "#0.00")

End Sub

Private Sub FindDirectPatientIncomeSelectedPeriod() 'Selected Period
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
TemCancellationRepayment = 0
TemTotalfee = 0
TemRefundRepayment = 0
TemNetTotal1 = 0

With DataEnvironment1.rssqlTemSu2
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)and (PaymentMode = 'Cash')"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
     TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
     If !refund = True Then
        If IsNull(!InstitutionRefund) = False Then TemRefundRepayment = TemRefundRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemRefundRepayment = TemRefundRepayment + (!personalrefund)
     End If
     If !cancelled = True Then
        If IsNull(!InstitutionRefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!InstitutionRefund)
        If IsNull(!personalrefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!personalrefund)
     End If
     .MoveNext
     Loop
   
    If .State = 1 Then .Close
End With

TemNetTotal1 = Val(TemTotalfee) - (Val(TemCancellationRepayment) + Val(TemRefundRepayment))
lblNetPatientcash.Caption = Format(TemNetTotal1, "#0.00")

End Sub

Private Sub FindAGentPatientIncome() 'Today
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
TemCancellationRepayment = 0
TemTotalfee = 0
TemRefundRepayment = 0
TemNetTotal2 = 0

With DataEnvironment1.rssqlTemSu2
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & Date & "#)and (PaymentMode = 'Agent')"
    .Open

    If .RecordCount = 0 Then Exit Sub

     Do While .EOF = False
     TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
     If !refund = True Then
     TemRefundRepayment = Val(TemRefundRepayment) + Val(!personalrefund) + Val(!InstitutionRefund)
     End If
     If !cancelled = True Then
     TemCancellationRepayment = Val(TemCancellationRepayment) + Val(!InstitutionRefund) + Val(!personalrefund)
     End If
     .MoveNext
     Loop

    If .State = 1 Then .Close
End With

TemNetTotal2 = Val(TemTotalfee) - (Val(TemCancellationRepayment) + Val(TemRefundRepayment))
lblNetAgentbookingCash.Caption = Format(TemNetTotal2, "#0.00")

End Sub

Private Sub FindAGentPatientIncomeSelectedDay() 'Selectedday
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
TemCancellationRepayment = 0
TemTotalfee = 0
TemRefundRepayment = 0
TemNetTotal2 = 0

With DataEnvironment1.rssqlTemSu2
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & DTPicker1.Value & "#)and (PaymentMode = 'Agent')"
    .Open

    If .RecordCount = 0 Then Exit Sub

     Do While .EOF = False
     TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
     If !refund = True Then
     TemRefundRepayment = Val(TemRefundRepayment) + Val(!personalrefund) + Val(!InstitutionRefund)
     End If
     If !cancelled = True Then
     TemCancellationRepayment = Val(TemCancellationRepayment) + Val(!InstitutionRefund) + Val(!personalrefund)
     End If
     .MoveNext
     Loop

    If .State = 1 Then .Close
End With

TemNetTotal2 = Val(TemTotalfee) - (Val(TemCancellationRepayment) + Val(TemRefundRepayment))
lblNetAgentbookingCash.Caption = Format(TemNetTotal2, "#0.00")

End Sub

Private Sub FindAGentPatientIncomeSelectedPeriod() 'Period
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
TemCancellationRepayment = 0
TemTotalfee = 0
TemRefundRepayment = 0
TemNetTotal2 = 0

With DataEnvironment1.rssqlTemSu2
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate Between #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)and (PaymentMode = 'Agent')"
    .Open

    If .RecordCount = 0 Then Exit Sub

     Do While .EOF = False
     TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
     If !refund = True Then
     TemRefundRepayment = Val(TemRefundRepayment) + Val(!personalrefund) + Val(!InstitutionRefund)
     End If
     If !cancelled = True Then
     TemCancellationRepayment = Val(TemCancellationRepayment) + Val(!InstitutionRefund) + Val(!personalrefund)
     End If
     .MoveNext
     Loop

    If .State = 1 Then .Close
End With

TemNetTotal2 = Val(TemTotalfee) - (Val(TemCancellationRepayment) + Val(TemRefundRepayment))
lblNetAgentbookingCash.Caption = Format(TemNetTotal2, "#0.00")

End Sub

Private Sub NetChannelingIncome()
TemNetChannelingincome = Val(TemNetTotal1) + Val(TemNetTotal2)
lblNetChannelingincome.Caption = Format(TemNetChannelingincome, "#0.00")
End Sub

Private Sub ButtonEx2_Click()

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
    .Sections("Section2").Controls.Item("rptLCashChannaling").Caption = Format(lblChannelingincome, "#0.00")
    .Sections("Section2").Controls.Item("rptLRefund").Caption = Format(lblRefund, "#0.00")
    .Sections("Section2").Controls.Item("rptlCancell").Caption = Format(lblCancel, "#0.00")
    .Sections("Section2").Controls.Item("rptlNetCashChanelling").Caption = Format(lblNetPatientcash, "#0.00")
    .Sections("Section2").Controls.Item("rptlAgentBookingChaneling").Caption = Format(lblNetAgentbookingCash, "#0.00")
    .Sections("Section2").Controls.Item("rptlNetChannlingIncome").Caption = Format(lblNetChannelingincome, "#0.00")
        If SSTab1.Tab = 0 Then
        .Sections("Section4").Controls.Item("rptFromdate").Caption = Format(Date, "dd/MM/YYY")
        .Sections("Section4").Controls.Item("RptToDate").Caption = Format(Date, "dd/MM/YYY")
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

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
Call Setcolours
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Call CalculateValues
End Sub

Private Sub CalculateValues()
If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
Call ClearValues

Call CashIncome
Call AgentIncome
Call CashExpence

Call CalculateTotals


Exit Sub

If SSTab1.Tab = 0 Then
Call ClearValues
Call FindTodayChanelingIncome
Call FindDirectPatientIncome
Call FindAGentPatientIncome
Call NetChannelingIncome

ElseIf SSTab1.Tab = 1 Then
Call ClearValues
Call FindSelectedDayChanelingIncome
Call FindDirectPatientIncomeSelectedDay
Call FindAGentPatientIncomeSelectedDay
Call NetChannelingIncome

ElseIf SSTab1.Tab = 2 Then
Call ClearValues
Call FindSelectedPeriodChanelingIncome
Call FindDirectPatientIncomeSelectedPeriod
Call FindAGentPatientIncomeSelectedPeriod
Call NetChannelingIncome
End If
End Sub

Private Sub CashIncome()
TemTotalfee = 0

With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
        Case 0:
            .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & Date & "#)and (PaymentMode = 'Cash')"
        Case 1:
            .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & DTPicker1.Value & "#)and (PaymentMode = 'Cash')"
        Case 2:
            .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)and (PaymentMode = 'Cash')"
        End Select
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
        TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
        .MoveNext
    Loop
    If .State = 1 Then .Close
End With
lblChannelingincome.Caption = Format(TemTotalfee, "#0.00")
End Sub

Private Sub AgentIncome()
TemTotalfee = 0

With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
        Case 0:
            .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & Date & "#)and (PaymentMode = 'Agent')"
        Case 1:
            .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & DTPicker1.Value & "#)and (PaymentMode = 'Agent')"
        Case 2:
            .Source = "Select tblPatientFacility.* From tblPatientFacility Where (User_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)and (PaymentMode = 'Agent')"
        End Select
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
        TemTotalfee = Val(TemTotalfee) + Val(!totalfee)
        .MoveNext
    Loop
    If .State = 1 Then .Close
End With
lblNetAgentbookingCash.Caption = Format(TemTotalfee, "#0.00")
End Sub

Private Sub CashExpence()
TemCancellationRepayment = 0
TemRefundRepayment = 0

With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
        Case 0:
            .Source = "Select tblPatientFacility.* From tblPatientFacility Where (repayUser_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & Date & "#)"
        Case 1:
            .Source = "Select tblPatientFacility.* From tblPatientFacility Where (repayUser_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate = #" & DTPicker1.Value & "#)"
        Case 2:
            .Source = "Select tblPatientFacility.* From tblPatientFacility Where (repayUser_ID = " & Val(DataComboUser.BoundText) & ") and (BookingDate Between  #" & DTPicker2.Value & "# and #" & DTPicker3.Value & "#)"
        End Select
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
        If !refund = True Then
           If IsNull(!InstitutionRefund) = False Then TemRefundRepayment = TemRefundRepayment + (!InstitutionRefund)
           If IsNull(!personalrefund) = False Then TemRefundRepayment = TemRefundRepayment + (!personalrefund)
        End If
        If !cancelled = True Then
           If IsNull(!InstitutionRefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!InstitutionRefund)
           If IsNull(!personalrefund) = False Then TemCancellationRepayment = TemCancellationRepayment + (!personalrefund)
        End If
        .MoveNext
    Loop
   
    If .State = 1 Then .Close
End With
lblRefund.Caption = Format(TemRefundRepayment, "#0.00")
lblCancel.Caption = Format(TemCancellationRepayment, "#0.00")
End Sub

Private Sub CalculateTotals()
Dim NetCash As Double
Dim NetChanneling As Double

NetCash = Val(lblChannelingincome.Caption) - (Val(lblRefund.Caption) + Val(lblCancel.Caption))
lblNetPatientcash.Caption = Format(NetCash, "#0.00")

NetChanneling = Val(lblNetAgentbookingCash.Caption) + NetCash
lblNetChannelingincome.Caption = Format(NetChanneling, "#0.00")


End Sub
