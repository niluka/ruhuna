VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAllAgentSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agents Summery"
   ClientHeight    =   5280
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
   Icon            =   "frmAllAgentSummery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8925
   Begin btButtonEx.ButtonEx btnBalance 
      Height          =   495
      Left            =   480
      TabIndex        =   29
      Top             =   4680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Current Balance"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   4680
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
   Begin VB.Frame FrameShiftSummary 
      Height          =   3255
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   7815
      Begin btButtonEx.ButtonEx bttnRepaidTotal 
         Height          =   375
         Left            =   4440
         TabIndex        =   24
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print R&e-paid Total"
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
      Begin btButtonEx.ButtonEx bttnPrintBalance 
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   1920
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print &Balance"
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
         Left            =   4440
         TabIndex        =   5
         Top             =   960
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
         Left            =   4440
         TabIndex        =   4
         Top             =   480
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
      Begin btButtonEx.ButtonEx btnSame 
         Height          =   375
         Left            =   6600
         TabIndex        =   27
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Same"
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
      Begin btButtonEx.ButtonEx btnDifferent 
         Height          =   375
         Left            =   6600
         TabIndex        =   28
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Different"
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
      Begin VB.Line Line2 
         X1              =   6600
         X2              =   6240
         Y1              =   1320
         Y2              =   1200
      End
      Begin VB.Line Line1 
         X1              =   6240
         X2              =   6600
         Y1              =   1200
         Y2              =   840
      End
      Begin VB.Label lblCancellations 
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
         Left            =   2760
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancellations"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lblAgentCashrefund 
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
         Left            =   2760
         TabIndex        =   23
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Repay to Agent"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Booking Value"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Receive"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   4680
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCashReceive 
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
         Left            =   2760
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblAgenBooking 
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
         Left            =   2760
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label lblBalance 
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
         Left            =   2760
         TabIndex        =   9
         Top             =   2400
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   7435
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Today"
      TabPicture(0)   =   "frmAllAgentSummery.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblToday"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected &Day"
      TabPicture(1)   =   "frmAllAgentSummery.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "DTPicker1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Selected &Period"
      TabPicture(2)   =   "frmAllAgentSummery.frx":047A
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
         Left            =   -71280
         TabIndex        =   1
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22872067
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   420
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22872067
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   420
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22872067
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
         Left            =   -74040
         TabIndex        =   20
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label11 
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
         Left            =   -72960
         TabIndex        =   19
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
         Left            =   840
         TabIndex        =   18
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
         Left            =   4440
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   3840
      TabIndex        =   21
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "frmAllAgentSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemCashTotal As Double
    Dim TemAgentBookingPatients As Double
    Dim TemRefund As Double
    Dim CSetPrinter As New cSetDfltPrinter
    Dim rsRepay As New ADODB.Recordset
    Dim temSQL As String

Private Sub btnBalance_Click()
    frmAgentBalance.Show
    frmAgentBalance.ZOrder 0
End Sub

Private Sub btnDifferent_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1.rssqlTem12
    Dim A As Integer
    If .State = 1 Then .Close
   
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "Select tblPatientFacility.*,tblInstitutions.* fROM tblPatientFacility Left Join tblInstitutions On tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where ( tblPatientFacility.BookingDate = '" & Date & "') and (tblPatientFacility.PaymentMode ='Agent')"
    .Open
    
    Case 1
    .Source = "Select tblPatientFacility.*,tblInstitutions.* fROM tblPatientFacility Left Join tblInstitutions On tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where ( tblPatientFacility.BookingDate = '" & DTPicker1 & "')and (tblPatientFacility.PaymentMode ='Agent')"
    .Open
    
    Case 2
    .Source = "Select tblPatientFacility.*,tblInstitutions.* fROM tblPatientFacility Left Join tblInstitutions On tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where ( tblPatientFacility.BookingDate between '" & DTPicker2 & "' and '" & DTPicker3 & "') and (tblPatientFacility.AppointmentDate > '" & DTPicker3 & "') and (tblPatientFacility.PaymentMode ='Agent')"
    .Open
    
    End Select
    
    If .RecordCount = 0 Then A = MsgBox("No Agent Booking to view", vbCritical + vbOKOnly, "No Data"): Exit Sub

    With DataReportAgentBookings
    
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
         
        .Sections("Section2").Controls.Item("rptleading1").Caption = ""
        .Sections("Section2").Controls.Item("RptCashierName").Caption = ""
 
        
        Set .DataSource = DataEnvironment1.rssqlTem12
        .Show
    End With

End With


End Sub

Private Sub btnSame_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1.rssqlTem12
    Dim A As Integer
    If .State = 1 Then .Close
   
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "Select tblPatientFacility.*,tblInstitutions.* fROM tblPatientFacility Left Join tblInstitutions On tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where ( tblPatientFacility.BookingDate = '" & Date & "') and (tblPatientFacility.PaymentMode ='Agent')"
    .Open
    
    Case 1
    .Source = "Select tblPatientFacility.*,tblInstitutions.* fROM tblPatientFacility Left Join tblInstitutions On tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where ( tblPatientFacility.BookingDate = '" & DTPicker1 & "')and (tblPatientFacility.PaymentMode ='Agent')"
    .Open
    
    Case 2
    .Source = "Select tblPatientFacility.*,tblInstitutions.* fROM tblPatientFacility Left Join tblInstitutions On tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where ( tblPatientFacility.BookingDate between '" & Format(DTPicker2, "dd MMMM yyyy") & "' and '" & Format(DTPicker3, "dd MMMM yyyy") & "') and ( tblPatientFacility.AppointmentDate between '" & Format(DTPicker2, "dd MMMM yyyy") & "' and '" & Format(DTPicker3, "dd MMMM yyyy") & "')  and (tblPatientFacility.PaymentMode ='Agent')"
    .Open
    
    End Select
    
    If .RecordCount = 0 Then A = MsgBox("No Agent Booking to view", vbCritical + vbOKOnly, "No Data"): Exit Sub

    With DataReportAgentBookings
    
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
         
        .Sections("Section2").Controls.Item("rptleading1").Caption = ""
        .Sections("Section2").Controls.Item("RptCashierName").Caption = ""
 
        
        Set .DataSource = DataEnvironment1.rssqlTem12
        .Show
    End With


End With

End Sub

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

bttnRepaidTotal.BackColor = BttnBackColour
bttnRepaidTotal.ForeColor = BttnForeColour

bttnPrintBalance.BackColor = BttnBackColour
bttnPrintBalance.ForeColor = BttnForeColour


FrameShiftSummary.BackColor = FrameBackColour
FrameShiftSummary.ForeColor = FrameForeColour

frmAllAgentSummery.BackColor = FrameBackColour
frmAllAgentSummery.ForeColor = FrameForeColour
End Sub

Private Sub FindAgentCashReceive()
    TemCashTotal = 0
    lblCashReceive.Caption = "0.00"
    
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        
        Select Case SSTab1.Tab
        
        Case 0
         .Source = "Select tblAgentCashSettle.* fROM tblAgentCashSettle Where (SettledDate = '" & Date & "')"
        .Open
        
        Case 1
        .Source = "Select tblAgentCashSettle.* fROM tblAgentCashSettle Where (SettledDate = '" & DTPicker1 & "')"
        .Open
        
        Case 2
        .Source = "Select tblAgentCashSettle.* fROM tblAgentCashSettle Where (SettledDate between '" & DTPicker2 & "' and '" & DTPicker3 & "')"
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

Private Sub FindAgentCancellation()
    TemCashTotal = 0
    lblCancellations.Caption = "0.00"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        Select Case SSTab1.Tab
            Case 0
                .Source = "Select tblAgentPaymentCancellation.* fROM tblAgentPaymentCancellation Where (Date between '" & Date & "' AND '" & Date & "'  )"
            Case 1
                .Source = "Select tblAgentPaymentCancellation.* fROM tblAgentPaymentCancellation Where (Date between '" & DTPicker1 & "' and '" & DTPicker1 & "'  )"
            Case 2
                .Source = "Select tblAgentPaymentCancellation.* fROM tblAgentPaymentCancellation Where (Date between '" & DTPicker2 & "' and '" & DTPicker3 & "')"
        End Select
        .Open
        If .RecordCount = 0 Then Exit Sub
        Do While .EOF = False
            TemCashTotal = Val(TemCashTotal) + Val(!Amount)
            .MoveNext
        Loop
    End With
    lblCancellations.Caption = Format(TemCashTotal, "0.00")
End Sub


Private Sub bttnPrintAgentBooking_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1.rssqlTem12
    Dim A As Integer
    If .State = 1 Then .Close
   
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "Select tblPatientFacility.*,tblInstitutions.* fROM tblPatientFacility Left Join tblInstitutions On tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where ( tblPatientFacility.BookingDate = '" & Date & "') and (tblPatientFacility.PaymentMode ='Agent')"
    .Open
    
    Case 1
    .Source = "Select tblPatientFacility.*,tblInstitutions.* fROM tblPatientFacility Left Join tblInstitutions On tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where ( tblPatientFacility.BookingDate = '" & DTPicker1 & "')and (tblPatientFacility.PaymentMode ='Agent')"
    .Open
    
    Case 2
    .Source = "Select tblPatientFacility.*,tblInstitutions.* fROM tblPatientFacility Left Join tblInstitutions On tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where ( tblPatientFacility.BookingDate between '" & DTPicker2 & "' and '" & DTPicker3 & "') and (tblPatientFacility.PaymentMode ='Agent')"
    .Open
    
    End Select
    
    If .RecordCount = 0 Then A = MsgBox("No Agent Booking to view", vbCritical + vbOKOnly, "No Data"): Exit Sub

    With DataReportAgentBookings
    
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
         
        .Sections("Section2").Controls.Item("rptleading1").Caption = ""
        .Sections("Section2").Controls.Item("RptCashierName").Caption = ""
 
        
        Set .DataSource = DataEnvironment1.rssqlTem12
        .Show
    End With



End With

End Sub

Private Sub bttnPrintBalance_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

Dim A As Integer

With DataEnvironment1.rssqlTem11

    If .State = 1 Then .Close
    
     .Source = "Select tblInstitutions.* fROM tblInstitutions Order by InstitutionName" 'Where InstitutionCredit <> 0"
     .Open
     
    If .RecordCount = 0 Then A = MsgBox("No Institution to view", vbCritical + vbOKOnly, "No Data"): Exit Sub

With dtrAgentBalance
        If HospitalDetails = True Then
            .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
            .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
            .Sections("section3").Controls.Item("lblAds").Caption = LongAd
            .Sections("Section2").Controls.Item("rptLDate1").Caption = Format(Date, DefaultLongDate)
        Else
            .Sections("Section4").Controls.Item("RptName").Caption = Empty
            .Sections("Section4").Controls.Item("RptAddress").Caption = Empty
            .Sections("section3").Controls.Item("lblAds").Caption = LongAd
            .Sections("Section2").Controls.Item("rptLDate1").Caption = Format(Date, DefaultLongDate)
        End If
        Set .DataSource = DataEnvironment1.rssqlTem11
        .Show
        
        End With

End With

End Sub

Private Sub bttnPrintCashReceive_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)

With DataEnvironment1.rssqlTem3
    If .State = 1 Then .Close
    
    Dim A As Integer
    
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "Select tblAgentCashSettle.*, tblInstitutions.* fROM tblAgentCashSettle Left Join tblInstitutions On tblAgentCashSettle.Institution_Id = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.SettledDate = '" & Date & "')"
    .Open
    
    Case 1
    .Source = "Select tblAgentCashSettle.*, tblInstitutions.* fROM tblAgentCashSettle Left Join tblInstitutions On tblAgentCashSettle.Institution_Id = tblInstitutions.Institution_ID Where (tblAgentCashSettle.SettledDate = '" & DTPicker1 & "')"
    .Open
    
    Case 2
    .Source = "SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID Where (tblAgentCashSettle.SettledDate between '" & DTPicker2 & "' and '" & DTPicker3 & "')"
    .Open
    
    End Select
    
    If .RecordCount = 0 Then A = MsgBox("No Cash receive to view", vbCritical + vbOKOnly, "No Data"): Exit Sub
    
    With dtrAgentCashReceive
        Set .DataSource = DataEnvironment1.rssqlTem3
    
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
         
        .Sections("Section2").Controls.Item("rptlHeding1").Caption = ""
        .Sections("Section2").Controls.Item("RptCashierName").Caption = ""
    
        
        .Show
    
    End With

End With

End Sub

Private Sub FindAgentBookingPatients()
TemAgentBookingPatients = 0
lblAgenBooking.Caption = "0.00"

With DataEnvironment1.rssqlTem2
    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (BookingDate = '" & Date & "') and (PaymentMode ='Agent')"
    .Open
    
    Case 1
    .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (BookingDate = '" & DTPicker1 & "')and (PaymentMode ='Agent')"
    .Open
    
    Case 2
    .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (BookingDate between '" & DTPicker2 & "' and '" & DTPicker3 & "') and (PaymentMode ='Agent')"
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
        .Source = "SELECT sum(tblInstitutionBalance.EBalance) as InsBal From tblInstitutionBalance where tblInstitutionBalance.Date = '" & Format(TemDate, "dd MMMM yyyy") & "'"
        .Open
        If .RecordCount = 0 Then Exit Sub
        If IsNull(!InsBal) = True Then Exit Sub
        TemInstutionBal = Val(!InsBal)
        lblBalance.Caption = Format(TemInstutionBal, "0.00")
    End With
End Sub

Private Sub DTPicker1_Change()
CalculateValues
End Sub

Private Sub DTPicker2_Change()
CalculateValues
End Sub

Private Sub DTPicker3_Change()
CalculateValues
End Sub

Private Sub CalculateValues()
    ClearValues
    FindAgentCashReceive
    FindAgentBookingPatients
    FindInstutionBalanc
    FindAgentCancellation
    FindAgentRefund
End Sub

Private Sub ClearValues()
    lblCashReceive.Caption = "0.00"
    lblAgenBooking.Caption = "0.00"
    lblBalance.Caption = "0.00"
    lblAgentCashrefund = "0.00"
    lblCancellations = "0.00"
End Sub


Private Sub Form_Load()
If SetPrinter = False Then Unload Me: Exit Sub
    Me.Top = (Screen.Height / 2) - (Me.Height)
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
DTPicker1 = Date
DTPicker2 = Date
DTPicker3 = Date
CalculateValues
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


Private Sub FindAgentRefund()
TemRefund = 0

With DataEnvironment1.rssqlTem12

    If .State = 1 Then .Close
    
    Select Case SSTab1.Tab
    
    Case 0
     .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (RefundToAgent = 1) and ( RepayDate = '" & Date & "') and (PaymentMode ='Agent') Order by PatientFacility_ID "
     .Open
    Case 1
     .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (RefundToAgent = 1) and ( RepayDate = '" & DTPicker1 & "')and (PaymentMode ='Agent')Order by PatientFacility_ID "
     .Open
    Case 2
     .Source = "Select tblPatientFacility.* fROM tblPatientFacility Where (RefundToAgent = 1) and ( RepayDate between '" & DTPicker2 & "' and '" & DTPicker3 & "') and (PaymentMode ='Agent')Order by PatientFacility_ID "
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

Private Sub bttnRepaidTotal_Click()
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
Dim A As Integer

TemRefund = 0

With rsRepay
    If .State = 1 Then .Close
    Select Case SSTab1.Tab
    Case 0
     temSQL = "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode FROM (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (RefundToAgent = 1) and ( RepayDate = '" & Format(Date, "dd MMMM yyyy") & "') and (PaymentMode ='Agent') Order by PatientFacility_ID "
     .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
    Case 1
     temSQL = "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode FROM (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (RefundToAgent = 1) and ( RepayDate = '" & Format(DTPicker1.Value, "dd MMMM yyyy") & "')and (PaymentMode ='Agent')Order by PatientFacility_ID "
    .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
    Case 2
     temSQL = "SELECT tblPatientFacility.*, tblPatientMainDetails.FirstName, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode FROM (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (RefundToAgent = 1) and ( RepayDate between '" & Format(DTPicker2.Value, "dd MMMM yyyy") & "' and '" & DTPicker3 & "') and (PaymentMode ='Agent')Order by PatientFacility_ID "
     .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
    End Select
    If .RecordCount = 0 Then A = MsgBox("No Agent Booking to view", vbCritical + vbOKOnly, "No Data"): Exit Sub
        
        With dtrAgentRefundSummery
            Set .DataSource = rsRepay
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
         
'        .Sections("Section2").Controls.Item("rptAgentName").Caption = dtcAgentName.Text
'        .Sections("Section2").Controls.Item("rptAgentCode").Caption = dtcAgentCode.Text
        
        
        .Show
        End With

 
End With

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)

Select Case SSTab1.Tab

    Case 0
    CalculateValues
    Case 1
    CalculateValues
    Case 2
    CalculateValues

End Select

End Sub

