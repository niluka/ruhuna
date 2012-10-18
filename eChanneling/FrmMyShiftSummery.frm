VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmMyShiftSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift Summery"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMyShiftSummery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   9885
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   33
      Top             =   6720
      Width           =   6735
      Begin btButtonEx.ButtonEx btnMyStaffBookings 
         Height          =   375
         Left            =   3240
         TabIndex        =   34
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "My Staff Bookings"
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
      Begin btButtonEx.ButtonEx btnMyTelephoneBookings 
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "My &Telephone Bookings"
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
      Begin btButtonEx.ButtonEx btnMySettledTelephoneBookings 
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "My Settled Telephone Bookings"
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
      Begin btButtonEx.ButtonEx btnMySettledStaffBookings 
         Height          =   375
         Left            =   3240
         TabIndex        =   37
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "My Settled Staff Bookings"
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
   End
   Begin VB.Frame FrameShiftSummary 
      Height          =   4455
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   9375
      Begin btButtonEx.ButtonEx bttnDetailSummary 
         Height          =   375
         Left            =   6360
         TabIndex        =   32
         Top             =   3960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Detail Su&mmery Print"
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
      Begin btButtonEx.ButtonEx bttnCashFromCreditChanneling 
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print &Settling Credit"
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
      Begin btButtonEx.ButtonEx bttnPrintAgentCash 
         Height          =   375
         Left            =   6360
         TabIndex        =   6
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print &Agent Payments"
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
         Left            =   6360
         TabIndex        =   10
         Top             =   3120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Summary &Print"
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
      Begin btButtonEx.ButtonEx bttnCashIncome 
         Height          =   375
         Left            =   6360
         TabIndex        =   5
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print Cash &Income from channelling"
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
      Begin btButtonEx.ButtonEx bttnCashRefunds 
         Height          =   375
         Left            =   6360
         TabIndex        =   8
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print Cash &Refunds"
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
      Begin btButtonEx.ButtonEx bttnDoctorPayments 
         Height          =   375
         Left            =   6360
         TabIndex        =   9
         Top             =   2520
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Print &Doctor Payments"
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
      Begin VB.Label lblCashFromCreditChanneling 
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
         Left            =   3120
         TabIndex        =   31
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Cash Cettling for Credit Patients "
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label lblAgentCashPayments 
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
         Left            =   3120
         TabIndex        =   28
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblNetCashCollection 
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
         Left            =   3000
         TabIndex        =   24
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label lblDoctorPayment 
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
         Left            =   4800
         TabIndex        =   23
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblRefund 
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
         Left            =   4800
         TabIndex        =   22
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblCashFromChanneling 
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
         Left            =   3120
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Net Cash Collection"
         Height          =   255
         Left            =   240
         TabIndex        =   20
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
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label6 
         Caption         =   "Doctor Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Expenses"
         Height          =   255
         Left            =   5040
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Income"
         Height          =   255
         Left            =   3480
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Refunds / Cancellations "
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Cash From Channeling"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "Agent Cash Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   2295
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   9975
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Today"
      TabPicture(0)   =   "FrmMyShiftSummery.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblToday"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected D&ay"
      TabPicture(1)   =   "FrmMyShiftSummery.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).Control(1)=   "Label11"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "S&elected Period"
      TabPicture(2)   =   "FrmMyShiftSummery.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label17"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label18"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DTPicker3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "DTPicker2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -72360
         TabIndex        =   2
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   160104451
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   160104451
         CurrentDate     =   39442
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   160104451
         CurrentDate     =   39442
      End
      Begin VB.Label Label18 
         Caption         =   "T&o"
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
         Left            =   3960
         TabIndex        =   27
         Top             =   480
         Width           =   495
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
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   735
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
         Left            =   -74040
         TabIndex        =   25
         Top             =   480
         Width           =   1815
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
         TabIndex        =   13
         Top             =   600
         Width           =   3975
      End
   End
   Begin MSDataListLib.DataCombo DataComboUser 
      Bindings        =   "FrmMyShiftSummery.frx":0496
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
      Style           =   2
      ListField       =   "StaffListedName"
      BoundColumn     =   "Staff_ID"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   7920
      TabIndex        =   11
      Top             =   7800
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
   Begin VB.Label Label1 
      Caption         =   "&User :"
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
      Left            =   600
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMyShiftSummery"
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
Dim Cse
Dim A
Private Sub Setcolours()
    bttnCashIncome.BackColor = BttnBackColour
    bttnCashIncome.ForeColor = BttnForeColour
    bttnPrintAgentCash.BackColor = BttnBackColour
    bttnPrintAgentCash.ForeColor = BttnForeColour
    bttnCashFromCreditChanneling.BackColor = BttnBackColour
    bttnCashFromCreditChanneling.ForeColor = BttnForeColour
    bttnCashRefunds.BackColor = BttnBackColour
    bttnCashRefunds.ForeColor = BttnForeColour
    bttnDoctorPayments.BackColor = BttnBackColour
    bttnDoctorPayments.ForeColor = BttnForeColour
    bttnPrintSummary.BackColor = BttnBackColour
    bttnPrintSummary.ForeColor = BttnForeColour
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
    FrameShiftSummary.BackColor = FrameBackColour
    FrameShiftSummary.ForeColor = FrameForeColour
    FrmShiftSummery.BackColor = FrameBackColour
    FrmShiftSummery.ForeColor = FrameForeColour
    Me.BackColor = FrameBackColour
    Me.ForeColor = FrameForeColour
End Sub

Private Sub btnMyTelephoneBookings_Click()
On Error GoTo ErrorHandler
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    Dim TemResponce As Long
    Dim RetVal As Integer
    Dim TemWhere As String
    
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    dtrShiftEndSummery.DataMember = Empty
    With DataEnvironment1.rsDayEndSummery
        If .State = 1 Then .Close
        
        TemWhere = " Where "
        
        
        If SSTab1.Tab = 0 Then
            TemWhere = TemWhere & " tblPatientFacility.bookingdate = '" & Date & "'  "
        ElseIf SSTab1.Tab = 1 Then
            TemWhere = TemWhere & " tblPatientFacility.bookingdate = '" & DTPicker1.Value & "' "
        ElseIf SSTab1.Tab = 2 Then
             TemWhere = TemWhere & " tblPatientFacility.bookingdate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "' "
        End If

        TemWhere = TemWhere & " And tblPatientFacility.user_ID = " & Val(DataComboUser.BoundText) & " "
        TemWhere = TemWhere & " And tblPatientFacility.paymentmethod_ID = 4 "
        TemWhere = TemWhere & " And tblPatientFacility.user_ID = " & Val(DataComboUser.BoundText) & " "
        
        .Source = "SELECT tblPatientFacility.*, tblStaff_Booked.StaffName as BStaffName, tblTitle.Title, tblDoctor.DoctorName, tblInstitutions.InstitutionName, tblStaff_CreditSettle.StaffName as CStaffName, tblStaff_Repay.StaffName as RStaffName FROM tblStaff AS tblStaff_CreditSettle RIGHT JOIN (tblStaff AS tblStaff_Repay RIGHT JOIN (tblInstitutions RIGHT JOIN ((tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) RIGHT JOIN (tblStaff AS tblStaff_Booked RIGHT JOIN tblPatientFacility ON tblStaff_Booked.Staff_ID = tblPatientFacility.User_ID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID) ON tblStaff_Repay.Staff_ID = tblPatientFacility.RepayUser_ID) ON tblStaff_CreditSettle.Staff_ID = tblPatientFacility.CreditSettleUser_ID " & _
        " where (bookingdate = '" & Date & "' and tblPatientFacility.user_ID = " & UserID & ") or " & _
        " (SettleCashDate = '" & Date & "' and tblPatientFacility.CreditSettleUser_ID = " & UserID & ") order by patientfacility_ID "
        .Open
    End With
    dtrShiftEndSummery.DataMember = "dayendsummery"
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Report Not Printed")
            Exit Sub
        Case FORM_SELECTED   ' 1
            If HospitalDetails = True Then
                dtrShiftEndSummery.Sections.Item("Section4").Controls.Item("lblInstitutionName").Caption = InstitutionName
                dtrShiftEndSummery.Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
            End If
            dtrShiftEndSummery.Sections.Item("Section4").Controls.Item("lblreport").Caption = "Current State of All Bookings on " & Format(Date, DefaultLongDate) & " by " & UserName
            dtrShiftEndSummery.Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
            dtrShiftEndSummery.Show
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added. Report NOT Printed", vbExclamation, "New Paper size")
            Exit Sub
    End Select
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub bttnCashFromCreditChanneling_Click()
If IsNumeric(DataComboUser.BoundText) = False Then
    Dim TemResponce As Integer
    TemResponce = MsgBox("Please select a user", vbCritical, "User?")
    DataComboUser.SetFocus
    Exit Sub
End If
With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    If SSTab1.Tab = 0 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.SettleCashDate = '" & Date & "') and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID ")
    ElseIf SSTab1.Tab = 1 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.SettleCashDate = '" & DTPicker1.Value & "') and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID")
    ElseIf SSTab1.Tab = 2 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.SettleCashDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID ")
    End If
    
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "No Transaction"): Exit Sub
    
    With dtrCredittBookingsPayment
        Set .DataSource = DataEnvironment1.rssqlTem10
        If HospitalDetails = True Then
            .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
            .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
            .Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
        Else
            .Sections("Section4").Controls.Item("RptName").Caption = Empty
            .Sections("Section4").Controls.Item("RptAddress").Caption = Empty
            .Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
        End If
        If SSTab1.Tab = 0 Then
        .Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        .Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
        ElseIf SSTab1.Tab = 1 Then
        .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
        .Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
        ElseIf SSTab1.Tab = 2 Then
        .Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
        .Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
        End If
        .Show
    End With
End With
End Sub

Private Sub bttnCashRefunds_Click()
If IsNumeric(DataComboUser.BoundText) = False Then
    Dim TemResponce As Integer
    TemResponce = MsgBox("Please select a user", vbCritical, "User?")
    DataComboUser.SetFocus
    Exit Sub
End If
With DataEnvironment1.rssqlCashireRepost
    If .State = 1 Then .Close
    If SSTab1.Tab = 0 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = 1) and ( ((tblPatientFacility.cancelled = 1) and (tblPatientFacility.RefundToPatient = 1) )or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate = '" & Date & "') and hospitalfacility_ID = 10 ")
    ElseIf SSTab1.Tab = 1 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = 1) and ( ((tblPatientFacility.cancelled = 1) and (tblPatientFacility.RefundToPatient = 1) ) or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate = '" & DTPicker1.Value & "') and hospitalfacility_ID = 10 ")
    ElseIf SSTab1.Tab = 2 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.RefundToPatient = 1) and ( ((tblPatientFacility.cancelled = 1) and (tblPatientFacility.RefundToPatient = 1) )or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and hospitalfacility_ID = 10  ")
    End If
    
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "No Transaction"): Exit Sub
    
        
    Set DataReportCashRefunds.DataSource = DataEnvironment1.rssqlCashireRepost
    If HospitalDetails = True Then
        DataReportCashRefunds.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        DataReportCashRefunds.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        DataReportCashRefunds.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    Else
        DataReportCashRefunds.Sections("Section4").Controls.Item("RptName").Caption = Empty
        DataReportCashRefunds.Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        DataReportCashRefunds.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    End If
    If SSTab1.Tab = 0 Then
    DataReportCashRefunds.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    ElseIf SSTab1.Tab = 1 Then
    DataReportCashRefunds.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
    DataReportCashRefunds.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
    End If
    
DataReportCashRefunds.Show

End With

End Sub

Private Sub bttnCashIncome_Click()
If IsNumeric(DataComboUser.BoundText) = False Then
    Dim TemResponce As Integer
    TemResponce = MsgBox("Please select a user", vbCritical, "User?")
    DataComboUser.SetFocus
    Exit Sub
End If
With DataEnvironment1.rssqlCashireRepost

    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = '" & Date & "') and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID ")
    ElseIf SSTab1.Tab = 1 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = '" & DTPicker1.Value & "')and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID")
    ElseIf SSTab1.Tab = 2 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID ")
    End If
    
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "No Transaction"): Exit Sub
  
    Set DataReportCashIncome.DataSource = DataEnvironment1.rssqlCashireRepost
    
    If HospitalDetails = True Then
        DataReportCashIncome.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        DataReportCashIncome.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        DataReportCashIncome.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    Else
        DataReportCashIncome.Sections("Section4").Controls.Item("RptName").Caption = Empty
        DataReportCashIncome.Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        DataReportCashIncome.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    End If
    If SSTab1.Tab = 0 Then
    DataReportCashIncome.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    DataReportCashIncome.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    ElseIf SSTab1.Tab = 1 Then
    DataReportCashIncome.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
    DataReportCashIncome.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
    DataReportCashIncome.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
    DataReportCashIncome.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
    End If
    
    DataReportCashIncome.Show

End With

End Sub


Private Sub bttnDetailSummary_Click()
Const PreSHape = "SHAPE {"
Const Sql = "SELECT tblPatientFacility.*, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode, tblPatientMainDetails.FirstName, tblDoctor.DoctorName FROM tblInstitutions RIGHT JOIN (tblPatientMainDetails RIGHT JOIN (tblDoctor RIGHT JOIN tblPatientFacility ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID"
Const PostSHape = "ORDER BY tblPatientFacility.PaymentMode DESC , tblPatientFacility.PatientFacility_ID}  AS sqlPDAF COMPUTE sqlPDAF, SUM(sqlPDAF.'PersonalDue') AS DocFee, SUM(sqlPDAF.'InstitutionDue') AS HosFee, ANY(sqlPDAF.'PaymentMode') AS PaymentMethodName, SUM(sqlPDAF.'TotalDue') AS TotFee BY 'PaymentMethod_Id'"
If IsNumeric(DataComboUser.BoundText) = False Then
    Dim TemResponce As Integer
    TemResponce = MsgBox("Please select a user", vbCritical, "User?")
    DataComboUser.SetFocus
    Exit Sub
End If
With DataEnvironment1
    If .rssqlPDAF_Grouping.State = 1 Then DataEnvironment1.rssqlPDAF_Grouping.Close
    Select Case SSTab1.Tab
    Case 0
        .Commands!sqlPDAF_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate = '" & Date & "')and (tblPatientFacility.User_ID = " & Val(DataComboUser.BoundText) & ") and hospitalfacility_ID = 10  " & PostSHape
        .sqlPDAF_Grouping
    Case 1
        .Commands!sqlPDAF_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate ='" & DTPicker1.Value & "')and (tblPatientFacility.User_ID = " & Val(DataComboUser.BoundText) & ") and hospitalfacility_ID = 10 " & PostSHape
        .sqlPDAF_Grouping
    Case 2
        .Commands!sqlPDAF_Grouping.CommandText = PreSHape & Sql & " Where (BookingDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and (tblPatientFacility.User_ID = " & Val(DataComboUser.BoundText) & ") and hospitalfacility_ID = 10 " & PostSHape
        .sqlPDAF_Grouping
    End Select
    If DataEnvironment1.rssqlPDAF_Grouping.RecordCount = 0 Then A = MsgBox("No Transaction to view", vbInformation + vbOKOnly, "No Transactions"): Exit Sub
    
    With dtrShifUserSummary
    
    Set .DataSource = DataEnvironment1
    
        If HospitalDetails = True Then
            .Sections("ReportHeader").Controls.Item("RptName").Caption = InstitutionName
            .Sections("ReportHeader").Controls.Item("RptAddress").Caption = InstitutionAddress
            .Sections("ReportHeader").Controls.Item("rptLHeding3").Caption = "Cashier Shift End Summary"
            .Sections("ReportFooter").Controls.Item("lblCashirerName").Caption = "Cashier Name :  " & DataComboUser.Text
        Else
            .Sections("ReportHeader").Controls.Item("RptName").Caption = Empty
            .Sections("ReportHeader").Controls.Item("RptAddress").Caption = Empty
            .Sections("ReportHeader").Controls.Item("rptLHeding3").Caption = "Cashier Shift End Summary"
            .Sections("ReportFooter").Controls.Item("lblCashirerName").Caption = "Cashier Name :  " & DataComboUser.Text
        End If
        
        If SSTab1.Tab = 0 Then
            .Sections("PageHeader").Controls.Item("rptDate").Caption = "On   " & Format(Date, DefaultLongDate)
        ElseIf SSTab1.Tab = 1 Then
            .Sections("PageHeader").Controls.Item("rptDate").Caption = "On   " & DTPicker1.Value
        ElseIf SSTab1.Tab = 2 Then
            .Sections("PageHeader").Controls.Item("rptDate").Caption = "Date From   " & DTPicker2.Value & "   To   " & DTPicker3.Value
        End If
        
        .Show
    End With
    
End With
End Sub

Private Sub bttnDoctorPayments_Click()
If IsNumeric(DataComboUser.BoundText) = False Then
    Dim TemResponce As Integer
    TemResponce = MsgBox("Please select a user", vbCritical, "User?")
    DataComboUser.SetFocus
    Exit Sub
End If
With DataEnvironment1.rssqlTem9
    If .State = 1 Then .Close
    If SSTab1.Tab = 0 Then
        .Open ("Select tblstaffpayment.*,tbldoctor.* From tblstaffpayment Left Join tbldoctor On tblstaffpayment.Staff_ID = tbldoctor.Doctor_ID Where  (tblstaffpayment.User_ID = " & Val(DataComboUser.BoundText) & " and (tblstaffpayment.PaidDate = '" & Date & "' ))Order by StaffPayment_ID")
    ElseIf SSTab1.Tab = 1 Then
        .Open ("Select tblstaffpayment.*,tbldoctor.* From tblstaffpayment Left Join tbldoctor On tblstaffpayment.Staff_ID = tbldoctor.Doctor_ID Where (tblstaffpayment.User_ID = " & Val(DataComboUser.BoundText) & " and (tblstaffpayment.PaidDate = '" & DTPicker1.Value & "'))Order by StaffPayment_ID")
    ElseIf SSTab1.Tab = 2 Then
        .Open ("Select tblstaffpayment.*,tbldoctor.* From tblstaffpayment Left Join tbldoctor On tblstaffpayment.Staff_ID = tbldoctor.Doctor_ID Where (tblstaffpayment.User_ID = " & Val(DataComboUser.BoundText) & " and (tblstaffpayment.PaidDate between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "'))Order by StaffPayment_ID")
    End If
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "No Transaction"): Exit Sub
    Set DataReportDoctorPayment.DataSource = DataEnvironment1.rssqlTem9
    If HospitalDetails = True Then
        DataReportDoctorPayment.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        DataReportDoctorPayment.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        DataReportDoctorPayment.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    Else
        DataReportDoctorPayment.Sections("Section4").Controls.Item("RptName").Caption = Empty
        DataReportDoctorPayment.Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        DataReportDoctorPayment.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    End If
    If SSTab1.Tab = 0 Then
    DataReportDoctorPayment.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    ElseIf SSTab1.Tab = 1 Then
    DataReportDoctorPayment.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
    DataReportDoctorPayment.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
    End If
    DataReportDoctorPayment.Show
End With
End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub



Private Sub bttnPrintAgentCash_Click()
If IsNumeric(DataComboUser.BoundText) = False Then
    Dim TemResponce As Integer
    TemResponce = MsgBox("Please select a user", vbCritical, "User?")
    DataComboUser.SetFocus
    Exit Sub
End If
With DataEnvironment1.rssqlAgentPayment1
    If .State = 1 Then .Close
    If SSTab1.Tab = 0 Then
        .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate = '" & Date & "')   ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    ElseIf SSTab1.Tab = 1 Then
         .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate = '" & DTPicker1.Value & "') ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    ElseIf SSTab1.Tab = 2 Then
        .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.User_ID = " & DataComboUser.BoundText & ") and (tblAgentCashSettle.SettledDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    End If
    If .RecordCount = 0 Then A = MsgBox("No Transaction to view this User", vbInformation + vbOKOnly, "No Transaction"): Exit Sub
    Set dtrAgentCashReceive.DataSource = DataEnvironment1.rssqlAgentPayment1
    If HospitalDetails = True Then
        dtrAgentCashReceive.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        dtrAgentCashReceive.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        dtrAgentCashReceive.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    Else
        dtrAgentCashReceive.Sections("Section4").Controls.Item("RptName").Caption = Empty
        dtrAgentCashReceive.Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        dtrAgentCashReceive.Sections("Section2").Controls.Item("RptCashierName").Caption = DataComboUser.Text
    End If
    If SSTab1.Tab = 0 Then
        dtrAgentCashReceive.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        dtrAgentCashReceive.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    ElseIf SSTab1.Tab = 1 Then
        dtrAgentCashReceive.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker1.Value
        dtrAgentCashReceive.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker1.Value
    ElseIf SSTab1.Tab = 2 Then
        dtrAgentCashReceive.Sections("Section2").Controls.Item("rptFromdate").Caption = DTPicker2.Value
        dtrAgentCashReceive.Sections("Section2").Controls.Item("RptToDate").Caption = DTPicker3.Value
    End If
    dtrAgentCashReceive.Show
End With
End Sub

Private Sub ClearValues()
    lblCashFromChanneling.Caption = "0.00"
    lblAgentCashPayments.Caption = "0.00"
    lblRefund.Caption = "0.00"
    lblDoctorPayment.Caption = "0.00"
    lblNetCashCollection.Caption = "0.00"
    lblCashFromCreditChanneling = "0.00"
End Sub

Private Sub bttnPrintSummary_Click()
If IsNumeric(DataComboUser.BoundText) = False Then
    Dim TemResponce As Integer
    TemResponce = MsgBox("Please select a user", vbCritical, "User?")
    DataComboUser.SetFocus
    Exit Sub
End If
With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    .Open "Select * From tblTem"
    Set dtrShiftEndCash.DataSource = DataEnvironment1.rssqlTemSu1
End With
With dtrShiftEndCash
    If HospitalDetails = True Then
        .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    Else
        .Sections("Section4").Controls.Item("RptName").Caption = Empty
        .Sections("Section4").Controls.Item("RptAddress").Caption = Empty
    End If
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
'Call CalculateValues
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
    On Error GoTo Error_Han
    DataComboUser.RowMember = Empty
    With DataEnvironment1.rssqlTem17
        If .State = 1 Then .Close
        .Source = "Select * From tblStaff Order by StaffName "
        DataComboUser.RowMember = "SQLTEM17"
    End With
    DataComboUser.BoundText = UserID
    DTPicker1 = Date
    DTPicker2 = Date
    DTPicker3 = Date
    SSTab1.Tab = 0
    If UserAuthority <> AuthorityOwner Then
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
    End If
    
    Exit Sub
Error_Han:
    A = MsgBox(Err.Number & vbNewLine & Err.Description, vbInformation + vbOKOnly, "Loding Error")
    Call Setcolours
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Call CalculateValues
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
lblDoctorPayment.Caption = Format(TemDoctorPayment, "0.00")

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
     TemDoctorPayment = TemDoctorPayment + !PaidAmount
     .MoveNext
    Loop
   
    If .State = 1 Then .Close
End With

lblDoctorPayment.Caption = Format(TemDoctorPayment, "0.00")

End Sub

Private Sub CashReceiveFromCreditBooking()

TemCashFromCreditCahnneling = 0
lblCashFromCreditChanneling.Caption = Format(TemCashFromCreditCahnneling, "0.00")

With DataEnvironment1.rssqlTem

    If .State = 1 Then .Close

    If SSTab1.Tab = 0 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.SettleCashDate = '" & Date & "') and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID  ")
    ElseIf SSTab1.Tab = 1 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.SettleCashDate = '" & DTPicker1.Value & "')and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID ")
    ElseIf SSTab1.Tab = 2 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.SettleCashDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "')and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID ")
    End If
   
    If .RecordCount = 0 Then Exit Sub
    
    While .EOF = False
        TemCashFromCreditCahnneling = Val(TemCashFromCreditCahnneling) + Val(!PersonalFee) + Val(!InstitutionFee) + Val(!otherfee)
        .MoveNext
    Wend
    
    

End With
lblCashFromCreditChanneling.Caption = Format(TemCashFromCreditCahnneling, "0.00")

End Sub

Private Sub AgentCashReceive()
TemAgentCash = 0
lblAgentCashPayments.Caption = Format(TemAgentCash, "###0.00")

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
    TemAgentCash = Val(TemAgentCash) + Val(!Cash)
    .MoveNext
    Loop
    
End With
lblAgentCashPayments.Caption = Format(TemAgentCash, "###0.00")
End Sub

Private Sub ChannelingCashIncome()

If IsNumeric(DataComboUser.BoundText) = False Then Exit Sub
CashFromCahnneling = 0
lblCashFromChanneling.Caption = Format(CashFromCahnneling, "0.00")

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    If SSTab1.Tab = 0 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = '" & Date & "') and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID ")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 1 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = '" & DTPicker1.Value & "') and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID")
    If .RecordCount = 0 Then Exit Sub
    ElseIf SSTab1.Tab = 2 Then
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & DataComboUser.BoundText & ") and (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "') and hospitalfacility_ID = 10  order by tblPatientFacility.patientfacility_ID ")
    If .RecordCount = 0 Then Exit Sub
    End If
    
    While .EOF = False
        CashFromCahnneling = CashFromCahnneling + !PersonalFee + !InstitutionFee + !otherfee
        .MoveNext
    Wend
End With

lblCashFromChanneling.Caption = Format(CashFromCahnneling, "0.00")

End Sub

Private Sub CashRefund()
temCashRefund = 0
lblRefund.Caption = Format(temCashRefund, "0.00")


With DataEnvironment1.rssqlTemSu1
    If .State = 1 Then .Close
    
    If SSTab1.Tab = 0 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ")and hospitalfacility_ID = 10  and (tblPatientFacility.RefundToPatient = 1) and ( (tblPatientFacility.cancelled = 1)or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate = '" & Date & "')")
    ElseIf SSTab1.Tab = 1 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ")and hospitalfacility_ID = 10  and (tblPatientFacility.RefundToPatient = 1) and ( (tblPatientFacility.cancelled = 1)or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate = '" & DTPicker1.Value & "')")
    ElseIf SSTab1.Tab = 2 Then
        .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & DataComboUser.BoundText & ")and hospitalfacility_ID = 10  and (tblPatientFacility.RefundToPatient = 1) and ( (tblPatientFacility.cancelled = 1)or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate Between  '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "')")
    End If
    
    If .RecordCount = 0 Then Exit Sub
    
     Do While .EOF = False
           If IsNull(!institutionrefund) = False Then temCashRefund = temCashRefund + (!institutionrefund)
           If IsNull(!Personalrefund) = False Then temCashRefund = temCashRefund + (!Personalrefund)
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
lblNetCashCollection.Caption = Format(TemNetChannelingincome, "0.00")

TemTotalCash = (CashFromCahnneling + TemAgentCash + TemCashFromCreditCahnneling)
TemTotalPayment = (temCashRefund + TemDoctorPayment)

TemNetChannelingincome = Val(TemTotalCash) - Val(TemTotalPayment)

lblNetCashCollection.Caption = Format(TemNetChannelingincome, "0.00")

End Sub
