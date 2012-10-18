VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAppointments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Appointments"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
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
   ScaleHeight     =   7185
   ScaleWidth      =   6810
   Begin VB.Frame Frame1 
      Caption         =   "Appointments"
      Height          =   2895
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   6255
      Begin VB.Label lblCancelledPatient 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblCancelledAgent 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "Paid to patient"
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Paid to agent"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblCashBookings 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblCreditBookings 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblAgentBookings 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Cash"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Credit Settling"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Agent"
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Total Bookings"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Cancelled"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Refunds"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblCancelled 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblRefunded 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Top             =   2280
         Width           =   2175
      End
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Exit"
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
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Selected Day"
      TabPicture(0)   =   "frmAppointments.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MonthView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected Period"
      TabPicture(1)   =   "frmAppointments.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MonthView3"
      Tab(1).Control(1)=   "MonthView2"
      Tab(1).ControlCount=   2
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2820
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   61997057
         CurrentDate     =   39525
      End
      Begin MSComCtl2.MonthView MonthView2 
         Height          =   2820
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   61997057
         CurrentDate     =   39525
      End
      Begin MSComCtl2.MonthView MonthView3 
         Height          =   2820
         Left            =   -71760
         TabIndex        =   3
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   61997057
         CurrentDate     =   39525
      End
   End
End
Attribute VB_Name = "frmAppointments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemSelect As String
Dim TemWhere As String
Dim TemOrder As String
Dim temSql As String


Private Sub Calculate()
If SSTab1.Tab = 0 Then
    Frame1.Caption = "Appointemnts on " & Format(MonthView1.Value, "DD MMMM YYYY")
Else
    Frame1.Caption = "Appointments from " & Format(MonthView2.Value, "DD MMMM YYYY") & " to " & Format(MonthView3.Value, "DD MMMM YYYY")
End If

Me.MousePointer = vbHourglass
DoEvents

With DataEnvironment1.rssqlTem
    TemSelect = "SELECT Count(tblPatientFacility.PatientFacility_ID) AS TotalPatients FROM tblPatientFacility "
    If SSTab1.Tab = 0 Then
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView1.Value & "' And '" & MonthView1.Value & "') and (FullyPaid = 1) )"
    Else
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView2.Value & "' And '" & MonthView3.Value & "')and (FullyPaid = 1))"
    End If
    TemOrder = ""
    temSql = TemSelect & TemWhere & TemOrder
    If .State = 1 Then .Close
    .Source = temSql
    .Open
    If Not IsNull(!TotalPatients) Then
        lblTotal.Caption = !TotalPatients
    Else
        lblTotal.Caption = 0
    End If
End With

With DataEnvironment1.rssqlTem
    TemSelect = "SELECT Count(tblPatientFacility.PatientFacility_ID) AS TotalPatients FROM tblPatientFacility "
    If SSTab1.Tab = 0 Then
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView1.Value & "' And '" & MonthView1.Value & "') AND ((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.PaymentMode)='Credit'))"
    Else
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView2.Value & "' And '" & MonthView3.Value & "') AND ((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.PaymentMode)='Credit'))"
    End If
    TemOrder = ""
    temSql = TemSelect & TemWhere & TemOrder
    If .State = 1 Then .Close
    .Source = temSql
    .Open
    If Not IsNull(!TotalPatients) Then
        lblCreditBookings.Caption = !TotalPatients
    Else
        lblCreditBookings.Caption = 0
    End If
End With

With DataEnvironment1.rssqlTem
    TemSelect = "SELECT Count(tblPatientFacility.PatientFacility_ID) AS TotalPatients FROM tblPatientFacility "
    If SSTab1.Tab = 0 Then
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView1.Value & "' And '" & MonthView1.Value & "') AND ((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.PaymentMode)='Agent'))"
    Else
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView2.Value & "' And '" & MonthView3.Value & "') AND ((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.PaymentMode)='Agent'))"
    End If
    TemOrder = ""
    temSql = TemSelect & TemWhere & TemOrder
    If .State = 1 Then .Close
    .Source = temSql
    .Open
    If Not IsNull(!TotalPatients) Then
        lblAgentBookings.Caption = !TotalPatients
    Else
        lblAgentBookings.Caption = 0
    End If
End With

With DataEnvironment1.rssqlTem
    TemSelect = "SELECT Count(tblPatientFacility.PatientFacility_ID) AS TotalPatients FROM tblPatientFacility "
    If SSTab1.Tab = 0 Then
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView1.Value & "' And '" & MonthView1.Value & "') AND ((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.PaymentMode)='Cash'))"
    Else
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView2.Value & "' And '" & MonthView3.Value & "') AND ((tblPatientFacility.FullyPaid)=1) AND ((tblPatientFacility.PaymentMode)='Cash'))"
    End If
    TemOrder = ""
    temSql = TemSelect & TemWhere & TemOrder
    If .State = 1 Then .Close
    .Source = temSql
    .Open
    If Not IsNull(!TotalPatients) Then
        lblCashBookings.Caption = !TotalPatients
    Else
        lblCashBookings.Caption = 0
    End If
End With


With DataEnvironment1.rssqlTem
    TemSelect = "SELECT Count(tblPatientFacility.PatientFacility_ID) AS TotalPatients FROM tblPatientFacility "
    If SSTab1.Tab = 0 Then
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView1.Value & "' And '" & MonthView1.Value & "') AND ((tblPatientFacility.Cancelled)=1))"
    Else
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView2.Value & "' And '" & MonthView3.Value & "') AND ((tblPatientFacility.Cancelled)=1))"
    End If
    TemOrder = ""
    temSql = TemSelect & TemWhere & TemOrder
    If .State = 1 Then .Close
    .Source = temSql
    .Open
    If Not IsNull(!TotalPatients) Then
        lblCancelled.Caption = !TotalPatients
    Else
        lblCancelled.Caption = 0
    End If
End With

With DataEnvironment1.rssqlTem
    TemSelect = "SELECT Count(tblPatientFacility.PatientFacility_ID) AS TotalPatients FROM tblPatientFacility "
    If SSTab1.Tab = 0 Then
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView1.Value & "' And '" & MonthView1.Value & "') AND ((tblPatientFacility.Cancelled)=1) AND ((tblPatientFacility.RefundToPatient)=1))"
    Else
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView2.Value & "' And '" & MonthView3.Value & "') AND ((tblPatientFacility.Cancelled)=1) AND ((tblPatientFacility.RefundToPatient)=1))"
    End If
    TemOrder = ""
    temSql = TemSelect & TemWhere & TemOrder
    If .State = 1 Then .Close
    .Source = temSql
    .Open
    If Not IsNull(!TotalPatients) Then
        lblCancelledPatient.Caption = !TotalPatients
    Else
        lblCancelledPatient.Caption = 0
    End If
End With

With DataEnvironment1.rssqlTem
    TemSelect = "SELECT Count(tblPatientFacility.PatientFacility_ID) AS TotalPatients FROM tblPatientFacility "
    If SSTab1.Tab = 0 Then
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView1.Value & "' And '" & MonthView1.Value & "') AND ((tblPatientFacility.Cancelled)=1) AND ((tblPatientFacility.RefundToAgent)=1))"
    Else
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView2.Value & "' And '" & MonthView3.Value & "') AND ((tblPatientFacility.Cancelled)=1) AND ((tblPatientFacility.RefundToAgent)=1))"
    End If
    TemOrder = ""
    temSql = TemSelect & TemWhere & TemOrder
    If .State = 1 Then .Close
    .Source = temSql
    .Open
    If Not IsNull(!TotalPatients) Then
        lblCancelledAgent.Caption = !TotalPatients
    Else
        lblCancelledAgent.Caption = 0
    End If
End With


With DataEnvironment1.rssqlTem
    TemSelect = "SELECT Count(tblPatientFacility.PatientFacility_ID) AS TotalPatients FROM tblPatientFacility "
    If SSTab1.Tab = 0 Then
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView1.Value & "' And '" & MonthView1.Value & "') AND ((tblPatientFacility.Refund)=1))"
    Else
        TemWhere = " WHERE (((tblPatientFacility.AppointmentDate) Between '" & MonthView2.Value & "' And '" & MonthView3.Value & "') AND ((tblPatientFacility.Refund)=1))"
    End If
    TemOrder = ""
    temSql = TemSelect & TemWhere & TemOrder
    If .State = 1 Then .Close
    .Source = temSql
    .Open
    If Not IsNull(!TotalPatients) Then
        lblRefunded.Caption = !TotalPatients
    Else
        lblRefunded.Caption = 0
    End If
End With

Me.MousePointer = vbDefault
DoEvents

End Sub

Private Sub ButtonEx1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    MonthView1.Value = Date
    MonthView2.Value = Date
    MonthView3.Value = Date
    Call Calculate
    If UserAuthority <> AuthorityOwner Then
        SSTab1.TabVisible(1) = False
        MonthView1.Enabled = False
        SSTab1.TabCaption(0) = "Today"
    End If
    
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Call Calculate
End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
    Call Calculate
End Sub

Private Sub MonthView3_DateClick(ByVal DateClicked As Date)
    Call Calculate
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call Calculate
End Sub

