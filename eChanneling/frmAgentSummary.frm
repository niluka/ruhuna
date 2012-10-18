VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmNewAgentSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agent Summary"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAgentSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   8955
   Begin btButtonEx.ButtonEx bttnAgentSummary 
      Height          =   375
      Left            =   2520
      TabIndex        =   26
      Top             =   7200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Agent Summary"
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
   Begin btButtonEx.ButtonEx bttnPrintSummary 
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   7200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Print Summary"
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
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   8415
      Begin MSDataListLib.DataCombo dtcAgent 
         Height          =   360
         Left            =   2280
         TabIndex        =   17
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label16 
         Caption         =   "Agent Name"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
   End
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   7200
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.Frame Frame1 
      Caption         =   "Date"
      Height          =   4215
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   7815
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   5280
         TabIndex        =   13
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Balance"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Line Line3 
         X1              =   5280
         X2              =   6840
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label lblTotalBooking 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   5280
         TabIndex        =   11
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Total Bookings"
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Line Line2 
         X1              =   3000
         X2              =   4200
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label lblRefund 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Refunds"
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblBooking 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Less   :"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Bookings"
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Total Cash"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblTotalCash 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   3000
         X2              =   4200
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblCashSettle 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Cash Receive"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   240
      TabIndex        =   18
      Top             =   1200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10398
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Today"
      TabPicture(0)   =   "frmAgentSummary.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDate"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Selected Day"
      TabPicture(1)   =   "frmAgentSummary.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPicker1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Period"
      TabPicture(2)   =   "frmAgentSummary.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DTPicker3"
      Tab(2).Control(1)=   "DTPicker2"
      Tab(2).Control(2)=   "Label4"
      Tab(2).Control(3)=   "Label3"
      Tab(2).ControlCount=   4
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   -69960
         TabIndex        =   25
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   156958721
         CurrentDate     =   39508
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -73560
         TabIndex        =   24
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   156958721
         CurrentDate     =   39508
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -71880
         TabIndex        =   23
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   156958721
         CurrentDate     =   39508
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         Height          =   375
         Left            =   -70440
         TabIndex        =   22
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Date From"
         Height          =   255
         Left            =   -74640
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblDate 
         Caption         =   "Label1"
         Height          =   375
         Left            =   3120
         TabIndex        =   20
         Top             =   600
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmNewAgentSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemtotalBooking As Double
Dim TemtotalRefund As Double
Dim TemCashSettle As Double

Private Sub FillAgent()
With DataEnvironment1.rscmmdTemAgent
    If .State = 1 Then .Close
    .Open "Select* From tblInstitutions Order By InstitutionName"
    Set dtcAgent.RowSource = DataEnvironment1.rscmmdTemAgent
    dtcAgent.BoundColumn = "Institution_ID"
    dtcAgent.ListField = "InstitutionName"

End With

End Sub

Private Sub bttnAgentSummary_Click()
If IsNumeric(dtcAgent.BoundText) = False Then Exit Sub

With DataEnvironment1
    If .rssqlTem10.State = 1 Then .rssqlTem10.Close
    
    .rssqlTem10.Open "Delete* From tblAgentDetail"
    
    If .rssqlTem10.State = 1 Then .rssqlTem10.Close
    .rssqlTem10.Open "Select* From tblAgentDetail"
    
    If .rssqlTem1.State = 1 Then .rssqlTem1.Close
    
    ''''''''''Agent Balance'''''''''''
    
''    .rssqlTem1.Open "Select* From tblInstitutionBalance Where Institution_Id = " & dtcAgent.BoundText & ""
''    MsgBox .rssqlTem1.RecordCount
''    If .rssqlTem1.State = 1 Then .rssqlTem1.Close
''
''    Do While .rssqlTem1.EOF = False
''
''
''    .rssqlTem1.MoveNext
''    Loop

    ''''''''''Case Settle'''''''''''''''
    .rssqlTem1.Open "SELECT Sum(tblAgentCashSettle.Cash) AS TotalCash, tblAgentCashSettle.Institution_ID, tblAgentCashSettle.SettledDate From tblAgentCashSettle Where (Institution_ID =" & dtcAgent.BoundText & ") GROUP BY tblAgentCashSettle.Institution_ID, tblAgentCashSettle.SettledDate ORDER BY tblAgentCashSettle.SettledDate"
    
        Do While .rssqlTem1.EOF = False
'            MsgBox "Cash Settle  " & .rssqlTem1.RecordCount & vbNewLine & .rssqlTem1!TotalCash & "    " & .rssqlTem1![SettledDate]
            
            .rssqlTem10.AddNew
            .rssqlTem10!Agent_ID = dtcAgent.BoundText
            .rssqlTem10!AgentName = dtcAgent.Text
            .rssqlTem10!CashSettle = .rssqlTem1![TotalCash]
            .rssqlTem10!Date = .rssqlTem1![SettledDate]
            .rssqlTem10.Update
            
            .rssqlTem1.MoveNext
        Loop
        
    If .rssqlTem1.State = 1 Then .rssqlTem1.Close
    
    '''''''''''Find Agent Booing'''''''''
    
    .rssqlTem1.Open "SELECT Sum(tblPatientFacility.TotalFee) AS AgentBooking, tblPatientFacility.Agent_ID, tblPatientFacility.BookingDate From tblPatientFacility WHERE (((Agent_ID) = " & dtcAgent.BoundText & ")and((tblPatientFacility.PaymentMode)='Agent'))GROUP BY tblPatientFacility.Agent_ID, tblPatientFacility.BookingDate ORDER BY tblPatientFacility.BookingDate"
    
        Do While .rssqlTem1.EOF = False
'            MsgBox "booking    " & .rssqlTem1.RecordCount & vbNewLine & .rssqlTem1!AgentBooking & "   " & .rssqlTem1![BookingDate]
            .rssqlTem10.AddNew
            .rssqlTem10!Agent_ID = dtcAgent.BoundText
            .rssqlTem10!AgentName = dtcAgent.Text
            .rssqlTem10!Booking = .rssqlTem1![AgentBooking]
            .rssqlTem10!Date = .rssqlTem1![BookingDate]
            .rssqlTem10.Update

            .rssqlTem1.MoveNext
        Loop
    
    If .rssqlTem1.State = 1 Then .rssqlTem1.Close
    
    '''''''''Find Agent Refund'''''''''''
    
     .rssqlTem1.Open "SELECT Sum(tblPatientFacility.TotalFee) AS AgentRefund, tblPatientFacility.Agent_ID, tblPatientFacility.BookingDate From tblPatientFacility WHERE (((Agent_ID) = " & dtcAgent.BoundText & ")and((tblPatientFacility.PaymentMode)='Agent') AND ((tblPatientFacility.Cancelled)=1) AND ((tblPatientFacility.RefundToAgent)=1))GROUP BY tblPatientFacility.Agent_ID, tblPatientFacility.BookingDate ORDER BY tblPatientFacility.BookingDate"
'    If .rssqlTem1.RecordCount = 0 Then Exit Sub
    
        Do While .rssqlTem1.EOF = False
'            MsgBox "Refund    " & .rssqlTem1.RecordCount & vbNewLine & .rssqlTem1!AgentRefund & "   " & .rssqlTem1![BookingDate]
            .rssqlTem10.AddNew
            .rssqlTem10!Agent_ID = dtcAgent.BoundText
            .rssqlTem10!AgentName = dtcAgent.Text
            .rssqlTem10!Refund = .rssqlTem1![AgentRefund]
            .rssqlTem10!Date = .rssqlTem1![BookingDate]
'            .rssqlTem10!DayEndBalance = dsdcs
            .rssqlTem10.Update

            .rssqlTem1.MoveNext
        Loop
        
    If .rssqlTem10.State = 1 Then .rssqlTem10.Close
    
    If .rssqlTem1.State = 1 Then .rssqlTem1.Close
    
    If .rscmdAgentDetailsView_Grouping.State = 1 Then .rscmdAgentDetailsView_Grouping.Close
    
    .rscmdAgentDetailsView_Grouping.Open " SHAPE {Select* From tblAgentDetail}  AS cmdAgentDetailsView COMPUTE cmdAgentDetailsView, SUM(cmdAgentDetailsView.'CashSettle') AS TotalCashSettle, SUM(cmdAgentDetailsView.'Booking') AS TotalBooking, SUM(cmdAgentDetailsView.'Refund') AS TotalRefund BY 'AgentName','Date'"
    Set dtrAgentDetailViewSummary.DataSource = DataEnvironment1
    With dtrAgentDetailViewSummary
    .Sections("PageHeader").Controls("lblAgentName").Caption = "Agent Name     " & dtcAgent.Text
'    .Sections("ReportFooter").Controls("lblBalance").Caption = (DataEnvironment1.rscmdAgentDetailsView_Grouping!TotalCashSettle + DataEnvironment1.rscmdAgentDetailsView_Grouping!TotalBooking) - (DataEnvironment1.rscmdAgentDetailsView_Grouping!TotalRefund)
'   Function3
    .Show
    End With
End With
End Sub

Private Sub bttnPrintSummary_Click()
    With DataEnvironment1.rssqlTemSu1
        If .State = 1 Then .Close
        .Open "Select * From tblTem"
        Set dtrNewAgentSummary.DataSource = DataEnvironment1.rssqlTemSu1
    End With
    
    With dtrNewAgentSummary
    
        If HospitalDetails = True Then
            .Sections("Section4").Controls.Item("lblinstitutionname").Caption = InstitutionName
            .Sections("Section4").Controls.Item("lblinstitutionaddress").Caption = InstitutionAddress
        End If
        .Sections("Section4").Controls.Item("lblreport").Caption = "Agent Summery"
        .Sections("Section4").Controls.Item("lblreportsub").Caption = Format(Date, DefaultLongDate)
'        .Sections("Section2").Controls.Item("lblcash").Caption = lblCashBookings.Caption
'        .Sections("Section2").Controls.Item("lblcredit").Caption = lblSettlingCredit.Caption
'        .Sections("Section2").Controls.Item("lblNetTotalBooking").Caption = lblAgentCashPayments.Caption
        .Sections("Section2").Controls.Item("lblTotalCash").Caption = lblTotalCash.Caption
        .Sections("Section2").Controls.Item("lblTotalbooking").Caption = lblBooking.Caption
        .Sections("Section2").Controls.Item("lblTotalRefund").Caption = lblRefund.Caption
        
        .Sections("Section2").Controls.Item("lblNetTotalBooking").Caption = lblTotalBooking.Caption
        .Sections("Section2").Controls.Item("lblnetcash").Caption = lblBalance.Caption
        .Show
        
    End With

End Sub

Private Sub ButtonEx1_Click()
Unload Me
End Sub

Private Sub dtcAgent_Click(Area As Integer)
If IsNumeric(dtcAgent.BoundText) = False Then Exit Sub
Call FindDetails
End Sub

Private Sub DTPicker1_Change()
If IsNumeric(dtcAgent.BoundText) = False Then Exit Sub
Call FindDetails

End Sub

Private Sub DTPicker2_Change()
If IsNumeric(dtcAgent.BoundText) = False Then Exit Sub
Call FindDetails

End Sub

Private Sub DTPicker3_Change()
If IsNumeric(dtcAgent.BoundText) = False Then Exit Sub
Call FindDetails

End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
lblDate.Caption = Date
DTPicker1.Value = Date
DTPicker2.Value = Date
DTPicker3.Value = Date
Call FillAgent
    If UserAuthority <> AuthorityOwner Then
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
    End If

End Sub

Private Sub FindDetails()
TemtotalBooking = 0
TemtotalRefund = 0
TemCashSettle = 0
lblBooking.Caption = "0.00"
lblRefund.Caption = "0.00"
lblTotalBooking = "0.00"
lblCashSettle.Caption = "0.00"
lblTotalCash.Caption = "0.00"

With DataEnvironment1.rssqlTem1
''''''''''''Booking'''''''''
    If .State = 1 Then .Close
    Select Case SSTab1.Tab
    
    Case 0
    .Open "SELECT* From  tblPatientFacility Where Agent_ID = " & dtcAgent.BoundText & " and PaymentMode = 'Agent'and BookingDate ='" & Date & "'"
    Case 1
    .Open "SELECT* From  tblPatientFacility Where Agent_ID = " & dtcAgent.BoundText & " and PaymentMode = 'Agent'and BookingDate ='" & DTPicker1.Value & "'"
    Case 2
    .Open "SELECT* From  tblPatientFacility Where Agent_ID = " & dtcAgent.BoundText & " and PaymentMode = 'Agent'and BookingDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "' "
    End Select
    
    Do While .EOF = False
        TemtotalBooking = Val(TemtotalBooking) + Val(!totalfee)
        .MoveNext
    Loop
    
    lblBooking.Caption = Format(TemtotalBooking, "0.00")
    
'''''''''''Refunds''''''''''''
    If .State = 1 Then .Close
    Select Case SSTab1.Tab
    Case 0
    .Open "SELECT* From  tblPatientFacility Where Agent_ID = " & dtcAgent.BoundText & " and PaymentMode = 'Agent' and cancelled = 1  and RefundToAgent = 1 and RepayDate ='" & Date & "'"
    Case 1
    .Open "SELECT* From  tblPatientFacility Where Agent_ID = " & dtcAgent.BoundText & " and PaymentMode = 'Agent' and cancelled = 1  and RefundToAgent = 1 and RepayDate = '" & DTPicker1.Value & "' "
    Case 2
    .Open "SELECT* From  tblPatientFacility Where Agent_ID = " & dtcAgent.BoundText & " and PaymentMode = 'Agent' and cancelled = 1  and RefundToAgent = 1 and RepayDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "' "
    End Select
    
    Do While .EOF = False
        TemtotalRefund = Val(TemtotalRefund) + Val(!totalfee)
        .MoveNext
    Loop
    
    lblRefund.Caption = Format(TemtotalRefund, "0.00")
    lblTotalBooking = Format(TemtotalBooking - TemtotalRefund, "0.00")
   
''''''''Cash Settle''''''''''''
    If .State = 1 Then .Close
    Select Case SSTab1.Tab
    
    Case 0
    .Open "Select* From tblAgentCashSettle Where Institution_ID = " & dtcAgent.BoundText & " and SettledDate = '" & Date & "'"
    Case 1
    .Open "Select* From tblAgentCashSettle Where Institution_ID = " & dtcAgent.BoundText & " and SettledDate = '" & DTPicker1.Value & "'"
    Case 2
    .Open "Select* From tblAgentCashSettle Where Institution_ID = " & dtcAgent.BoundText & " and SettledDate Between '" & DTPicker2.Value & "' and '" & DTPicker3.Value & "'"
    End Select
    
    Do While .EOF = False
    
    TemCashSettle = TemCashSettle + !Cash
    .MoveNext
    Loop
    
    If .State = 1 Then .Close
    
    lblCashSettle.Caption = Format(TemCashSettle, "0.00")
    lblTotalCash.Caption = Format(TemCashSettle, "0.00")
End With

'''''''Balance'''''''''''''''''
    lblBalance.Caption = Format(Val(lblTotalCash.Caption) - Val(lblTotalBooking.Caption), "0.00")
    
End Sub

