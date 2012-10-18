VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmSingleAgentTransactions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transactions of a Single Agents of a Selected Period"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7545
   Begin MSDataListLib.DataCombo DataComboAgent 
      Height          =   360
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6165
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin btButtonEx.ButtonEx bttnExit 
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   157155331
      CurrentDate     =   39515
   End
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnSearch 
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Serch"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   157155331
      CurrentDate     =   39515
   End
   Begin MSDataListLib.DataCombo DataComboAgentCode 
      Height          =   360
      Left            =   4800
      TabIndex        =   10
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label3 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Agent"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSingleAgentTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bttnExit_Click()
Unload Me
End Sub

Private Sub bttnPrint_Click()
With dtrSingleAgentDetails
    .Sections("Section4").Controls("lblDate").Caption = "Date From   :  " & DTPicker1.Value & "Date From   :  " & DTPicker2.Value
    .Sections("Section4").Controls("lblReportSub").Caption = "Agent Name  :  " & DataComboAgent.Text

    With DataEnvironment1.rssqlTem11
        If .State = 1 Then .Close
        .Open "Select* From tblTemSingleAgentTransctions Order By Date"
    
    End With
        Set dtrSingleAgentDetails.DataSource = DataEnvironment1.rssqlTem11
    
        .Show
End With
End Sub

Private Sub WriteToTemTable()
Dim TemDate As Date
Dim TemDate1 As Date
Dim TemDate2 As Date

If DTPicker1.Value > DTPicker2.Value Then
    TemDate1 = DTPicker2.Value
    TemDate2 = DTPicker1.Value
Else
    TemDate2 = DTPicker2.Value
    TemDate1 = DTPicker1.Value
End If

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "Truncate Table tblTemSingleAgentTransctions"
    .Open
    If .State = 1 Then .Close
    .Source = "SELECT tblTemSingleAgentTransctions.* FROM tblTemSingleAgentTransctions"
    .Open
    TemDate = TemDate1
    While TemDate < TemDate2 + 1
        .AddNew
        !Date = TemDate
        .Update
        TemDate = TemDate + 1
    Wend

    If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
    DataEnvironment1.rssqlTem1.Source = "SELECT tblInstitutionBalance.* From tblInstitutionBalance Where Date between '" & TemDate1 & "' and '" & TemDate2 & "' and institution_ID =" & Val(DataComboAgent.BoundText)
    DataEnvironment1.rssqlTem1.Open
    If DataEnvironment1.rssqlTem1.RecordCount = 0 Then Exit Sub
    While DataEnvironment1.rssqlTem1.EOF = False
        If .State = 1 Then .Close
        .Source = "SELECT tblTemSingleAgentTransctions.* FROM tblTemSingleAgentTransctions where Date = '" & DataEnvironment1.rssqlTem1!Date & "'"
        .Open
        If .RecordCount <> 0 Then
            !StartingBalance = DataEnvironment1.rssqlTem1!SBalance
            !EndingBalance = DataEnvironment1.rssqlTem1!EBalance
            .Update
        End If
        DataEnvironment1.rssqlTem1.MoveNext
    Wend
    
   
    If .State = 1 Then .Close
    .Source = "SELECT tblTemSingleAgentTransctions.* FROM tblTemSingleAgentTransctions"
    .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    While .EOF = False
        If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
        DataEnvironment1.rssqlTem1.Source = "SELECT sum (Totalfee) as TotalGrand FROM tblPatientFacility WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent')) and tblPatientFacility.bookingdate = '" & !Date & "' and tblPatientFacility.agent_ID = " & DataComboAgent.BoundText
        DataEnvironment1.rssqlTem1.Open
        If DataEnvironment1.rssqlTem1.RecordCount <> 0 Then
            If IsNull(DataEnvironment1.rssqlTem1!TotalGrand) = False Then
                !BookingValue = DataEnvironment1.rssqlTem1!TotalGrand
            Else
                !BookingValue = 0
            End If
            DataEnvironment1.rssqlTem1.Update
        End If
        .MoveNext
    Wend
    
    If .State = 1 Then .Close
    .Source = "SELECT tblTemSingleAgentTransctions.* FROM tblTemSingleAgentTransctions"
    .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    While .EOF = False
        If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
        DataEnvironment1.rssqlTem1.Source = "SELECT sum (Totalrefund) as TotalGrand FROM tblPatientFacility WHERE (((tblPatientFacility.RefundToAgent)= 1) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent')) and tblPatientFacility.RepayDate = '" & !Date & "' and tblPatientFacility.Agent_ID = " & DataComboAgent.BoundText
        DataEnvironment1.rssqlTem1.Open
        If DataEnvironment1.rssqlTem1.RecordCount <> 0 Then
            If IsNull(DataEnvironment1.rssqlTem1!TotalGrand) = False Then
                !RefundValue = DataEnvironment1.rssqlTem1!TotalGrand
            Else
                !RefundValue = 0
            End If
            DataEnvironment1.rssqlTem1.Update
        End If
        .MoveNext
    Wend
    
    
    If .State = 1 Then .Close
    .Source = "SELECT tblTemSingleAgentTransctions.* FROM tblTemSingleAgentTransctions"
    .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    While .EOF = False
        If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
        DataEnvironment1.rssqlTem1.Source = "SELECT sum(cash) as TotalGrand FROM tblagentcashsettle   where SettledDate = '" & !Date & "' and Institution_ID = " & DataComboAgent.BoundText
        DataEnvironment1.rssqlTem1.Open
        If DataEnvironment1.rssqlTem1.RecordCount <> 0 Then
            If IsNull(DataEnvironment1.rssqlTem1!TotalGrand) = False Then
                !CashPayment = DataEnvironment1.rssqlTem1!TotalGrand
            Else
                !CashPayment = 0
            End If
            DataEnvironment1.rssqlTem1.Update
        End If
        .MoveNext
    Wend
    
End With
End Sub

Private Sub bttnSearch_Click()
    Call WriteToTemTable
    Call ViewData
End Sub

Private Sub ViewData()
'    .Open "Select agentcode as Code,AgentName As [Agent Name],StartingBalance as [Start Bal],CashPayment as Payment,BookingValue As Booking,RefundValue as Refund,EndingBalance as [End Bal] From tblTemAllAgentTransctions Order By AgentName"

    With DataEnvironment1.rssqlTem11
        If .State = 1 Then .Close
        .Open "Select Date,StartingBalance as [Start Bal],CashPayment as Payment,BookingValue As Booking,RefundValue as Refund,EndingBalance as [End Bal] From tblTemSingleAgentTransctions Order By Date"
        Set MSHFlexGrid1.DataSource = DataEnvironment1.rssqlTem11
    End With
End Sub

Private Sub DataComboAgent_Change()
DataComboAgentCode.BoundText = DataComboAgent.BoundText
End Sub


Private Sub DataComboAgentCode_Change()
    DataComboAgent.BoundText = DataComboAgentCode.BoundText
End Sub


Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    If UserAuthority <> AuthorityOwner Then
        DTPicker1.Enabled = False
        DTPicker2.Enabled = False
    End If
    
        With DataEnvironment1
            DataComboAgentCode.RowMember = Empty
            DataComboAgentCode.ListField = Empty
            DataComboAgentCode.BoundColumn = Empty
            If .rssqlTemAgents2.State = 1 Then .rssqlTemAgents2.Close
            .Commands!SqlTemAgentS2.CommandText = "SELECT tblInstitutions.institutioncode , tblInstitutions.institution_ID From tblInstitutions ORDER BY tblInstitutions.InstitutionCode"
            .SqlTemAgentS2
            Set DataComboAgentCode.RowSource = DataEnvironment1
            DataComboAgentCode.RowMember = "sqlTemAgents2"
            DataComboAgentCode.ListField = "InstitutionCode"
            DataComboAgentCode.BoundColumn = "Institution_ID"
        End With
        With DataEnvironment1
            DataComboAgent.RowMember = Empty
            DataComboAgent.ListField = Empty
            DataComboAgent.BoundColumn = Empty
            If .rssqlTemAgents1.State = 1 Then .rssqlTemAgents1.Close
            .Commands!sqlTemAgents1.CommandText = "SELECT tblInstitutions.institutionname , tblinstitutions.institution_ID From tblInstitutions ORDER BY tblInstitutions.institutionname"
            .sqlTemAgents1
            Set DataComboAgent.RowSource = DataEnvironment1
            DataComboAgent.RowMember = "sqlTemAgents1"
            DataComboAgent.ListField = "InstitutionName"
            DataComboAgent.BoundColumn = "Institution_ID"
        End With



End Sub
