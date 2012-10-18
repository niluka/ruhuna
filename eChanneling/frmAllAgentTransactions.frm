VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmAllAgentTransactions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transactions of All Agents of a Selected date"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11985
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
   ScaleHeight     =   5940
   ScaleWidth      =   11985
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7858
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin btButtonEx.ButtonEx bttnExit 
      Height          =   495
      Left            =   10680
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   156631043
      CurrentDate     =   39515
   End
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   9360
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
      Left            =   3000
      TabIndex        =   3
      Top             =   120
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
End
Attribute VB_Name = "frmAllAgentTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bttnExit_Click()
Unload Me
End Sub

Private Sub bttnPrint_Click()
With dtrAllAgentDetails
    .Sections("Section4").Controls("lblFromDate").Caption = "Date     : " & DTPicker1.Value
    
    With DataEnvironment1.rssqlTem11
    If .State = 1 Then .Close
    .Open "Select* From tblTemAllAgentTransctions Order By AgentName"
    End With
    
    Set dtrAllAgentDetails.DataSource = DataEnvironment1.rssqlTem11
    .Show
    
End With
End Sub

Private Sub WriteToTemTable()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "Truncate table tblTemAllAgentTransctions"
    .Open
    If .State = 1 Then .Close
    .Source = "SELECT tblTemAllAgentTransctions.* FROM tblTemAllAgentTransctions"
    .Open
    If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
    DataEnvironment1.rssqlTem1.Source = "select tblInstitutions.* FROM tblInstitutions order by InstitutionName"
    DataEnvironment1.rssqlTem1.Open
    If DataEnvironment1.rssqlTem1.RecordCount = 0 Then Exit Sub
    While DataEnvironment1.rssqlTem1.EOF = False
        .AddNew
        !Agent_ID = DataEnvironment1.rssqlTem1!Institution_Id
        !agentcode = DataEnvironment1.rssqlTem1!InstitutionCode
        !AgentName = DataEnvironment1.rssqlTem1!InstitutionName
        .Update
        DataEnvironment1.rssqlTem1.MoveNext
    Wend
    
    If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
    DataEnvironment1.rssqlTem1.Source = "SELECT tblInstitutionBalance.* From tblInstitutionBalance Where (((tblInstitutionBalance.Date) = '" & DTPicker1.Value & "')) "
    DataEnvironment1.rssqlTem1.Open
    If DataEnvironment1.rssqlTem1.RecordCount = 0 Then Exit Sub
    While DataEnvironment1.rssqlTem1.EOF = False
        If .State = 1 Then .Close
        .Source = "SELECT tblTemAllAgentTransctions.* FROM tblTemAllAgentTransctions where Agent_ID =" & DataEnvironment1.rssqlTem1!Institution_Id
        .Open
        If .RecordCount <> 0 Then
            !StartingBalance = DataEnvironment1.rssqlTem1!SBalance
            !EndingBalance = DataEnvironment1.rssqlTem1!EBalance
            .Update
        End If
        DataEnvironment1.rssqlTem1.MoveNext
    Wend
    
    If .State = 1 Then .Close
    .Source = "SELECT tblTemAllAgentTransctions.* FROM tblTemAllAgentTransctions"
    .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    While .EOF = False
        If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
        DataEnvironment1.rssqlTem1.Source = "SELECT sum (Totalfee) as TotalGrand FROM tblPatientFacility WHERE (((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent')) and tblPatientFacility.bookingdate = '" & DTPicker1.Value & "' and tblPatientFacility.agent_ID = " & !Agent_ID
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
    .Source = "SELECT tblTemAllAgentTransctions.* FROM tblTemAllAgentTransctions"
    .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    While .EOF = False
        If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
        DataEnvironment1.rssqlTem1.Source = "SELECT sum (Totalrefund) as TotalGrand FROM tblPatientFacility WHERE (((tblPatientFacility.RefundToAgent)= 1) AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent')) and tblPatientFacility.RepayDate = '" & DTPicker1.Value & "' and tblPatientFacility.Agent_ID = " & !Agent_ID
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
    .Source = "SELECT tblTemAllAgentTransctions.* FROM tblTemAllAgentTransctions"
    .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    While .EOF = False
        If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
        DataEnvironment1.rssqlTem1.Source = "SELECT sum(cash) as TotalGrand FROM tblagentcashsettle   where SettledDate = '" & DTPicker1.Value & "' and Institution_ID = " & !Agent_ID
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
    
    If DTPicker1.Value = Date Then
        If .State = 1 Then .Close
        .Source = "SELECT tblTemAllAgentTransctions.* FROM tblTemAllAgentTransctions"
        .Open
        Do While .EOF = False
                !EndingBalance = (!StartingBalance + !CashPayment) - (!BookingValue - !RefundValue)
                .Update
        .MoveNext
        Loop
    End If

End With
End Sub

Private Sub bttnSearch_Click()
    Call WriteToTemTable
    Call ViewData
End Sub

Private Sub ViewData()
MSHFlexGrid1.ColWidth(1, 0) = 4500

    With DataEnvironment1.rssqlTem11
    If .State = 1 Then .Close
    .Open "Select agentcode as Code,AgentName As [Agent Name],StartingBalance as [Start Bal],CashPayment as Payment,BookingValue As Booking,RefundValue as Refund,EndingBalance as [End Bal] From tblTemAllAgentTransctions Order By AgentName"
    
    
    Set MSHFlexGrid1.DataSource = DataEnvironment1.rssqlTem11
    
    End With
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    If UserAuthority <> AuthorityOwner Then
        DTPicker1.Enabled = False
    End If

End Sub
