VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmCalculateAgentBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Agent Balance"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
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
   ScaleHeight     =   7365
   ScaleWidth      =   5970
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
   Begin btButtonEx.ButtonEx bttnCalculate 
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   3000
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Calculate"
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
   Begin VB.TextBox txtCalculatedBalance 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   3600
      Width           =   3615
   End
   Begin VB.TextBox txtCurrentBalance 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   5280
      Width           =   3615
   End
   Begin VB.TextBox txtStartingBalance 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   62128131
      CurrentDate     =   39637
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   62128131
      CurrentDate     =   39637
   End
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   5760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save New Balance"
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
   Begin MSDataListLib.DataCombo dtcAgentCode 
      Height          =   360
      Left            =   1920
      TabIndex        =   12
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin MSDataListLib.DataCombo dtcAgentName 
      Height          =   360
      Left            =   1920
      TabIndex        =   13
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Cod&e"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent &Name"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Calculated Balance"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Current Balance"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Starting Balance"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "End Date"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Start Date"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmCalculateAgentBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsAgent As New ADODB.Recordset
    Dim rsAgentBalance As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
    Dim temSql As String
    Dim rsViewAgent As New ADODB.Recordset
    Dim rsViewCode As New ADODB.Recordset
    
Private Sub bttnCalculate_Click()
    Dim tr As Integer
    If IsNumeric(dtcAgentName.BoundText) = False Then
        MsgBox "No Agent"
        dtcAgentName.SetFocus
        Exit Sub
    End If
    If dtpStart.Value >= dtpEnd.Value Then
        MsgBox "Wrong Date"
        dtpStart.SetFocus
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Dim i As Integer
    Dim Balance As Double
    Balance = Val(txtStartingBalance.Text)
    With rsAgentBalance
        If .State = 1 Then .Close
        temSql = "Delete  from tblInstitutionBalance Where Institution_ID = " & dtcAgentName.BoundText
        .Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
    End With
    For i = 0 To DateDiff("d", dtpStart.Value, dtpEnd.Value)
        With rsAgentBalance
            If .State = 1 Then .Close
            temSql = "SELECT * from tblInstitutionBalance"
            .Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
            .AddNew
            !Institution_Id = dtcAgentName.BoundText
            !Date = Format(dtpStart.Value + i, "dd MMMM yyyy")
            !SBalance = Balance
            With rsTem
                If .State = 1 Then .Close
                temSql = "SELECT Sum([Cash]) AS TotalCash fROM tblAgentCashSettle Where (Institution_ID = " & dtcAgentName.BoundText & ")  and (SettledDate = '" & Format((dtpStart.Value + i), "dd MMMM yyyy") & "')"
                .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    If IsNull(!TotalCash) = False Then
                        Balance = Balance + !TotalCash
                    End If
                End If
                If .State = 1 Then .Close
                temSql = "SELECT Sum([TotalFee]) AS TotalCash fROM tblPatientFacility Where (tblPatientFacility.Agent_ID = " & dtcAgentName.BoundText & ")  and ( tblPatientFacility.BookingDate = '" & Format((dtpStart.Value + i), "dd MMMM yyyy") & "')   and (tblPatientFacility.PaymentMode ='Agent')"
                .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    If IsNull(!TotalCash) = False Then
                        Balance = Balance - !TotalCash
                    End If
                End If
                If .State = 1 Then .Close
                temSql = "SELECT Sum([TotalRefund]) AS TotalCash fROM tblPatientFacility Where (tblPatientFacility.Agent_ID = " & dtcAgentName.BoundText & ")  and ( tblPatientFacility.BookingDate = '" & Format((dtpStart.Value + i), "dd MMMM yyyy") & "')   and (tblPatientFacility.RefundToAgent = 1 )"
                .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    If IsNull(!TotalCash) = False Then
                        Balance = Balance + !TotalCash
                    End If
                End If
            End With
            !EBalance = Balance
            .Update
        End With
    Next
    Me.MousePointer = vbDefault
    txtCalculatedBalance.Text = Format(Balance, "0.00")
End Sub

Private Sub bttnSave_Click()
    Dim tr As Integer
    tr = MsgBox("Are You sure?", vbCritical + vbYesNo, "Change Balance?")
    If tr = vbNo Then Exit Sub
    With rsAgent
        If .State = 1 Then .Close
        temSql = "SELECT * from tblInstitutions where Institution_ID = " & Val(dtcAgentName.BoundText)
        .Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !InstitutionCredit = txtCurrentBalance.Text
            .Update
        End If
        .Close
    End With
End Sub

Private Sub dtcAgentCode_Change()
    If IsNumeric(dtcAgentCode.BoundText) = False Then Exit Sub
    dtcAgentName.BoundText = dtcAgentCode.BoundText
End Sub

Private Sub dtcAgentName_Change()
    If IsNumeric(dtcAgentName.BoundText) = False Then Exit Sub
    dtcAgentCode.BoundText = dtcAgentName.BoundText
    With rsAgent
        If .State = 1 Then .Close
        temSql = "SELECT * from tblInstitutions where Institution_ID = " & Val(dtcAgentName.BoundText)
        .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtCurrentBalance.Text = Format(!InstitutionCredit, "0.00")
        End If
        .Close
    End With
End Sub

Private Sub Form_Load()
    Call FillCOmbos
End Sub

Private Sub FillCOmbos()
    With rsViewAgent
        If .State = 1 Then .Close
        temSql = "SELECT * from tblInstitutions order by InstitutionName"
        .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
    End With
    With dtcAgentName
        Set .RowSource = rsViewAgent
        .BoundColumn = "Institution_ID"
        .ListField = "InstitutionName"
    End With
    With rsViewCode
        If .State = 1 Then .Close
        temSql = "SELECT * from tblInstitutions order by InstitutionCode"
        .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
    End With
    With dtcAgentCode
        Set .RowSource = rsViewCode
        .BoundColumn = "Institution_ID"
        .ListField = "InstitutionCode"
    End With
End Sub
