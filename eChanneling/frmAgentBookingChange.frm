VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAgentBookingChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agent Booking Change"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
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
   ScaleHeight     =   4215
   ScaleWidth      =   5835
   Begin VB.TextBox txtTotalFee 
      Height          =   360
      Left            =   1920
      TabIndex        =   11
      Top             =   2640
      Width           =   2415
   End
   Begin MSDataListLib.DataCombo dtcPreviousAgent 
      Height          =   360
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnSearch 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Search"
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
   Begin VB.TextBox txtID 
      Height          =   360
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin MSDataListLib.DataCombo dtcPreviousAgentCode 
      Height          =   360
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcNewAgent 
      Height          =   360
      Left            =   1920
      TabIndex        =   7
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcNewAgentCode 
      Height          =   360
      Left            =   4800
      TabIndex        =   8
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnChange 
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Change"
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
      Left            =   4440
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&C&lose"
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
   Begin VB.Label Label4 
      Caption         =   "Total Fee"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "&New Agent"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "&Previous Agent"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "&Booking ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmAgentBookingChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsAgent As New ADODB.Recordset
    Dim rsBooking As New ADODB.Recordset
    Dim rsPF As New ADODB.Recordset
    Dim temSql As String
    
Private Sub bttnChange_Click()
    Dim i As Integer
    If IsNumeric(dtcPreviousAgent.BoundText) = False Then
        MsgBox "Wrong Search"
        txtID.SetFocus
        SendKeys "{home}+{End}"
        Exit Sub
    End If
    If IsNumeric(dtcNewAgent.BoundText) = False Then
        MsgBox "No New Agent"
        dtcNewAgent.SetFocus
        Exit Sub
    End If
    With rsAgent
        temSql = "SELECT * FROM tblInstitutions WHERE Institution_ID = " & Val(dtcPreviousAgent.BoundText)
        If .State = 1 Then .Close
        .Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !InstitutionCredit = !InstitutionCredit + Val(txtTotalFee.Text)
            .Update
        End If
        temSql = "SELECT * FROM tblInstitutions WHERE Institution_ID = " & Val(dtcNewAgent.BoundText)
        If .State = 1 Then .Close
        .Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !InstitutionCredit = !InstitutionCredit - Val(txtTotalFee.Text)
            .Update
        End If
        .Close
    End With
    With rsPF
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPatientFacility WHERE PatientFacility_ID = " & Val(txtID.Text)
        .Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Agent_ID = Val(dtcNewAgent.BoundText)
            .Update
        End If
        .Close
    End With
    MsgBox "Change Successfull"
    txtID.Text = Empty
    txtTotalFee.Text = Empty
    bttnChange.Enabled = False
    dtcNewAgent.Text = Empty
    dtcPreviousAgent.Text = Empty
    dtcNewAgentCode.Text = Empty
    dtcPreviousAgentCode.Text = Empty
    txtID.SetFocus
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnSearch_Click()
    With rsPF
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPatientFacility WHERE PatientFacility_ID = " & Val(txtID.Text)
        .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!Agent_ID) = True Then
                MsgBox "Not an Agent Booking"
                txtID.SetFocus
                SendKeys "{Home}+{End}"
                bttnChange.Enabled = False
                Exit Sub
            End If
            If !cancelled = True Then
                MsgBox "Cancelled Booking"
                txtID.SetFocus
                SendKeys "{Home}+{End}"
                bttnChange.Enabled = False
                Exit Sub
            End If
            If !FullyPaid = 0 Then
                MsgBox "Not Fully Paid Booking"
                txtID.SetFocus
                SendKeys "{Home}+{End}"
                bttnChange.Enabled = False
                Exit Sub
            End If
            If !Refund = True Then
                MsgBox "Refunded Booking"
                txtID.SetFocus
                SendKeys "{Home}+{End}"
                bttnChange.Enabled = False
                Exit Sub
            End If
                dtcPreviousAgent.BoundText = !Agent_ID
                txtTotalFee.Text = Format(!TotalDue, "0.00")
                bttnChange.Enabled = True
        Else
                MsgBox "Not an Agent Booking"
                txtID.SetFocus
                SendKeys "{Home}+{End}"
                bttnChange.Enabled = False
                Exit Sub
        End If
        .Close
    End With
    
End Sub

Private Sub FillCOmbos()
    With DataEnvironment1
        If .rssqlTemAgents2.State = 1 Then .rssqlTemAgents2.Close
        .Commands!SqlTemAgentS2.CommandText = "SELECT tblInstitutions.institutioncode , tblInstitutions.institution_ID From tblInstitutions ORDER BY tblInstitutions.InstitutionCode"
        .SqlTemAgentS2
        Set dtcPreviousAgentCode.RowSource = DataEnvironment1
        Set dtcNewAgentCode.RowSource = DataEnvironment1
        dtcPreviousAgentCode.RowMember = "sqlTemAgents2"
        dtcPreviousAgentCode.ListField = "InstitutionCode"
        dtcPreviousAgentCode.BoundColumn = "Institution_ID"
        dtcNewAgentCode.RowMember = "sqlTemAgents2"
        dtcNewAgentCode.ListField = "InstitutionCode"
        dtcNewAgentCode.BoundColumn = "Institution_ID"
        
        If .rssqlTemAgents1.State = 1 Then .rssqlTemAgents1.Close
        .Commands!sqlTemAgents1.CommandText = "SELECT tblInstitutions.institutionname , tblinstitutions.institution_ID From tblInstitutions ORDER BY tblInstitutions.institutionname"
        .sqlTemAgents1
        Set dtcPreviousAgent.RowSource = DataEnvironment1
        Set dtcNewAgent.RowSource = DataEnvironment1
        dtcPreviousAgent.RowMember = "sqlTemAgents1"
        dtcPreviousAgent.ListField = "InstitutionName"
        dtcPreviousAgent.BoundColumn = "Institution_ID"
        dtcNewAgent.RowMember = "sqlTemAgents1"
        dtcNewAgent.ListField = "InstitutionName"
        dtcNewAgent.BoundColumn = "Institution_ID"
    End With
End Sub

Private Sub dtcNewAgent_Change()
    dtcNewAgentCode.BoundText = dtcNewAgent.BoundText
End Sub

Private Sub dtcNewAgentCode_Change()
    dtcNewAgent.BoundText = dtcNewAgentCode.BoundText
End Sub

Private Sub dtcPreviousAgent_Change()
    dtcPreviousAgentCode.BoundText = dtcPreviousAgent.BoundText
End Sub

Private Sub dtcPreviousAgentCode_Change()
    dtcPreviousAgent.BoundText = dtcPreviousAgentCode.BoundText
End Sub

Private Sub Form_Load()
    Call FillCOmbos
End Sub

Private Sub txtID_Change()
    bttnChange.Enabled = False
End Sub
