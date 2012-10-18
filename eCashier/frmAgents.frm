VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAgents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agents"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12045
   ClipControls    =   0   'False
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
   ScaleHeight     =   8115
   ScaleWidth      =   12045
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   4920
      TabIndex        =   22
      Top             =   6720
      Width           =   6975
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Ca&ncel"
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Save"
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
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Save"
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
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   240
      TabIndex        =   21
      Top             =   6720
      Width           =   4575
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Edit"
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
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Add"
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
   Begin VB.Frame Frame2 
      Height          =   6495
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   4575
      Begin MSDataListLib.DataCombo dtcAgent 
         Height          =   5940
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   10478
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   4920
      TabIndex        =   13
      Top             =   240
      Width           =   6975
      Begin VB.TextBox txtCode 
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtAddress 
         Height          =   1215
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox txtTelephone 
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   2640
         Width           =   4095
      End
      Begin VB.TextBox txtFax 
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   3600
         Width           =   4095
      End
      Begin VB.TextBox txtOther 
         Height          =   2295
         Left            =   2520
         TabIndex        =   8
         Top             =   4080
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "Code"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Telephone"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Fax"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Email"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Other Details"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   4080
         Width           =   2295
      End
   End
   Begin btButtonEx.ButtonEx bttnCLose 
      Height          =   375
      Left            =   9840
      TabIndex        =   12
      Top             =   7560
      Width           =   1935
      _ExtentX        =   3413
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
End
Attribute VB_Name = "frmAgents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewAgent As New ADODB.Recordset
    Dim rsViewCity As New ADODB.Recordset
    Dim rsAgent As New ADODB.Recordset
    Dim A As Integer
    Dim TemAgentId As Long

Private Sub bttnCancel_Click()
    ClearValues
    BeforeAddEdit
    dtcAgent.Text = Empty
    dtcAgent.SetFocus
End Sub

Private Sub EditAgent()
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
    On Error GoTo ErrorHandler
    With rsAgent
        If .State = 1 Then .Close
        .Open "Select* From tblAgent Where AgentID = " & TemAgentId & "", cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then Exit Sub
        !Agent = Trim(txtName.Text)
        !Address = txtAddress.Text
        If Trim(txtCode.Text) <> "" Then
            !Code = txtCode.Text
        End If
        If txtTelephone.Text <> "" Then
            !Telephone = txtTelephone.Text
        End If
        If txtFax.Text <> "" Then
            !Fax = txtFax.Text
        End If
        If txtEmail.Text <> "" Then
            !Email = txtEmail.Text
        End If
        If txtOther.Text <> "" Then
            !Comments = txtOther.Text
        End If
        
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call FillAgent
        dtcAgent.SetFocus
        dtcAgent.Text = Empty
        Exit Sub
    
ErrorHandler:
        A = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcAgent.Text = Empty
        dtcAgent.SetFocus
    End With
        
End Sub

Private Sub bttnChange_Click()
    Call EditAgent
End Sub

Private Sub bttnSave_Click()
    Call SaveAgent
End Sub

Private Sub dtcAgent_Click(Area As Integer)
    If IsNumeric(dtcAgent.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    Call FillAgent
    Call FillCity
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub FillAgent()
    With rsViewAgent
        If .State = 1 Then .Close
        .Open "Select* From tblAgent Order By Agent", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtcAgent.RowSource = rsViewAgent
        dtcAgent.BoundColumn = "AgentID"
        dtcAgent.ListField = "Agent"
    End With
End Sub

Private Sub FillCity()
'    With rsViewCity
'        If .State = 1 Then .Close
'        .Open "Select* From tblCity Order By City", cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount = 0 Then Exit Sub
'        Set dtcCity.RowSource = rsViewCity
'        dtcCity.BoundColumn = "CityID"
'        dtcCity.ListField = "City"
'    End With
End Sub

Private Sub bttnAdd_Click()
    ClearValues
    AfterAdd
    txtName.Text = dtcAgent.Text
    dtcAgent.Text = Empty
    txtName.SetFocus
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    AfterEdit
    txtName.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub SaveAgent()
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
'    On Error GoTo ErrorHandler
    With rsAgent
        If .State = 1 Then .Close
        .Open "Select* From tblAgent", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Agent = Trim(txtName.Text)
        !Address = txtAddress.Text
        If txtTelephone.Text <> "" Then
            !Telephone = txtTelephone.Text
        End If
        If Trim(txtCode.Text) <> "" Then
            !Code = txtCode.Text
        End If
        If txtFax.Text <> "" Then
            !Fax = txtFax.Text
        End If
        If txtEmail.Text <> "" Then
            !Email = txtEmail.Text
        End If
        If txtOther.Text <> "" Then
            !Comments = txtOther.Text
        End If
        

        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call FillAgent
        dtcAgent.SetFocus
        dtcAgent.Text = Empty
        Exit Sub
ErrorHandler:
        A = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcAgent.SetFocus
        dtcAgent.Text = Empty
    End With
End Sub


Private Sub AfterAdd()
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
End Sub

Private Sub AfterEdit()
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
End Sub

Private Sub BeforeAddEdit()
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    bttnSave.Visible = False
    bttnCancel.Visible = False
    bttnChange.Visible = False
    Frame1.Enabled = False
    Frame2.Enabled = True
End Sub

Private Sub NoName()
    Dim TemResponce As Integer
    TemResponce = MsgBox("No Such Health Scheme Supplier found among the records", , "No Record")
    Exit Sub
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    txtAddress.Text = Empty
    txtTelephone.Text = Empty
    txtFax.Text = Empty
    txtEmail.Text = Empty
    txtOther.Text = Empty
    txtCode.Text = Empty
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcAgent.BoundText) Then Exit Sub
    With rsAgent
        If .State = 1 Then .Close
        .Open "Select* From tblAgent Where AgentID = " & dtcAgent.BoundText & "", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Call ClearValues
        If Not IsNull(!Agent) Then txtName.Text = !Agent
        If Not IsNull(!Address) Then txtAddress.Text = !Address
        If Not IsNull(!Telephone) Then txtTelephone.Text = !Telephone
        If Not IsNull(!Fax) Then txtFax.Text = !Fax
        If Not IsNull(!Email) Then txtEmail.Text = !Email
        If Not IsNull(!Comments) Then txtOther.Text = !Comments
        If Not IsNull(!Code) Then txtCode.Text = !Code
        TemAgentId = !AgentID
        If .RecordCount = 0 Then Exit Sub
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsViewAgent.State = 1 Then rsViewAgent.Close: Set rsViewAgent = Nothing
    If rsViewCity.State = 1 Then rsViewCity.Close: Set rsViewCity = Nothing
End Sub
