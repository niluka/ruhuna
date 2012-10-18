VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form FrmInstitutions1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Institution Details"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   10680
   Begin VB.TextBox txtSearch 
      Height          =   360
      Left            =   120
      TabIndex        =   26
      Top             =   240
      Width           =   4215
   End
   Begin VB.Frame framInstitution 
      Height          =   7575
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtName 
         Height          =   360
         Left            =   2040
         TabIndex        =   13
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtAddress 
         Height          =   1215
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtTel 
         Height          =   360
         Left            =   2040
         TabIndex        =   11
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtFax 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox txtComment 
         Height          =   735
         Left            =   2040
         TabIndex        =   9
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txtAccount 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   6120
         Width           =   3495
      End
      Begin VB.TextBox txtCredit 
         Height          =   360
         Left            =   2040
         TabIndex        =   7
         Top             =   6720
         Width           =   3495
      End
      Begin VB.TextBox txtEmail 
         Height          =   360
         Left            =   2040
         TabIndex        =   6
         Top             =   3720
         Width           =   3495
      End
      Begin VB.CheckBox CheckAgent 
         Caption         =   "Is an agent"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DataComboBank 
         Bindings        =   "FormInstitutions1.frx":0000
         Height          =   360
         Left            =   2040
         TabIndex        =   14
         Top             =   5640
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "BankName"
         BoundColumn     =   "Bank_ID"
         Text            =   ""
         Object.DataMember      =   "sqlBank"
      End
      Begin MSDataListLib.DataCombo DataComboPaymenyMethod 
         Bindings        =   "FormInstitutions1.frx":001F
         Height          =   360
         Left            =   2040
         TabIndex        =   15
         Top             =   5160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "PaymentMethod"
         BoundColumn     =   "PaymentMethod_ID"
         Text            =   ""
         Object.DataMember      =   "sqlPaymentMethod"
      End
      Begin VB.Label Label1 
         Caption         =   "Institution Name"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Institution Address"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Institution Tel:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Institution Fax"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Institution Comments"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Payment Method"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   5160
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Institution  Bank"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   5760
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Institution Account"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   6240
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Institution Credit"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "E - Mail"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   1575
      End
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   495
      Left            =   9000
      TabIndex        =   0
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9000
      TabIndex        =   1
      Top             =   8760
      Width           =   1335
      _ExtentX        =   2355
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
   Begin btButtonEx.ButtonEx bttnChange 
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   8640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&hange"
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
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Sa&ve"
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
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   495
      Left            =   2760
      TabIndex        =   27
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Edit"
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
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   6855
      Left            =   120
      TabIndex        =   28
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   12091
      _Version        =   393216
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
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Add"
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
End
Attribute VB_Name = "FrmInstitutions1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemInstitutionID As Long
Dim FromGrid As Boolean

Private Sub bttnAdd_Click()
    Call AfterAdd
    Call ClearValues
End Sub

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub bttnChange_Click()
    Dim TemResponce As Byte
    If Trim(txtName.Text) = "" Then
        TemResponce = MsgBox("Please enter the name of the institution", vbCritical + vbOKOnly, "No Name")
        txtName.SetFocus
        Exit Sub
    End If
    Call EditData
    Call ClearValues
    Call FormatGrid
    Call FillGrid
    Call BeforeAddEdit
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    FromGrid = True
    
    Call AfterEdit
End Sub

Private Sub bttnSave_Click()
    Dim TemResponce As Byte
    If Trim(txtName.Text) = "" Then
        TemResponce = MsgBox("Please enter the name of the institution", vbCritical + vbOKOnly, "No Name")
        txtName.SetFocus
        Exit Sub
    End If
    Call SaveData
    Call FormatGrid
    Call FillGrid
    Call ClearValues
    Call AfterAdd
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call FillGrid
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub BeforeAddEdit()
    bttnEdit.Enabled = True
    bttnAdd.Enabled = True
    
    bttnSave.Visible = False
    bttnChange.Visible = False
    bttnCancel.Visible = False
    
    framInstitution.Enabled = False
    Grid1.Enabled = True
    
    FromGrid = False

    
End Sub

Private Sub AfterAdd()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
    
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    
    framInstitution.Enabled = True
    Grid1.Enabled = False
End Sub
Private Sub AfterEdit()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
    
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    
    framInstitution.Enabled = True
    Grid1.Enabled = True
End Sub

Private Sub SaveData()
    'On Error GoTo ErrorHandler
    With DataEnvironment1.rssqlInstitutions
        If .State = 1 Then .Close
        .Source = "SELECT tblInstitutions.* FROM tblInstitutions ORDER BY InstitutionName"
        .Open
        .AddNew
        !InstitutionName = Trim(txtName.Text)
        !InstitutionAddress = Trim(txtAddress.Text)
        !InstitutionTelephone = Trim(txtTel.Text)
        !InstitutionFax = Trim(txtFax.Text)
        !InstitutionEmail = Trim(txtEmail.Text)
        !InstitutionComments = Trim(txtComment.Text)
        If IsNumeric(DataComboPaymenyMethod.BoundText) Then !InstitutionPaymentMethod_ID = DataComboPaymenyMethod.BoundText
        If IsNumeric(DataComboBank.BoundText) Then !InstitutionBank_ID = DataComboBank.BoundText
        !InstitutionAccount = Trim(txtAccount.Text)
        If CheckAgent.Value = 1 Then
            !InstitutionIsAnAgent = True
        Else
            !InstitutionIsAnAgent = False
        End If
        .Update
        .Close
    Exit Sub
    
    
ErrorHandler:
     MsgBox Err.Description
    .CancelUpdate
    End With
    
End Sub


Private Sub EditData()
    'On Error GoTo ErrorHandler
    With DataEnvironment1.rssqlInstitutions
        If .State = 1 Then .Close
        .Source = "SELECT tblInstitutions.* FROM tblInstitutions where Institution_ID = " & TemInstitutionID
        If .RecordCount = 0 Then Exit Sub
        !InstitutionName = Trim(txtName.Text)
        !InstitutionAddress = Trim(txtAddress.Text)
        !InstitutionTelephone = Trim(txtTel.Text)
        !InstitutionFax = Trim(txtFax.Text)
        !InstitutionEmail = Trim(txtEmail.Text)
        !InstitutionComments = Trim(txtComment.Text)
        If IsNumeric(DataComboPaymenyMethod.BoundText) Then !InstitutionPaymentMethod_ID = DataComboPaymenyMethod.BoundText
        If IsNumeric(DataComboBank.BoundText) Then !InstitutionBank_ID = DataComboBank.BoundText
        !InstitutionAccount = Trim(txtAccount.Text)
        If CheckAgent.Value = 1 Then
            !InstitutionIsAnAgent = True
        Else
            !InstitutionIsAnAgent = False
        End If
        .Update
        .Close
    Exit Sub
ErrorHandler:
     MsgBox Err.Description
    .CancelUpdate
    End With
    
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    txtAddress.Text = Empty
    txtTel.Text = Empty
    txtFax.Text = Empty
    DataComboPaymenyMethod.Text = Empty
    DataComboBank.Text = Empty
    txtAccount.Text = Empty
    txtComment.Text = Empty
    txtAccount.Text = Empty
    txtCredit.Text = Empty
    CheckAgent.Value = 0
End Sub


Private Sub GetData()
    Call ClearValues
    If Grid1.Row < 1 Then Exit Sub
    Grid1.Col = 2
    If IsNumeric(Grid1.Text) = False Then Exit Sub
    TemInstitutionID = Val(Grid1.Text)
    With DataEnvironment1.rssqlInstitutions
        If .State = 1 Then .Close
        .Source = "SELECT tblInstitutions.* FROM tblInstitutions where Institution_ID = " & TemInstitutionID
        .Open
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!InstitutionName) Then
            txtName.Text = !InstitutionName
        End If
        If Not IsNull(!InstitutionAddress) Then
            txtAddress.Text = !InstitutionAddress
        End If
        If Not IsNull(!InstitutionTelephone) Then
            txtTel.Text = !InstitutionTelephone
        End If
        If Not IsNull(!InstitutionFax) Then
            txtFax.Text = !InstitutionFax
        End If
        If Not IsNull(!InstitutionEmail) Then
            txtEmail.Text = !InstitutionEmail
        End If
        If Not IsNull(!InstitutionComments) Then
            txtComment.Text = !InstitutionComments
        End If
        If Not IsNull(!InstitutionPaymentMethod) Then
            DataComboPaymenyMethod.BoundText = !InstitutionPaymentMethod_ID
        End If
        If Not IsNull(!InstitutionBank) Then
            DataComboBank.BoundText = !InstitutionBank_ID
        End If
        If Not IsNull(!InstitutionAccount) Then
            txtAccount.Text = !InstitutionAccount
        End If
        If Not IsNull(!InstitutionCredit) Then
            txtCredit.Text = !InstitutionCredit
        End If
        If !InstitutionIsAnAgent = True Then CheckAgent.Value = 1
End With

End Sub

Private Sub FormatGrid()
    Dim BorderMargin As Long
    BorderMargin = 100
    With Grid1
        .Clear
        .Cols = 3
        .Rows = 1
        .ColWidth(0) = 600
        .ColWidth(2) = 1
        .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + BorderMargin)
        .Row = 0
        .Col = 0
        .Text = "NO"
        .CellAlignment = 6
        .Col = 1
        .Text = "Institution Name"
        .Col = 2
        .Text = "ID"
        .CellAlignment = 6
    End With
End Sub


Private Sub FillGrid()
    Dim NowRow As Long
    With DataEnvironment1.rssqlInstitutions
    If .State = 1 Then .Close
    .Source = "SELECT tblInstitutions.* FROM tblInstitutions ORDER BY InstitutionName"
    .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
        Do While .EOF = False
            If Not IsNull(!InstitutionName) Then
            NowRow = NowRow + 1
            Grid1.Rows = NowRow + 1
            Grid1.Row = NowRow
            Grid1.Col = 0
            Grid1.CellAlignment = 7
            Grid1.Text = NowRow
            Grid1.Col = 1
            Grid1.CellAlignment = 1
            Grid1.Text = !InstitutionName
            Grid1.Col = 2
            Grid1.CellAlignment = 7
            Grid1.Text = !institution_ID
            End If
        .MoveNext
        Loop
    End With
End Sub



Private Sub Grid1_Click()
    If Grid1.Rows < 1 Then Exit Sub
    Grid1.Col = 2
    If Not IsNumeric(Grid1.Text) Then Exit Sub
    Call GetData
    Grid1.Col = 0
    Grid1.ColSel = Grid1.Cols - 1
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then Grid1_Click
End Sub

Private Sub txtSearch_Change()
    
' **************************************

    If FromGrid = True Then Exit Sub
    Dim TemFRows As Long
    Dim TemNowRow As Long
    Dim TemArray As Long
    Dim SearchSuccess As Boolean
    Dim TemLength As Single
    TemFRows = Grid1.Rows
    Grid1.Col = 1
    SearchSuccess = False
    If Len(txtSearch.Text) = 0 Then GoTo MeasureSuccess
    For TemArray = 1 To (TemFRows - 1)
        Grid1.Row = TemArray
        If Len(txtSearch.Text) > Len(Grid1.Text) Then
            GoTo FinishLoop
        Else
            TemLength = Len(txtSearch.Text)
        End If
        If UCase(Left((Grid1.Text), TemLength)) = UCase(txtSearch.Text) Then
            SearchSuccess = True
            Exit For
        Else
            SearchSuccess = False
        End If
FinishLoop:
    Next
    
MeasureSuccess:
    
    If SearchSuccess = True Then
        Grid1.TopRow = TemArray
        Grid1.Row = TemArray
        Grid1.Col = 0
        Grid1.ColSel = (Grid1.Cols - 1)
        bttnEdit.Enabled = True
        bttnAdd.Enabled = False
        Grid1.Col = 2
        TemInstitutionID = Grid1.Text
        Call GetData
        Grid1.Col = 0
        Grid1.ColSel = Grid1.Cols - 1
    Else
        Grid1.TopRow = 1
        Grid1.Row = 0
        Grid1.Col = 0
        Grid1.ColSel = 0
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
    End If
'**************************************
End Sub


