VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmInstitutions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Institution Details"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   7
      Top             =   0
      Width           =   1000
      Begin VB.PictureBox Picture10 
         BackColor       =   &H000000FF&
         Height          =   1000
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   23
         Top             =   0
         Width           =   1000
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H000000FF&
         Height          =   1000
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   21
         Top             =   0
         Width           =   1000
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H000000FF&
         Height          =   1000
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   19
         Top             =   0
         Width           =   1000
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H000000FF&
         Height          =   1000
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   17
         Top             =   0
         Width           =   1000
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H000000FF&
         Height          =   1000
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   15
         Top             =   0
         Width           =   1000
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H000000FF&
         Height          =   1000
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   13
         Top             =   0
         Width           =   1000
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H000000FF&
         Height          =   1000
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   11
         Top             =   0
         Width           =   1000
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         Height          =   1000
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   9
         Top             =   0
         Width           =   1000
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   5
      Top             =   0
      Width           =   1000
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   495
      Left            =   8880
      TabIndex        =   1
      Top             =   7920
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
      Left            =   8880
      TabIndex        =   2
      Top             =   8640
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
      Left            =   6240
      TabIndex        =   3
      Top             =   8520
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
End
Attribute VB_Name = "frmInstitutions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemInstitutionID As Long

Private Sub bttnAdd_Click()
    Call AfterAdd
    Call Clearvalues
End Sub

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
    Call Clearvalues
End Sub

Private Sub bttnChange_Click()
    Dim TemResponce As Byte
    If Trim(txtName.Text) = "" Then
        TemResponce = MsgBox("Please enter the name of the institution", vbCritical + vbOKOnly, "No Name")
        txtName.SetFocus
        Exit Sub
    End If
    Call EditData
    Call Clearvalues
    Call FormatGrid
    Call FillGrid
    Call BeforeAddEdit
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
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
    Call Clearvalues
    Call AfterAdd
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call FillGrid
    Call BeforeAddEdit
    Call Clearvalues
End Sub

Private Sub BeforeAddEdit()
    bttnEdit.Enabled = True
    bttnAdd.Enabled = True
    
    bttnSave.Visible = False
    bttnChange.Visible = False
    bttnCancel.Visible = False
    
    framInstitution.Enabled = False
    grid1.Enabled = True
End Sub

Private Sub AfterAdd()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
    
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    
    framInstitution.Enabled = True
    grid1.Enabled = False
End Sub
Private Sub AfterEdit()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
    
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    
    framInstitution.Enabled = True
    grid1.Enabled = True
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
        !InstitutionFax = Trim(txtFaxText)
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

Private Sub Clearvalues()
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
    Call Clearvalues
    If grid1.Row < 1 Then Exit Sub
    grid1.Col = 2
    If IsNumeric(grid1.Text) = False Then Exit Sub
    TemInstitutionID = Val(grid1.Text)
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
    With grid1
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
            grid1.Rows = NowRow + 1
            grid1.Row = NowRow
            grid1.Col = 0
            grid1.CellAlignment = 7
            grid1.Text = NowRow
            grid1.Col = 1
            grid1.CellAlignment = 1
            grid1.Text = !InstitutionName
            grid1.Col = 2
            grid1.CellAlignment = 7
            grid1.Text = !institution_id
            End If
        .MoveNext
        Loop
    End With
End Sub



Private Sub grid1_Click()
    If grid1.Rows < 1 Then Exit Sub
    grid1.Col = 2
    If Not IsNumeric(grid1.Text) Then Exit Sub
    Call GetData
    grid1.Col = 0
    grid1.ColSel = grid1.Cols - 1
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then grid1_Click
End Sub

Private Sub txtSearch_Change()
    
' **************************************

    If FromGrid = True Then Exit Sub
    Dim TemFRows As Long
    Dim TemNowRow As Long
    Dim TemArray As Long
    Dim SearchSuccess As Boolean
    Dim TemLength As Single
    TemFRows = grid1.Rows
    grid1.Col = 1
    SearchSuccess = False
    If Len(txtSearch.Text) = 0 Then GoTo MeasureSuccess
    For TemArray = 1 To (TemFRows - 1)
        grid1.Row = TemArray
        If Len(txtSearch.Text) > Len(grid1.Text) Then
            GoTo FinishLoop
        Else
            TemLength = Len(txtSearch.Text)
        End If
        If UCase(Left((grid1.Text), TemLength)) = UCase(txtSearch.Text) Then
            SearchSuccess = True
            Exit For
        Else
            SearchSuccess = False
        End If
FinishLoop:
    Next
    
MeasureSuccess:
    
    If SearchSuccess = True Then
        grid1.TopRow = TemArray
        grid1.Row = TemArray
        grid1.Col = 0
        grid1.ColSel = (grid1.Cols - 1)
        bttnEdit.Enabled = True
        bttnAdd.Enabled = False
        grid1.Col = 2
        TemInstitutionID = grid1.Text
        Call GetData
        grid1.Col = 0
        grid1.ColSel = grid1.Cols - 1
    Else
        grid1.TopRow = 1
        grid1.Row = 0
        grid1.Col = 0
        grid1.ColSel = 0
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
    End If
'**************************************
End Sub

