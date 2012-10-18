VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPatientsDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patients details"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
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
   ScaleHeight     =   5865
   ScaleWidth      =   6720
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Add"
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
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin MSDataListLib.DataCombo dtcPID 
         Height          =   360
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtNIC 
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   5280
         Width           =   3975
      End
      Begin VB.TextBox txtPhone 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   4680
         Width           =   3975
      End
      Begin VB.TextBox txtAddress 
         Height          =   1815
         Left            =   1320
         TabIndex        =   10
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox txtSName 
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtOName 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox txtFName 
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "Patient ID"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "NIC No"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Other Name"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Surname"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Phone"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   4680
         Width           =   1575
      End
   End
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Edit"
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Delete"
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
      Left            =   5760
      TabIndex        =   16
      Top             =   3120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save"
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
      Left            =   5760
      TabIndex        =   17
      Top             =   3120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save"
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
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cancel"
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
      Left            =   5760
      TabIndex        =   19
      Top             =   5040
      Width           =   855
      _ExtentX        =   1508
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
End
Attribute VB_Name = "frmPatientsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsPatientsDetails As New ADODB.Recordset
    Dim rsViewPatientsDetails As New ADODB.Recordset
    Dim temSql As String
    Dim temPatientID As Long

Private Sub BeforeAddEdit()
    dtcPID.Enabled = True
    bttnAdd.Enabled = True
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    
    txtFName.Enabled = False
    txtOName.Enabled = False
    txtSName.Enabled = False
    txtAddress.Enabled = False
    txtPhone.Enabled = False
    txtNIC.Enabled = False

    
    bttnSave.Enabled = False
    bttnChange.Enabled = False
    bttnCancel.Enabled = False
    
    bttnSave.Visible = False
    bttnChange.Visible = False
End Sub

Private Sub AfterAdd()
    dtcPID.Enabled = True
    bttnAdd.Enabled = True
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    
    txtFName.Enabled = True
    txtOName.Enabled = True
    txtSName.Enabled = True
    txtAddress.Enabled = True
    txtPhone.Enabled = True
    txtNIC.Enabled = True
    
    bttnSave.Enabled = True
    bttnChange.Enabled = False
    bttnCancel.Enabled = True
    
    bttnSave.Visible = True
    bttnChange.Visible = False
End Sub

Private Sub AfterEdit()
    dtcPID.Enabled = True
    bttnAdd.Enabled = True
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    
    txtFName.Enabled = True
    txtOName.Enabled = True
    txtSName.Enabled = True
    txtAddress.Enabled = True
    txtPhone.Enabled = True
    txtNIC.Enabled = True
    
    bttnSave.Enabled = False
    bttnChange.Enabled = True
    bttnCancel.Enabled = True
    
    bttnSave.Visible = False
    bttnChange.Visible = True
End Sub

Private Sub bttnAdd_Click()
    
    Call AfterAdd
    txtFName.Text = dtcPID
    dtcPID.Text = Empty
    Call ClearValues
End Sub

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub bttnChange_Click()
    Call ChangeDetails
    Call BeforeAddEdit
    Call ClearValues
    Call FillCombos
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDelete_Click()
    If IsNumeric(dtcPID.BoundText) = False Then Exit Sub
    With rsPatientsDetails
    temSql = "SELECT * FROM tblPatientMainDetails WHERE PatientID = " & dtcPID.BoundText
    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
    .Delete adAffectCurrent
    .Close
    End With
    Call ClearValues
    dtcPID.Text = Empty
    Call FillCombos
End Sub

Private Sub bttnEdit_Click()
    Call AfterEdit
End Sub

Private Sub bttnSave_Click()
    Call SaveDetails
    Call BeforeAddEdit
    Call ClearValues
    Call FillCombos
End Sub



Private Sub dtcPID_Click(Area As Integer)
    If IsNumeric(dtcPID.BoundText) = True Then
        bttnEdit.Enabled = True
        bttnDelete.Enabled = True
        bttnAdd.Enabled = False
        Call ClearValues
        Call DisplayDetails
    Else
        Call ClearValues
        bttnEdit.Enabled = False
        bttnDelete.Enabled = False
        bttnAdd.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Call BeforeAddEdit
    FillCombos
End Sub

Private Sub FillCombos()
    With rsViewPatientsDetails
    temSql = "SELECT * FROM tblPatientMainDetails"
    If .State = 1 Then rsViewPatientsDetails.Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    
    With dtcPID
    Set .RowSource = Nothing
    Set .RowSource = rsViewPatientsDetails
    .ListField = "PatientID"
    .BoundColumn = "PatientID"
    End With
End Sub

Private Sub SaveDetails()
    On Error Resume Next

    With rsViewPatientsDetails
    If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPatientMainDetails"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !FirstName = txtFName.Text
        !OtherNames = txtOName.Text
        !Surname = txtSName.Text
        !Address = txtAddress.Text
        !Phone = txtPhone.Text
        !NICNo = txtNIC.Text
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        temPatientID = !NewID
        .Close
    End With
End Sub

Private Sub ChangeDetails()
    On Error Resume Next
    With rsPatientsDetails
    If .State = 1 Then .Close
    temSql = "SELECT * FROM tblPatientMainDetailstblStaff1 WHERE PatientID = " & dtcPID.BoundText
    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        !FirstName = txtFName.Text
        !OtherNames = txtOName.Text
        !Surname = txtSName.Text
        !Address = txtAddress.Text
        !Phone = txtPhone.Text
        !NICNo = txtNIC.Text
    .Update
    .Close
    End With
End Sub

Private Sub DisplayDetails()
    With rsPatientsDetails
    temSql = "SELECT * FROM tblPatientMainDetails WHERE PatientID = " & dtcPID.BoundText
    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        txtFName.Text = !FirstName
        If IsNull(!OtherNames) = False Then txtOName.Text = !OtherNames
        If IsNull(!Surname) = False Then txtSName.Text = !Surname
        If IsNull(!Address) = False Then txtAddress.Text = !Address
        If IsNull(!Phone) = False Then txtPhone.Text = !Phone
        If IsNull(!NICNo) = False Then txtNIC.Text = !NICNo
    .Close
    End With
End Sub

Private Sub ClearValues()
    txtFName.Text = Empty
    txtOName.Text = Empty
    txtSName.Text = Empty
    txtAddress.Text = Empty
    txtPhone.Text = Empty
    txtNIC.Text = Empty
End Sub


