VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmHealthSchemeSuppliers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Health Scheme Suppliers"
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
   Moveable        =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   12045
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   4920
      TabIndex        =   26
      Top             =   6720
      Width           =   6975
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   4920
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   12
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
      TabIndex        =   25
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
      TabIndex        =   24
      Top             =   240
      Width           =   4575
      Begin MSDataListLib.DataCombo dtcHealthSchemeSupplier 
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
      TabIndex        =   15
      Top             =   240
      Width           =   6975
      Begin VB.TextBox txtProfitMargin 
         Height          =   375
         Left            =   2520
         TabIndex        =   27
         Top             =   4560
         Width           =   4095
      End
      Begin MSDataListLib.DataCombo dtcCity 
         Height          =   360
         Left            =   2520
         TabIndex        =   5
         Top             =   2160
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
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
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtTelephone 
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   2640
         Width           =   4095
      End
      Begin VB.TextBox txtFax 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   3600
         Width           =   4095
      End
      Begin VB.TextBox txtWebsite 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   4080
         Width           =   4095
      End
      Begin VB.TextBox txtOther 
         Height          =   1335
         Left            =   2520
         TabIndex        =   10
         Top             =   5040
         Width           =   4095
      End
      Begin VB.Label Label9 
         Caption         =   "Profit Margin"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "City"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Telephone"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Fax"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Email"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Website"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Other Details"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   5040
         Width           =   2295
      End
   End
   Begin btButtonEx.ButtonEx bttnCLose 
      Height          =   375
      Left            =   9840
      TabIndex        =   14
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
Attribute VB_Name = "frmHealthSchemeSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewHealthSchemeSupplier As New ADODB.Recordset
    Dim rsViewCity As New ADODB.Recordset
    Dim rsHealthSchemeSupplier As New ADODB.Recordset
    Dim A As Integer
    Dim TemHealthSchemeSupplierId As Long

Private Sub bttnCancel_Click()
    ClearValues
    BeforeAddEdit
    dtcHealthSchemeSupplier.Text = Empty
    dtcHealthSchemeSupplier.SetFocus
End Sub

Private Sub EditHealthSchemeSupplier()
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
    On Error Resume Next
    With rsHealthSchemeSupplier
        If .State = 1 Then .Close
        .Open "Select* From tblHealthSchemeSuppliers Where HealthSchemeSupplierID = " & TemHealthSchemeSupplierId & "", cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then Exit Sub
        !HealthSchemeSupplierName = Trim(txtName.Text)
        !HealthSchemeSupplierAddress = txtAddress.Text
        !HealthSchemeSuppliercityID = Val(dtcCity.BoundText)
        !HealthSchemeSupplierTelephone = txtTelephone.Text
        !HealthSchemeSupplierFax = txtFax.Text
        !HealthSchemeSupplierEmail = txtEmail.Text
        !HealthSchemeSupplierWebsite = txtWebsite.Text
        !HealthSchemeSupplierComments = txtOther.Text
        !ProfitMargin = Val(txtProfitMargin.Text)
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call FillHealthSchemeSupplier
        dtcHealthSchemeSupplier.SetFocus
        dtcHealthSchemeSupplier.Text = Empty
        Exit Sub
    
ErrorHandler:
        A = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcHealthSchemeSupplier.Text = Empty
        dtcHealthSchemeSupplier.SetFocus
    End With
        
End Sub

Private Sub bttnChange_Click()
    Call EditHealthSchemeSupplier
End Sub

Private Sub bttnSave_Click()
    Call SaveHealthSchemeSupplier
End Sub

Private Sub dtcHealthSchemeSupplier_Click(Area As Integer)
    If IsNumeric(dtcHealthSchemeSupplier.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    Call FillHealthSchemeSupplier
    Call FillCity
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub FillHealthSchemeSupplier()
    With rsViewHealthSchemeSupplier
        If .State = 1 Then .Close
        .Open "Select* From tblHealthSchemeSuppliers Order By HealthSchemeSupplierName", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtcHealthSchemeSupplier.RowSource = rsViewHealthSchemeSupplier
        dtcHealthSchemeSupplier.BoundColumn = "HealthSchemeSupplierID"
        dtcHealthSchemeSupplier.ListField = "HealthSchemeSupplierName"
    End With
End Sub

Private Sub FillCity()
    With rsViewCity
        If .State = 1 Then .Close
        .Open "Select* From tblCity Order By City", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtcCity.RowSource = rsViewCity
        dtcCity.BoundColumn = "CityID"
        dtcCity.ListField = "City"
    End With
End Sub

Private Sub bttnAdd_Click()
    ClearValues
    AfterAdd
    txtName.Text = dtcHealthSchemeSupplier.Text
    dtcHealthSchemeSupplier.Text = Empty
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

Private Sub SaveHealthSchemeSupplier()
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
    On Error Resume Next
    With rsHealthSchemeSupplier
        If .State = 1 Then .Close
        .Open "Select* From tblHealthSchemeSuppliers", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !HealthSchemeSupplierName = Trim(txtName.Text)
        !HealthSchemeSupplierAddress = txtAddress.Text
        !HealthSchemeSuppliercityID = Val(dtcCity.BoundText)
        !HealthSchemeSupplierTelephone = txtTelephone.Text
        !HealthSchemeSupplierFax = txtFax.Text
        !HealthSchemeSupplierEmail = txtEmail.Text
        !HealthSchemeSupplierWebsite = txtWebsite.Text
        !HealthSchemeSupplierComments = txtOther.Text
        !ProfitMargin = Val(txtProfitMargin.Text)
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call FillHealthSchemeSupplier
        dtcHealthSchemeSupplier.SetFocus
        dtcHealthSchemeSupplier.Text = Empty
        Exit Sub
ErrorHandler:
        A = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcHealthSchemeSupplier.SetFocus
        dtcHealthSchemeSupplier.Text = Empty
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
    dtcCity.Text = Empty
    txtTelephone.Text = Empty
    txtFax.Text = Empty
    txtEmail.Text = Empty
    txtWebsite.Text = Empty
    txtOther.Text = Empty
    txtProfitMargin.Text = Empty
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcHealthSchemeSupplier.BoundText) Then Exit Sub
    With rsHealthSchemeSupplier
        If .State = 1 Then .Close
        .Open "Select* From tblHealthSchemeSuppliers Where HealthSchemeSupplierID = " & dtcHealthSchemeSupplier.BoundText & "", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Call ClearValues
        If Not IsNull(!HealthSchemeSupplierName) Then txtName.Text = !HealthSchemeSupplierName
        If Not IsNull(!HealthSchemeSupplierAddress) Then txtAddress.Text = !HealthSchemeSupplierAddress
        If Not IsNull(!HealthSchemeSuppliercityID) Then dtcCity.BoundText = Val(!HealthSchemeSuppliercityID)
        If Not IsNull(!HealthSchemeSupplierTelephone) Then txtTelephone.Text = !HealthSchemeSupplierTelephone
        If Not IsNull(!HealthSchemeSupplierFax) Then txtFax.Text = !HealthSchemeSupplierFax
        If Not IsNull(!HealthSchemeSupplierEmail) Then txtEmail.Text = !HealthSchemeSupplierEmail
        If Not IsNull(!HealthSchemeSupplierWebsite) Then txtWebsite.Text = !HealthSchemeSupplierWebsite
        If Not IsNull(!HealthSchemeSupplierComments) Then txtOther.Text = !HealthSchemeSupplierComments
        If Not IsNull(!ProfitMargin) Then txtProfitMargin.Text = !ProfitMargin
        TemHealthSchemeSupplierId = !HealthSchemeSupplierID
        If .RecordCount = 0 Then Exit Sub
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsViewHealthSchemeSupplier.State = 1 Then rsViewHealthSchemeSupplier.Close: Set rsViewHealthSchemeSupplier = Nothing
    If rsViewCity.State = 1 Then rsViewCity.Close: Set rsViewCity = Nothing
End Sub
