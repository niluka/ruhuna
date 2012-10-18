VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmManufacturers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manufacturers"
   ClientHeight    =   8640
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
   ScaleHeight     =   8640
   ScaleWidth      =   12045
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   5040
      TabIndex        =   26
      Top             =   7080
      Width           =   6855
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   240
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
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   4680
         TabIndex        =   13
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   240
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
      Top             =   7080
      Width           =   4695
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
   End
   Begin VB.Frame Frame2 
      Height          =   6855
      Left            =   240
      TabIndex        =   24
      Top             =   240
      Width           =   4695
      Begin MSDataListLib.DataCombo dtcManufacture 
         Height          =   6180
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   10901
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   5040
      TabIndex        =   15
      Top             =   240
      Width           =   6855
      Begin MSDataListLib.DataCombo dtcCountry 
         Height          =   360
         Left            =   2520
         TabIndex        =   5
         Top             =   2280
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
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtAddress 
         Height          =   1200
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtTelephone 
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   2880
         Width           =   4095
      End
      Begin VB.TextBox txtFax 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   3360
         Width           =   4095
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   3840
         Width           =   4095
      End
      Begin VB.TextBox txtWebsite 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   4320
         Width           =   4095
      End
      Begin VB.TextBox txtOther 
         Height          =   1095
         Left            =   2520
         TabIndex        =   10
         Top             =   4800
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Manufacturer Name"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Manufacturer Address"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Manufacturer Country"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Manufacturer Telephone"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Manufacturer Fax"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Manufacturer Email"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Manufacturer Website"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Other Details"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   4800
         Width           =   2295
      End
   End
   Begin btButtonEx.ButtonEx bttnCLose 
      Height          =   375
      Left            =   9840
      TabIndex        =   14
      Top             =   8040
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
End
Attribute VB_Name = "frmManufacturers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewCountry As New ADODB.Recordset
    Dim rsViewManufacture As New ADODB.Recordset
    Dim rsManufacture As New ADODB.Recordset
    Dim NowROw As Long
    Dim TemManufacturerID As Long

Private Sub bttnCancel_Click()
    ClearValues
    BeforeAddEdit
    dtcManufacture.SetFocus
    dtcManufacture.Text = Empty
End Sub

Private Sub bttnChange_Click()
    Dim TemResponce As Integer
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
    On Error Resume Next
    With rsManufacture
        If .State = 1 Then .Close
        .Source = "Select* From tblManufacturer Where (ManufacturerID =" & TemManufacturerID & ")"
        .Open , cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then Exit Sub
        !ManufacturerName = txtName.Text
        !ManufacturerAddress = txtAddress.Text
        !manufacturerCountryID = Val(dtcCountry.BoundText)
        !ManufacturerTelephone = txtTelephone.Text
        !ManufacturerFax = txtFax.Text
        !ManufacturerEmail = txtEmail.Text
        !ManufacturerWebsite = txtWebsite.Text
        !ManufacturerComments = txtOther.Text
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call ViewMnufacture
        dtcManufacture.SetFocus
        dtcManufacture.Text = Empty
        Exit Sub
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Edit Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcManufacture.SetFocus
        dtcManufacture.Text = Empty
    End With
End Sub

Private Sub dtcManufacture_Click(Area As Integer)
    If IsNumeric(dtcManufacture.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    Call ViewCountry
    Call ViewMnufacture
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub ViewMnufacture()
    With rsViewManufacture
    If .State = 1 Then .Close
        .Source = "Select* From tblManufacturer Order By ManufacturerName"
        .Open , cnnStores, adOpenStatic, adLockReadOnly
        Set dtcManufacture.RowSource = rsViewManufacture
        dtcManufacture.BoundColumn = "ManufacturerID"
        dtcManufacture.ListField = "ManufacturerName"
    End With
End Sub

Private Sub ViewCountry()
    With rsViewCountry
        If .State = 1 Then .Close
        .Source = "Select* From tblCountry Order By Country"
        .Open , cnnStores, adOpenStatic, adLockReadOnly
        Set dtcCountry.RowSource = rsViewCountry
        dtcCountry.BoundColumn = "CountryID"
        dtcCountry.ListField = "Country"
    End With
End Sub
Private Sub bttnAdd_Click()
    ClearValues
    AfterAdd
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

Private Sub bttnSave_Click()
    Dim TemResponce As Integer
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
    On Error Resume Next
    With rsManufacture
        If .State = 1 Then .Close
        .Source = "Select tblManufacturer.* From tblManufacturer"
        .Open , cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ManufacturerName = txtName.Text
        !ManufacturerAddress = txtAddress.Text
        !manufacturerCountryID = Val(dtcCountry.BoundText)
        !ManufacturerTelephone = txtTelephone.Text
        !ManufacturerFax = txtFax.Text
        !ManufacturerEmail = txtEmail.Text
        !ManufacturerWebsite = txtWebsite.Text
        !ManufacturerComments = txtOther.Text
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call ViewMnufacture
        dtcManufacture.SetFocus
        dtcManufacture.Text = Empty
    Exit Sub
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcManufacture.SetFocus
        dtcManufacture.Text = Empty
    End With
End Sub

Private Sub AfterAdd()
    bttnSave.Visible = True
    bttnCancel.Visible = True
    bttnChange.Visible = False
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
End Sub

Private Sub AfterEdit()
    bttnSave.Visible = False
    bttnCancel.Visible = True
    bttnChange.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
End Sub

Private Sub BeforeAddEdit()
    bttnSave.Visible = False
    bttnCancel.Visible = False
    bttnChange.Visible = False
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    Frame1.Enabled = False
    Frame2.Enabled = True
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
    txtAddress.Text = Empty
    dtcCountry.Text = Empty
    txtTelephone.Text = Empty
    txtFax.Text = Empty
    txtEmail.Text = Empty
    txtWebsite.Text = Empty
    txtOther.Text = Empty
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcManufacture) = False Then Exit Sub
    With rsManufacture
        If .State = 1 Then .Close
        .Source = "Select* From tblManufacturer Where (ManufacturerID =" & dtcManufacture.BoundText & ")"
        .Open , cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Call ClearValues
        If Not (!ManufacturerName) = "" Then txtName.Text = !ManufacturerName
        If Not (!ManufacturerAddress) = "" Then txtAddress.Text = !ManufacturerAddress
        If Not (!manufacturerCountryID) = "" Then dtcCountry.BoundText = !manufacturerCountryID
        If Not (!ManufacturerTelephone) = "" Then txtTelephone.Text = !ManufacturerTelephone
        If Not (!ManufacturerFax) = "" Then txtFax.Text = !ManufacturerFax
        If Not (!ManufacturerFax) = "" Then txtEmail.Text = !ManufacturerEmail
        If Not (!ManufacturerWebsite) = "" Then txtWebsite.Text = !ManufacturerWebsite
        If Not (!ManufacturerComments) = "" Then txtOther.Text = !ManufacturerComments
        TemManufacturerID = !ManufacturerID
        If .State = 1 Then .Close
    End With
End Sub

Private Sub NoName()
    Dim TemResponce As Integer
    TemResponce = MsgBox("You have not entered a Manufacturers Name to save", vbCritical + vbOKOnly, "No Name")
    txtName.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsViewManufacture.State = 1 Then rsViewManufacture.Close: Set rsViewManufacture = Nothing
    If rsViewCountry.State = 1 Then rsViewCountry.Close: Set rsViewCountry = Nothing
End Sub
