VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmImporters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importers"
   ClientHeight    =   8205
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
   ScaleHeight     =   8205
   ScaleWidth      =   12045
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   240
      TabIndex        =   26
      Top             =   6600
      Width           =   4575
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   360
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
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
   Begin VB.Frame framSerch 
      Height          =   6495
      Left            =   240
      TabIndex        =   25
      Top             =   120
      Width           =   4575
      Begin MSDataListLib.DataCombo dtcImporter 
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
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   4920
      TabIndex        =   24
      Top             =   6600
      Width           =   6855
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
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
         Left            =   4680
         TabIndex        =   13
         Top             =   360
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
         Top             =   360
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
   Begin VB.Frame framAdd 
      Height          =   6495
      Left            =   4920
      TabIndex        =   15
      Top             =   120
      Width           =   6855
      Begin MSDataListLib.DataCombo dtcCountry 
         Height          =   360
         Left            =   2640
         TabIndex        =   5
         Top             =   2520
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
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtAddress 
         Height          =   1575
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtTelephone 
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox txtFax 
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   3600
         Width           =   4095
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   4080
         Width           =   4095
      End
      Begin VB.TextBox txtWebsite 
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   4560
         Width           =   4095
      End
      Begin VB.TextBox txtOther 
         Height          =   1095
         Left            =   2640
         TabIndex        =   10
         Top             =   5040
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Importer's Name"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Importer's Address"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Importer's Country"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Importer's Telephone"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Importer's Fax"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Importer's Email"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Importer's Website"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Other Details"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   5040
         Width           =   2295
      End
   End
   Begin btButtonEx.ButtonEx bttnCLose 
      Height          =   375
      Left            =   9600
      TabIndex        =   14
      Top             =   7680
      Width           =   2055
      _ExtentX        =   3625
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
Attribute VB_Name = "frmImporters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewImpoter As New ADODB.Recordset
    Dim rsViewCity As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
    Dim TemImpoterId As Long
    Dim A As Byte

Private Sub bttnCancel_Click()
    ClearValues
    BeforeAddEdit
    dtcImporter.SetFocus
    dtcImporter.Text = Empty
End Sub

Private Sub bttnChange_Click()
    Dim TemResponce As Integer
    If Trim(txtName.Text) = "" Then A = MsgBox("Please Select Name From List", vbCritical + vbOKOnly, "No Name"): Exit Sub
    On Error Resume Next
    With rsTem
        If .State = 1 Then .Close
        .Open "Select* From tblImporter Where ImporterID = " & dtcImporter.BoundText & " ", cnnStores, 3, 3
        If .RecordCount = 0 Then Exit Sub
        !ImporterName = Trim(txtName.Text)
        !importerAddress = txtAddress.Text
        !ImporterCountryID = Val(dtcCountry.BoundText)
        !importerTelephone = txtTelephone.Text
        !importerFax = txtFax.Text
        !importerEmail = txtEmail.Text
        !importerWebsite = txtWebsite.Text
        !importerComments = txtOther.Text
        .Update
        If .State = 1 Then .Close
        Call FillImporterCombo
        BeforeAddEdit
        ClearValues
        dtcImporter.Text = Empty
        dtcImporter.SetFocus
    Exit Sub
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
    End With
End Sub

Private Sub dtcImporter_Click(Area As Integer)
    If Not IsNumeric(dtcImporter.BoundText) Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    Call FillImporterCombo
    Call FillCountry
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub FillImporterCombo()
    With rsViewImpoter
        If .State = 1 Then .Close
        .Open "Select* From tblImporter Order By ImporterName", cnnStores, 3, 1
        If .RecordCount = 0 Then Exit Sub
        Set dtcImporter.RowSource = rsViewImpoter
        dtcImporter.BoundColumn = "ImporterID"
        dtcImporter.ListField = "ImporterName"
    End With
End Sub

Private Sub FillCountry()
    With rsViewCity
        If .State = 1 Then .Close
        .Open "Select* From tblCountry Order By Country", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtcCountry.RowSource = rsViewCity
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
    SendKeys "{Home}+{end}"
End Sub

Private Sub bttnSave_Click()
    Dim TemResponce As Integer
    If Trim(txtName.Text) = "" Then A = MsgBox("Please Enter Name", vbCritical + vbOKOnly, "No Name"): Exit Sub
    On Error Resume Next
    With rsTem
        If .State = 1 Then .Close
            .Open "Select* From tblImporter Order By ImporterName", cnnStores, 3, 3
            .AddNew
            !ImporterName = Trim(txtName.Text)
            !importerAddress = txtAddress.Text
            !ImporterCountryID = Val(dtcCountry.BoundText)
            !importerTelephone = txtTelephone.Text
            !importerFax = txtFax.Text
            !importerEmail = txtEmail.Text
            !importerWebsite = txtWebsite.Text
            !importerComments = txtOther.Text
            .Update
        If .State = 1 Then .Close
        Call FillImporterCombo
        BeforeAddEdit
        ClearValues
        dtcImporter.Text = Empty
        dtcImporter.SetFocus
        Exit Sub
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcImporter.Text = Empty
        dtcImporter.SetFocus
    End With
End Sub


Private Sub AfterAdd()
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnSave.Enabled = True
    bttnCancel.Enabled = True
    bttnChange.Enabled = False
    framAdd.Enabled = True
    framSerch.Enabled = False
End Sub

Private Sub AfterEdit()
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnSave.Enabled = False
    bttnCancel.Enabled = True
    bttnChange.Enabled = True
    framAdd.Enabled = True
    framSerch.Enabled = False
End Sub

Private Sub BeforeAddEdit()
    bttnAdd.Visible = True
    bttnEdit.Visible = True
    bttnSave.Visible = False
    bttnCancel.Visible = False
    bttnChange.Visible = False
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    bttnSave.Enabled = False
    bttnCancel.Enabled = False
    framAdd.Enabled = False
    framSerch.Enabled = True
End Sub

Function ClearValues()
    txtName.Text = Empty
    txtAddress.Text = Empty
    dtcCountry.BoundText = Empty
    txtTelephone.Text = Empty
    txtFax.Text = Empty
    txtEmail.Text = Empty
    txtWebsite.Text = Empty
    txtOther.Text = Empty
End Function

Private Sub DisplaySelected()
    If Not IsNumeric(dtcImporter.BoundText) Then Exit Sub
    With rsTem
        If .State = 1 Then .Close
        .Open "Select* From tblImporter Where ImporterID = " & dtcImporter.BoundText & " ", cnnStores, 3, 3
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!ImporterName) Then txtName.Text = !ImporterName
        If Not IsNull(!importerAddress) Then txtAddress.Text = !importerAddress
        If Not IsNull(!ImporterCountryID) Then dtcCountry.BoundText = !ImporterCountryID
        If Not IsNull(!importerTelephone) Then txtTelephone.Text = !importerTelephone
        If Not IsNull(!importerFax) Then txtFax.Text = !importerFax
        If Not IsNull(!importerEmail) Then txtEmail.Text = !importerEmail
        If Not IsNull(!importerWebsite) Then txtWebsite.Text = !importerWebsite
        If Not IsNull(!importerComments) Then txtOther.Text = !importerComments
        TemImpoterId = !ImporterID
    End With
End Sub
