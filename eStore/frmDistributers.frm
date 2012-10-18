VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDistributers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distributers"
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
      Begin MSDataListLib.DataCombo dtcDistributor 
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
         Height          =   975
         Left            =   2520
         TabIndex        =   10
         Top             =   4560
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Distributer Name"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Distributer Address"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Distributer City"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Distributer Telephone"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Distributer Fax"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Distributer Email"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Distributer Website"
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
         Top             =   4560
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
Attribute VB_Name = "frmDistributers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewDistributor As New ADODB.Recordset
    Dim rsViewCity As New ADODB.Recordset
    Dim rsDistributor As New ADODB.Recordset
    Dim A As Integer
    Dim TemDistributorId As Long

Private Sub bttnCancel_Click()
    ClearValues
    BeforeAddEdit
    dtcDistributor.Text = Empty
    dtcDistributor.SetFocus
End Sub

Private Sub EditDistributor()
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
    On Error Resume Next
    On Error Resume Next
    With rsDistributor
        If .State = 1 Then .Close
        .Open "Select* From tblDistrubutor Where DistributorID = " & TemDistributorId & "", cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then Exit Sub
        !DistributorName = Trim(txtName.Text)
        !DistributorAddress = txtAddress.Text
        !distributorcityID = Val(dtcCity.BoundText)
        !DistributorTelephone = txtTelephone.Text
        !DistributorFax = txtFax.Text
        !distributorEmail = txtEmail.Text
        !distributorWebsite = txtWebsite.Text
        !distributorComments = txtOther.Text
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call Filldistributor
        dtcDistributor.SetFocus
        dtcDistributor.Text = Empty
        Exit Sub
    
ErrorHandler:
        A = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcDistributor.Text = Empty
        dtcDistributor.SetFocus
    End With
        
End Sub

Private Sub bttnChange_Click()
    Call EditDistributor
End Sub

Private Sub bttnSave_Click()
    Call SaveDistributor
End Sub

Private Sub dtcDistributor_Click(Area As Integer)
    If IsNumeric(dtcDistributor.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    Call Filldistributor
    Call FillCity
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub Filldistributor()
    With rsViewDistributor
        If .State = 1 Then .Close
        .Open "Select* From tblDistrubutor Order By DistributorName", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtcDistributor.RowSource = rsViewDistributor
        dtcDistributor.BoundColumn = "DistributorID"
        dtcDistributor.ListField = "DistributorName"
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

Private Sub SaveDistributor()
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
    On Error Resume Next
    With rsDistributor
        If .State = 1 Then .Close
        .Open "Select* From tblDistrubutor", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !DistributorName = Trim(txtName.Text)
        !DistributorAddress = txtAddress.Text
        !distributorcityID = Val(dtcCity.BoundText)
        !DistributorTelephone = txtTelephone.Text
        !DistributorFax = txtFax.Text
        !distributorEmail = txtEmail.Text
        !distributorWebsite = txtWebsite.Text
        !distributorComments = txtOther.Text
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call Filldistributor
        dtcDistributor.SetFocus
        dtcDistributor.Text = Empty
        Exit Sub
ErrorHandler:
        A = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcDistributor.SetFocus
        dtcDistributor.Text = Empty
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
    TemResponce = MsgBox("No Such distributor found among the records", , "No Record")
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
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcDistributor.BoundText) Then Exit Sub
    With rsDistributor
        If .State = 1 Then .Close
        .Open "Select* From tblDistrubutor Where DistributorID = " & dtcDistributor.BoundText & "", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Call ClearValues
        If Not (!DistributorName) = "" Then txtName.Text = !DistributorName
        If Not (!DistributorAddress) = "" Then txtAddress.Text = !DistributorAddress
        If Not (!distributorcityID) = "" Then dtcCity.BoundText = Val(!distributorcityID)
        If Not (!DistributorTelephone) = "" Then txtTelephone.Text = !DistributorTelephone
        If Not (!DistributorFax) = "" Then txtFax.Text = !DistributorFax
        If Not (!distributorEmail) = "" Then txtEmail.Text = !distributorEmail
        If Not (!distributorWebsite) = "" Then txtWebsite.Text = !distributorWebsite
        If Not (!distributorComments) = "" Then txtOther.Text = !distributorComments
        TemDistributorId = !DistributorID
        If .RecordCount = 0 Then Exit Sub
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsViewDistributor.State = 1 Then rsViewDistributor.Close: Set rsViewDistributor = Nothing
    If rsViewCity.State = 1 Then rsViewCity.Close: Set rsViewCity = Nothing
End Sub
