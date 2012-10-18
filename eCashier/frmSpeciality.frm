VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSpeciality 
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
      TabIndex        =   12
      Top             =   6720
      Width           =   6975
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   4920
         TabIndex        =   6
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
         TabIndex        =   4
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
         TabIndex        =   5
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   240
      Width           =   4575
      Begin MSDataListLib.DataCombo dtcSpeciality 
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
      TabIndex        =   8
      Top             =   240
      Width           =   6975
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
   End
   Begin btButtonEx.ButtonEx bttnCLose 
      Height          =   375
      Left            =   9840
      TabIndex        =   7
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
Attribute VB_Name = "frmSpeciality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsViewSpeciality As New ADODB.Recordset
    Dim rsSpeciality As New ADODB.Recordset
    Dim A As Integer
    Dim TemSpecialityId As Long

Private Sub bttnCancel_Click()
    ClearValues
    BeforeAddEdit
    dtcSpeciality.Text = Empty
    dtcSpeciality.SetFocus
End Sub

Private Sub EditSpeciality()
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
    On Error GoTo ErrorHandler
    With rsSpeciality
        If .State = 1 Then .Close
        .Open "Select* From tblSpeciality Where SpecialityID = " & TemSpecialityId & "", cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount = 0 Then Exit Sub
        !Speciality = Trim(txtName.Text)
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call FillSpeciality
        dtcSpeciality.SetFocus
        dtcSpeciality.Text = Empty
        Exit Sub
    
ErrorHandler:
        A = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcSpeciality.Text = Empty
        dtcSpeciality.SetFocus
    End With
        
End Sub

Private Sub bttnChange_Click()
    Call EditSpeciality
End Sub

Private Sub bttnSave_Click()
    Call SaveSpeciality
End Sub

Private Sub dtcSpeciality_Click(Area As Integer)
    If IsNumeric(dtcSpeciality.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    Call FillSpeciality
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub FillSpeciality()
    With rsViewSpeciality
        If .State = 1 Then .Close
        .Open "Select* From tblSpeciality Order By Speciality", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Set dtcSpeciality.RowSource = rsViewSpeciality
        dtcSpeciality.BoundColumn = "SpecialityID"
        dtcSpeciality.ListField = "Speciality"
    End With
End Sub

Private Sub bttnAdd_Click()
    ClearValues
    AfterAdd
    txtName.Text = dtcSpeciality.Text
    dtcSpeciality.Text = Empty
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

Private Sub SaveSpeciality()
    If Trim(txtName.Text) = "" Then Call NoName: Exit Sub
'    On Error GoTo ErrorHandler
    With rsSpeciality
        If .State = 1 Then .Close
        .Open "Select* From tblSpeciality", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Speciality = Trim(txtName.Text)
        .Update
        If .State = 1 Then .Close
        BeforeAddEdit
        ClearValues
        Call FillSpeciality
        dtcSpeciality.SetFocus
        dtcSpeciality.Text = Empty
        Exit Sub
ErrorHandler:
        A = MsgBox(Err.Number & vbNewLine & Err.Description & vbTab & Me.Caption, vbCritical + vbOKOnly, "Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
        dtcSpeciality.SetFocus
        dtcSpeciality.Text = Empty
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
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcSpeciality.BoundText) Then Exit Sub
    With rsSpeciality
        If .State = 1 Then .Close
        .Open "Select* From tblSpeciality Where SpecialityID = " & dtcSpeciality.BoundText & "", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount = 0 Then Exit Sub
        Call ClearValues
        If Not IsNull(!Speciality) Then txtName.Text = !Speciality
        TemSpecialityId = !SpecialityID
        If .RecordCount = 0 Then Exit Sub
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsViewSpeciality.State = 1 Then rsViewSpeciality.Close: Set rsViewSpeciality = Nothing
End Sub
