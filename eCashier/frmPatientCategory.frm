VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPatientCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Category"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
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
   ScaleHeight     =   4230
   ScaleWidth      =   10410
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   5880
      TabIndex        =   12
      Top             =   2160
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtSurchage 
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   1680
      Width           =   3495
   End
   Begin VB.OptionButton optOutPatient 
      Caption         =   "&Outpatient"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   1320
      Width           =   2895
   End
   Begin VB.OptionButton optInward 
      Caption         =   "&Inward Patient"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtCat 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   480
      Width           =   3495
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Delete"
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
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin MSDataListLib.DataCombo cmbCat 
      Height          =   3060
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5398
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9000
      TabIndex        =   15
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   375
      Left            =   7800
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "&Payment Method"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "S&urcharge"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Ca&tegory"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Pa&tient Category"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmPatientCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbCat.BoundText) = False Then
        temText = cmbCat.Text
    Else
        temText = Empty
    End If
    cmbCat.Text = Empty
    Call EditMode
    txtCat.Text = temText
    txtCat.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbCat.Text = Empty
    cmbCat.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbCat.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatientCategory where PatientCategoryID = " & Val(cmbCat.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
            MsgBox "Deleted"
        Else
            MsgBox "Nothing to Delete"
        End If
        .Close
    End With
    Set rsTem = Nothing
    Call FillCombos
    cmbCat.SetFocus
    cmbCat.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbCat.BoundText) = False Then Exit Sub
    Call EditMode
    txtCat.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtCat.Text) = Empty Then
        MsgBox "You have not entered a Patient Category"
        txtCat.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbCat.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbCat.Text = Empty
    cmbCat.SetFocus
End Sub

Private Sub cmbCat_Change()
    Call ClearValues
    If IsNumeric(cmbCat.BoundText) = True Then Call DisplayDetails
End Sub


'Private Sub SetColours()
'    Me.ForeColor = DefaultColourScheme.LabelForeColour
'    Me.BackColor = DefaultColourScheme.LabelBackColour
'
'    On Error Resume Next
'
'    Dim MyControl As Control
'
'    For Each MyControl In Controls
'        If InStr(UCase(MyControl.Name), "BTN") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.ButtonForeColour
'            MyControl.BackColor = DefaultColourScheme.ButtonBackColour
'            MyControl.BorderColor = DefaultColourScheme.ButtonBorderColour
'        ElseIf InStr(UCase(MyControl.Name), "LST") > 0 Then
'
'        ElseIf InStr(UCase(MyControl.Name), "TXTID") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
'            MyControl.BackColor = DefaultColourScheme.LabelBackColour
'        ElseIf InStr(UCase(MyControl.Name), "CMB") > 0 Then
'
'        ElseIf InStr(UCase(MyControl.Name), "TXT") > 0 Then
'
'        Else
'            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
'            MyControl.BackColor = DefaultColourScheme.LabelBackColour
'        End If
'    Next
'
'End Sub

Private Sub Form_Load()
    Me.Top = GetSetting(App.EXEName, Me.Name, "Top", Me.Top)
    Me.Left = GetSetting(App.EXEName, Me.Name, "Left", Me.Left)
'    Call SetColours
    Call SelectMode
    Call FillCombos
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbCat.Enabled = False
    
    txtCat.Enabled = True
    txtSurchage.Enabled = True
    cmbPaymentMethod.Enabled = True
    optInward.Enabled = True
    optOutPatient.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbCat.Enabled = True
    
    txtCat.Enabled = False
    txtSurchage.Enabled = False
    cmbPaymentMethod.Enabled = False
    optInward.Enabled = False
    optOutPatient.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtCat.Text = Empty
    txtSurchage.Text = Empty
    cmbPaymentMethod.Text = Empty
    optInward.Value = False
    optOutPatient.Value = False
End Sub

Private Sub SaveNew():    On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatientCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !PatientCategory = txtCat.Text
        !Surcharge = Val(txtSurchage.Text)
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        !IndoorPatient = optInward.Value
        !OutdoorPatient = optOutPatient.Value
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub SaveOld():    On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatientCategory where PatientCategoryID = " & Val(cmbCat.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        !PatientCategory = txtCat.Text
        !Surcharge = Val(txtSurchage.Text)
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        !IndoorPatient = optInward.Value
        !OutdoorPatient = optOutPatient.Value
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillAnyCombo cmbCat, "PatientCategory", True
    Dim PM As New clsFillCombos
    PM.FillAnyCombo cmbPaymentMethod, "PaymentMethod", False
End Sub

Private Sub DisplayDetails(): On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatientCategory where PatientCategoryID = " & Val(cmbCat.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtCat.Text = !PatientCategory
            txtSurchage.Text = !Surcharge
            optInward.Value = !IndoorPatient
            optOutPatient.Value = !OutdoorPatient
            cmbPaymentMethod.BoundText = !PaymentMethodID
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "Top", Me.Top
    SaveSetting App.EXEName, Me.Name, "Left", Me.Left
End Sub

