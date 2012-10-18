VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmServiceCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Category"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   6735
   ScaleWidth      =   11760
   Begin VB.CheckBox chkExpence 
      Caption         =   "Expence"
      Height          =   255
      Left            =   6480
      TabIndex        =   25
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CheckBox chkR 
      Caption         =   "For Roentgents "
      Height          =   255
      Left            =   6480
      TabIndex        =   24
      Top             =   4440
      Width           =   3495
   End
   Begin VB.CheckBox chkHST 
      Caption         =   "For HST"
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   4080
      Width           =   3495
   End
   Begin VB.CheckBox chkMT 
      Caption         =   "For MT"
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   3720
      Width           =   3495
   End
   Begin VB.CheckBox chkLab 
      Caption         =   "For Lab"
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   3360
      Width           =   3495
   End
   Begin VB.CheckBox chkGSB 
      Caption         =   "For GSB"
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   2640
      Width           =   3495
   End
   Begin VB.CheckBox chkBHT 
      Caption         =   "For BHT"
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CheckBox chkOPD 
      Caption         =   "For OPD"
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox txtInwardSurcharge 
      Height          =   360
      Left            =   6480
      TabIndex        =   10
      Top             =   1320
      Width           =   4575
   End
   Begin VB.CheckBox chkToMedicineCharge 
      Caption         =   "To &Medicine charge"
      Height          =   255
      Left            =   8520
      TabIndex        =   12
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CheckBox chkCanChange 
      Caption         =   "Can C&hange"
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox txtFee 
      Height          =   360
      Left            =   6480
      TabIndex        =   8
      Top             =   840
      Width           =   4575
   End
   Begin VB.TextBox txtComments 
      Height          =   720
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   5160
      Width           =   4575
   End
   Begin VB.TextBox txtServiceCategory 
      Height          =   360
      Left            =   6480
      TabIndex        =   6
      Top             =   360
      Width           =   4575
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add"
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
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Edit"
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Delete"
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
   Begin MSDataListLib.DataCombo cmbServiceCategory 
      Height          =   5460
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9631
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   10200
      TabIndex        =   23
      Top             =   6480
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
      Left            =   9000
      TabIndex        =   22
      Top             =   6000
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
      Left            =   7680
      TabIndex        =   21
      Top             =   6000
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
      Caption         =   "&Inward Surcharge"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "&Fee"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Ca&tegory"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Service Ca&tegory"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmServiceCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbServiceCategory.BoundText) = False Then
        temText = cmbServiceCategory.Text
    Else
        temText = Empty
    End If
    cmbServiceCategory.Text = Empty
    Call EditMode
    txtServiceCategory.Text = temText
    txtServiceCategory.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbServiceCategory.Text = Empty
    cmbServiceCategory.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbServiceCategory.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblServiceCategory where ServiceCategoryID = " & Val(cmbServiceCategory.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
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
    cmbServiceCategory.SetFocus
    cmbServiceCategory.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbServiceCategory.BoundText) = False Then Exit Sub
    Call EditMode
    txtServiceCategory.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtServiceCategory.Text) = Empty Then
        MsgBox "You have not entered an ServiceCategory"
        txtServiceCategory.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbServiceCategory.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbServiceCategory.Text = Empty
    cmbServiceCategory.SetFocus
End Sub

Private Sub cmbServiceCategory_Change()
    Call ClearValues
    If IsNumeric(cmbServiceCategory.BoundText) = True Then Call displayDetails
End Sub


Private Sub cmbServiceCategory_Click(Area As Integer)
    Call ClearValues
    If IsNumeric(cmbServiceCategory.BoundText) = True Then Call displayDetails
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
    cmbServiceCategory.Enabled = False
    
    txtServiceCategory.Enabled = True
    txtComments.Enabled = True
    txtFee.Enabled = True
    chkCanChange.Enabled = True
    chkToMedicineCharge.Enabled = True
    txtInwardSurcharge.Enabled = True
    chkBHT.Enabled = True
    chkGSB.Enabled = True
    chkOPD.Enabled = True
    chkLab.Enabled = True
    chkMT.Enabled = True
    chkHST.Enabled = True
    chkR.Enabled = True
    chkExpence.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbServiceCategory.Enabled = True
    
    txtServiceCategory.Enabled = False
    txtComments.Enabled = False
    txtFee.Enabled = False
    chkCanChange.Enabled = False
    chkToMedicineCharge.Enabled = False
    txtInwardSurcharge.Enabled = False
    chkBHT.Enabled = False
    chkGSB.Enabled = False
    chkOPD.Enabled = False
    chkLab.Enabled = False
    chkMT.Enabled = False
    chkHST.Enabled = False
    chkR.Enabled = False
    chkExpence.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtServiceCategory.Text = Empty
    txtComments.Text = Empty
    chkCanChange.Value = 0
    txtFee.Text = Empty
    chkToMedicineCharge.Value = 0
    txtInwardSurcharge.Text = Empty
    chkBHT.Value = 0
    chkGSB.Value = 0
    chkOPD.Value = 0
    chkLab.Value = 0
    chkMT.Value = 0
    chkHST.Value = 0
    chkR.Value = 0
    chkExpence.Value = 0
End Sub

Private Sub SaveNew():    On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblServiceCategory"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ServiceCategory = txtServiceCategory.Text
        !Comments = txtComments.Text
        !Fee = Val(txtFee.Text)
        If chkCanChange.Value = 1 Then
            !CanChange = True
        Else
            !CanChange = False
        End If
        If chkToMedicineCharge.Value = 1 Then
            !ToMedicineCharge = True
        Else
            !ToMedicineCharge = False
        End If
        
        If chkOPD.Value = 1 Then
            !ForOPD = True
        Else
            !ForOPD = False
        End If

        If chkLab.Value = 1 Then
            !ForLab = True
        Else
            !ForLab = False
        End If
        
        If chkBHT.Value = 1 Then
            !ForBHT = True
        Else
            !ForBHT = False
        End If
        
        If chkGSB.Value = 1 Then
            !ForGSB = True
        Else
            !ForGSB = False
        End If

        If chkMT.Value = 1 Then
            !ForMT = True
        Else
            !ForMT = False
        End If
        
        If chkHST.Value = 1 Then
            !ForHST = True
        Else
            !ForHST = False
        End If
        
        If chkR.Value = 1 Then
            !ForR = True
        Else
            !ForR = False
        End If

        If chkExpence.Value = 1 Then
            !ForExpence = True
        Else
            !ForExpence = False
        End If
        
        !InwardSurcharge = Val(txtInwardSurcharge.Text)
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub SaveOld():    On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblServiceCategory where ServiceCategoryID = " & Val(cmbServiceCategory.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        !ServiceCategory = txtServiceCategory.Text
        !Comments = txtComments.Text
        !Fee = Val(txtFee.Text)
        If chkCanChange.Value = 1 Then
            !CanChange = True
        Else
            !CanChange = False
        End If
        If chkToMedicineCharge.Value = 1 Then
            !ToMedicineCharge = True
        Else
            !ToMedicineCharge = False
        End If
        
        If chkOPD.Value = 1 Then
            !ForOPD = True
        Else
            !ForOPD = False
        End If

        If chkLab.Value = 1 Then
            !ForLab = True
        Else
            !ForLab = False
        End If
        
        If chkBHT.Value = 1 Then
            !ForBHT = True
        Else
            !ForBHT = False
        End If
        
        If chkGSB.Value = 1 Then
            !ForGSB = True
        Else
            !ForGSB = False
        End If

        If chkMT.Value = 1 Then
            !ForMT = True
        Else
            !ForMT = False
        End If
        
        If chkHST.Value = 1 Then
            !ForHST = True
        Else
            !ForHST = False
        End If

        If chkR.Value = 1 Then
            !ForR = True
        Else
            !ForR = False
        End If
        If chkExpence.Value = 1 Then
            !ForExpence = True
        Else
            !ForExpence = False
        End If
        
        !InwardSurcharge = Val(txtInwardSurcharge.Text)
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillAnyCombo cmbServiceCategory, "ServiceCategory", True
End Sub

Private Sub displayDetails(): On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblServiceCategory where ServiceCategoryID = " & Val(cmbServiceCategory.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtServiceCategory.Text = !ServiceCategory
            txtComments.Text = Format(!Comments, "0")
            txtFee.Text = Format(!Fee, "0.00")
            If !CanChange = True Then
                chkCanChange.Value = 1
            Else
                chkCanChange.Value = 0
            End If
            If !ToMedicineCharge = True Then
                chkToMedicineCharge.Value = 1
            Else
                chkToMedicineCharge.Value = 0
            End If
            
            If !ForOPD = True Then
                chkOPD.Value = 1
            Else
                chkOPD.Value = 0
            End If
            
            If !ForHST = True Then
                chkHST.Value = 1
            Else
                chkHST.Value = 0
            End If

            If !ForMT = True Then
                chkMT.Value = 1
            Else
                chkMT.Value = 0
            End If

            If !ForGSB = True Then
                chkGSB.Value = 1
            Else
                chkGSB.Value = 0
            End If

            If !ForBHT = True Then
                chkBHT.Value = 1
            Else
                chkBHT.Value = 0
            End If

            If !ForLab = True Then
                chkLab.Value = 1
            Else
                chkLab.Value = 0
            End If
            If !ForR = True Then
                chkR.Value = 1
            Else
                chkR.Value = 0
            End If
            If !ForExpence = True Then
                chkExpence.Value = 1
            Else
                chkExpence.Value = 0
            End If
            
            If Not IsNull(!InwardSurcharge) Then
                txtInwardSurcharge.Text = !InwardSurcharge
            Else
                txtInwardSurcharge.Text = 0
            End If
            
            
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "Top", Me.Top
    SaveSetting App.EXEName, Me.Name, "Left", Me.Left
End Sub
