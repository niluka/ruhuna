VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmServiceSubCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service subcategory"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
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
   ScaleHeight     =   7455
   ScaleWidth      =   10905
   Begin VB.CheckBox chkHourlyRate 
      Caption         =   "Hourly Rate"
      Height          =   255
      Left            =   8160
      TabIndex        =   28
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CheckBox chkExpence 
      Caption         =   "Expence"
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   4920
      Width           =   3495
   End
   Begin VB.CheckBox chkR 
      Caption         =   "For Roentgents "
      Height          =   255
      Left            =   6120
      TabIndex        =   26
      Top             =   4560
      Width           =   3495
   End
   Begin VB.CheckBox chkOPD 
      Caption         =   "For OPD"
      Height          =   255
      Left            =   6120
      TabIndex        =   25
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CheckBox chkBHT 
      Caption         =   "For BHT"
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   2400
      Width           =   3495
   End
   Begin VB.CheckBox chkGSB 
      Caption         =   "For GSB"
      Height          =   255
      Left            =   6120
      TabIndex        =   23
      Top             =   2760
      Width           =   3495
   End
   Begin VB.CheckBox chkLab 
      Caption         =   "For Lab"
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   3480
      Width           =   3495
   End
   Begin VB.CheckBox chkMT 
      Caption         =   "For MT"
      Height          =   255
      Left            =   6120
      TabIndex        =   21
      Top             =   3840
      Width           =   3495
   End
   Begin VB.CheckBox chkHST 
      Caption         =   "For HST"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   4200
      Width           =   3495
   End
   Begin VB.CheckBox chkToMedicineCharge 
      Caption         =   "To Medicine Charge"
      Height          =   255
      Left            =   8160
      TabIndex        =   19
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CheckBox chkCanChange 
      Caption         =   "Can C&hange"
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtFee 
      Height          =   360
      Left            =   6120
      TabIndex        =   12
      Top             =   1440
      Width           =   4575
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtComments 
      Height          =   720
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   5280
      Width           =   4575
   End
   Begin VB.TextBox txtSubCategory 
      Height          =   360
      Left            =   6120
      TabIndex        =   8
      Top             =   480
      Width           =   4575
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6840
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
      TabIndex        =   5
      Top             =   6840
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
      TabIndex        =   6
      Top             =   6840
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
   Begin MSDataListLib.DataCombo cmbSubCategory 
      Height          =   5700
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10054
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9240
      TabIndex        =   18
      Top             =   6960
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
      Left            =   8040
      TabIndex        =   17
      Top             =   6360
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
      Left            =   6720
      TabIndex        =   16
      Top             =   6360
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
   Begin MSDataListLib.DataCombo cmbECategory 
      Height          =   360
      Left            =   6120
      TabIndex        =   10
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label5 
      Caption         =   "&Fee"
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Cate&gory"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Service S&ubcategory"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label11 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "S&ubcategory"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Service Cate&gory"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmServiceSubCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsSC As New ADODB.Recordset
    
Private Sub btnAdd_Click()
    On Error Resume Next
    
    Dim temText As String
    If IsNumeric(cmbSubCategory.BoundText) = False Then
        temText = cmbSubCategory.Text
    Else
        temText = Empty
    End If
    cmbSubCategory.Text = Empty
    Call EditMode
    cmbECategory.BoundText = Val(cmbCategory.BoundText)
    txtSubCategory.Text = temText
    txtSubCategory.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbSubCategory.Text = Empty
    cmbSubCategory.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbSubCategory.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSubCategory where ServiceSubCategoryID = " & Val(cmbSubCategory.BoundText)
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
    cmbSubCategory.SetFocus
    cmbSubCategory.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbSubCategory.BoundText) = False Then Exit Sub
    Call EditMode
    txtSubCategory.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtSubCategory.Text) = Empty Then
        MsgBox "You have not entered an ServiceCategory"
        txtSubCategory.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbECategory.BoundText) = False Then
        MsgBox "Please select a service category"
        cmbECategory.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSubCategory.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbSubCategory.Text = Empty
    cmbSubCategory.SetFocus
End Sub

Private Sub cmbCategory_Change()
    With rsSC
        If .State = 1 Then .Close
        If IsNumeric(cmbCategory.BoundText) = True Then
            temSql = "Select * from tblServiceSubCategory where Deleted = 0 AND ServiceCategoryID = " & Val(cmbCategory.BoundText) & " Order by ServiceSubcategory"
        Else
            temSql = "Select * from tblServiceSubCategory where Deleted = 0 ORDER BY Servicesubcategory"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSubCategory
        Set .RowSource = rsSC
        .ListField = "ServiceSubCategory"
        .BoundColumn = "ServiceSubCategoryID"
        .Text = Empty
    End With
End Sub

Private Sub cmbCategory_Click(Area As Integer)
    With rsSC
        If .State = 1 Then .Close
        If IsNumeric(cmbCategory.BoundText) = True Then
            temSql = "Select * from tblServiceSubCategory where Deleted = 0 AND ServiceCategoryID = " & Val(cmbCategory.BoundText) & " Order by ServiceSubcategory"
        Else
            temSql = "Select * from tblServiceSubCategory where Deleted = 0 ORDER BY Servicesubcategory"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSubCategory
        Set .RowSource = rsSC
        .ListField = "ServiceSubCategory"
        .BoundColumn = "ServiceSubCategoryID"
    End With
End Sub

Private Sub cmbSubCategory_Change()
    Call ClearValues
    If IsNumeric(cmbSubCategory.BoundText) = True Then Call DisplayDetails
End Sub


Private Sub cmbSubCategory_Click(Area As Integer)
    Call ClearValues
    If IsNumeric(cmbSubCategory.BoundText) = True Then Call DisplayDetails
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
    
    Dim SCat As New clsFillCombos
    SCat.FillAnyCombo cmbCategory, "ServiceCategory", True
    Dim ECat As New clsFillCombos
    ECat.FillAnyCombo cmbECategory, "ServiceCategory", True

'    Call SetColours
    Call SelectMode
    Call FillCombos
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbSubCategory.Enabled = False
    cmbCategory.Enabled = False
    
    txtSubCategory.Enabled = True
    txtComments.Enabled = True
    cmbECategory.Enabled = True
    txtFee.Enabled = True
    chkCanChange.Enabled = True
    chkToMedicineCharge.Enabled = True
    chkHourlyRate.Enabled = True
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
    cmbSubCategory.Enabled = True
    cmbCategory.Enabled = True
    
    txtSubCategory.Enabled = False
    txtComments.Enabled = False
    txtFee.Enabled = False
    chkCanChange.Enabled = False
    chkToMedicineCharge.Enabled = False
    chkHourlyRate.Enabled = False
    chkBHT.Enabled = False
    chkGSB.Enabled = False
    chkOPD.Enabled = False
    chkLab.Enabled = False
    chkMT.Enabled = False
    chkHST.Enabled = False
    chkR.Enabled = False
    chkExpence.Enabled = False
    
    cmbECategory.Enabled = False
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtSubCategory.Text = Empty
    txtComments.Text = Empty
    cmbECategory.Text = Empty
    txtFee.Text = Empty
    chkCanChange.Value = 0
    chkToMedicineCharge.Value = 0
    chkHourlyRate.Value = 0
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
        temSql = "Select * from tblServiceSubCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ServiceCategoryID = Val(cmbECategory.BoundText)
        !ServiceSubcategory = txtSubCategory.Text
        !Comments = txtComments.Text
        !Fee = Val(txtFee.Text)
        If chkCanChange.Value = 1 Then
            !CanChange = True
        Else
            !CanChange = False
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
        
        If chkToMedicineCharge.Value = 1 Then
            !ToMedicineCharge = True
        Else
            !ToMedicineCharge = False
        End If

        If chkHourlyRate.Value = 1 Then
            !HourlyRate = True
        Else
            !HourlyRate = False
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
                
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub SaveOld():    On Error Resume Next

    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSubCategory where ServiceSubCategoryID = " & Val(cmbSubCategory.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        !ServiceCategoryID = Val(cmbECategory.BoundText)
        !ServiceSubcategory = txtSubCategory.Text
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
        
        If chkHourlyRate.Value = 1 Then
            !HourlyRate = True
        Else
            !HourlyRate = False
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
        
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim SCat As New clsFillCombos
    SCat.FillAnyCombo cmbSubCategory, "ServiceSubCategory", True
End Sub

Private Sub DisplayDetails(): On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSubCategory where ServiceSubCategoryID = " & Val(cmbSubCategory.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtSubCategory.Text = !ServiceSubcategory
            cmbECategory.BoundText = !ServiceCategoryID
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
            If !HourlyRate = True Then
                chkHourlyRate.Value = 1
            Else
                chkHourlyRate.Value = 0
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
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "Top", Me.Top
    SaveSetting App.EXEName, Me.Name, "Left", Me.Left
End Sub
