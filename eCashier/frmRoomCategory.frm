VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmRoomCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Category"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11235
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
   ScaleHeight     =   5130
   ScaleWidth      =   11235
   Begin VB.CheckBox chkICUNursing 
      Caption         =   "ICU Nursing"
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtSurcharge 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6480
      TabIndex        =   11
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox txtDiscountForCash 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6480
      TabIndex        =   9
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox txtGeneralCharge 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6480
      TabIndex        =   7
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtComments 
      Height          =   840
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2760
      Width           =   4575
   End
   Begin VB.TextBox txtRoomCategory 
      Height          =   360
      Left            =   6480
      TabIndex        =   6
      Top             =   480
      Width           =   4575
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4560
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
      Top             =   4560
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
      Top             =   4560
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
   Begin MSDataListLib.DataCombo cmbRoomCategory 
      Height          =   4020
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7091
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9840
      TabIndex        =   18
      Top             =   4320
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
      Left            =   8640
      TabIndex        =   16
      Top             =   3720
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
      Left            =   7320
      TabIndex        =   15
      Top             =   3720
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
   Begin VB.Label Label5 
      Caption         =   "Sur. for Credit"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Discount for Cash"
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "General Charge"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Comments"
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Category"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Room Category"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmRoomCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbRoomCategory.BoundText) = False Then
        temText = cmbRoomCategory.Text
    Else
        temText = Empty
    End If
    cmbRoomCategory.Text = Empty
    Call EditMode
    txtRoomCategory.Text = temText
    txtRoomCategory.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbRoomCategory.Text = Empty
    cmbRoomCategory.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbRoomCategory.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblRoomCategory where RoomCategoryID = " & Val(cmbRoomCategory.BoundText)
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
    cmbRoomCategory.SetFocus
    cmbRoomCategory.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbRoomCategory.BoundText) = False Then Exit Sub
    Call EditMode
    txtRoomCategory.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtRoomCategory.Text) = Empty Then
        MsgBox "You have not entered an RoomCategory"
        txtRoomCategory.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbRoomCategory.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbRoomCategory.Text = Empty
    cmbRoomCategory.SetFocus
End Sub

Private Sub cmbRoomCategory_Change()
    Call ClearValues
    If IsNumeric(cmbRoomCategory.BoundText) = True Then Call DisplayDetails
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
    cmbRoomCategory.Enabled = False
    
    txtRoomCategory.Enabled = True
    txtComments.Enabled = True
    txtGeneralCharge.Enabled = True
    txtDiscountForCash.Enabled = True
    txtSurcharge.Enabled = True
    chkICUNursing.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbRoomCategory.Enabled = True
    
    txtRoomCategory.Enabled = False
    txtComments.Enabled = False
    txtGeneralCharge.Enabled = False
    txtDiscountForCash.Enabled = False
    txtSurcharge.Enabled = False
    chkICUNursing.Enabled = False
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtRoomCategory.Text = Empty
    txtComments.Text = Empty
    txtGeneralCharge.Text = Empty
    txtDiscountForCash.Text = Empty
    txtSurcharge.Text = Empty
    chkICUNursing.Value = 0
End Sub

Private Sub SaveNew():    On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblRoomCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !RoomCategory = txtRoomCategory.Text
        !Comments = txtComments.Text
        !GeneralCharge = Val(txtGeneralCharge.Text)
        !DiscountForCash = Val(txtDiscountForCash.Text)
        !SurchargeForCredit = Val(txtSurcharge.Text)
        If chkICUNursing.Value = 1 Then
            !ICUNursing = True
        Else
            !ICUNursing = False
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
        temSql = "Select * from tblRoomCategory where RoomCategoryID = " & Val(cmbRoomCategory.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        !RoomCategory = txtRoomCategory.Text
        !Comments = txtComments.Text
        !GeneralCharge = Val(txtGeneralCharge.Text)
        !DiscountForCash = Val(txtDiscountForCash.Text)
        !SurchargeForCredit = Val(txtSurcharge.Text)
        If chkICUNursing.Value = 1 Then
            !ICUNursing = True
        Else
            !ICUNursing = False
        End If
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillAnyCombo cmbRoomCategory, "RoomCategory", True
End Sub

Private Sub DisplayDetails(): On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblRoomCategory where RoomCategoryID = " & Val(cmbRoomCategory.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtRoomCategory.Text = !RoomCategory
            txtComments.Text = Format(!Comments, "0")
            txtGeneralCharge.Text = Format(!GeneralCharge, "0.00")
            txtDiscountForCash.Text = Format(!DiscountForCash, "0.00")
            txtSurcharge.Text = Format(!SurchargeForCredit, "0.00")
            If !ICUNursing = True Then
                chkICUNursing.Value = 1
            Else
                chkICUNursing.Value = 0
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
