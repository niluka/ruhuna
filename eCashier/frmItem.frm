VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Services"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
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
   ScaleHeight     =   5820
   ScaleWidth      =   11460
   Begin VB.TextBox txtProfessionalCharge 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6720
      TabIndex        =   16
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox txtHospitalCharge 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6720
      TabIndex        =   15
      Top             =   2400
      Width           =   4575
   End
   Begin VB.TextBox txtTotalCharge 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6720
      TabIndex        =   14
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox txtComments 
      Height          =   840
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3360
      Width           =   4575
   End
   Begin VB.TextBox txtItem 
      Height          =   360
      Left            =   6720
      TabIndex        =   6
      Top             =   480
      Width           =   4575
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4800
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
      Top             =   4800
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
      Top             =   4800
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
   Begin MSDataListLib.DataCombo cmbItem 
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
      Left            =   10200
      TabIndex        =   13
      Top             =   5400
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
      TabIndex        =   12
      Top             =   4800
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
      TabIndex        =   11
      Top             =   4800
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
   Begin MSDataListLib.DataCombo cmbSaleCategory 
      Height          =   360
      Left            =   6720
      TabIndex        =   8
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff 
      Height          =   360
      Left            =   6720
      TabIndex        =   17
      Top             =   1440
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label7 
      Caption         =   "Staff"
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Professioanl Charge"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Hospital Charge"
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Total Charge"
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Sale Category"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Comments"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Service"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Services"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbItem.BoundText) = False Then
        temText = cmbItem.Text
    Else
        temText = Empty
    End If
    cmbItem.Text = Empty
    Call EditMode
    txtItem.Text = temText
    txtItem.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbItem.Text = Empty
    cmbItem.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbItem.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem where ItemID = " & Val(cmbItem.BoundText)
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
    cmbItem.SetFocus
    cmbItem.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbItem.BoundText) = False Then Exit Sub
    Call EditMode
    txtItem.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtItem.Text) = Empty Then
        MsgBox "You have not entered a service"
        txtItem.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSaleCategory.BoundText) = False Then
        MsgBox "Please select a sale category"
        cmbSaleCategory.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbItem.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbItem.Text = Empty
    cmbItem.SetFocus
End Sub


Private Sub cmbItem_Change()
    Call ClearValues
    If IsNumeric(cmbItem.BoundText) = True Then Call DisplayDetails
End Sub


Private Sub Form_Load()
    Call SelectMode
    Call FillCombos
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbItem.Enabled = False
    
    txtItem.Enabled = True
    txtProfessionalCharge.Enabled = True
    txtHospitalCharge.Enabled = True
    txtTotalCharge.Enabled = True
    cmbStaff.Enabled = True
    txtComments.Enabled = True
    cmbSaleCategory.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbItem.Enabled = True
    
    txtItem.Enabled = False
    txtComments.Enabled = False
    cmbSaleCategory.Enabled = False
    txtProfessionalCharge.Enabled = False
    txtHospitalCharge.Enabled = False
    txtTotalCharge.Enabled = False
    cmbStaff.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtItem.Text = Empty
    txtComments.Text = Empty
    cmbSaleCategory.Text = Empty
    cmbStaff.Text = Empty
    txtProfessionalCharge.Text = Empty
    txtHospitalCharge.Text = Empty
    txtTotalCharge.Text = Empty
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    Dim temItem As New clsItem
    temItem.ID = Val(cmbSaleCategory.BoundText)
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Item = txtItem.Text
        !Comments = txtComments.Text
        !SaleCategoryID = Val(cmbSaleCategory.BoundText)
        !GroupID = temItem.GroupID
        !SubGroupID = temItem.SubGroupID
        !IsTradeName = True
        .Update
        !TradeNameID = !ItemID
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    Dim temItem As New clsItem
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem where ItemID = " & Val(cmbItem.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        !Item = txtItem.Text
        !Comments = txtComments.Text
        !GenericNameID = Val(cmbSaleCategory.BoundText)
        !GroupID = temItem.GroupID
        !SubGroupID = temItem.SubGroupID
        !TradeNameID = !ItemID
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim It As New clsFillCombos
    It.FillBoolCombo cmbItem, "Item", "Item", "IsTradeName", True
    Dim Generic As New clsFillCombos
    Generic.FillBoolCombo cmbSaleCategory, "Item", "Item", "IsGenericName", True
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem where ItemID = " & Val(cmbItem.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtItem.Text = !Item
            txtComments.Text = !Comments
            cmbSaleCategory.BoundText = !GenericNameID
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub
