VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmServiceSecession 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Secession"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11385
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
   ScaleHeight     =   4605
   ScaleWidth      =   11385
   Begin VB.TextBox txtDuration 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6720
      TabIndex        =   12
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox txtMaxNo 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6720
      TabIndex        =   14
      Top             =   2400
      Width           =   4575
   End
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   1440
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Format          =   70909954
      CurrentDate     =   39995
   End
   Begin VB.TextBox txtComments 
      Height          =   840
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox txtServiceSecession 
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
      Top             =   4080
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
      Top             =   4080
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
      Top             =   4080
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
   Begin MSDataListLib.DataCombo cmbServiceSecession 
      Height          =   3540
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6244
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9960
      TabIndex        =   19
      Top             =   5040
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
      Left            =   8880
      TabIndex        =   18
      Top             =   3960
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
      Left            =   7560
      TabIndex        =   17
      Top             =   3960
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
   Begin MSDataListLib.DataCombo cmbServiceCategory 
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
   Begin VB.Label Label7 
      Caption         =   "Max No"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Start Time"
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Duration in Minutes"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Service category"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Comments"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Secession"
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Secession Name"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmServiceSecession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsSC As New ADODB.Recordset
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbServiceSecession.BoundText) = False Then
        temText = cmbServiceSecession.Text
    Else
        temText = Empty
    End If
    cmbServiceSecession.Text = Empty
    Call EditMode
    txtServiceSecession.Text = temText
    txtServiceSecession.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbServiceSecession.Text = Empty
    cmbServiceSecession.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbServiceSecession.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSecession where ServiceSecessionID = " & Val(cmbServiceSecession.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            !DeletedDateTime = Now
            .Update
            MsgBox "Deleted"
        Else
            MsgBox "Nothing to Delete"
        End If
        .Close
    End With
    Set rsTem = Nothing
    Call FillCombos
    cmbServiceSecession.SetFocus
    cmbServiceSecession.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbServiceSecession.BoundText) = False Then Exit Sub
    Call EditMode
    txtServiceSecession.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtServiceSecession.Text) = Empty Then
        MsgBox "You have not entered a Service Secession"
        txtServiceSecession.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbServiceSecession.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbServiceSecession.Text = Empty
    cmbServiceSecession.SetFocus
End Sub

'Private Sub cmbServiceCategory_Change()
'    With rsSC
'        If .State = 1 Then .Close
'        temSql = "Select * from tblServiceSubCategory where ServiceCategoryID = " & Val(cmbServiceCategory.BoundText) & " Order by ServiceSubcategory"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            cmbServiceSubcategory.Visible = True
'        Else
'            cmbServiceSubcategory.Visible = False
'        End If
'    End With
'    With cmbServiceSubcategory
'        Set .RowSource = rsSC
'        .ListField = "ServiceSubcategory"
'        .BoundColumn = "ServiceSubcategoryID"
'    End With
'End Sub

Private Sub cmbServiceSecession_Change()
    Call ClearValues
    If IsNumeric(cmbServiceSecession.BoundText) = True Then Call DisplayDetails
End Sub

'Private Sub cmbServiceSubcategory_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then
'        cmbServiceSubcategory.Text = Empty
'    End If
'End Sub

Private Sub cmbServiceCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbServiceCategory.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call SelectMode
    Call FillCombos
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbServiceSecession.Enabled = False
    
    txtServiceSecession.Enabled = True
    txtComments.Enabled = True
    txtMaxNo.Enabled = True
    txtDuration.Enabled = True
    dtpStart.Enabled = True
    cmbServiceCategory.Enabled = True
'    cmbServiceSubcategory.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbServiceSecession.Enabled = True
    
    txtServiceSecession.Enabled = False
    txtComments.Enabled = False
    txtMaxNo.Enabled = False
    txtDuration.Enabled = False
    dtpStart.Enabled = False
    cmbServiceCategory.Enabled = False
'    cmbServiceSubcategory.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtServiceSecession.Text = Empty
    txtComments.Text = Empty
    cmbServiceCategory.Text = Empty
'    cmbServiceSubcategory.Text = Empty
    txtMaxNo.Text = Empty
    txtDuration.Text = Empty
    dtpStart.Value = "00:00:00"
End Sub

Private Sub SaveNew():    On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSecession"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ServiceSecession = txtServiceSecession.Text
        !Comments = txtComments.Text
        !ServiceCategoryID = Val(cmbServiceCategory.BoundText)
        !ServiceSubcategoryID = 0 'Val(cmbServiceSubcategory.BoundText)
        !AddedDate = Date
        !AddedTime = Now
        !AddedUserID = UserID
        !MaxNo = Val(txtMaxNo.Text)
        !StartTime = dtpStart.Value
        !DurationMinutes = Val(txtDuration.Text)
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub SaveOld():    On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSecession where ServiceSecessionID = " & Val(cmbServiceSecession.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !ServiceSecession = txtServiceSecession.Text
            !Comments = txtComments.Text
            !ServiceCategoryID = Val(cmbServiceCategory.BoundText)
            !ServiceSubcategoryID = 0 'Val(cmbServiceSubcategory.BoundText)
            !AddedDate = Date
            !AddedTime = Now
            !AddedUserID = UserID
            !MaxNo = Val(txtMaxNo.Text)
            !StartTime = dtpStart.Value
            !DurationMinutes = Val(txtDuration.Text)
            .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim It As New clsFillCombos
    It.FillAnyCombo cmbServiceSecession, "ServiceSecession", True
    Dim Gp As New clsFillCombos
    Gp.FillAnyCombo cmbServiceCategory, "ServiceCategory", True
'    Dim SGp As New clsFillCombos
'    SGp.FillAnyCombo cmbServiceSubcategory, "ServiceSubcategory", True
End Sub

Private Sub DisplayDetails(): On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSecession where ServiceSecessionID = " & Val(cmbServiceSecession.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtServiceSecession.Text = !ServiceSecession
            txtComments.Text = !Comments
            cmbServiceCategory.BoundText = !ServiceCategoryID
'            cmbServiceSubcategory.BoundText = !ServiceSubcategoryID
            txtMaxNo.Text = !MaxNo
            dtpStart.Value = !StartTime
            If Not IsNull(!DurationMinutes) Then txtDuration.Text = !DurationMinutes
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub
