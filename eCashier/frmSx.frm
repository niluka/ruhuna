VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Surgery"
   ClientHeight    =   6930
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
   ScaleHeight     =   6930
   ScaleWidth      =   11760
   Begin VB.TextBox txtComments 
      Height          =   5040
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   4575
   End
   Begin VB.TextBox txtSx 
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
   Begin MSDataListLib.DataCombo cmbSx 
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
   Begin VB.Label Label11 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Surgery"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Surgery"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmSx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbSx.BoundText) = False Then
        temText = cmbSx.Text
    Else
        temText = Empty
    End If
    cmbSx.Text = Empty
    Call EditMode
    txtSx.Text = temText
    txtSx.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbSx.Text = Empty
    cmbSx.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbSx.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblSx where SxID = " & Val(cmbSx.BoundText)
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
    cmbSx.SetFocus
    cmbSx.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbSx.BoundText) = False Then Exit Sub
    Call EditMode
    txtSx.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtSx.Text) = Empty Then
        MsgBox "You have not entered an Sx"
        txtSx.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSx.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbSx.Text = Empty
    cmbSx.SetFocus
End Sub

Private Sub cmbSx_Change()
    Call ClearValues
    If IsNumeric(cmbSx.BoundText) = True Then Call DisplayDetails
End Sub


Private Sub cmbSx_Click(Area As Integer)
    Call ClearValues
    If IsNumeric(cmbSx.BoundText) = True Then Call DisplayDetails
End Sub

Private Sub Form_Load()
    Me.Top = GetSetting(App.EXEName, Me.Name, "Top", Me.Top)
    Me.Left = GetSetting(App.EXEName, Me.Name, "Left", Me.Left)
    Call SelectMode
    Call FillCombos
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbSx.Enabled = False
    
    txtSx.Enabled = True
    txtComments.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbSx.Enabled = True
    
    txtSx.Enabled = False
    txtComments.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtSx.Text = Empty
    txtComments.Text = Empty
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblSx"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Sx = txtSx.Text
        !Comments = txtComments.Text
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblSx where SxID = " & Val(cmbSx.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        !Sx = txtSx.Text
        !Comments = txtComments.Text
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillAnyCombo cmbSx, "Sx", True
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblSx where SxID = " & Val(cmbSx.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtSx.Text = !Sx
            txtComments.Text = Format(!Comments, "")
            
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "Top", Me.Top
    SaveSetting App.EXEName, Me.Name, "Left", Me.Left
End Sub
