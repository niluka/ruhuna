VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAddNewIxName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Investigation Names"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
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
   ScaleHeight     =   4440
   ScaleWidth      =   11055
   Begin VB.Frame FrameIx 
      Height          =   3375
      Left            =   4320
      TabIndex        =   11
      Top             =   240
      Width           =   6495
      Begin VB.TextBox txtIx 
         Height          =   375
         Left            =   2040
         MaxLength       =   249
         TabIndex        =   4
         Top             =   600
         Width           =   4215
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   495
         Left            =   2400
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   495
         Left            =   2400
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
         Height          =   495
         Left            =   4920
         TabIndex        =   9
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
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
      Begin MSDataListLib.DataCombo dtcIxCatogery 
         Height          =   360
         Left            =   2040
         TabIndex        =   5
         Top             =   1200
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo dtcIxSubcatogery 
         Height          =   360
         Left            =   2040
         TabIndex        =   6
         Top             =   1800
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label3 
         Caption         =   "Subcatogery"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Catogery"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Investigation"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin MSDataListLib.DataCombo dtcIx 
      Height          =   3300
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5821
      _Version        =   393216
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Delete"
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9600
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
Attribute VB_Name = "frmAddNewIxName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsIxList As New ADODB.Recordset
    Dim rsIxCatogeryList As New ADODB.Recordset
    Dim rsIxSubcatogeryList As New ADODB.Recordset
    

Private Sub bttnAdd_Click()
    Call ClearAll
    Call AfterAdd
    txtIx.Text = dtcIx.Text
    txtIx.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub bttnCancel_Click()
    Call ClearAll
    Call BeforeAddEdit
End Sub

Private Sub bttnChange_Click()
    Dim TR As Integer
    If Trim(txtIx.Text) = "" Then
        TR = MsgBox("You have not entered a name for the Investigation to add", vbCritical, "No Ix Name")
        txtIx.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(dtcIxCatogery.BoundText) Then
        TR = MsgBox("You have not selected the catogery", vbCritical, "No Catogery")
        dtcIxCatogery.SetFocus
        Exit Sub
    End If
    Dim rsIx As New ADODB.Recordset
    With rsIx
        If .State = 1 Then .Close
        .Open "SELECT * from tblIx where IX_ID = " & dtcIx.BoundText, dbHospital, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !ix = Trim(txtIx.Text)
            !ixcatogery_ID = dtcIxCatogery.BoundText
            If IsNumeric(dtcIxSubcatogery.BoundText) Then !ixsubcatogery_ID = dtcIxSubcatogery.BoundText
            .Update
            .Close
        End If
        Set rsIx = Nothing
        Call ClearAll
        Call BeforeAddEdit
        Call FillIxCombo
    Exit Sub
EH:
        TR = MsgBox("An Error Occured" & vbNewLine & Err.Description, vbCritical, "Error")
        .Close
        Set rsIx = Nothing
        Call ClearAll
        Call BeforeAddEdit
        Exit Sub
    End With
End Sub

Private Sub bttnDelete_Click()
    If Not IsNumeric(dtcIx.BoundText) Then Exit Sub
    Dim TR As Integer
    TR = MsgBox("Are you sure you want to DELETE " & dtcIx.Text & "?", vbCritical + vbYesNo, "DELETE?")
    If TR = vbNo Then Exit Sub
    TR = MsgBox("This Delete will lead to loss of Investigation results causing malfunctioning of the program. Are you sure you want to prodeed with the delete?", vbCritical + vbYesNo, "Are you sure?")
    If vbNo Then Exit Sub
    On Error GoTo EH
    Dim rsIxDelete As New ADODB.Recordset
    rsIxDelete.Open "Delete from tblIx where ix_ID = " & dtcIx.BoundText
    If rsIxDelete.State = 1 Then rsIxDelete.Close
    Set rsIxDelete = Nothing
    Call ClearAll
    Call FillIxCombo
    Exit Sub
EH:
    TR = MsgBox("An unknown Error Occured", vbCritical, "NOT deleted")
    If rsIxDelete.State = 1 Then rsIxDelete.Close
    Set rsIxDelete = Nothing
    Exit Sub
End Sub

Private Sub bttnEdit_Click()
    Call AfterEdit
    txtIx.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub bttnSave_Click()
    Dim TR As Integer
    If Trim(txtIx.Text) = "" Then
        TR = MsgBox("You have not entered a name for the Investigation to add", vbCritical, "No Ix Name")
        txtIx.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(dtcIxCatogery.BoundText) Then
        TR = MsgBox("You have not selected the catogery", vbCritical, "No Catogery")
        dtcIxCatogery.SetFocus
        Exit Sub
    End If
    Dim rsIx As New ADODB.Recordset
    With rsIx
        If .State = 1 Then .Close
        .Open "SELECT * from tblIx where IX = '" & Trim(txtIx.Text) & "'", dbHospital, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            TR = MsgBox(txtIx.Text & " already excist in the database.", vbInformation, "Duplicate")
            txtIx.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        .Close
        .Open "SELECT * from tblix", dbHospital, adOpenDynamic, adLockOptimistic
        .AddNew
        !ix = Trim(txtIx.Text)
        !ixcatogery_ID = dtcIxCatogery.BoundText
        If IsNumeric(dtcIxSubcatogery.BoundText) Then !ixsubcatogery_ID = dtcIxSubcatogery.BoundText
        .Update
        .Close
        Set rsIx = Nothing
        Call ClearAll
        Call BeforeAddEdit
        Call FillIxCombo
    Exit Sub
EH:
        TR = MsgBox("An Error Occured" & vbNewLine & Err.Description, vbCritical, "Error")
        .Close
        Set rsIx = Nothing
        Call ClearAll
        Call BeforeAddEdit
        Exit Sub
    End With
End Sub

Private Sub dtcIx_Change()
    If IsNumeric(dtcIx.BoundText) Then
        bttnAdd.Enabled = False
        bttnEdit.Enabled = True
        bttnDelete.Enabled = True
        Call GetIxDetails
    Else
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
        bttnDelete.Enabled = False
        Call ClearAll
    End If
End Sub

Private Sub GetIxDetails()
    If Not IsNumeric(dtcIx.BoundText) Then Exit Sub
    On Error GoTo EH
    Dim rsGetIxDetails As New ADODB.Recordset
    With rsGetIxDetails
        If .State = 1 Then .Close
        .Open "SELECT * from tblix where Ix_ID = " & dtcIx.BoundText, dbHospital, adOpenStatic, adLockPessimistic
        txtIx.Text = !ix
        dtcIxCatogery.BoundText = !ixcatogery_ID
        dtcIxSubcatogery.BoundText = !ixsubcatogery_ID
        .Close
    End With
    Set rsGetIxDetails = Nothing
    Exit Sub
EH:
    Dim TR As Integer
    TR = MsgBox("An ERROR occured", vbCritical, "ERROR")
    rsGetIxDetails.Close
    Set rsGetIxDetails = Nothing
    Exit Sub
End Sub

Private Sub dtcIxCatogery_Click(Area As Integer)
    If IsNumeric(dtcIxCatogery.BoundText) Then
        Call FillIxSubcatogeryCombo
    End If
End Sub

Private Sub Form_Load()
    Call FillIxCombo
    Call FillIxCatogeryCombo
    Call FillIxSubcatogeryCombo
End Sub

Private Sub FillIxCombo()
    With rsIxList
        If .State = 1 Then .Close
        .Open "select * from tblix order by ix", dbHospital, adOpenStatic, adLockReadOnly
        Set dtcIx.RowSource = rsIxList
        dtcIx.ListField = "Ix"
        dtcIx.BoundColumn = "Ix_ID"
    End With
End Sub

Private Sub FillIxCatogeryCombo()
    With rsIxCatogeryList
        If .State = 1 Then .Close
        .Open "select * from tblixcatogery order by ixcatogery", dbHospital, adOpenStatic, adLockReadOnly
        Set dtcIxCatogery.RowSource = rsIxCatogeryList
        dtcIxCatogery.ListField = "IxCatogery"
        dtcIxCatogery.BoundColumn = "Ixcatogery_ID"
    End With
End Sub

Private Sub FillIxSubcatogeryCombo()
    If Not IsNumeric(dtcIxCatogery.BoundText) Then
        dtcIxSubcatogery.ListField = Empty
        dtcIxSubcatogery.BoundColumn = Empty
        Exit Sub
    End If
    With rsIxSubcatogeryList
        If .State = 1 Then .Close
        .Open "select * from tblixsubcatogery where ixcatogery_ID = " & dtcIxCatogery.BoundText & " order by ixsubcatogery", dbHospital, adOpenStatic, adLockReadOnly
        Set dtcIxSubcatogery.RowSource = rsIxSubcatogeryList
        dtcIxSubcatogery.ListField = "Ixsubcatogery"
        dtcIxSubcatogery.BoundColumn = "Ixsubcatogery_ID"
    End With
End Sub

Private Sub CloseAll()
    If rsIxList.State = 1 Then rsIxList.Close
    If rsIxCatogeryList.State = 1 Then rsIxCatogeryList.Close
    If rsIxSubcatogeryList.State = 1 Then rsIxSubcatogeryList.Close
    Set rsIxList = Nothing
    Set rsIxCatogeryList = Nothing
    Set rsIxSubcatogeryList = Nothing
End Sub

Private Sub BeforeAddEdit()
    bttnAdd.Enabled = True
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    dtcIx.Enabled = True
    FrameIx.Enabled = False
End Sub

Private Sub AfterAdd()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    dtcIx.Enabled = False
    FrameIx.Enabled = True
    bttnSave.Visible = True
    bttnCancel.Visible = False
End Sub

Private Sub AfterEdit()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    dtcIx.Enabled = False
    FrameIx.Enabled = True
    bttnSave.Visible = True
    bttnCancel.Visible = False
End Sub

Private Sub ClearAll()
    txtIx.Text = Empty
    dtcIxCatogery.Text = Empty
    dtcIxSubcatogery.Text = Empty
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseAll
End Sub
