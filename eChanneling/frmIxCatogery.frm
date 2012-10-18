VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmIxCatogery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Investigation Catogeries"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
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
   ScaleHeight     =   4425
   ScaleWidth      =   9075
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtIxCatogery 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   4455
      End
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   2520
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
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   2520
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
         Left            =   3120
         TabIndex        =   7
         Top             =   2520
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
      Begin VB.Label Label1 
         Caption         =   "Investigation Catogery"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3015
      End
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   495
      Left            =   120
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
   Begin MSDataListLib.DataCombo dtcIxCatogery 
      Height          =   3540
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6244
      _Version        =   393216
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   495
      Left            =   1440
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
      Left            =   2760
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
   Begin btButtonEx.ButtonEx bttnExit 
      Height          =   495
      Left            =   7800
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Exit"
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
Attribute VB_Name = "frmIxCatogery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsIxCatogeryList As New ADODB.Recordset

Private Sub bttnAdd_Click()
    Call ClearValues
    Call AfterAdd
    txtIxCatogery.Text = dtcIxCatogery.Text
End Sub

Private Sub ClearValues()
    txtIxCatogery.Text = Empty
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
End Sub

Private Sub bttnChange_Click()
    Dim TR As Integer
    If Trim(txtIxCatogery) = "" Then
        TR = MsgBox("You have not entered a investigation catogery name", vbCritical, "Catogery?")
        txtIxCatogery.SetFocus
        Exit Sub
    End If
    Dim rsIxCatogeryChange As New ADODB.Recordset
    On Error GoTo EH
    With rsIxCatogeryChange
        If .State = 1 Then .Close
        .Open "Select * from tblixcatogery where (ixcatogery_ID = " & dtcIxCatogery.BoundText & ")", dbHospital, adOpenKeyset, adLockOptimistic
        If .RecordCount < 1 Then
            TR = MsgBox("This catogery does not exist", vbInformation, "Not Exist")
            txtIxCatogery.SetFocus
            SendKeys "{home}+{end}"
            .Close
            Set rsIxCatogeryChange = Nothing
            Exit Sub
        Else
            !ixcatogery = txtIxCatogery.Text
            .Update
            .Close
        End If
        Set rsIxCatogeryChange = Nothing
        Call FillIxCatogery
        Call BeforeAddEdit
        Call ClearValues
        Exit Sub
EH:
        TR = MsgBox("An error has occured, Please contact Lakmedipro" & vbNewLine & Err.Description, vbCritical, "Error")
        If .State = 1 Then .Close
        Set rsIxCatogeryChange = Nothing
    End With
    Exit Sub
End Sub

Private Sub bttnDelete_Click()
    Dim TR As Integer
    TR = MsgBox("Are you sure you want to delete " & dtcIxCatogery.Text & " ?", vbYesNo, "Delete?")
    If TR = vbNo Then Exit Sub
    TR = MsgBox("Deleting this record will lead to loss of the details of all investigations details, will prevent the functining of the program as well. Are you still sure you want to delete this record?", vbCritical + vbYesNo, "DELETE? Are you sure?")
    If TR = vbNo Then Exit Sub
    Dim rsIxCatogeryDelete As New ADODB.Recordset
    On Error GoTo EH
    With rsIxCatogeryDelete
        If .State = 1 Then .Close
        .Open "delete * from tblixcatogery where ixcatogery_ID = " & dtcIxCatogery.BoundText, dbHospital, adOpenStatic, adLockOptimistic
        If .State = 1 Then .Close
        Set rsIxCatogeryDelete = Nothing
    Call FillIxCatogery
    Call BeforeAddEdit
    Call ClearValues
    Exit Sub
EH:
    TR = MsgBox("An error has occured, Please contact Lakmedipro" & vbNewLine & Err.Description, vbCritical, "Error")
    If .State = 1 Then .Close
    Set rsIxCatogeryDelete = Nothing
    End With
    Exit Sub
End Sub

Private Sub bttnEdit_Click()
    Call AfterEdit
End Sub

Private Sub bttnExit_Click()
    Unload Me
End Sub

Private Sub bttnSave_Click()
    Dim TR As Integer
    If Trim(txtIxCatogery) = "" Then
        TR = MsgBox("You have not entered a investigation catogery name", vbCritical, "Catogery?")
        txtIxCatogery.SetFocus
        Exit Sub
    End If
    Dim rsIxCatogerySave As New ADODB.Recordset
    On Error GoTo EH
    With rsIxCatogerySave
        If .State = 1 Then .Close
        .Open "SELECT tblIxCatogery.* From tblIxCatogery where tblIxCatogery.IxCatogery = '" & Trim(txtIxCatogery.Text) & "'", dbHospital, adOpenStatic, adLockOptimistic
        If .RecordCount >= 1 Then
            TR = MsgBox("This catogery already exist", vbInformation, "Already Exist")
            txtIxCatogery.SetFocus
            SendKeys "{home}+{end}"
            .Close
            Set rsIxCatogerySave = Nothing
            Exit Sub
        Else
            If .State = 1 Then .Close
            .Open "Select * from tblixcatogery", dbHospital, adOpenDynamic, adLockOptimistic
            .AddNew
            !ixcatogery = txtIxCatogery.Text
            .Update
            .Close
        End If
        Set rsIxCatogerySave = Nothing
        Call FillIxCatogery
        Call BeforeAddEdit
        Call ClearValues
        Exit Sub
EH:
        TR = MsgBox("An error has occured, Please contact Lakmedipro" & vbNewLine & Err.Description, vbCritical, "Error")
        If .State = 1 Then .Close
        Set rsIxCatogerySave = Nothing
    End With
    Exit Sub
End Sub

Private Sub dtcIxCatogery_Change()
    If IsNumeric(dtcIxCatogery.BoundText) = True Then
        bttnAdd.Enabled = False
        bttnEdit.Enabled = True
        bttnDelete.Enabled = True
        Call DisplayDetails
    Else
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
        bttnDelete.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Call FillIxCatogery
End Sub

Private Sub DisplayDetails()
    If IsNumeric(dtcIxCatogery.BoundText) = False Then Exit Sub
    Dim rsIxCatogery As New ADODB.Recordset
    With rsIxCatogery
        If .State = 1 Then .Close
        .Open "Select * from tblixcatogery where ixcatogery_ID =" & dtcIxCatogery.BoundText, dbHospital, adOpenStatic, adLockReadOnly
        If .RecordCount >= 1 Then
            txtIxCatogery.Text = !ixcatogery
        End If
        .Close
    End With
End Sub

Private Sub FillIxCatogery()
    With rsIxCatogeryList
        If .State = 1 Then .Close
        .Open "SELECT tblIxCatogery.* From tblIxCatogery ORDER BY tblIxCatogery.IxCatogery", dbHospital, adOpenStatic, adLockReadOnly
        Set dtcIxCatogery.RowSource = rsIxCatogeryList
        dtcIxCatogery.ListField = "IxCatogery"
        dtcIxCatogery.BoundColumn = "IxCatogery_ID"
    End With
    dtcIxCatogery.Text = Empty
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CloseAll
End Sub

Private Sub CloseAll()
    If rsIxCatogeryList.State = 1 Then rsIxCatogeryList.Close
    Set rsIxCatogeryList = Nothing
End Sub

Private Sub BeforeAddEdit()
    Frame1.Enabled = False
    bttnAdd.Enabled = True
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    dtcIxCatogery.Enabled = True
End Sub

Private Sub AfterAdd()
    Frame1.Enabled = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    dtcIxCatogery.Enabled = False
    bttnSave.Visible = True
    bttnChange.Visible = False
End Sub

Private Sub AfterEdit()
    Frame1.Enabled = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    dtcIxCatogery.Enabled = False
    bttnSave.Visible = False
    bttnChange.Visible = True
End Sub
