VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDiscardCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discard Categories"
   ClientHeight    =   5220
   ClientLeft      =   1860
   ClientTop       =   1380
   ClientWidth     =   10665
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
   Moveable        =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10665
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   3735
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   3960
      TabIndex        =   13
      Top             =   2280
      Width           =   6615
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   4680
         TabIndex        =   7
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Cancel"
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
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3735
      Begin MSDataListLib.DataCombo dtcCategoryName 
         Height          =   3780
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6668
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtComments 
         Height          =   735
         Left            =   1800
         TabIndex        =   4
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox txtCategoryName 
         Height          =   360
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Category Name"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Comments"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "C&lose"
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
Attribute VB_Name = "frmDiscardCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsCategory As New ADODB.Recordset
    Dim rsViewCategory As New ADODB.Recordset
    Dim TemCategoryID As Long

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub dtcCategoryName_Click(Area As Integer)
    If IsNumeric(dtcCategoryName.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    Call FillCategoryName
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub FillCategoryName()
    With rsViewCategory
        If .State = 1 Then .Close
        .Open "Select* From tblDiscardCategory Order By DiscardCategory", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Set dtcCategoryName.RowSource = rsViewCategory
            dtcCategoryName.BoundColumn = "DiscardCategoryID"
            dtcCategoryName.ListField = "DiscardCategory"
        End If
    End With
End Sub

Private Sub bttnAdd_Click()
    Call ClearValues
    Call AfterAdd
    txtCategoryName.SetFocus
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    Call AfterEdit
    txtCategoryName.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

Private Sub bttnChange_Click()
    Dim TemResponce As Integer
    If Trim(txtCategoryName.Text) = "" Then NoName: Exit Sub
    With rsCategory
    On Error Resume Next
        If .State = 1 Then .Close
        .Open "Select * From tblDiscardCategory Where (DiscardCategoryID = " & TemCategoryID & ")", cnnStores, 3, 3
        If .RecordCount > 0 Then
            !DiscardCategory = txtCategoryName.Text
            !Comments = txtComments.Text
            .Update
        End If
        FillCategoryName
        BeforeAddEdit
        ClearValues
        dtcCategoryName.Text = Empty
        dtcCategoryName.SetFocus
        If .State = 1 Then .Close
        Exit Sub
    
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Update Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
    End With
End Sub

Private Sub bttnSave_Click()
    Dim TemResponce As Integer
    If Trim(txtCategoryName.Text) = "" Then NoName: Exit Sub
    With rsCategory
    On Error Resume Next
        If .State = 1 Then .Close
        .Open "Select * From tblDiscardCategory", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !DiscardCategory = txtCategoryName.Text
        !Comments = txtComments.Text
        .Update
        FillCategoryName
        BeforeAddEdit
        ClearValues
        dtcCategoryName.Text = Empty
        dtcCategoryName.SetFocus
        If .State = 1 Then .Close
        Exit Sub
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        .CancelUpdate
        ClearValues
        BeforeAddEdit
        
        If .State = 1 Then .Close
        
End With


End Sub

Private Sub NoName()
    Dim TemResponce As Integer
    TemResponce = MsgBox("You must enter a Discard Category name to save", vbCritical, "No Name")
    txtCategoryName.SetFocus
End Sub

Private Sub AfterAdd()
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnSave.Enabled = True
    bttnCancel.Enabled = True
    bttnChange.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
End Sub

Private Sub AfterEdit()
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnSave.Enabled = False
    bttnCancel.Enabled = True
    bttnChange.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = False
End Sub

Private Sub BeforeAddEdit()
    bttnAdd.Visible = True
    bttnEdit.Visible = True
    bttnSave.Visible = False
    bttnCancel.Visible = False
    bttnChange.Visible = False
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    bttnSave.Enabled = False
    bttnCancel.Enabled = False
    Frame1.Enabled = False
    Frame2.Enabled = True
End Sub

Private Function ClearValues()
    txtCategoryName.Text = Empty
    txtComments.Text = Empty
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsCategory.State = 1 Then rsCategory.Close: Set rsCategory = Nothing
    If rsViewCategory.State = 1 Then rsViewCategory.Close: Set rsViewCategory = Nothing
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcCategoryName.BoundText) Then Exit Sub
    With rsCategory
        If .State = 1 Then .Close
        .Open "Select tblDiscardCategory.* From tblDiscardCategory Where (DiscardCategoryID = " & dtcCategoryName.BoundText & ")", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Call ClearValues
            If Not (!DiscardCategory) = "" Then txtCategoryName.Text = !DiscardCategory
            If Not (!Comments) = "" Then txtComments.Text = !Comments
            TemCategoryID = !DiscardCategoryID
        Else
            Call ClearValues
            txtCategoryName.Text = Empty
            txtComments.Text = Empty
            TemCategoryID = Empty
        End If
        If .State = 1 Then .Close
    End With
End Sub


