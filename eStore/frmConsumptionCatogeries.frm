VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmConsumptionCatogeries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consumption Catogeries"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
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
   ScaleHeight     =   5160
   ScaleWidth      =   10965
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cl&ose"
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
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   3840
      TabIndex        =   12
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtcomment 
         Height          =   1575
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   5175
      End
      Begin VB.TextBox txtCatogery 
         Height          =   360
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label2 
         Caption         =   "Comments"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Catogery Name"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   3615
      Begin MSDataListLib.DataCombo dtcCatogeryName 
         Height          =   4020
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   7091
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   3840
      TabIndex        =   10
      Top             =   2760
      Width           =   6855
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   240
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
         Left            =   4920
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   4320
      Width           =   3615
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
         Width           =   1455
         _ExtentX        =   2566
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
End
Attribute VB_Name = "frmConsumptionCatogeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsCatogery As New ADODB.Recordset
    Dim rsViewCatogeryName As New ADODB.Recordset
    Dim FRows As Long
    Dim NowROw As Long
    Dim TemConsumptionCatogeryID As Long

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
    Call ClearValues
    dtcCatogeryName.Text = Empty
    dtcCatogeryName.SetFocus
End Sub

Private Sub dtcCatogeryName_Click(Area As Integer)
    If IsNumeric(dtcCatogeryName.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    FillCategoryCombo
    BeforeAddEdit
    ClearValues
End Sub

Private Sub bttnAdd_Click()
    ClearValues
    AfterAdd
    txtCatogery.SetFocus
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    AfterEdit
    txtCatogery.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

Private Sub bttnChange_Click()
    Dim TemResponce As Integer
    If txtCatogery.Text = "" Then NoName: Exit Sub
    With rsCatogery
    On Error Resume Next
        If .State = 1 Then .Close
        .Open "Select* From tblConsumptionCategory Where (ConsumptionCategoryID = " & TemConsumptionCatogeryID & ")", cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !ConsumptionCategory = Trim(txtCatogery.Text)
            !Comments = txtcomment.Text
            .Update
        End If
        If .State = 1 Then .Close
        FillCategoryCombo
        BeforeAddEdit
        ClearValues
        dtcCatogeryName.Text = Empty
        dtcCatogeryName.SetFocus
        Exit Sub

ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
    End With
End Sub

Private Sub bttnSave_Click()
    Dim TemResponce As Integer
    If Trim(txtCatogery.Text) = "" Then NoName: Exit Sub
    With rsCatogery
    On Error Resume Next
        If .State = 1 Then .Close
        .Open "Select* From tblConsumptionCategory", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ConsumptionCategory = Trim(txtCatogery.Text)
        !Comments = txtcomment.Text
        .Update
        If .State = 1 Then .Close
        FillCategoryCombo
        BeforeAddEdit
        ClearValues
        dtcCatogeryName.Text = Empty
        dtcCatogeryName.SetFocus
        Exit Sub
    
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
    End With
End Sub

Private Sub NoName()
    Dim TemResponce As Integer
    TemResponce = MsgBox("You have not entered a Generic Name to save", vbCritical, "No Name")
    txtCatogery.SetFocus
End Sub


Private Sub FillCategoryCombo()
    With rsViewCatogeryName
        If .State = 1 Then .Close
        .Open "Select* From tblConsumptionCategory Order By ConsumptionCategory", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Set dtcCatogeryName.RowSource = rsViewCatogeryName
            dtcCatogeryName.ListField = "ConsumptionCategory"
            dtcCatogeryName.BoundColumn = "ConsumptionCategoryID"
        End If
    End With
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
    Frame2.Enabled = False
    Frame1.Enabled = True
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
    Frame2.Enabled = False
    Frame1.Enabled = True
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
    Frame2.Enabled = True
    Frame1.Enabled = False
End Sub

Private Sub ClearValues()
    txtCatogery.Text = Empty
    txtcomment.Text = Empty
    TemConsumptionCatogeryID = Empty
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsCatogery.State = 1 Then rsCatogery.Close: Set rsCatogery = Nothing
    If rsViewCatogeryName.State = 1 Then rsViewCatogeryName.Close: Set rsViewCatogeryName = Nothing
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcCatogeryName.BoundText) Then Exit Sub
    With rsCatogery
        If .State = 1 Then .Close
        .Open "Select* From tblConsumptionCategory Where (ConsumptionCategoryID = " & dtcCatogeryName.BoundText & ")", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Call ClearValues
            txtCatogery.Text = !ConsumptionCategory
            If Not IsNull(!Comments) Then txtcomment.Text = !Comments
            TemConsumptionCatogeryID = !ConsumptionCategoryID
        End If
        If .State = 1 Then .Close
    End With
End Sub
