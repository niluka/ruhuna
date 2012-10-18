VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDepartment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department Names"
   ClientHeight    =   5235
   ClientLeft      =   2130
   ClientTop       =   1635
   ClientWidth     =   10920
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
   ScaleHeight     =   5235
   ScaleWidth      =   10920
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   3615
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         TabIndex        =   14
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
      Height          =   855
      Left            =   3960
      TabIndex        =   8
      Top             =   3360
      Width           =   6855
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   360
         TabIndex        =   9
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
         Left            =   4800
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Change"
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
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin MSDataListLib.DataCombo dtcDepartment 
         Height          =   3780
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   6668
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtDepartment 
         Height          =   360
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtcomment 
         Height          =   1695
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Department Name"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Comments"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   4440
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
Attribute VB_Name = "frmDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDepartment As New ADODB.Recordset
Dim rsViewDepartment As New ADODB.Recordset

Dim FRows As Long
Dim NowRow As Long
Dim TemDepartmentID As Long


Private Sub bttnCancel_Click()
Call BeforeAddEdit
Call ClearValues
End Sub

Private Sub dtcDepartment_Click(Area As Integer)
If IsNumeric(dtcDepartment.BoundText) = False Then Exit Sub
Call DisplaySelected
End Sub

Private Sub Form_Load()
FillDepartmentCmbo
BeforeAddEdit
ClearValues
End Sub

Private Sub bttnAdd_Click()
ClearValues
AfterAdd
End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnEdit_Click()
AfterEdit
End Sub

Private Sub bttnChange_Click()
Dim TemResponce As Integer
If txtDepartment.Text = "" Then NoName: Exit Sub

With rsDepartment

On Error GoTo ErrorHandler

    If .State = 1 Then .Close
    .Open "Select tblStore.* From tblStore Where (StoreID = " & TemDepartmentID & ")", cnnStores, adOpenStatic, adLockOptimistic
    
    If .RecordCount = 0 Then Exit Sub

    !Store = Trim(txtDepartment.Text)
    !StoreDescreption = txtcomment.Text
    .Update
    
    If .State = 1 Then .Close
    
    
    FillDepartmentCmbo
    BeforeAddEdit
    ClearValues
    
    Exit Sub

ErrorHandler:
    TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
    .CancelUpdate
    ClearValues
    BeforeAddEdit
    
    If .State = 1 Then .Close

End With

End Sub

Private Sub bttnSave_Click()
Dim TemResponce As Integer
If Trim(txtDepartment.Text) = "" Then NoName: Exit Sub

With rsDepartment

On Error GoTo ErrorHandler
    If .State = 1 Then .Close
    .Open "Select tblStore.* From tblStore", cnnStores, adOpenStatic, adLockOptimistic
  
    
    .AddNew
    !Store = Trim(txtDepartment.Text)
    !StoreDescreption = txtcomment.Text
    .Update
    
    If .State = 1 Then .Close

    FillDepartmentCmbo
    BeforeAddEdit
    ClearValues
    
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
TemResponce = MsgBox("You have not entered a Generic Name to save", vbCritical, "No Name")
txtDepartment.SetFocus
End Sub


Private Sub FillDepartmentCmbo()

With rsViewDepartment
    If .State = 1 Then .Close
    .Open "Select tblStore.* From tblStore Order By Store", cnnStores, adOpenStatic, adLockReadOnly

    If .RecordCount = 0 Then Exit Sub
    Set dtcDepartment.RowSource = rsViewDepartment
    dtcDepartment.ListField = "Store"
    dtcDepartment.BoundColumn = "StoreID"
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
txtDepartment.Text = Empty
txtcomment.Text = Empty
TemDepartmentID = Empty
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If rsDepartment.State = 1 Then rsDepartment.Close: Set rsDepartment = Nothing
If rsViewDepartment.State = 1 Then rsViewDepartment.Close: Set rsViewDepartment = Nothing
End Sub

Private Sub DisplaySelected()
If Not IsNumeric(dtcDepartment.BoundText) Then Exit Sub

With rsDepartment

    If .State = 1 Then .Close
    .Open "Select tblStore.* From tblStore Where (StoreID = " & dtcDepartment.BoundText & ")", cnnStores, adOpenStatic, adLockReadOnly

    If .RecordCount = 0 Then Exit Sub
    Call ClearValues
    
    txtDepartment.Text = !Store
    If Not IsNull(!StoreDescreption) Then txtcomment.Text = !StoreDescreption
    TemDepartmentID = !StoreID
    
    If .State = 1 Then .Close
End With

End Sub

