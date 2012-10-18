VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmGenericName 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generic Names"
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
      Begin MSDataListLib.DataCombo dtcGenaricName 
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
      Begin VB.TextBox txtGName 
         Height          =   360
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox txtcomment 
         Height          =   1695
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Generic Name"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Comments"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
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
Attribute VB_Name = "frmGenericName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Gname As New ADODB.Recordset
Dim ViewGenaricName As New ADODB.Recordset

Dim FRows As Long
Dim NowRow As Long
Dim TemGenericNameID As Long


Private Sub bttnCancel_Click()
Call BeforeAddEdit
Call ClearValues
End Sub

Private Sub dtcGenaricName_Click(Area As Integer)
If IsNumeric(dtcGenaricName.BoundText) = False Then Exit Sub
Call DisplaySelected
End Sub

Private Sub Form_Load()
FillGenaricCombo
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
If txtGName.Text = "" Then NoName: Exit Sub

With Gname

On Error GoTo ErrorHandler

    If .State = 1 Then .Close
    .Open "Select tblGenericName.* From tblGenericName Where (GenericNameID = " & TemGenericNameID & ")", cnnStores, adOpenStatic, adLockOptimistic
    
    If .RecordCount = 0 Then Exit Sub

    !GenericName = Trim(txtGName.Text)
    !Comments = txtcomment.Text
    .Update
    
    If .State = 1 Then .Close
    
    
    FillGenaricCombo
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
If Trim(txtGName.Text) = "" Then NoName: Exit Sub

With Gname

On Error GoTo ErrorHandler
    If .State = 1 Then .Close
    .Open "Select tblGenericName.* From tblGenericName", cnnStores, adOpenStatic, adLockOptimistic
  
    
    .AddNew
    !GenericName = Trim(txtGName.Text)
    !Comments = txtcomment.Text
    .Update
    
    If .State = 1 Then .Close

    FillGenaricCombo
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
txtGName.SetFocus
End Sub


Private Sub FillGenaricCombo()

With ViewGenaricName

    If .State = 1 Then .Close
    .Open "Select tblGenericName.* From tblGenericName Order By GenericName", cnnStores, adOpenStatic, adLockReadOnly

    If .RecordCount = 0 Then Exit Sub
    Set dtcGenaricName.RowSource = ViewGenaricName
    dtcGenaricName.ListField = "GenericName"
    dtcGenaricName.BoundColumn = "GENERICNAMEID"
            
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
    txtGName.Text = Empty
    txtcomment.Text = Empty
    TemGenericNameID = Empty
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Gname.State = 1 Then Gname.Close: Set Gname = Nothing
If ViewGenaricName.State = 1 Then ViewGenaricName.Close: Set ViewGenaricName = Nothing
End Sub

Private Sub DisplaySelected()
If Not IsNumeric(dtcGenaricName.BoundText) Then Exit Sub

With Gname

    If .State = 1 Then .Close
    .Open "Select tblGenericName.* From tblGenericName Where (GenericNameID = " & dtcGenaricName.BoundText & ")", cnnStores, adOpenStatic, adLockReadOnly

    If .RecordCount = 0 Then Exit Sub
    Call ClearValues
    
    txtGName.Text = !GenericName
    If Not IsNull(!Comments) Then txtcomment.Text = !Comments
    TemGenericNameID = !GenericNameID
    
    If .State = 1 Then .Close
End With

End Sub

