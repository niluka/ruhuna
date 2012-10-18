VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmRoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rooms"
   ClientHeight    =   5175
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
   ScaleHeight     =   5175
   ScaleWidth      =   10665
   Begin VB.TextBox txtCategoryName 
      BackColor       =   &H00C0FFC0&
      Height          =   360
      Left            =   5640
      TabIndex        =   5
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox txtComments 
      BackColor       =   &H00C0FFC0&
      Height          =   2055
      Left            =   5640
      TabIndex        =   9
      Top             =   1440
      Width           =   4575
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   8640
      TabIndex        =   12
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   12648384
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
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   12648384
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
      Left            =   2040
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   12648384
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
   Begin MSDataListLib.DataCombo dtcRoom 
      Height          =   3780
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6668
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      BackColor       =   12648384
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbRoomCategory 
      Height          =   360
      Left            =   5640
      TabIndex        =   7
      Top             =   960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      BackColor       =   12648384
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   12648384
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
      Left            =   8040
      TabIndex        =   11
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   12648384
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
      Left            =   6120
      TabIndex        =   10
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   12648384
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
   Begin VB.Label Label3 
      Caption         =   "&Room Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "&Room Name"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Room &Category"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmRoom"
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

Private Sub dtcRoom_Click(Area As Integer)
    If IsNumeric(dtcRoom.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub FillCombos()
    With rsViewCategory
        If .State = 1 Then .Close
        .Open "Select* From tblRoom Order By Room", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Set dtcRoom.RowSource = rsViewCategory
            dtcRoom.BoundColumn = "RoomID"
            dtcRoom.ListField = "Room"
        End If
    End With
    Dim RC As New clsFillCombos
    RC.FillAnyCombo cmbRoomCategory, "RoomCategory", True
End Sub

Private Sub bttnAdd_Click()
    Call ClearValues
    Call AfterAdd
    txtCategoryName.SetFocus
    txtCategoryName.Text = dtcRoom.Text
    SendKeys "{home}+{end}"
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
    On Error Resume Next

    Dim TemResponce As Integer
    If Trim(txtCategoryName.Text) = "" Then NoName: Exit Sub
    With rsCategory
    On Error GoTo ErrorHandler
        If .State = 1 Then .Close
        .Open "Select * From tblRoom Where (RoomID = " & TemCategoryID & ")", cnnStores, 3, 3
        If .RecordCount > 0 Then
            !Room = txtCategoryName.Text
            !RoomCategoryID = Val(cmbRoomCategory.BoundText)
            !Comments = txtComments.Text
            .Update
        End If
        FillCombos
        BeforeAddEdit
        ClearValues
        dtcRoom.Text = Empty
        dtcRoom.SetFocus
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
    On Error Resume Next
    Dim TemResponce As Integer
    If Trim(txtCategoryName.Text) = "" Then NoName: Exit Sub
    With rsCategory
    On Error GoTo ErrorHandler
        If .State = 1 Then .Close
        .Open "Select * From tblRoom", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Room = txtCategoryName.Text
        !RoomCategoryID = Val(cmbRoomCategory.BoundText)
        !Comments = txtComments.Text
        .Update
        FillCombos
        BeforeAddEdit
        ClearValues
        dtcRoom.Text = Empty
        dtcRoom.SetFocus
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
    cmbRoomCategory.Enabled = True
    dtcRoom.Enabled = False
End Sub

Private Sub AfterEdit()
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnSave.Enabled = False
    bttnCancel.Enabled = True
    cmbRoomCategory.Enabled = True
    bttnChange.Enabled = True
    dtcRoom.Enabled = False
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
    cmbRoomCategory.Enabled = False
    dtcRoom.Enabled = True
End Sub

Private Function ClearValues()
    txtCategoryName.Text = Empty
    txtComments.Text = Empty
    cmbRoomCategory.Text = Empty
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsCategory.State = 1 Then rsCategory.Close: Set rsCategory = Nothing
    If rsViewCategory.State = 1 Then rsViewCategory.Close: Set rsViewCategory = Nothing
End Sub

Private Sub DisplaySelected()
    If Not IsNumeric(dtcRoom.BoundText) Then Exit Sub
    With rsCategory
        If .State = 1 Then .Close
        .Open "Select tblRoom.* From tblRoom Where (RoomID = " & dtcRoom.BoundText & ")", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Call ClearValues
            If Not (!Room) = "" Then txtCategoryName.Text = !Room
            If Not (!Comments) = "" Then txtComments.Text = !Comments
            If Not IsNull(!RoomCategoryID) Then cmbRoomCategory.BoundText = !RoomCategoryID
            TemCategoryID = !RoomID
        Else
            Call ClearValues
            txtCategoryName.Text = Empty
            txtComments.Text = Empty
            cmbRoomCategory.Text = Empty
            TemCategoryID = Empty
        End If
        If .State = 1 Then .Close
    End With
End Sub


