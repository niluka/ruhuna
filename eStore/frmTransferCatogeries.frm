VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmTransferCatogeries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Catogeries"
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   3480
      Width           =   6615
      Begin btButtonEx.ButtonEx bttnSave 
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
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   4680
         TabIndex        =   8
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
         TabIndex        =   7
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
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3735
      Begin MSDataListLib.DataCombo dtcTransaction 
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
      Height          =   3375
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtTransferName 
         Height          =   360
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   4575
      End
      Begin MSDataListLib.DataCombo dtcIssueStores 
         Height          =   360
         Left            =   1800
         TabIndex        =   4
         Top             =   1200
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo dtcReceiveStores 
         Height          =   360
         Left            =   1800
         TabIndex        =   5
         Top             =   1680
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "&Receive Stores"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "&Transfer Name"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "&Issue Stores"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1695
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   8640
      TabIndex        =   9
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
Attribute VB_Name = "frmTransferCatogeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTransaction As New ADODB.Recordset
    Dim rsViewTransaction As New ADODB.Recordset
    Dim rsViewStores As New ADODB.Recordset
    Dim TemTransactionID As Long

Private Sub bttnCancel_Click()
    Call BeforeAddEdit
    Call ClearValues
    dtcTransaction.Text = Empty
    dtcTransaction.SetFocus
End Sub

Private Sub dtctransaction_Click(Area As Integer)
    If IsNumeric(dtcTransaction.BoundText) = False Then Exit Sub
    Call DisplaySelected
End Sub

Private Sub Form_Load()
    Call FillTrancationName
    Call FillStores
    Call BeforeAddEdit
    Call ClearValues
End Sub

Private Sub FillTrancationName()
    With rsViewTransaction
        If .State = 1 Then .Close
        .Open "Select* From tblTransferCategory Order By TransferCategory", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Set dtcTransaction.RowSource = rsViewTransaction
            dtcTransaction.BoundColumn = "TransferCategoryID"
            dtcTransaction.ListField = "TransferCategory"
        End If
    End With
End Sub

Private Sub FillStores()
    With rsViewStores
        If .State = 1 Then .Close
        .Open "Select* From tblStore Order By Store", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Set dtcIssueStores.RowSource = rsViewStores
            dtcIssueStores.BoundColumn = "StoreID"
            dtcIssueStores.ListField = "Store"
            Set dtcReceiveStores.RowSource = rsViewStores
            dtcReceiveStores.BoundColumn = "StoreID"
            dtcReceiveStores.ListField = "Store"
        End If
    End With
End Sub


Private Sub bttnAdd_Click()
    Call ClearValues
    Call AfterAdd
    txtTransferName.SetFocus
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    Call AfterEdit
    txtTransferName.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

Private Sub bttnChange_Click()
    If CheckEmptys = False Then Exit Sub
    Call Editdata
End Sub

Private Function CheckEmptys() As Boolean
    Dim TemResponce As Integer
    CheckEmptys = False
    If Trim(txtTransferName.Text) = "" Then
        TemResponce = MsgBox("You must enter a Transaction name to save", vbCritical, "No Name")
        txtTransferName.SetFocus
        Exit Function
    End If
    If (dtcIssueStores.BoundText) = Empty Then
        TemResponce = MsgBox("You must select a issue stores to save", vbCritical, "No Issue Store")
        dtcIssueStores.SetFocus
        Exit Function
    End If
    
    If (dtcReceiveStores.BoundText) = Empty Then
        TemResponce = MsgBox("You must select a Receive stores to save", vbCritical, "No Receive Store")
        dtcReceiveStores.SetFocus
        Exit Function
    End If
    CheckEmptys = True
End Function


Private Sub Editdata()
    Dim TemResponce As Integer
    With rsTransaction
    On Error Resume Next
        If .State = 1 Then .Close
        .Open "Select tblTransferCategory.* From tblTransferCategory Where (TransferCategoryID = " & Val(TemTransactionID) & ")", cnnStores, 3, 3
        If .RecordCount > 0 Then
            !TransferCategory = Trim(txtTransferName.Text)
            !IStoreID = Val(dtcIssueStores.BoundText)
            !RStoreID = Val(dtcReceiveStores.BoundText)
            .Update
        End If
        FillTrancationName
        BeforeAddEdit
        ClearValues
        If .State = 1 Then .Close
        dtcTransaction.Text = Empty
        dtcTransaction.SetFocus
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
    If CheckEmptys = False Then Exit Sub
    Call SaveData
End Sub

Private Sub SaveData()
    Dim TemResponce As Integer
    With rsTransaction
    On Error Resume Next
        If .State = 1 Then .Close
        .Open "Select* From tblTransferCategory", cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !TransferCategory = Trim(txtTransferName.Text)
        !IStoreID = Val(dtcIssueStores.BoundText)
        !RStoreID = Val(dtcReceiveStores.BoundText)
        .Update
        FillTrancationName
        BeforeAddEdit
        ClearValues
        dtcTransaction.Text = Empty
        dtcTransaction.SetFocus
        If .State = 1 Then .Close
        Exit Sub
    
ErrorHandler:
        TemResponce = MsgBox(Err.Number & vbNewLine & Err.Description & Me.Caption, vbCritical + vbOKOnly, "Save Error")
        If .State = 1 Then .CancelUpdate
        ClearValues
        BeforeAddEdit
        If .State = 1 Then .Close
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

Private Sub ClearValues()
    txtTransferName.Text = Empty
    dtcIssueStores.BoundText = Empty
    dtcReceiveStores.BoundText = Empty
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsTransaction.State = 1 Then rsTransaction.Close: Set rsTransaction = Nothing
    If rsViewTransaction.State = 1 Then rsViewTransaction.Close: Set rsViewTransaction = Nothing
    If rsViewStores.State = 1 Then rsViewStores.Close: Set rsViewStores = Nothing
End Sub


Private Sub DisplaySelected()
    If Not IsNumeric(dtcTransaction.BoundText) Then Exit Sub
    With rsTransaction
        If .State = 1 Then .Close
        .Open "Select tblTransferCategory.* From tblTransferCategory Where (TransferCategoryID = " & dtcTransaction.BoundText & ")", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Call ClearValues
            If Not (!TransferCategory) = "" Then txtTransferName.Text = !TransferCategory
            If Not (!IStoreID) = "" Then dtcIssueStores.BoundText = !IStoreID
            If Not (!RStoreID) = "" Then dtcReceiveStores.BoundText = !RStoreID
            TemTransactionID = !TransferCategoryID
        Else
            Call ClearValues
            txtTransferName.Text = Empty
            dtcIssueStores.BoundText = Empty
            dtcReceiveStores.BoundText = Empty
            TemTransactionID = 0
        End If
        If .State = 1 Then .Close
    End With
End Sub
