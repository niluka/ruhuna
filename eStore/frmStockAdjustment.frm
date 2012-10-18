VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmStockAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Adjustment"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
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
   ScaleHeight     =   6975
   ScaleWidth      =   11700
   Begin VB.TextBox txtStoreID 
      Height          =   360
      Left            =   5640
      TabIndex        =   31
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtBatchID 
      Height          =   360
      Left            =   4320
      TabIndex        =   30
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frameAdjust 
      Caption         =   "Adjust"
      Height          =   5655
      Left            =   6720
      TabIndex        =   7
      Top             =   0
      Width           =   4815
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   4920
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.TextBox txtAdjustment 
         BackColor       =   &H8000000F&
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox txtAStock 
         Height          =   360
         Left            =   1680
         TabIndex        =   15
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtCStock 
         BackColor       =   &H8000000F&
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   3135
      End
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   5160
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSDataListLib.DataCombo dtcCheckedStaff 
         Height          =   360
         Left            =   1440
         TabIndex        =   27
         Top             =   2280
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcIssueStaff 
         Height          =   360
         Left            =   1440
         TabIndex        =   28
         Top             =   1800
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Locked          =   -1  'True
         BackColor       =   -2147483633
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcCategoryName 
         Height          =   360
         Left            =   1440
         TabIndex        =   29
         Top             =   2760
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcDepartment 
         Height          =   360
         Left            =   1440
         TabIndex        =   32
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   5160
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSDataListLib.DataCombo dtcBatch 
         Height          =   360
         Left            =   1440
         TabIndex        =   36
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label9 
         Caption         =   "Reason"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Checked by"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "User"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblIUnit2 
         Caption         =   "Unit"
         Height          =   255
         Left            =   3000
         TabIndex        =   20
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Adjustment"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label lblIUnit1 
         Caption         =   "Unit"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblIUnit 
         Caption         =   "Unit"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Department"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Batch"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Item"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Actual Stock"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Computer Stock"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   1935
      End
   End
   Begin MSDataListLib.DataCombo dtcCatogery 
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItem 
      Height          =   360
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCode 
      Height          =   360
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid GridTotal 
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7435
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10200
      TabIndex        =   23
      Top             =   5760
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
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
      Left            =   1440
      TabIndex        =   34
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label Label35 
      Caption         =   "Code"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "Item"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label28 
      Caption         =   "Catogery"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmStockAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim NewItem As New Item
    
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCategory As New ADODB.Recordset
    Dim rsCode As New ADODB.Recordset
    Dim rsAdjustmentCategory As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsDepartment As New ADODB.Recordset
    Dim rsBatch As New ADODB.Recordset
    
    Dim rsTemBatchStock As New ADODB.Recordset
    Dim rsTemBatch As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    Dim rsAdjustment As New ADODB.Recordset

Private Sub bttnAdd_Click()
    Call ClearAddValues
    Call AfterAdd
    dtcDepartment.SetFocus
End Sub



Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
    dtcCatogery.SetFocus
End Sub

Private Sub bttnChange_Click()
    If CanAdd = False Then Exit Sub
    Call AddData
    Call BeforeAddEdit
    Call ClearValues
    dtcCatogery.SetFocus
End Sub

Private Sub AddData()
    With rsTemBatchStock
        If .State = 1 Then .Close
        temSql = "SELECT tblBatchStock.BatchStockID, tblBatchStock.BatchID, tblBatchStock.StoreID, tblBatchStock.Stock " & _
                    "From tblBatchStock " & _
                    "WHERE (((tblBatchStock.BatchID)=" & dtcBatch.BoundText & ") AND ((tblBatchStock.StoreID)=" & dtcDepartment.BoundText & ")) "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Stock = Val(txtAStock.Text)
        Else
            .AddNew
            !BatchID = dtcBatch.BoundText
            !StoreID = dtcDepartment.BoundText
            !Stock = Val(txtAStock.Text)
        End If
        .Update
        .Close
    End With
    With rsAdjustment
        If .State = 1 Then .Close
        temSql = "SELECT * from tblAdjustment"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Date = Date
        !Time = Now
        !StoreID = Val(dtcDepartment.BoundText)
        !CategoryID = Val(dtcCategoryName.BoundText)
        !ItemID = Val(dtcItem.BoundText)
        !BatchID = Val(txtBatchID.Text)
        !Amount = Val(txtAdjustment.Text)
        !StaffID = UserID
        .Update
        .Close
    End With
End Sub

Private Sub ChangeData()

End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    Call AfterEdit
    dtcCategoryName.SetFocus
End Sub

Private Sub bttnSave_Click()
    If CanAdd = False Then Exit Sub
    Call AddData
    Call BeforeAddEdit
    Call ClearValues
    dtcCatogery.SetFocus
End Sub

Private Sub dtcCatogery_Change()
    If IsNumeric(dtcCatogery.BoundText) Then
        ListSelectedItems
    Else
        ListAllItems
    End If
    dtcItem.Text = Empty
    dtcCode.Text = Empty
End Sub


Private Sub ListSelectedItems()
With rsItem
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by display"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "Display"
    .BoundColumn = "ItemID"
End With
With rsCode
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by code"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcCode
    Set .RowSource = rsCode
    .ListField = "Code"
    .BoundColumn = "ItemID"
End With

End Sub

Private Sub ListAllItems()
With rsItem
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem order by display"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "display"
    .BoundColumn = "ItemID"
End With
With rsCode
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem order by code"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcCode
    Set .RowSource = rsCode
    .ListField = "Code"
    .BoundColumn = "ItemID"
End With
End Sub

Private Sub dtcCatogery_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtcCatogery.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcItem.SetFocus
    End If
End Sub


Private Sub dtcItem_Click(Area As Integer)
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    dtcCode.BoundText = dtcItem.BoundText
    NewItem.ID = Val(dtcItem.BoundText)
    lblIUnit.Caption = NewItem.IUnit
    lblIUnit1.Caption = NewItem.IUnit
    lblIUnit2.Caption = NewItem.IUnit
    FillStocks (Val(dtcItem.BoundText))
    FillBatch
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call SetValues
    Call BeforeAddEdit
End Sub

Private Sub SetValues()
    dtcIssueStaff.BoundText = UserID
    dtcCheckedStaff.BoundText = UserID
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
End Sub

Private Sub FillCombos()
    With rsDepartment
        If .State = 1 Then .Close
        .Open "Select tblStore.* From tblStore Order By Store", cnnStores, adOpenStatic, adLockReadOnly
    
        If .RecordCount = 0 Then Exit Sub
        Set dtcDepartment.RowSource = rsDepartment
        dtcDepartment.ListField = "Store"
        dtcDepartment.BoundColumn = "StoreID"
    End With

    With rsAdjustmentCategory
        If .State = 1 Then .Close
        .Open "Select* From tblAdjustmentCategory Order By AdjustmentCategory", cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Set dtcCategoryName.RowSource = rsAdjustmentCategory
            dtcCategoryName.BoundColumn = "AdjustmentCategoryID"
            dtcCategoryName.ListField = "AdjustmentCategory"
        End If
    End With
    With rsStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblstaff order by listedname"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcIssueStaff
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With dtcCheckedStaff
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With rsItem
        If .State = 1 Then .Close
        temSql = "SELECT * from tblitem order by display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcItem
        Set .RowSource = rsItem
        .ListField = "display"
        .BoundColumn = "ItemID"
    End With
    With rsItemCategory
        If .State = 1 Then .Close
        temSql = "SELECT * from tblItemCategory order by ItemCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCatogery
        Set .RowSource = rsItemCategory
        .ListField = "ItemCategory"
        .BoundColumn = "ItemCategoryID"
    End With
    With rsCode
        If .State = 1 Then .Close
        temSql = "SELECT * from tblitem order by code"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCode
        Set .RowSource = rsCode
        .ListField = "code"
        .BoundColumn = "ItemID"
    End With

End Sub


Private Sub FormatStocks()
    With GridTotal
        .Visible = False
        .Cols = 7
        .Rows = 1
        .Row = 0
        .FixedCols = 0
        .Col = 0
        .CellAlignment = 4
        .Text = "Batch"
        .Col = 1
        .CellAlignment = 4
        .Text = "Stock (" & NewItem.IUnit & ")"
        .Col = 2
        .CellAlignment = 4
        .Text = "Expiary"
        .Col = 3
        .CellAlignment = 4
        .Text = "Department"
        .ColWidth(1) = 1600
        .ColWidth(2) = 1600
        .ColWidth(3) = 1600
        .ColWidth(4) = 1
        .ColWidth(5) = 1
        .ColWidth(6) = 1
        .ColWidth(0) = .Width - (.ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 100)
    End With
    '   0   Batch
    '   1   Stock String
    '   2   Expiary
    '   3   Dept
    '   4   BatchID
    '   5   StoreID
    '   6   Stock Double
End Sub

Private Sub FillStocks(ByVal ItemID As Long)
    Call FormatStocks
    With rsTemStore
        If .State = 1 Then .Close
        temSql = "SELECT tblBatch.*, tblBatchStock.*, tblStore.* " & _
                    " FROM tblStore RIGHT JOIN (tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) ON tblStore.StoreID = tblBatchStock.StoreID " & _
                    " WHERE tblBatch.ItemID=" & ItemID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                GridTotal.Rows = GridTotal.Rows + 1
                GridTotal.Row = GridTotal.Rows - 1
                GridTotal.Col = 0
                GridTotal.CellAlignment = 1
                GridTotal.Text = !Batch
                GridTotal.Col = 1
                GridTotal.CellAlignment = 7
                If Not IsNull(!Stock) Then
                    GridTotal.Text = !Stock
                Else
                    GridTotal.Text = 0
                End If
                GridTotal.Col = 2
                GridTotal.CellAlignment = 1
                GridTotal.Text = Format(!DOE, ShortDateFormat)
                GridTotal.Col = 3
                GridTotal.CellAlignment = 1
                If Not IsNull(!Store) Then
                    GridTotal.Text = !Store
                Else
                    GridTotal.Text = Empty
                End If
                GridTotal.Col = 4
                If Not IsNull(![BatchID]) Then GridTotal.Text = ![BatchID]
                GridTotal.Col = 5
                If Not IsNull(![StoreID]) Then GridTotal.Text = ![StoreID]
                GridTotal.Col = 6
                If Not IsNull(!Stock) Then GridTotal.Text = !Stock
                .MoveNext
            Wend
        End If
        GridTotal.Visible = True
        .Close
    End With
    '   0   Batch
    '   1   Stock String
    '   2   Expiary
    '   3   Dept
    '   4   BatchID
    '   5   StoreID
    '   6   Stock Double
End Sub


Private Function CanAdd() As Boolean
    CanAdd = False
    Dim tr As Integer
        If IsNumeric(dtcDepartment.BoundText) = False Then
            tr = MsgBox("You have not entered the department", vbCritical, "Department?")
            dtcDepartment.SetFocus
            Exit Function
        End If
        If Trim(dtcCategoryName.Text) = Empty Then
            tr = MsgBox("You have not entered the reason", vbCritical, "Expiary Date")
            dtcCategoryName.SetFocus
            Exit Function
        End If
        
    CanAdd = True
End Function

Private Sub AfterAdd()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    dtcItem.Enabled = False
    dtcCatogery.Enabled = False
    dtcCode.Enabled = False
    
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnSave.Enabled = True
    bttnChange.Enabled = False
    txtAdjustment.Enabled = True
    txtAStock.Enabled = True
    dtcBatch.Enabled = True
    txtCStock.Enabled = True
    dtcCheckedStaff.Enabled = True
    dtcDepartment.Enabled = True
    dtcIssueStaff.Enabled = True
    dtcCategoryName.Enabled = True
End Sub

Private Sub ClearValues()
    txtAStock.Text = Empty
    dtcBatch.Text = Empty
    txtCStock.Text = Empty
    dtcDepartment.Text = Empty
    txtItem.Text = Empty
    txtBatchID.Text = Empty
    txtAdjustment.Text = Empty
    txtStoreID.Text = Empty
End Sub

Private Sub ClearAddValues()
    txtAStock.Text = Empty
    txtCStock.Text = Empty
    txtAdjustment.Text = Empty
End Sub


Private Sub BeforeAddEdit()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    dtcItem.Enabled = True
    dtcCatogery.Enabled = True
    dtcCode.Enabled = True
    
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnSave.Enabled = False
    bttnChange.Enabled = False
    txtAdjustment.Enabled = False
    txtAStock.Enabled = False
    dtcBatch.Enabled = False
    txtCStock.Enabled = False
    dtcCheckedStaff.Enabled = False
    dtcDepartment.Enabled = False
    dtcIssueStaff.Enabled = False
    dtcCategoryName.Enabled = False
End Sub

Private Sub AfterEdit()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    dtcItem.Enabled = False
    dtcCatogery.Enabled = False
    dtcCode.Enabled = False
    
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnSave.Enabled = False
    bttnChange.Enabled = True
    txtAdjustment.Enabled = True
    txtAStock.Enabled = True
    dtcBatch.Enabled = False
    txtCStock.Enabled = True
    dtcCheckedStaff.Enabled = True
    If Trim(dtcDepartment.Text) = "" Then
        dtcDepartment.Enabled = True
    Else
        dtcDepartment.Enabled = False
    End If
    dtcIssueStaff.Enabled = True
    dtcCategoryName.Enabled = True
End Sub

Private Sub DisplayDetails()
        With GridTotal
            dtcDepartment.BoundText = Val(.TextMatrix(.Row, 5))
            txtBatchID.Text = .TextMatrix(.Row, 4)
            txtStoreID.Text = .TextMatrix(.Row, 5)
            txtItem.Text = dtcItem.BoundText
            dtcBatch.BoundText = .TextMatrix(.Row, 4)
            txtCStock.Text = .TextMatrix(.Row, 6)
            txtAStock.Text = .TextMatrix(.Row, 6)
            txtAdjustment.Text = 0
        End With
End Sub

Private Sub GridTotal_Click()
    With GridTotal
        If .Rows > 1 And .Row >= 1 Then
            Call DisplayDetails
            bttnEdit.Enabled = True
            bttnAdd.Enabled = True
        End If
    End With
End Sub

Private Sub txtAStock_Change()
    txtAdjustment.Text = Val(txtAStock.Text) - Val(txtCStock.Text)
End Sub

Private Sub FillBatch()
    With rsBatch
        If .State = 1 Then .Close
            temSql = "SELECT tblBatch.Batch, tblBatch.BatchID, tblBatch.DOE, tblBatchStock.Stock " & _
                        "FROM tblBatch LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID " & _
                        "WHERE tblBatch.ItemID=" & Val(dtcItem.BoundText) & " " & _
                        "ORDER BY tblBatch.DOE"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcBatch
        Set .RowSource = rsBatch
        .ListField = "Batch"
        .BoundColumn = "BatchID"
    End With
End Sub

