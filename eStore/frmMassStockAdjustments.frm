VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMassStockAdjustments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mass Stock Adjustment"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13215
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
   ScaleHeight     =   8835
   ScaleWidth      =   13215
   Begin VB.CheckBox chkShowZero 
      Caption         =   "Only Show Items with no stocks"
      Height          =   255
      Left            =   4320
      TabIndex        =   19
      Top             =   8280
      Width           =   3735
   End
   Begin VB.CheckBox chkShowAll 
      Caption         =   "Show Items with no stocks"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   8280
      Width           =   3735
   End
   Begin btButtonEx.ButtonEx btnFill 
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Fill"
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
   Begin VB.TextBox txtCellText 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6600
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order By"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   6840
      Width           =   4215
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optQuentity 
         Caption         =   "Quentity"
         Height          =   240
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Order By"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   4215
      Begin VB.OptionButton optDescending 
         Caption         =   "Descinding"
         Height          =   240
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   11640
      TabIndex        =   6
      Top             =   8160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin MSFlexGridLib.MSFlexGrid GridStock 
      Height          =   6135
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   10821
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCheckedStaff 
      Height          =   360
      Left            =   6000
      TabIndex        =   11
      Top             =   7320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcIssueStaff 
      Height          =   360
      Left            =   6000
      TabIndex        =   12
      Top             =   6840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
      BackColor       =   -2147483633
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCategoryName 
      Height          =   360
      Left            =   6000
      TabIndex        =   13
      Top             =   7800
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label6 
      Caption         =   "User"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Checked by"
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Reason"
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMassStockAdjustments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim temSelect As String
    Dim temFrom As String
    Dim temWhere As String
    Dim temGroupBy As String
    Dim temOrderBY As String
    Dim rsBatchStock As New ADODB.Recordset
    Dim temTopic As String
    Dim temSubTopic As String
    Dim i As Integer
    Dim TotalValue As Double
    Dim rsViewCategory As New ADODB.Recordset
    Dim temRow As Long
    Dim temCol As Long
    Dim temCellText As String
    Dim temBoxText As String
    Dim temText As String
    
    
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
    
    
Private Sub AdjustStock(ItemID As Long, BatchID As Long, OldStock As Double, NewStock As Double)
    With rsTemBatchStock
        If .State = 1 Then .Close
        temSql = "SELECT tblBatchStock.BatchStockID, tblBatchStock.BatchID, tblBatchStock.StoreID, tblBatchStock.Stock " & _
                    "From tblBatchStock " & _
                    "WHERE (((tblBatchStock.BatchID)=" & BatchID & ") AND ((tblBatchStock.StoreID)=" & UserStoreID & ")) "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Stock = NewStock
        Else
            .AddNew
            !BatchID = BatchID
            !StoreID = UserStoreID
            !Stock = NewStock
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
        !Time = Time
        !StoreID = UserStoreID
        !CategoryID = Val(dtcCategoryName.BoundText)
        !ItemID = ItemID
        !BatchID = BatchID
        !Amount = NewStock - OldStock
        .Update
        .Close
    End With
End Sub
    
    
Private Sub btnFill_Click()
    GridStock.Col = 1
    GridStock.Row = 1
    GridStock.Col = 2
    GridStock.Row = 2
    
    If IsNumeric(cmbCategory.BoundText) = False Then
        MsgBox "Please select a item category"
        cmbCategory.SetFocus
        Exit Sub
    End If
    If IsNumeric(dtcCategoryName.BoundText) = False Then
        MsgBox "Please select a reason for stock adjustment"
        dtcCategoryName.SetFocus
        Exit Sub
    End If
    
    Call FormatGrid
    Call FillGrid

End Sub

Private Sub GridStock_EnterCell()
    txtCellText.Visible = False
    If GridStock.Row = 0 Then
    
    ElseIf GridStock.CellWidth < 2 Then
    
    ElseIf GridStock.Visible = False Then Exit Sub
    
    Else
        temRow = GridStock.Row
        temCol = GridStock.Col
        temCellText = GridStock.TextMatrix(temRow, temCol)
        txtCellText.Top = GridStock.Top + GridStock.CellTop
        txtCellText.Left = GridStock.Left + GridStock.CellLeft
        txtCellText.Height = GridStock.CellHeight - 60
        txtCellText.Width = GridStock.CellWidth
        'txtCellText.BackColor = GridStock.CellBackColor
        txtCellText.Alignment = GridStock.CellAlignment
        txtCellText.Text = temCellText
        txtCellText.Visible = True
        On Error Resume Next
        txtCellText.SetFocus
        SendKeys "{Home}+{end}"
    End If
    
    If GridStock.Col = 2 Then
        txtCellText.Locked = False
        Call CalculateSettling
    Else
        txtCellText.Locked = True
    End If

End Sub

Private Sub CalculateSettling()
    Beep
End Sub

Private Sub GridStock_LeaveCell()
    txtCellText.Visible = False
    If GridStock.Row = 0 Then
    
    ElseIf temRow = 0 Then Exit Sub
    
    ElseIf GridStock.Col = 0 Or GridStock.Col = 1 Or GridStock.Col = 3 Then
    
    ElseIf GridStock.CellWidth < 2 Then
    
    ElseIf GridStock.Visible = False Then
        Exit Sub
    Else
        temBoxText = txtCellText.Text
        If GridStock.TextMatrix(temRow, temCol) <> txtCellText.Text Then
            AdjustStock Val(GridStock.TextMatrix(temRow, 4)), Val(GridStock.TextMatrix(temRow, 3)), GridStock.TextMatrix(temRow, 2), Val(txtCellText.Text)
            GridStock.TextMatrix(temRow, temCol) = temBoxText
        End If
    End If
End Sub

Private Sub GridStock_Scroll()
    txtCellText.Visible = False
End Sub

Private Sub txtCellText_KeyDown(KeyCode As Integer, Shift As Integer)
    With GridStock
        If KeyCode = vbKeyReturn Then
            If temCol < .Cols - 1 Then
                .Col = temCol + 1
            Else
                .Col = 1
                .Row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyEscape Then
            txtCellText.Text = temText
        ElseIf KeyCode = vbKeyTab Then
            If temCol < .Cols - 1 Then
                .Col = temCol + 1
            Else
                .Col = 1
                .Row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyUp Then
            If temRow > 1 Then
                .Row = temRow - 1
            End If
        ElseIf KeyCode = vbKeyDown Then
            If temRow < .Rows - 1 Then
                .Row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If temCol > 1 Then
                .Col = temCol - 1
            End If
        ElseIf KeyCode = vbKeyRight Then
            If temCol < .Cols - 1 Then
                .Col = temCol + 1
            End If
        End If
    End With
End Sub

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
Private Sub FormatGrid()
    With GridStock
        .Clear
        
        .Rows = 1
        .Cols = 5
        
        .Row = 0
        
        .Col = 0
        .Text = "Item"
        .CellAlignment = 4
        
        .Col = 1
        .Text = "Batch"
        .CellAlignment = 4
        
        .Col = 2
        .Text = "Quentity"
        .CellAlignment = 4
        
        .Col = 3
        .Text = "BatchID"
        
        .Col = 4
        .Text = "ItemID"
        
        
        
        .ColWidth(0) = 5500
        .ColWidth(1) = 2100
        .ColWidth(2) = 2400
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        
    End With
End Sub

Private Sub FillGrid()
    Screen.MousePointer = vbHourglass
    DoEvents
    With rsBatchStock
        If .State = 1 Then .Close
        temSelect = "SELECT tblItem.Display, tblItem.ItemID, tblBatch.BatchID, tblBatch.Batch, tblBatch.DOE, tblBatchStock.Stock, tblCurrentPurchasePrice.PPrice*tblBatchStock.Stock AS StockValue"
        temFrom = "FROM ((tblBatch RIGHT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID) LEFT JOIN tblItem ON tblBatch.ItemID = tblItem.ItemID) LEFT JOIN tblCurrentPurchasePrice ON tblItem.ItemID = tblCurrentPurchasePrice.ItemID"
        If IsNumeric(cmbCategory.BoundText) = False Then
            If chkShowAll.Value = 1 Then
                temWhere = "WHERE (((tblBatchStock.Stock)>=0))"
            ElseIf chkShowZero.Value = 1 Then
                temWhere = "WHERE (((tblBatchStock.Stock)=0))"
            Else
                temWhere = "WHERE (((tblBatchStock.Stock)>0))"
            End If
            
        Else
            If chkShowAll.Value = 1 Then
                temWhere = "WHERE (((tblBatchStock.Stock)>=0) And (tblItem.ItemCategoryID = " & Val(cmbCategory.BoundText) & " ) )"
            ElseIf chkShowZero.Value = 1 Then
                temWhere = "WHERE (((tblBatchStock.Stock)=0) And (tblItem.ItemCategoryID = " & Val(cmbCategory.BoundText) & " ) )"
            Else
                temWhere = "WHERE (((tblBatchStock.Stock)>0) And (tblItem.ItemCategoryID = " & Val(cmbCategory.BoundText) & " ) )"
            End If
        End If
        If optItem.Value = True Then
            temOrderBY = "ORDER BY tblItem.Display"
        ElseIf optQuentity.Value = True Then
            temOrderBY = "ORDER BY tblBatchStock.Stock"
        End If
        If optDescending.Value = True Then temOrderBY = temOrderBY & " DESC"
        temSql = temSelect & " " & temFrom & " " & temWhere & " " & temOrderBY
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        i = 0
        TotalValue = 0
        If .RecordCount > 0 Then
            .MoveLast
            GridStock.Rows = .RecordCount + 1
            .MoveFirst
            While .EOF = False
                i = i + 1
                If Not IsNull(!Display) Then GridStock.TextMatrix(i, 0) = !Display
                If Not IsNull(!Batch) Then GridStock.TextMatrix(i, 1) = !Batch
                If Not IsNull(!Stock) Then GridStock.TextMatrix(i, 2) = !Stock
                If Not IsNull(!BatchID) Then GridStock.TextMatrix(i, 3) = !BatchID
                If Not IsNull(!ItemID) Then GridStock.TextMatrix(i, 4) = !ItemID
                .MoveNext
            Wend
        End If
    End With
    Screen.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    dtcIssueStaff.BoundText = UserID
End Sub

Private Sub FillCombos()
    With rsViewCategory
        If .State = 1 Then .Close
        temSql = "Select * from tblItemCategory order by ItemCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCategory
        Set .RowSource = rsViewCategory
        .ListField = "ItemCategory"
        .BoundColumn = "ItemCategoryID"
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
End Sub

Private Sub optAscending_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optDescending_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optExpiary_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optItem_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optQuentity_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optValue_Click()
    Call FormatGrid
    Call FillGrid
End Sub

