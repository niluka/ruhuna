VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPriceAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price Adjustment"
   ClientHeight    =   4695
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
   ScaleHeight     =   4695
   ScaleWidth      =   11700
   Begin VB.TextBox txtStoreID 
      Height          =   360
      Left            =   4560
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtBatchID 
      Height          =   360
      Left            =   3240
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frameAdjust 
      Caption         =   "Adjust"
      Height          =   3855
      Left            =   6720
      TabIndex        =   7
      Top             =   120
      Width           =   4815
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   2640
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
      Begin VB.TextBox txtNPrice 
         Height          =   360
         Left            =   1800
         TabIndex        =   13
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtCPrice 
         BackColor       =   &H8000000F&
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   3135
      End
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   3000
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   3000
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
      Begin VB.Label Label3 
         Caption         =   "Item"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "New Sale Price"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Current Sale Price"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
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
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4260
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10200
      TabIndex        =   16
      Top             =   4080
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
      TabIndex        =   19
      Top             =   4080
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
      TabIndex        =   20
      Top             =   4080
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
Attribute VB_Name = "frmPriceAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemSql As String
    Dim NewItem As New Item
    
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCatogery As New ADODB.Recordset
    Dim rsCode As New ADODB.Recordset
    Dim rsSPrice As New ADODB.Recordset

Private Sub bttnAdd_Click()
    Call ClearAddValues
    Call AfterAdd
    txtNPrice.SetFocus
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
    txtNPrice.SetFocus
End Sub

Private Sub bttnChange_Click()
    If CanAdd = False Then Exit Sub
    Call AddData
    Call BeforeAddEdit
    Call ClearValues
    dtcCatogery.SetFocus
End Sub

Private Sub AddData()
    With rsTem
        If .State = 1 Then .Close
        TemSql = "Select* From tblCurrentSalePrice"
        .Open TemSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !ItemID = Val(dtcItem.BoundText)
            !SPrice = Val(txtNPrice.Text)
            !setdate = Date
            !SetTime = Time
            !StaffID = UserID
            .Update
        If .State = 1 Then .Close
    End With
End Sub

Private Sub ChangeData()

End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnEdit_Click()
    Call AfterEdit
    txtNPrice.SetFocus
End Sub

Private Sub bttnSave_Click()
    If CanAdd = False Then Exit Sub
    Call AddData
    Call BeforeAddEdit
    Call ClearValues
    dtcCatogery.SetFocus
End Sub

Private Sub dtcItem_Click(Area As Integer)
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    dtcCode.BoundText = dtcItem.BoundText
    NewItem.ID = Val(dtcItem.BoundText)
    FillPrice (Val(dtcItem.BoundText))
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call SetValues
    Call BeforeAddEdit
End Sub

Private Sub SetValues()
    bttnEdit.Enabled = False
End Sub

Private Sub FillCombos()
    With rsItem
        If .State = 1 Then .Close
        TemSql = "SELECT * from tblitem order by display"
        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcItem
        Set .RowSource = rsItem
        .ListField = "display"
        .BoundColumn = "ItemID"
    End With
    With rsItemCatogery
        If .State = 1 Then .Close
        TemSql = "SELECT * from tblitemcatogery order by itemcatogery"
        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCatogery
        Set .RowSource = rsItemCatogery
        .ListField = "ItemCatogery"
        .BoundColumn = "ItemCategoryID"
    End With
    With rsCode
        If .State = 1 Then .Close
        TemSql = "SELECT * from tblitem order by code"
        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCode
        Set .RowSource = rsCode
        .ListField = "code"
        .BoundColumn = "ItemID"
    End With

End Sub


Private Sub FormatPrices()
    With GridSPrice
        .Cols = 2
        .Rows = 1
        .FixedCols = 0
        
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Starting Date"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Sales Price per " & NewItem.IUnit
        
        .ColWidth(0) = (.Width - 100) / 2
        .ColWidth(1) = (.Width - 100) / 2
        
    End With

    '   0   Batch
    '   1   Stock String
    '   2   Expiary
    '   3   Dept
    '   4   BatchID
    '   5   StoreID
    '   6   Stock Double
End Sub

Private Sub FillPrice(ByVal ItemID As Long)

    FormatPrices

    With rsTemPrice
        If .State = 1 Then .Close
        TemSql = "SELECT tblCurrentSalePrice.SetDate, tblCurrentSalePrice.SPrice FROM tblCurrentSalePrice WHERE (((tblCurrentSalePrice.ItemID)=" & ItemID & ") AND ((tblCurrentSalePrice.SetDate) Between #" & Format(dtpPFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpPTo.Value, "dd MMMM yyyy") & "#)) ORDER BY tblCurrentSalePrice.SetDate DESC"
        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                With GridSPrice
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = 0
                    .CellAlignment = 1
                    .Text = Format(rsTemPrice!setdate, LongDateFormat)
                    .Col = 1
                    .CellAlignment = 7
                    .Text = Format(rsTemPrice!SPrice, "#,#00.00")
                End With
                .MoveNext
            Wend
        End If
    End With


End Sub


Private Function CanAdd() As Boolean
    CanAdd = False
    Dim TR As Integer
        If IsNumeric(dtcDepartment.BoundText) = False Then
            TR = MsgBox("You have not entered the department", vbCritical, "Department?")
            dtcDepartment.SetFocus
            Exit Function
        End If
        If Trim(dtcCategoryName.Text) = Empty Then
            TR = MsgBox("You have not entered the reason", vbCritical, "Expiary Date")
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
            TemSql = "SELECT tblBatch.Batch, tblBatch.BatchID, tblBatch.DOE, tblBatchStock.Stock " & _
                        "FROM tblBatch LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID " & _
                        "WHERE tblBatch.ItemID=" & Val(dtcItem.BoundText) & " " & _
                        "ORDER BY tblBatch.DOE"
        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcBatch
        Set .RowSource = rsBatch
        .ListField = "Batch"
        .BoundColumn = "BatchID"
    End With
End Sub

