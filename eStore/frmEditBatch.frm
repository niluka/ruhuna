VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditBatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11025
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
   ScaleHeight     =   5370
   ScaleWidth      =   11025
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   9240
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add Batch"
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
   Begin VB.TextBox txtBatch 
      Height          =   375
      Left            =   8880
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo dtcCatogery 
      Height          =   360
      Left            =   1200
      TabIndex        =   0
      Top             =   240
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
      Left            =   1200
      TabIndex        =   1
      Top             =   720
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
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid GridTotal 
      Height          =   3495
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6165
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpDOM 
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   20840451
      CurrentDate     =   39545
   End
   Begin MSComCtl2.DTPicker dtpDOE 
      Height          =   375
      Left            =   8880
      TabIndex        =   9
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   20840451
      CurrentDate     =   39545
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9240
      TabIndex        =   14
      Top             =   4800
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.Label Label36 
      Caption         =   "Date of Expiary"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label37 
      Caption         =   "Date of Manufacture"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label38 
      Caption         =   "Batch"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label28 
      Caption         =   "Catogery"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "Item"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label35 
      Caption         =   "Code"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frmEditBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Dim NewItem As New Item
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCategory As New ADODB.Recordset
    Dim rsCode As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    

Private Sub bttnAdd_Click()
    If CanAdd = False Then Exit Sub
        Dim ThisBatch As Long
        ThisBatch = BatchExist(Trim(txtBatch.Text), Val(dtcItem.BoundText))
        If ThisBatch = 0 Then
            ThisBatch = AddBatch(Trim(txtBatch.Text), Val(dtcItem.BoundText), dtpDOM.Value, dtpDOE.Value)
        End If
    FormatStocks
    FillStocks (Val(dtcItem.BoundText))
    txtBatch.Text = Empty
    dtpDOE.Value = Date
    dtpDOM.Value = Date
End Sub



Private Sub bttnClose_Click()
    Unload Me
End Sub


Private Sub dtcCatogery_Click(Area As Integer)
    If IsNumeric(dtcCatogery.BoundText) = False Then
        ListAllItems
    Else
        ListSelectedItems
    End If
End Sub

Private Sub ListSelectedItems()
    With rsItem
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by display"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcItem
        Set .RowSource = rsItem
        .ListField = "Display"
        .BoundColumn = "ItemID"
    End With
    With rsCode
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by code"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
    temSQL = "SELECT * from tblitem order by display"
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "display"
    .BoundColumn = "ItemID"
End With
With rsCode
    If .State = 1 Then .Close
    temSQL = "SELECT * from tblitem order by code"
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
    End If
End Sub


Private Sub dtcItem_Click(Area As Integer)
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    dtcCode.BoundText = dtcItem.BoundText
    NewItem.ID = Val(dtcItem.BoundText)
    FillStocks (Val(dtcItem.BoundText))
End Sub

Private Sub Form_Load()
    Call FillCombos
End Sub

Private Sub FillCombos()
    With rsItem
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblitem order by display"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcItem
        Set .RowSource = rsItem
        .ListField = "display"
        .BoundColumn = "ItemID"
    End With
    With rsItemCategory
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblItemCategory order by ItemCategory"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCatogery
        Set .RowSource = rsItemCategory
        .ListField = "ItemCategory"
        .BoundColumn = "ItemCategoryID"
    End With
    With rsCode
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblitem order by code"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
        .Cols = 4
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
        .ColWidth(0) = .Width - (.ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 100)
    End With
End Sub

Private Sub FillStocks(ByVal ItemID As Long)
    Call FormatStocks
    With rsTemStore
        If .State = 1 Then .Close
        temSQL = "SELECT tblBatch.Batch, tblBatch.DOE, tblBatchStock.Stock, tblStore.Store, tblBatch.ItemID " & _
                    " FROM tblStore RIGHT JOIN (tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) ON tblStore.StoreID = tblBatchStock.StoreID " & _
                    " WHERE tblBatch.ItemID=" & ItemID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
                .MoveNext
            Wend
        End If
        GridTotal.Visible = True
        .Close
    End With
    
End Sub


Private Function CanAdd() As Boolean
    CanAdd = False
    Dim tr As Integer
        If IsNumeric(dtcItem.BoundText) = False Then
            tr = MsgBox("You have not entered the item to add", vbCritical, "Item?")
            dtcItem.SetFocus
            Exit Function
        End If
        If dtpDOE.Value = Date Then
            tr = MsgBox("You have not entered a Date of Expiary", vbCritical, "Expiary Date")
            dtpDOE.SetFocus
            Exit Function
        End If
        If Trim(txtBatch.Text) = Empty Then
            tr = MsgBox("You have not entered a Batch number", vbCritical, "Expiary Date")
            dtpDOE.SetFocus
            Exit Function
        End If
    CanAdd = True
End Function

