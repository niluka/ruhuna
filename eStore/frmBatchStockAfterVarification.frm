VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBatchStockAfterVarification 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch Stock - After Verification"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
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
   ScaleHeight     =   8880
   ScaleWidth      =   11400
   Begin VB.Frame Frame2 
      Caption         =   "Order By"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   8160
      Width           =   4215
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optDescending 
         Caption         =   "Descinding"
         Height          =   240
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order By"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   7440
      Width           =   4215
      Begin VB.OptionButton optExpiary 
         Caption         =   "Expiary"
         Height          =   240
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optValue 
         Caption         =   "Value"
         Height          =   240
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optQuentity 
         Caption         =   "Quentity"
         Height          =   240
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9960
      TabIndex        =   2
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
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   8640
      TabIndex        =   1
      Top             =   8160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
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
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11880
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   1560
      TabIndex        =   13
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnFill 
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Label Label2 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   7560
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Total Value"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   7560
      Width           =   1575
   End
End
Attribute VB_Name = "frmBatchStockAfterVarification"
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
        .Text = "Expiary"
        .CellAlignment = 4
        .Col = 3
        .Text = "Quentity"
        .CellAlignment = 4
        .Col = 4
        .Text = "Value"
        .CellAlignment = 4
        .ColWidth(0) = 5500
        .ColWidth(1) = 1100
        .ColWidth(2) = 1400
        .ColWidth(3) = 1100
        .ColWidth(4) = 1400
    
    End With
End Sub

Private Sub FillGrid()
    Screen.MousePointer = vbHourglass
    DoEvents
    With rsBatchStock
        If .State = 1 Then .Close
        temSelect = "SELECT tblItem.Display, tblBatch.Batch, tblCurrentPurchasePrice.PPrice,  tblBatch.DOE, tblBatchStock.Stock, tblCurrentPurchasePrice.PPrice*tblBatchStock.Stock AS StockValue"
        temFrom = "FROM ((tblBatch RIGHT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID) LEFT JOIN tblItem ON tblBatch.ItemID = tblItem.ItemID) LEFT JOIN tblCurrentPurchasePrice ON tblItem.ItemID = tblCurrentPurchasePrice.ItemID"
        If IsNumeric(cmbCategory.BoundText) = False Then
            temWhere = "WHERE (((tblBatchStock.Stock)>0))"
        Else
            temWhere = "WHERE (((tblBatchStock.Stock)>0) And (tblItem.ItemCategoryID = " & Val(cmbCategory.BoundText) & " ) )"
        End If
        If optItem.Value = True Then
            temOrderBY = "ORDER BY tblItem.Display"
        ElseIf optValue.Value = True Then
            temOrderBY = "ORDER BY tblCurrentPurchasePrice.PPrice*tblBatchStock.Stock"
        ElseIf optQuentity.Value = True Then
            temOrderBY = "ORDER BY tblBatchStock.Stock"
        ElseIf optExpiary.Value = True Then
            temOrderBY = "ORDER BY tblBatch.DOE"
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
                If Not IsNull(!DOE) Then GridStock.TextMatrix(i, 2) = Format(!DOE, "MMMM yyyy")
                If Not IsNull(!Stock) Then GridStock.TextMatrix(i, 3) = !Stock
                If Not IsNull(!StockValue) Then GridStock.TextMatrix(i, 4) = Format(!StockValue, "#,##0.00")
                If Not IsNull(!StockValue) Then TotalValue = TotalValue + !StockValue
                .MoveNext
            Wend
        End If
    End With
    lblValue.Caption = Format(TotalValue, "#,##0.00")
    Screen.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub btnFill_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    Dim RetVal As Integer
    Dim TemResponce As Integer
    Dim CSetPrinter As New cSetDfltPrinter
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
            With dtrBatchStockAfterVerification
                Set .DataSource = rsBatchStock
                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
                If IsNumeric(cmbCategory.BoundText) = True Then
                    temTopic = UserStore & " - " & cmbCategory.Text
                Else
                    temTopic = UserStore
                End If
                .Sections("Section4").Controls.Item("lblContact").Caption = temTopic
                temTopic = "Batch Stock - After Verification "
                temSubTopic = "Date : " & Format(Date, "dd MMMM yyyy") & "         Time : " & Time
                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
                .Caption = temTopic & " - " & temSubTopic
                .Sections("Section1").Controls.Item("txtItem").DataField = "Display"
                .Sections("Section1").Controls.Item("txtQuentity").DataField = "Stock"
                .Sections("Section1").Controls.Item("txtValue").DataField = "StockValue"
                .Sections("Section1").Controls.Item("txtRate").DataField = "PPrice"
                .Sections("Section1").Controls.Item("txtDOE").DataField = "DOE"
                .Sections("Section1").Controls.Item("txtBatch").DataField = "Batch"
                
                .Sections("Section5").Controls.Item("funValue").DataField = "StockValue"
                .Show
            End With
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
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
