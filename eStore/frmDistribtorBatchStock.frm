VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDistributorBatchStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distributor-vice Stock"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13755
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
   ScaleHeight     =   8955
   ScaleWidth      =   13755
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   72941571
      CurrentDate     =   39885
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   72941571
      CurrentDate     =   39885
   End
   Begin btButtonEx.ButtonEx btnCalculate 
      Height          =   375
      Left            =   6600
      TabIndex        =   15
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Claculate"
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
      Left            =   12360
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
      Left            =   11040
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
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   11033
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbDistributor 
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
   Begin VB.Label lblPurchase 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   6600
      TabIndex        =   23
      Top             =   8520
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Total Purchase Value"
      Height          =   375
      Left            =   5040
      TabIndex        =   22
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Total Sale Value"
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label lblSale 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   1695
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
      Left            =   6600
      TabIndex        =   12
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Total Stock Value"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   7560
      Width           =   1575
   End
End
Attribute VB_Name = "frmDistributorBatchStock"
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
    Dim TotalSale As Double
    Dim TotalPurchase As Double
    Dim rsViewCategory As New ADODB.Recordset
    
Private Sub FormatGrid()
    With GridStock
        .Clear
        .Rows = 1
        .Cols = 8
        .Row = 0
        
        .Col = 0
        .Text = "Item"
        .CellAlignment = 4
        .Col = 1
        .Text = "Purchase Price"
        .CellAlignment = 4
        .Col = 2
        .Text = "Stock Quentity"
        .CellAlignment = 4
        .Col = 3
        .Text = "Stock Value"
        .CellAlignment = 4
        .Col = 4
        .Text = "Sale Quentity"
        .Col = 5
        .Text = "Sale Value"
        .Col = 6
        .Text = "Purchase Quentity"
        .Col = 7
        .Text = "Purchase Value"
        .ColWidth(0) = 4300
        .ColWidth(1) = 1000
        .ColWidth(2) = 1300
        .ColWidth(3) = 1300
        .ColWidth(4) = 1300
        .ColWidth(5) = 1300
        .ColWidth(6) = 1300
        .ColWidth(7) = 1300
    End With
End Sub

Private Sub FillGrid()
    Screen.MousePointer = vbHourglass
    DoEvents
    With rsBatchStock
        If .State = 1 Then .Close
        temSelect = "SELECT tblItem.Display, tblCurrentPurchasePrice.PPrice, tblItem.ItemID, Sum(tblBatchStock.Stock) AS SumOfStock "
        temFrom = "FROM ((tblItem LEFT JOIN (tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) ON tblItem.ItemID = tblBatch.ItemID) LEFT JOIN tblCurrentPurchasePrice ON tblItem.ItemID = tblCurrentPurchasePrice.ItemID) RIGHT JOIN tblItemCategory ON tblItem.ItemCategoryID = tblItemCategory.ItemCategoryID "
        If IsNumeric(cmbDistributor.BoundText) = False Then
            temWhere = " "
        Else
            temWhere = "WHERE tblItem.ItemCategoryID = " & Val(cmbDistributor.BoundText) & " "
        End If
        
        temGroupBy = "GROUP BY tblItem.Display, tblCurrentPurchasePrice.PPrice, tblItem.ItemID HAVING (((tblItem.Display) Is Not Null) AND ((Sum(tblBatchStock.Stock))>0)) "
        
        If optItem.Value = True Then
            temOrderBY = "ORDER BY tblItem.Display"
        ElseIf optValue.Value = True Then
            temOrderBY = "ORDER BY tblCurrentPurchasePrice.PPrice * sum(tblBatchStock.Stock)"
        ElseIf optQuentity.Value = True Then
            temOrderBY = "ORDER BY  Sum(tblBatchStock.Stock)"
'        ElseIf optExpiary.Value = True Then
'            temOrderBY = "ORDER BY tblBatch.DOE"
        End If
        
        
        
        If optDescending.Value = True Then temOrderBY = temOrderBY & " DESC"
        temSql = temSelect & " " & temFrom & " " & temWhere & " " & temGroupBy & " " & temOrderBY
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        i = 0
        TotalValue = 0
        TotalSale = 0
        TotalPurchase = 0
        If .RecordCount > 0 Then
            .MoveLast
            GridStock.Rows = .RecordCount + 1
            .MoveFirst
            While .EOF = False
                i = i + 1
                If Not IsNull(!Display) Then GridStock.TextMatrix(i, 0) = !Display
                If Not IsNull(!PPrice) Then GridStock.TextMatrix(i, 1) = Format(!PPrice, "0.00")
                If Not IsNull(!SumOfStock) Then GridStock.TextMatrix(i, 2) = !SumOfStock
                If Not IsNull(!SumOfStock) And Not IsNull(!PPrice) Then
                    GridStock.TextMatrix(i, 3) = Format(!SumOfStock * !PPrice, "#,##0.00")
                    TotalValue = TotalValue + !SumOfStock * !PPrice
                End If
                If Not IsNull(!ItemID) Then
                    GridStock.TextMatrix(i, 4) = CalculateSale(!ItemID, dtpFrom.Value, dtpTo.Value)
                    GridStock.TextMatrix(i, 5) = Format(CalculateSalePrice(!ItemID, dtpFrom.Value, dtpTo.Value), "0.00")
                    TotalSale = TotalSale + Val(GridStock.TextMatrix(i, 5))
                    
                    GridStock.TextMatrix(i, 6) = CalculatePurchase(!ItemID, dtpFrom.Value, dtpTo.Value)
                    GridStock.TextMatrix(i, 7) = Format(CalculatePurchaseValue(!ItemID, dtpFrom.Value, dtpTo.Value), "0.00")
                    TotalPurchase = TotalPurchase + Val(GridStock.TextMatrix(i, 7))
                    
                End If
                .MoveNext
            Wend
        End If
    End With
    lblValue.Caption = Format(TotalValue, "#,##0.00")
    lblSale.Caption = Format(TotalSale, "#,##0.00")
    lblPurchase.Caption = Format(TotalPurchase, "#,##0.00")
    Screen.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub btnCalculate_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    Dim RetVal As Integer
    Dim TemResponce As Integer
    Dim CsetPrinter As New cSetDfltPrinter
    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)

    GridPrint GridStock, "Item Stock - " & cmbDistributor.Text, Format(Date, LongDateFormat)
    Printer.EndDoc

End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
End Sub

Private Sub FillCombos()
    With rsViewCategory
        If .State = 1 Then .Close
        temSql = "Select* From tblDistrubutor Order By DistributorName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbDistributor
        Set .RowSource = rsViewCategory
        .ListField = "DistributorName"
        .BoundColumn = "DistributorID"
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
