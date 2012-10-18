VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAllItemBHTIssue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Item Issue to BHT"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9780
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
   ScaleHeight     =   7770
   ScaleWidth      =   9780
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   8400
      TabIndex        =   16
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Process"
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
   Begin MSDataListLib.DataCombo dtcUnit 
      Height          =   360
      Left            =   1320
      TabIndex        =   10
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8400
      TabIndex        =   6
      Top             =   7200
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   7200
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
   Begin MSFlexGridLib.MSFlexGrid GridIssue 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8493
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   100401155
      CurrentDate     =   29224
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   100401155
      CurrentDate     =   29224
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   7200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Excel"
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
   Begin VB.Label Label8 
      Caption         =   "Net Value"
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   6720
      Width           =   2655
   End
   Begin VB.Label lblNetValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Total Return"
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label lblTotalReturn 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   7680
      TabIndex        =   11
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "BHT"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblTotalValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Total Value"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmAllItemBHTIssue"
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
    Dim i As Integer
    Dim TotalValue As Double
    Dim TotalReturn As Double
    Dim temTopic As String
    Dim temSubTopic As String
    Dim CsetPrinter As New cSetDfltPrinter
    
    Dim rsItemIssie As New ADODB.Recordset
    Dim rsViewUnit As New ADODB.Recordset
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel GridIssue, "Issue of Medicins to BHT - " & dtcUnit.Text
End Sub

Private Sub btnPrint_Click()
    Dim RetVal As Integer
    Dim TemResponce As Integer
    Dim CsetPrinter As New cSetDfltPrinter
    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    GridPrint GridIssue, ThisReportFormat, "All Item Issue - " & dtcUnit.Text, Format(Date, LongDateFormat)
    Printer.EndDoc
End Sub


Private Sub btnProcess_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
    Call FillCombos
    Call FormatGrid
    GetCommonSettings Me
End Sub

Private Sub FillCombos()
    With rsViewUnit
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblBHT Order by BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcUnit
        Set .RowSource = rsViewUnit
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
End Sub

Private Sub FormatGrid()
    With GridIssue
        .Clear
        
        .Rows = 1
        .Cols = 5
        
        .FixedCols = 0
        
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Item"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Quentity"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Value"
        
        .Col = 3
        .CellAlignment = 4
        .Text = "Return Quentity"
        
        .Col = 4
        .CellAlignment = 4
        .Text = "Return Value"
        
    End With
End Sub

Private Sub FillGrid()
    Dim MyItemSaleReturn As ItemSaleAndReturn
    If IsNumeric(dtcUnit.BoundText) = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    DoEvents
    With rsItemIssie
        temSql = "SELECT DISTINCT dbo.tblItem.ItemID, dbo.tblItem.Display " & _
                    "FROM dbo.tblReturn LEFT OUTER JOIN " & _
                        "dbo.tblSale RIGHT OUTER JOIN " & _
                        "dbo.tblSaleBill ON dbo.tblSale.SaleBillID = dbo.tblSaleBill.SaleBillID LEFT OUTER JOIN " & _
                        "dbo.tblItem ON dbo.tblSale.ItemID = dbo.tblItem.ItemID ON dbo.tblReturn.ItemID = dbo.tblItem.ItemID RIGHT OUTER JOIN " & _
                        "dbo.tblReturnBill ON dbo.tblReturn.ReturnBillID = dbo.tblReturnBill.ReturnBillID " & _
                        "WHERE     (dbo.tblReturn.Date BETWEEN CONVERT(DATETIME, '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "', 102) AND CONVERT(DATETIME, '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "', 102)) OR " & _
                        "(dbo.tblSale.Date BETWEEN CONVERT(DATETIME, '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "', 102) AND CONVERT(DATETIME, '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "', 102)) " & _
                        "ORDER BY dbo.tblItem.Display"
                        
                        
        temSql = "SELECT DISTINCT TOP 100 PERCENT dbo.tblItem.ItemID, dbo.tblItem.Display " & _
                    "FROM         dbo.tblSale RIGHT OUTER JOIN " & _
                      "dbo.tblSaleBill ON dbo.tblSale.SaleBillID = dbo.tblSaleBill.SaleBillID LEFT OUTER JOIN " & _
                      "dbo.tblItem ON dbo.tblSale.ItemID = dbo.tblItem.ItemID " & _
                        "WHERE     " & _
                        "tblSale.Date BETWEEN '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' " & _
                        "ORDER BY dbo.tblItem.Display"
                        
                        
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        TotalValue = 0
        TotalReturn = 0
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            GridIssue.Rows = .RecordCount + 1
            i = 1
            While .EOF = False
                If IsNull(!ItemID) = False Then
                    MyItemSaleReturn = PeriodSale(dtpFromDate.Value, dtpToDate.Value, !ItemID, Val(dtcUnit.BoundText))
                    If MyItemSaleReturn.SaleValue <> 0 Or MyItemSaleReturn.ReturnValue <> 0 Then
                        If Not IsNull(!Display) Then GridIssue.TextMatrix(i, 0) = !Display
                        GridIssue.TextMatrix(i, 1) = MyItemSaleReturn.SaleQuentity
                        GridIssue.TextMatrix(i, 2) = Format(MyItemSaleReturn.SaleValue, "#,##0.00")
                        TotalValue = TotalValue + MyItemSaleReturn.SaleValue
                        GridIssue.TextMatrix(i, 3) = MyItemSaleReturn.ReturnQuentity
                        GridIssue.TextMatrix(i, 4) = Format(MyItemSaleReturn.ReturnValue, "#,##0.00")
                        TotalReturn = TotalReturn + MyItemSaleReturn.ReturnValue
                        i = i + 1
                    End If
                End If
                .MoveNext
            Wend
        End If
        GridIssue.Rows = i
        lblTotalValue.Caption = Format(TotalValue, "#,##0.00")
        lblTotalReturn.Caption = Format(TotalReturn, "#,##0.00")
        lblNetValue.Caption = Format(TotalValue - TotalReturn, "#,##0.00")
    End With
    With GridIssue
        .Rows = .Rows + 4
        
        .Row = i + 1
        .Col = 0
        .Text = "Total Value"
        .Col = 2
        .Text = Format(TotalValue, "#,##0.00")
        
        .Row = i + 2
        .Col = 0
        .Text = "Total Return"
        .Col = 2
        .Text = Format(TotalReturn, "#,##0.00")
        
        .Row = i + 3
        .Col = 0
        .Text = "Net Value"
        .Col = 2
        .Text = Format(TotalValue - TotalReturn, "#,##0.00")
        
    
    End With
    
    
    Screen.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub
