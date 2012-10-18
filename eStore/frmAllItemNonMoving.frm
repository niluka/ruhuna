VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAllItemNonMoving 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Item Issue"
   ClientHeight    =   7260
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
   ScaleHeight     =   7260
   ScaleWidth      =   9780
   Begin VB.CheckBox chkInStock 
      Caption         =   "Show Items with stocks only"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6720
      Value           =   1  'Checked
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid GridIssue 
      Height          =   5535
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   66715651
      CurrentDate     =   29224
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   66715651
      CurrentDate     =   29224
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   6720
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
      TabIndex        =   6
      Top             =   6720
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
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
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
Attribute VB_Name = "frmAllItemNonMoving"
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
    Dim temTopic As String
    Dim temSubTopic As String
    Dim CSetPrinter As New cSetDfltPrinter
    
    Dim rsItemIssie As New ADODB.Recordset
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Dim RetVal As Integer
    Dim TemResponce As Integer
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
            With dtrItem
                Set .DataSource = rsItemIssie
                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
                temTopic = "All Non Moving Items"
                If dtpFromDate.Value = dtpToDate.Value Then
                    temSubTopic = "On " & Format(dtpFromDate.Value, LongDateFormat)
                Else
                    temSubTopic = "From " & Format(dtpFromDate.Value, LongDateFormat) & " to " & Format(dtpToDate.Value, LongDateFormat)
                End If
                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
                .Caption = temTopic & " - " & temSubTopic
                .Sections("Section1").Controls.Item("txtItem").DataField = "Display"
                .Sections("Section1").Controls.Item("txtQty").DataField = "SumOfStock"
                .Show
            End With
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select

End Sub


Private Sub chkInStock_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpFromDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpToDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub



Private Sub Form_Load()
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FormatGrid()
    With GridIssue
        .Clear
        
        .Rows = 1
        .Cols = 3
        
        .FixedCols = 0
        
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Item"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Stock Qty"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Stock Value"
        
        .ColWidth(0) = .Width - 2400
        .ColWidth(1) = 2000
        .ColWidth(2) = 1
        
        
        
    End With
End Sub

Private Sub FillGrid()
    Screen.MousePointer = vbHourglass
    DoEvents
    With rsItemIssie
        If chkInStock.Value = 1 Then
            temSelect = "SELECT tblItem.Display, Sum(tblBatchStock.Stock) AS SumOfStock "
            temFrom = "FROM tblBatchStock RIGHT JOIN (tblBatch RIGHT JOIN ((tblSale LEFT JOIN tblSaleBill ON tblSale.SaleBillID = tblSaleBill.SaleBillID) RIGHT JOIN tblItem ON tblSale.ItemID = tblItem.ItemID) ON tblBatch.ItemID = tblItem.ItemID) ON tblBatchStock.BatchID = tblBatch.BatchID "
            temWhere = "WHERE (((tblSaleBill.Date) Is Null)) OR (((tblSaleBill.Date) Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "'))"
            temGroupBy = "GROUP BY tblItem.Display"
            temOrderBY = "HAVING (((Sum(tblSale.Amount)) Is Null Or (Sum(tblSale.Amount))=0) AND ((Sum(tblBatchStock.Stock))>0))"
        Else
            temSelect = "SELECT tblItem.Display, Sum(tblBatchStock.Stock) AS SumOfStock "
            temFrom = "FROM tblBatchStock RIGHT JOIN (tblBatch RIGHT JOIN ((tblSale LEFT JOIN tblSaleBill ON tblSale.SaleBillID = tblSaleBill.SaleBillID) RIGHT JOIN tblItem ON tblSale.ItemID = tblItem.ItemID) ON tblBatch.ItemID = tblItem.ItemID) ON tblBatchStock.BatchID = tblBatch.BatchID "
            temWhere = "WHERE (((tblSaleBill.Date) Is Null)) OR (((tblSaleBill.Date) Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "'))"
            temGroupBy = "GROUP BY tblItem.Display"
            temOrderBY = "HAVING (((Sum(tblSale.Amount)) Is Null)) OR (((Sum(tblSale.Amount))=0))"
        End If
        If .State = 1 Then .Close
        temSql = temSelect & " " & temFrom & " " & temWhere & " " & temGroupBy & " " & temOrderBY
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        TotalValue = 0
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            GridIssue.Rows = .RecordCount + 1
            i = 1
            While .EOF = False
                If Not IsNull(!Display) Then GridIssue.TextMatrix(i, 0) = !Display
                If Not IsNull(!SumOfStock) Then GridIssue.TextMatrix(i, 1) = !SumOfStock
                i = i + 1
                .MoveNext
            Wend
        End If
    End With
    Screen.MousePointer = vbDefault
    DoEvents
End Sub


