VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAllItemMoving 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Item Issue"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
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
   ScaleHeight     =   7650
   ScaleWidth      =   10305
   Begin MSComctlLib.Slider Slider 
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   600
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   661
      _Version        =   393216
      Min             =   1
      Max             =   100
      SelStart        =   10
      TickFrequency   =   5
      Value           =   10
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   2160
      TabIndex        =   12
      Top             =   6240
      Width           =   1935
      Begin VB.OptionButton optDesc 
         Caption         =   "&Fast Moving"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optAcs 
         Caption         =   "&Slow Moving"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   6240
      Width           =   1935
      Begin VB.OptionButton optQty 
         Caption         =   "&Quentity-vice"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optVal 
         Caption         =   "&Value-vice"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   6960
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
      Left            =   7680
      TabIndex        =   5
      Top             =   6960
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
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8916
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
      Format          =   66781187
      CurrentDate     =   29224
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   66781187
      CurrentDate     =   29224
   End
   Begin VB.Label Label4 
      Caption         =   "Percentage to display"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblTotalValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Total Value"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   6480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   4080
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
Attribute VB_Name = "frmAllItemMoving"
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
            With dtrItemIssueStock
                Set .DataSource = rsItemIssie
                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
                If optAcs.Value = True Then
                    temTopic = temTopic & "Fast Moving Items - "
                ElseIf optDesc.Value = True Then
                    temTopic = temTopic & "Slow Moving Items - "
                End If
                temTopic = temTopic & "Top " & Slider.Value & "% Item Issues - "
                
                If dtpFromDate.Value = dtpToDate.Value Then
                    temSubTopic = "On " & Format(dtpFromDate.Value, LongDateFormat)
                Else
                    temSubTopic = "From " & Format(dtpFromDate.Value, LongDateFormat) & " to " & Format(dtpToDate.Value, LongDateFormat)
                End If
                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
                .Caption = temTopic & " - " & temSubTopic
                .Sections("Section1").Controls.Item("txtItem").DataField = "Display"
                .Sections("Section1").Controls.Item("txtQuentity").DataField = "SumOfAmount"
                .Sections("Section1").Controls.Item("txtValue").DataField = "SumOfPrice"
                .Sections("Section1").Controls.Item("txtStock").DataField = "SumOfStock"
                .Sections("Section5").Controls.Item("funValue").DataField = "SumOfPrice"
                .Show
            End With
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select

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
        .Cols = 4
        
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
        .Text = "Stock"
        
        
        .ColWidth(0) = 5500
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        
        
    End With
End Sub

Private Sub FillGrid()
    Screen.MousePointer = vbHourglass
    DoEvents
    With rsItemIssie
        temSelect = "SELECT TOP " & Slider.Value & " PERCENT  tblItem.Display, Sum(tblSale.Price) AS SumOfPrice, Sum(tblSale.Amount) AS SumOfAmount, Sum(tblBatchStock.Stock) AS SumOfStock "
        temFrom = "FROM (((tblSale RIGHT JOIN tblSaleBill ON tblSale.SaleBillID = tblSaleBill.SaleBillID) LEFT JOIN tblItem ON tblSale.ItemID = tblItem.ItemID) LEFT JOIN tblBatch ON tblItem.ItemID = tblBatch.ItemID) LEFT JOIN tblBatchStock ON tblBatch.BatchID = tblBatchStock.BatchID "
        temWhere = "WHERE (((tblSaleBill.Date) Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "') AND ((tblSale.Amount)>0))"
        temGroupBy = "GROUP BY tblItem.Display"
        If optQty.Value = True Then
            temOrderBY = "ORDER BY Sum(tblSale.Amount)"
        ElseIf optVal.Value = True Then
            temOrderBY = "ORDER BY Sum(tblSale.Price)"
        End If
        If optDesc.Value = True Then temOrderBY = temOrderBY & " " & " DESC"
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
                If Not IsNull(!SumOfAmount) Then GridIssue.TextMatrix(i, 1) = !SumOfAmount
                If Not IsNull(!SumOfPrice) Then GridIssue.TextMatrix(i, 2) = Format(!SumOfPrice, "#,##0.00")
                If Not IsNull(!SumOfStock) Then GridIssue.TextMatrix(i, 3) = !SumOfStock
                i = i + 1
                TotalValue = TotalValue + !SumOfPrice
                .MoveNext
            Wend
        End If
        lblTotalValue.Caption = Format(TotalValue, "#,##0.00")
    End With
    Screen.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub optAcs_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optDesc_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optItem_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optNonMoving_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optQty_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optVal_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Slider_Change()
    Call FormatGrid
    Call FillGrid
End Sub
