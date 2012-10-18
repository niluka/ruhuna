VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAllItemHospitalIssue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Item Issue to Hospital Units"
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
   Begin MSDataListLib.DataCombo dtcUnit 
      Height          =   360
      Left            =   720
      TabIndex        =   17
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   2160
      TabIndex        =   13
      Top             =   5880
      Width           =   1935
      Begin VB.OptionButton optDesc 
         Caption         =   "&Descending"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton optAcs 
         Caption         =   "&Ascending"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   1935
      Begin VB.OptionButton optItem 
         Caption         =   "&Item-vice"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
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
         Top             =   960
         Width           =   1695
      End
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8400
      TabIndex        =   6
      Top             =   6600
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
      Top             =   6600
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
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16449539
      CurrentDate     =   29224
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16449539
      CurrentDate     =   29224
   End
   Begin VB.Label Label4 
      Caption         =   "Unit"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblTotalValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Total Value"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   3600
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
Attribute VB_Name = "frmAllItemHospitalIssue"
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
    Dim CsetPrinter As New cSetDfltPrinter
    
    Dim rsItemIssie As New ADODB.Recordset
    Dim rsViewUnit As New ADODB.Recordset
    
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Dim RetVal As Integer
    Dim TemResponce As Integer
    If IsNumeric(dtcUnit.BoundText) = False Then Exit Sub
    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
            With dtrItemIssue
                Set .DataSource = rsItemIssie
                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
                temTopic = "All Item Issues - " & dtcUnit.Text
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
                .Sections("Section5").Controls.Item("funValue").DataField = "SumOfPrice"
                .Show
            End With
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select

End Sub


Private Sub dtcUnit_Change()
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
    Call FillCombos
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FillCombos()
    With rsViewUnit
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblStore Order by Store"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcUnit
        Set .RowSource = rsViewUnit
        .ListField = "Store"
        .BoundColumn = "StoreID"
    End With
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
        .Text = "Quentity"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Value"
        
        .ColWidth(0) = 5500
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        
        
        
    End With
End Sub

Private Sub FillGrid()
    If IsNumeric(dtcUnit.BoundText) = False Then Exit Sub
    Screen.MousePointer = vbHourglass
    DoEvents
    With rsItemIssie
        temSelect = "SELECT tblItem.Display, Sum(tblSale.Price) AS SumOfPrice, Sum(tblSale.Amount) AS SumOfAmount "
        temFrom = "FROM (tblSale RIGHT JOIN tblSaleBill ON tblSale.SaleBillID = tblSaleBill.SaleBillID) LEFT JOIN tblItem ON tblSale.ItemID = tblItem.ItemID "
        temWhere = "WHERE (((tblSaleBill.Date) Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "') AND ((tblSale.BilledUnitID)=" & Val(dtcUnit.BoundText) & ") AND ((tblSale.Amount)>0))"
        temGroupBy = "GROUP BY tblItem.Display"
        If optItem.Value = True Then
            temOrderBY = "ORDER BY tblItem.Display"
        ElseIf optQty.Value = True Then
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

Private Sub optQty_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optVal_Click()
    Call FormatGrid
    Call FillGrid
End Sub
