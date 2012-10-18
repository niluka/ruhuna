VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShiftEndCreditSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift End Cash Sale"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   15015
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   13680
      TabIndex        =   6
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   12360
      TabIndex        =   5
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   70254595
      CurrentDate     =   29224
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   12360
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   70254595
      CurrentDate     =   29224
   End
   Begin MSFlexGridLib.MSFlexGrid GridSales 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5318
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid GridReturn 
      Height          =   3015
      Left            =   7560
      TabIndex        =   15
      Top             =   1200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5318
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo dtcStaff 
      Height          =   360
      Left            =   1440
      TabIndex        =   19
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label9 
      Caption         =   "User"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Cash Refunds"
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Cash Collection"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   2415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4680
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label8 
      Caption         =   "Net Cash Collection"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label lblNetCash 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Total Cash Collection"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lblRefunds 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Total Cost"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblTotalCost 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Total Cash Refunds"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   11760
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmShiftEndCreditSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsSale As New ADODB.Recordset
    Dim CSetPrinter As New cSetDfltPrinter
    Dim rsReport As New ADODB.Recordset
    Dim rsReport1 As New ADODB.Recordset
    Dim rsViewStaff As New ADODB.Recordset

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub FillCombos()
    With rsViewStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblstaff order by listedname"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcStaff
        Set .RowSource = rsViewStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
End Sub

Private Sub bttnPrint_Click()
    Dim TemResponce As Long
    Dim RetVal As Integer
    If IsNumeric(dtcStaff.BoundText) = False Then
        TemResponce = MsgBox("You have not selected the staff member", vbCritical, "Name")
        dtcStaff.SetFocus
        Exit Sub
    End If
    
    
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
            Call SelectPrint
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select
End Sub

Private Sub SelectPrint()
    With rsReport
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleBill.SaleBillID, tblSaleBill.NetCost, tblSaleBill.NetPrice, tblSaleBill.Time, tblSaleBill.Date, tblStaff.Name, tblSaleBill.NetPrice " & _
                    "FROM (tblSaleBill LEFT JOIN tblStaff ON tblSaleBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblSaleBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID " & _
                    "WHERE (((tblSaleBill.Date) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblPaymentMethod.PaymentMethod)='Credit')  AND ((tblSaleBill.StaffID)= " & dtcStaff.BoundText & " ) )"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrShiftEndCashSale
        Set .DataSource = rsReport
        .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
        .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
        .Sections("Section4").Controls.Item("lblTopic").Caption = "Shift End Credit Sale - " & dtcStaff.Text
        .Sections("Section4").Controls.Item("lblSubTopic").Caption = "Date " & Format(dtpFrom.Value, LongDateFormat)
        .Show
    End With
    With rsReport1
        If .State = 1 Then .Close
        temSql = "SELECT tblReturnBill.ReturnBillID, tblReturnBill.SaleBillID, tblReturnBill.NetCost, tblReturnBill.Date, tblReturnBill.Time, tblStaff.Name, tblReturnBill.NetPrice " & _
                    "FROM tblSaleBill RIGHT JOIN ((tblReturnBill LEFT JOIN tblStaff ON tblReturnBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblReturnBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) ON tblSaleBill.SaleBillID = tblReturnBill.SaleBillID " & _
                    "WHERE (((tblReturnBill.Date) Between '" & dtpFrom.Value & "' And '" & dtpTo.Value & "') AND ((tblPaymentMethod.PaymentMethod)='Credit')   AND ((tblReturnBill.StaffID)= " & dtcStaff.BoundText & " )  )"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrShiftEndCashReturn
        Set .DataSource = rsReport1
        .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
        .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
        .Sections("Section4").Controls.Item("lblTopic").Caption = "Shift End Credit Return - " & dtcStaff.Text
        .Sections("Section4").Controls.Item("lblSubTopic").Caption = "Date " & Format(dtpFrom.Value, LongDateFormat)
        .Show
    End With
    
End Sub


Private Sub dtcStaff_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpFrom_Change()
    dtpTo.Value = dtpFrom.Value
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Call FillCombos
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FormatGrid()
    Call FormatSaleGrid
    Call FormatReturnGrid
End Sub

Private Sub FormatSaleGrid()
    With GridSales
        .Clear
        
        .Rows = 1
        .Cols = 5
        
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Bill ID"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Date"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Time"
        
        .Col = 3
        .CellAlignment = 4
        .Text = "Staff"
        
        .Col = 4
        .CellAlignment = 4
        .Text = "Amount"
        
        
    .ColWidth(0) = 600
    .ColWidth(1) = 1400
    .ColWidth(2) = 1400
    .ColWidth(3) = 1800
    .ColWidth(4) = 1400

    
    End With
    
    
    lblTotal.Caption = "0.00"
    lblTotalCost.Caption = "0.00"
    
End Sub

Private Sub FormatReturnGrid()
    With GridReturn
        .Clear
        
        .Rows = 1
        .Cols = 6
        .FixedCols = 2
        
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Bill ID"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Bill ID"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Date"
        
        .Col = 3
        .CellAlignment = 4
        .Text = "Time"
        
        .Col = 4
        .CellAlignment = 4
        .Text = "Staff"
        
        .Col = 5
        .CellAlignment = 4
        .Text = "Amount"
        
        
    .ColWidth(0) = 600
    .ColWidth(1) = 600
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 1800
    .ColWidth(5) = 1400

    
    End With
    
    
    lblRefunds.Caption = "0.00"
    
End Sub


Private Sub FillSaleGrid()
    If IsNumeric(dtcStaff.BoundText) = False Then Exit Sub
    With rsSale
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleBill.SaleBillID, tblSaleBill.NetCost, tblSaleBill.Date, tblSaleBill.Time, tblStaff.Name, tblSaleBill.NetPrice " & _
                    "FROM (tblSaleBill LEFT JOIN tblStaff ON tblSaleBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblSaleBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID " & _
                    "WHERE (((tblSaleBill.Date) Between '" & dtpFrom.Value & "' And '" & dtpTo.Value & "') AND ((tblPaymentMethod.PaymentMethod)='Credit')  AND ((tblSaleBill.StaffID)=" & dtcStaff.BoundText & ") )"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            GridSales.Rows = .RecordCount + 1
            .MoveFirst
            Dim i As Integer
            Dim TCash As Double
            Dim TCost As Double
            While .EOF = False
                i = i + 1
                GridSales.TextMatrix(i, 0) = !SaleBillID
                GridSales.TextMatrix(i, 1) = !Date
                GridSales.TextMatrix(i, 2) = !Time
                GridSales.TextMatrix(i, 3) = !Name
                GridSales.TextMatrix(i, 4) = Format(!NetPrice, "0.00")
                
                TCash = TCash + !NetPrice
                If Not IsNull(!NetCost) Then
                    TCost = TCost + !NetCost
                End If
                .MoveNext
            Wend
        End If
    End With
    lblTotal.Caption = Format(TCash, "0.00")
    lblTotalCost.Caption = Format(TCost, "0.00")
End Sub

Private Sub FillGrid()
    Call FillSaleGrid
    Call FillReturnGrid
    lblNetCash.Caption = Format(Val(lblTotal.Caption) - Val(lblRefunds.Caption), "0.00")
End Sub

Private Sub FillReturnGrid()
    If IsNumeric(dtcStaff.BoundText) = False Then Exit Sub
    With rsSale
        If .State = 1 Then .Close
        temSql = "SELECT tblReturnBill.ReturnBillID, tblReturnBill.SaleBillID, tblReturnBill.NetCost, tblReturnBill.Date, tblReturnBill.Time, tblStaff.Name, tblReturnBill.NetPrice " & _
                    "FROM tblSaleBill RIGHT JOIN ((tblReturnBill LEFT JOIN tblStaff ON tblReturnBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblReturnBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) ON tblSaleBill.SaleBillID = tblReturnBill.SaleBillID " & _
                    "WHERE (((tblReturnBill.Date) Between '" & dtpFrom.Value & "' And '" & dtpTo.Value & "') AND ((tblPaymentMethod.PaymentMethod)='Credit') AND ((tblReturnBill.StaffID)=" & dtcStaff.BoundText & ") ) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            GridReturn.Rows = .RecordCount + 1
            .MoveFirst
            Dim i As Integer
            Dim TCash As Double
            Dim TCost As Double
            While .EOF = False
                i = i + 1
                GridReturn.TextMatrix(i, 0) = ![SaleBillID]
                GridReturn.TextMatrix(i, 1) = ![ReturnBillID]
                GridReturn.TextMatrix(i, 2) = Format(![Date], ShortDateFormat)
                GridReturn.TextMatrix(i, 3) = ![Time]
                GridReturn.TextMatrix(i, 4) = !Name
                GridReturn.TextMatrix(i, 5) = Format(!NetPrice, "0.00")
                TCash = TCash + !NetPrice
                If Not IsNull(!NetCost) Then
                    TCost = TCost + !NetCost
                End If
                .MoveNext
            Wend
        End If
    End With
    lblRefunds.Caption = Format(TCash, "0.00")
End Sub

