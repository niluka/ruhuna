VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPastPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Past Purchases"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14325
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
   ScaleHeight     =   7920
   ScaleWidth      =   14325
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   12840
      TabIndex        =   5
      Top             =   7320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16744576
      Caption         =   "Close"
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
   Begin MSFlexGridLib.MSFlexGrid gridBills 
      Height          =   5775
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   10186
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16744576
      CalendarTitleForeColor=   16744576
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   72744963
      CurrentDate     =   39772
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16744576
      CalendarTitleForeColor=   16744576
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   72744963
      CurrentDate     =   39772
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   11400
      TabIndex        =   6
      Top             =   7320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16744576
      Caption         =   "Print"
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
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPastPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsRefill As New ADODB.Recordset
    Dim CsetPrinter As New cSetDfltPrinter

Private Sub btnPrint_Click()
    Dim temRow As Integer
    temRow = gridBills.Row
    If temRow < 1 Then Exit Sub
    If IsNumeric(gridBills.TextMatrix(temRow, 0)) = False Then Exit Sub
    Dim temRefillBillID As Long
    temRefillBillID = (gridBills.TextMatrix(temRow, 0))
    Dim RetVal As Integer
    Dim TemResponce     As Integer
    If Dataenvironment1.rscmmdGoodReceive.State = 1 Then Dataenvironment1.rscmmdGoodReceive.Close
    Dataenvironment1.rscmmdGoodReceive.Source = "SELECT tblItem.Display, tblRefill.DOE, tblRefill.Amount, tblRefill.FreeAmount, tblRefill.PPrice, tblRefill.Price, tblRefill.SPrice, tblRefill.LastPPrice " & _
            " FROM tblRefill LEFT JOIN tblItem ON tblRefill.ItemID = tblItem.ItemID " & _
            " WHERE (((tblRefill.RefillBillID)= " & temRefillBillID & ") AND ((tblRefill.Amount) > 0))"
    Dataenvironment1.rscmmdGoodReceive.Open
    If Dataenvironment1.rscmmdGoodReceive.RecordCount > 0 Then
        CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
        If SelectForm(ReportPaperName, Me.hdc) = 1 Then
            With dtrPurchase
                Set .DataSource = Dataenvironment1.rscmmdGoodReceive
                .Sections("Section4").Controls("lblName").Caption = HospitalName
                .Sections("Section4").Controls("lblContact").Caption = HospitalAddress
                .Sections("Section4").Controls("lblTopic").Caption = "Good Receive Note"
                .Sections("Section4").Controls("lblSUbtopic").Caption = Empty
                .Sections("Section4").Controls("lblTo").Caption = gridBills.TextMatrix(temRow, 4)
                .Sections("Section4").Controls("lblAddress").Caption = gridBills.TextMatrix(temRow, 5)
                .Sections("Section4").Controls("lblTel").Caption = gridBills.TextMatrix(temRow, 6)
                .Sections("Section4").Controls("lblFax").Caption = gridBills.TextMatrix(temRow, 7)
                .Sections("Section4").Controls("lblDate").Caption = Format(gridBills.TextMatrix(temRow, 1), LongDateFormat)
                .Sections("Section4").Controls("lblRefillID").Caption = temRefillBillID
                .Sections("Section4").Controls("lblInvoiceDate").Caption = Format(gridBills.TextMatrix(temRow, 1), LongDateFormat)
                .Sections("Section4").Controls("lblInvoiceNo").Caption = gridBills.TextMatrix(temRow, 2)
                .Sections("Section5").Controls("lblPayee").Caption = gridBills.TextMatrix(temRow, 4)
                .Sections("Section5").Controls("lblTotalAmount").Caption = gridBills.TextMatrix(temRow, 8)
                .Sections("Section5").Controls("lblDiscount").Caption = gridBills.TextMatrix(temRow, 9)
                .Sections("Section5").Controls("lblNetTotal").Caption = Format(Val(gridBills.TextMatrix(temRow, 10)), "#,##0.00")
                .Sections("Section5").Controls("lblOperatedBy").Caption = gridBills.TextMatrix(temRow, 11)
                .Show
            End With
        Else
           MsgBox "Printer Error"
        End If
    End If
        
End Sub

Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FormatGrid()
    With gridBills
        .Clear
        .Rows = 1
        .Cols = 12
        
        .Col = 0
        .Text = "Bill ID"
        
        .Col = 1
        .Text = "Date"
        
        .Col = 2
        .Text = "Invoice No"
        
        .Col = 3
        .Text = "User"
        
        .Col = 4
        .Text = "Supplier"
        
        .Col = 5
        .Text = "Addres"
        
        .Col = 6
        .Text = "Telephone"
        
        .Col = 7
        .Text = "Fax"
        
        .Col = 8
        .Text = "Gross Total"
        
        .Col = 9
        .Text = "Discount"
        
        .Col = 10
        .Text = "Net Total"
        
        .Col = 11
        .Text = "Operator"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 1200
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 3600
        .ColWidth(5) = 1
        .ColWidth(6) = 1
        .ColWidth(7) = 1
        .ColWidth(8) = 1400
        .ColWidth(9) = 1200
        .ColWidth(10) = 1400
        .ColWidth(11) = 1400
    End With
End Sub

Private Sub FillGrid()
    Dim i As Integer
    With rsRefill
        If .State = 1 Then .Close
        temSql = "SELECT tblRefillBill.RefillBillID, tblRefillBill.StaffID, tblRefillBill.InvoiceDate, tblRefillBill.InvoiceNo, tblStaff.Name, tblDistrubutor.DistributorName, tblRefillBill.NetPrice, tblRefillBill.InvoiceDate, tblDistrubutor.DistributorAddress, tblDistrubutor.DistributorTelephone, tblDistrubutor.DistributorFax, tblRefillBill.Price, tblRefillBill.Discount " & _
                    "FROM tblStaff RIGHT JOIN (tblDistrubutor RIGHT JOIN tblRefillBill ON tblDistrubutor.DistributorID = tblRefillBill.DistributorID) ON tblStaff.StaffID = tblRefillBill.StaffID " & _
                    "WHERE (((tblRefillBill.InvoiceDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'))" & _
                    "ORDER BY tblRefillBill.RefillBillID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            gridBills.Rows = .RecordCount + 1
            For i = 1 To .RecordCount
                gridBills.TextMatrix(i, 0) = !RefillBillID
                gridBills.TextMatrix(i, 1) = !InvoiceDate
                gridBills.TextMatrix(i, 2) = !InvoiceNo
                gridBills.TextMatrix(i, 3) = !Name
                gridBills.TextMatrix(i, 4) = !DistributorName
                gridBills.TextMatrix(i, 5) = !DistributorAddress
                gridBills.TextMatrix(i, 6) = !DistributorTelephone
                gridBills.TextMatrix(i, 7) = !DistributorFax
                gridBills.TextMatrix(i, 8) = !Price
                gridBills.TextMatrix(i, 9) = !Discount
                gridBills.TextMatrix(i, 10) = !NetPrice
                gridBills.TextMatrix(i, 11) = getUserName(!StaffID)
                .MoveNext
            Next
        End If
        gridBills.Row = 0
        .Close
    End With
End Sub

Private Sub PrintBill(ByVal RefillBillID As Long)

End Sub


Private Sub gridBills_Click()
    With gridBills
        .Col = .Cols - 1
        .ColSel = 0
    End With
End Sub
