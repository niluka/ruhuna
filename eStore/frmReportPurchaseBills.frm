VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReportPurchaseBills 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Bills"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12285
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
   ScaleHeight     =   7800
   ScaleWidth      =   12285
   Begin btButtonEx.ButtonEx bttnClose 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10920
      TabIndex        =   11
      Top             =   7320
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
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   10080
      TabIndex        =   10
      Top             =   6840
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo dtcSupplier 
      Height          =   360
      Left            =   1800
      TabIndex        =   7
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   21889027
      CurrentDate     =   39697
   End
   Begin MSFlexGridLib.MSFlexGrid GridPurchase 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   8916
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   21889027
      CurrentDate     =   39697
   End
   Begin MSDataListLib.DataCombo dtcPayment 
      Height          =   360
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Default         =   -1  'True
      Height          =   375
      Left            =   9600
      TabIndex        =   13
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin VB.Label Label5 
      Caption         =   "Total Value"
      Height          =   375
      Left            =   7560
      TabIndex        =   9
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Payment Method"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Supplier"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmReportPurchaseBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsPurchase As New ADODB.Recordset
    Dim rsViewSupplier As New ADODB.Recordset
    Dim rsViewPayment As New ADODB.Recordset
    
    Dim temSql As String
    Dim i As Integer
    

Private Sub btnExcel_Click()
    GridToExcel GridPurchase, "Purchase Bills", "From " & Format(dtpFrom.Value, "dd MMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub dtcPayment_Change()
    FillGrid
End Sub

Private Sub dtcSupplier_Change()
    FillGrid
End Sub

Private Sub dtpFrom_Change()
    FillGrid
End Sub

Private Sub dtpTo_Change()
    FillGrid
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    FillCombos
    FillGrid
End Sub

Private Sub FillCombos()
    With rsViewSupplier
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblDistrubutor order by DistributorName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With rsViewPayment
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPaymentMethod order by PaymentMethod"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcPayment
        Set .RowSource = rsViewPayment
        .ListField = "PaymentMethod"
        .BoundColumn = "PaymentMethodID"
    End With
    With dtcSupplier
        Set .RowSource = rsViewSupplier
        .ListField = "DistributorName"
        .BoundColumn = "DistributorID"
    End With
End Sub

Private Sub FillGrid()
    Dim Total As Double
    With GridPurchase
        .Rows = 1
        .Cols = 12
        
        .Row = 0
        
        .Col = 0
        .Text = "No"
        
        .Col = 1
        .Text = "GRN No"
        
        .Col = 2
        .Text = "Supplier"
        
        .Col = 3
        .Text = "Date"
        
        .Col = 4
        .Text = "Supplier ID"
        
        .Col = 5
        .Text = "Invoice Date"
        
        .Col = 6
        .Text = "Payment"
        
        .Col = 7
        .Text = "BillID"
        
        .Col = 8
        .Text = "Invoice No"
        
        .Col = 9
        .Text = "GRN/Return Date"
        
        .Col = 10
        .Text = "Value"
        
        .Col = 11
        .Text = "ID"
        
        .ColWidth(0) = 600
        .ColWidth(1) = 800
        .ColWidth(2) = 3000
        .ColWidth(3) = 1200
        .ColWidth(4) = 1400
        .ColWidth(5) = 1400
        .ColWidth(6) = 1400
        .ColWidth(7) = 1600
        .ColWidth(8) = 1400
        .ColWidth(9) = 1400
        .ColWidth(10) = 1600
        
        .ColWidth(11) = 1

        With rsPurchase
            temSql = "SELECT tblDistrubutor.DistributorName, tblRefillBill.InvoiceDate, tblRefillBill.InvoiceNo, tblRefillBill.Date, tblRefillBill.RefillBillID, tblPaymentMethod.PaymentMethod, tblRefillBill.NetPrice, tblRefillBill.DistributorID, tblPaymentMethod.PaymentMethod, tblRefillBill.Date "
            temSql = temSql & "FROM (tblRefillBill LEFT JOIN tblDistrubutor ON tblRefillBill.DistributorID = tblDistrubutor.DistributorID) LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID "
            temSql = temSql & "WHERE tblRefillBill.Cancelled = 0 AND tblRefillBill.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
            If IsNumeric(dtcSupplier.BoundText) = True Then
                temSql = temSql & " And tblRefillBill.DistributorID = " & Val(dtcSupplier.BoundText)
            End If
            If IsNumeric(dtcPayment.BoundText) = True Then
                temSql = temSql & " AND tblPaymentMethod.PaymentMethod= '" & dtcPayment.Text & "'"
            End If
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                .MoveLast
                .MoveFirst
                GridPurchase.Rows = .RecordCount + 1
                i = 1
                While .EOF = False
                    
                    If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(i, 0) = i '!RefillBillID
                    
                    If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(i, 1) = !RefillBillID
                    
                    If Not IsNull(!DistributorName) Then GridPurchase.TextMatrix(i, 2) = !DistributorName
                    
                    
                    
                    If Not IsNull(!Date) Then GridPurchase.TextMatrix(i, 3) = Format(!Date, "dd MMMM yyyy")

                    If Not IsNull(!DistributorID) Then GridPurchase.TextMatrix(i, 4) = i

                    If Not IsNull(!InvoiceDate) Then GridPurchase.TextMatrix(i, 5) = !InvoiceDate

                    If Not IsNull(!PaymentMethod) Then GridPurchase.TextMatrix(i, 6) = !PaymentMethod

                    If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(i, 7) = !RefillBillID
                    
                    If Not IsNull(!InvoiceNo) Then GridPurchase.TextMatrix(i, 8) = !InvoiceNo
                    
                    If Not IsNull(!Date) Then GridPurchase.TextMatrix(i, 9) = Format(!Date, "dd MMMM yyyy")
                    
                    If Not IsNull(!NetPrice) Then
                        GridPurchase.TextMatrix(i, 10) = Format(!NetPrice, "#,##0.00")
                        Total = Total + !NetPrice
                    End If
                    i = i + 1
                    .MoveNext
                Wend
            End If
        End With
        
        
        With rsPurchase
            temSql = "SELECT     dbo.tblDistrubutor.DistributorName, dbo.tblRefillBill.InvoiceDate, dbo.tblRefillBill.InvoiceNo, dbo.tblRefillBill.Date, dbo.tblRefillBill.RefillBillID, dbo.tblPaymentMethod.PaymentMethod, dbo.tblRefillBill.NetPrice, dbo.tblRefillBill.DistributorID, dbo.tblRefillBill.ReturnedDate, dbo.tblRefillBill.ReturnedTime, dbo.tblRefillBill.ReturnedValue "
            temSql = temSql & "FROM         dbo.tblDistrubutor RIGHT OUTER JOIN dbo.tblRefillBill LEFT OUTER JOIN dbo.tblPaymentMethod ON dbo.tblRefillBill.RepayPaymentMethodID = dbo.tblPaymentMethod.PaymentMethodID ON dbo.tblDistrubutor.DistributorID = dbo.tblRefillBill.DistributorID "
            temSql = temSql & "WHERE tblRefillBill.Returned = 1 AND tblRefillBill.ReturnedDate Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
            If IsNumeric(dtcSupplier.BoundText) = True Then
                temSql = temSql & " And tblRefillBill.DistributorID = " & Val(dtcSupplier.BoundText)
            End If
            If IsNumeric(dtcPayment.BoundText) = True Then
                temSql = temSql & " AND tblPaymentMethod.PaymentMethod= '" & dtcPayment.Text & "'"
            End If
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                .MoveLast
                .MoveFirst
                i = GridPurchase.Rows + 1
                GridPurchase.Rows = GridPurchase.Rows + 1 + .RecordCount + 1
                
                While .EOF = False
'                    If Not IsNull(!DistributorID) Then GridPurchase.TextMatrix(i, 0) = i - 2
'                    If Not IsNull(!DistributorName) Then GridPurchase.TextMatrix(i, 1) = !DistributorName
'                    If Not IsNull(!InvoiceNo) Then GridPurchase.TextMatrix(i, 2) = !InvoiceNo
'                    If Not IsNull(!InvoiceDate) Then GridPurchase.TextMatrix(i, 3) = !InvoiceDate
'                    If Not IsNull(!ReturnedDate) Then GridPurchase.TextMatrix(i, 4) = Format(!ReturnedDate, "dd MMMM yyyy")
'                    If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(i, 5) = !RefillBillID
'                    If Not IsNull(!PaymentMethod) Then GridPurchase.TextMatrix(i, 6) = !PaymentMethod
                    
                    If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(i, 0) = i '!RefillBillID
                    If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(i, 1) = !RefillBillID
                    
                    If Not IsNull(!DistributorName) Then GridPurchase.TextMatrix(i, 2) = !DistributorName
                    
                    If Not IsNull(!ReturnedDate) Then GridPurchase.TextMatrix(i, 3) = Format(!ReturnedDate, "dd MMMM yyyy")

                    If Not IsNull(!DistributorID) Then GridPurchase.TextMatrix(i, 4) = i

                    If Not IsNull(!InvoiceDate) Then GridPurchase.TextMatrix(i, 5) = !InvoiceDate

                    If Not IsNull(!PaymentMethod) Then GridPurchase.TextMatrix(i, 6) = !PaymentMethod

                    If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(i, 7) = !RefillBillID
                    
                    If Not IsNull(!InvoiceNo) Then GridPurchase.TextMatrix(i, 8) = !InvoiceNo
                    
                    If Not IsNull(!ReturnedDate) Then GridPurchase.TextMatrix(i, 9) = Format(!ReturnedDate, "dd MMMM yyyy")
                    
                    If Not IsNull(!ReturnedValue) Then
                        GridPurchase.TextMatrix(i, 10) = Format(0 - !ReturnedValue, "#,##0.00")
                        Total = Total - !ReturnedValue
                    End If
                    If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(i, 7) = !RefillBillID
                    i = i + 1
                    .MoveNext
                Wend
            End If
        End With
                
        
    End With
    txtTotal.Text = Format(Total, "#,##0.00")
End Sub

