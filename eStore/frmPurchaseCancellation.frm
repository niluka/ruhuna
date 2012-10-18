VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPurchaseCancellation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Cancellations"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15585
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   15585
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   8760
      Width           =   1815
   End
   Begin VB.TextBox txtDataEntry 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   9240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   255
      Caption         =   "&Cancel Purchase"
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
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   8493
      _Version        =   393216
      WordWrap        =   -1  'True
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   13800
      TabIndex        =   0
      Top             =   9240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   255
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
   Begin MSDataListLib.DataCombo dtcRePayment 
      Height          =   360
      Left            =   2400
      TabIndex        =   4
      Top             =   8280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcChecked 
      Height          =   360
      Left            =   2400
      TabIndex        =   24
      Top             =   7680
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcStaff 
      Height          =   360
      Left            =   2400
      TabIndex        =   25
      Top             =   7200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblSupplierID 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   29
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblDiscount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   11880
      TabIndex        =   28
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label21 
      Caption         =   "Checked by"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Received by"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Discount"
      Height          =   255
      Left            =   10320
      TabIndex        =   23
      Top             =   7860
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Net Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   22
      Top             =   8460
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Gross Total"
      Height          =   255
      Left            =   10320
      TabIndex        =   21
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label lblNetTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   20
      Top             =   8340
      Width           =   1935
   End
   Begin VB.Label lblGrossTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   11880
      TabIndex        =   19
      Top             =   7140
      Width           =   1935
   End
   Begin VB.Label lblChecked 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label9 
      Caption         =   "Checked By"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblReceived 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Received By"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblSupplier 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblRefillBillID 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Refill Bill ID"
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Time"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Re-Payment Method"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   8280
      Width           =   2055
   End
End
Attribute VB_Name = "frmPurchaseCancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
    Dim TemOrderBillID As Long
    Dim TemDistributorId As Long
    Dim TemDistributorOrderID As Long
    Dim EditingData As Boolean
    Dim TemContent(22) As String
    Dim CurrentRow As Integer
    Dim TemCellContent As String
    
    Dim rsRefillBill As New ADODB.Recordset
    
    Dim NewItem As New Item
    
    Dim rsStaff As New ADODB.Recordset
    Dim rsSPrice As New ADODB.Recordset
    Dim rsPPrice As New ADODB.Recordset
    Dim rsCC As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCategory As New ADODB.Recordset
    Dim rsCode As New ADODB.Recordset
    Dim rsBanks As New ADODB.Recordset
    Dim rsCreditCards As New ADODB.Recordset
    Dim rsCities As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsDistributor As New ADODB.Recordset
    
    Dim rsTemBatch As New ADODB.Recordset
    Dim rsTemOrder As New ADODB.Recordset
    Dim rsTemPrice As New ADODB.Recordset
    Dim rsTemDistributor As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    Dim rsTemOrderBill As New ADODB.Recordset
    Dim rsTemDistributorOrder As New ADODB.Recordset
    Dim rsTemRefill As New ADODB.Recordset
    Dim rsTemRefillBill As New ADODB.Recordset
    Dim rsTemCash As New ADODB.Recordset
    Dim rsTemCredit As New ADODB.Recordset
    Dim rsTemCheque As New ADODB.Recordset
    
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub dtcRePayment_Change()
    If dtcRePayment.Text = "" Or dtcRePayment.Text = "Cash" Or dtcRePayment.Text = "Credit" Or dtcRePayment.Text = "Cheque" Then
    
    Else
        MsgBox "Only Cash , Credit or Cheque Payments only"
        dtcRePayment.Text = Empty
        dtcRePayment.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    Call FillGrid
    Call GetValues
    dtcStaff.BoundText = UserID
    GridItem.RowHeight(0) = GridItem.RowHeight(0) * 3
End Sub

Private Sub GetValues()
    With rsTemRefillBill
        temSql = "SELECT tblRefillBill.Date,  tblRefillBill.PaymentMethodID,tblRefillBill.Time, tblRefillBill.RefillBillID, tblDistrubutor.DistributorName, tblDistrubutor.DistributorID, tblStaffReceivedBy.Name as RName, tblStaffCheckedBy.Name AS CName, tblRefillBill.Price, tblRefillBill.Discount, tblRefillBill.NetPrice " & _
                    "FROM ((tblRefillBill LEFT JOIN tblDistrubutor ON tblRefillBill.DistributorID = tblDistrubutor.DistributorID) LEFT JOIN tblStaff AS tblStaffCheckedBy ON tblRefillBill.CheckedStaffID = tblStaffCheckedBy.StaffID) LEFT JOIN tblStaff AS tblStaffReceivedBy ON tblRefillBill.StaffID = tblStaffReceivedBy.StaffID " & _
                    "Where (((tblRefillBill.RefillBillID) = " & TxRefillBillID & " ))"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!Date) Then lblDate.Caption = Format(!Date, LongDateFormat)
            If Not IsNull(!Price) Then lblGrossTotal.Caption = Format(!Price, "0.00")
            If Not IsNull(!Discount) Then lblDiscount.Caption = Format(!Discount, "0.00")
            If Not IsNull(!NetPrice) Then lblNetTotal.Caption = Format(!NetPrice, "0.00")
            If Not IsNull(!RefillBillID) Then lblRefillBillID.Caption = !RefillBillID
            If Not IsNull(!DistributorID) Then lblSupplierID.Caption = !DistributorID
            If Not IsNull(!Time) Then lblTime.Caption = !Time
            If Not IsNull(!DistributorName) Then lblSupplier.Caption = !DistributorName
            If Not IsNull(![RName]) Then lblReceived.Caption = ![RName]
            If Not IsNull(![CName]) Then lblChecked.Caption = ![CName]
            If Not IsNull(![PaymentMethodID]) Then dtcRePayment.BoundText = ![PaymentMethodID]
        End If
        .Close
    End With
End Sub

Private Sub FillCombos()
    With rsStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblstaff order by listedname"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcStaff
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With dtcChecked
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With rsCC
        If .State = 1 Then .Close
        temSql = "SELECT * from tblpaymentMethod " & _
                    "ORDER BY PaymentMethod"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcRePayment
        Set .RowSource = rsCC
        .ListField = "PaymentMethod"
        .BoundColumn = "PaymentMethodID"
    End With
End Sub
    
Private Sub FormatGrid()
    EditingData = False
    With GridItem
        .Cols = 22
        .Rows = 1
        .Row = 0
        .Col = 0
        .FixedCols = 0
        
        Dim i As Integer
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            Select Case i
                Case 0:     .Text = "No"
                            .ColWidth(i) = 450
                Case 1:     .Text = "Item"
                            .ColWidth(i) = 4000
                Case 5:     .Text = "Purchased"
                            .ColWidth(i) = 1200
                Case 6:     .Text = "Unit"
                            .ColWidth(i) = 1100
                Case 7:     .Text = "Free"
                            .ColWidth(i) = 1100
                Case 8:     .Text = "Unit"
                            .ColWidth(i) = 1100
                Case 9:     .Text = "Batch"
                            .ColWidth(i) = 1400
                Case 10:     .Text = "Pruchase Price Unit"
                            .ColWidth(i) = 1100
                Case 11:     .Text = "Slaes Price Per Unit"
                            .ColWidth(i) = 1100
                Case 19:    .ColWidth(i) = 1100
                            .Text = "Total Pruchase Value"
                Case 20:    .ColWidth(i) = 1600
                            .Text = "Exp. Date"

                Case Else:  .ColWidth(i) = 1
            End Select
        Next i
    End With
    '   0   No
    '   1   Item
    '   2   ItemID
    '   3   PackUnitID
    '   4   IssueUnitID
    '   5   PurchaseQuentity
    '   6   PUnit
    '   7   FreeQuentity
    '   8   PUnit
    '   9   Batch
    '   10  Purchase Price
    '   11  Sales Price
    '   12  Sales Margin
    '   13
    '   14
    '   15  IPurchased
    '   16  IFreePurchased
    '   17  IUnitsPerPack
    '   18  Display Price
    '   19  Actual Price
    '   20  DOM
    '   21  DOE
    EditingData = True
End Sub

Private Sub FillGrid()
    With rsRefillBill
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.Display, tblRefill.ItemID, tblRefill.RefillID, tblRefill.Price, tblItem.IssueUnitID, tblItem.PackUnitID, tblRefill.Amount, tblRefill.FreeAmount, tblBatch.Batch, tblBatch.BatchID, tblRefill.PPrice, tblRefill.SPrice, tblRefill.PackPPrice, tblRefill.DOE, tblIssueUnit.IssueUnit " & _
                    "FROM ((((tblRefillBill LEFT JOIN tblRefill ON tblRefillBill.RefillBillID = tblRefill.RefillBillID) LEFT JOIN tblItem ON tblRefill.ItemID = tblItem.ItemID) LEFT JOIN tblPackUnit ON tblItem.PackUnitID = tblPackUnit.PackUnitID) LEFT JOIN tblIssueUnit ON tblItem.IssueUnitID = tblIssueUnit.IssueUnitID) LEFT JOIN tblBatch ON tblRefill.BatchID = tblBatch.BatchID " & _
                    "WHERE (((tblRefillBill.RefillBillID)= " & TxRefillBillID & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                GridItem.Rows = GridItem.Rows + 1
                GridItem.Row = GridItem.Rows - 1
                GridItem.TextMatrix(GridItem.Row, 0) = GridItem.Row
                
                GridItem.TextMatrix(GridItem.Row, 1) = !Display
                GridItem.TextMatrix(GridItem.Row, 2) = !ItemID
                GridItem.TextMatrix(GridItem.Row, 3) = !PackUnitID
                GridItem.TextMatrix(GridItem.Row, 4) = !IssueUnitID
                GridItem.TextMatrix(GridItem.Row, 5) = !Amount
                GridItem.TextMatrix(GridItem.Row, 6) = !IssueUnit
                GridItem.TextMatrix(GridItem.Row, 7) = !FreeAmount
                GridItem.TextMatrix(GridItem.Row, 8) = !IssueUnit
                GridItem.TextMatrix(GridItem.Row, 9) = !Batch
                GridItem.TextMatrix(GridItem.Row, 10) = !PPrice
                GridItem.TextMatrix(GridItem.Row, 11) = !SPrice
                'GridItem.TextMatrix(GridItem.Row, 12) = !ItemID
                GridItem.TextMatrix(GridItem.Row, 13) = !BatchID
                'GridItem.TextMatrix(GridItem.Row, 14) = !IssueUnitID
                'GridItem.TextMatrix(GridItem.Row, 15) = !Display
                'GridItem.TextMatrix(GridItem.Row, 16) = !Display
                'GridItem.TextMatrix(GridItem.Row, 17) = !ItemID
                GridItem.TextMatrix(GridItem.Row, 18) = Format(!Price, "0.00")
                GridItem.TextMatrix(GridItem.Row, 19) = !Price
                If Not IsNull(!DOE) Then GridItem.TextMatrix(GridItem.Row, 20) = !DOE
                GridItem.TextMatrix(GridItem.Row, 21) = !RefillID
    '   0   No
    '   1   Item
    '   2   ItemID
    '   3   PackUnitID
    '   4   IssueUnitID
    '   5   PurchaseQuentity
    '   6   PUnit
    '   7   FreeQuentity
    '   8   PUnit
    '   9   Batch
    '   10  Purchase Price
    '   11  Sales Price
    '   12  Sales Margin
    '   13  BatchID
    '   14
    '   15  IPurchased
    '   16  IFreePurchased
    '   17  IUnitsPerPack
    '   18  Display Price
    '   19  Actual Price
    '   20  DOE
    '   21  RefillID
                .MoveNext
            Wend
        End If
    End With
End Sub

Private Sub bttnCancel_Click()
'On Error GoTo eh:
    If IsNumeric(dtcRePayment.BoundText) = False Then
        MsgBox "No Repayment Method"
        dtcRePayment.SetFocus
        Exit Sub
    End If
    
    With rsTemRefillBill
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblRefillBill where refillbillid = " & Val(lblRefillBillID.Caption)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Cancelled = True
            !CancelledUserID = UserID
            !CancelledDate = Date
            !CancelledTime = Now
            !CancelledCUserID = Val(dtcChecked.BoundText)
            !RepayPaymentMethodID = Val(dtcRePayment.BoundText)
            If dtcRePayment.Text = "Cash" Then
                !ReceivedCashID = ReceiveCash
            ElseIf dtcRePayment.Text = "Credit" Then
                !ReceivedCreditID = ReceiveCredit
            ElseIf dtcRePayment.Text = "Cheque" Then
                !ReceivedChequeID = ReceiveCheque
            End If
            !CancelledValue = Val(lblNetTotal.Caption)
            .Update
        End If
        .Close
    End With
    Dim i As Integer
    For i = 1 To GridItem.Rows - 1
        With rsTemRefill
            If .State = 1 Then .Close
            temSql = "Select * from tblRefill where RefillID = " & GridItem.TextMatrix(i, 21)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Returned = True
                !ReturnedUserID = UserID
                !ReturnedCUserID = UserID
                !ReturnedDate = Date
                !ReturnedTime = Now
                !ReturnedAmount = Val(GridItem.TextMatrix(i, 5)) + Val(GridItem.TextMatrix(i, 7))
                !ReturnedValue = Val(GridItem.TextMatrix(i, 19))
                If ConsumeStocks(UserStoreID, GridItem.TextMatrix(i, 13), (Val(GridItem.TextMatrix(i, 5)) + Val(GridItem.TextMatrix(i, 7)))) = True Then
                Else
                End If
                .Update
            End If
        End With
    Next
    If chkPrint.Value = 1 Then Call PrintBill
    MsgBox "Successfull Cancelled"
    Unload Me
    Exit Sub
eh:
    MsgBox "Error during Cancellation. Please contact Lakmedipro 077 3177874"
    Unload Me
End Sub
   
   
Private Function ConsumeStocks(ByVal IStoreIDValue As Long, ByVal BatchIDValue As Long, ByVal Quentity As Double) As Boolean
    Dim tr As Integer
    On Error GoTo eh
    ConsumeStocks = False
    With rsTemBatch
        If .State = 1 Then .Close
        temSql = "SELECT * from tblBatchstock where batchid = " & BatchIDValue & " AND StoreID = " & IStoreIDValue & " ORDER BY tblBatchstock.Stock DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount < 1 Then
            tr = MsgBox("There is no such drug batch", vbCritical, "Error")
            .Close
            Exit Function
        End If
        If !Stock < Quentity Then
            tr = MsgBox("There are no enough stocks in you store to transfer to another store", vbCritical, "No Enough Stocks")
            .Close
            Exit Function
        End If
        !Stock = !Stock - Quentity
        .Update
        .Close
    ConsumeStocks = True
    Exit Function

eh:
    If .State = 1 Then
        .CancelUpdate
        .Close
    End If
    tr = MsgBox("Could not deduct stocks from your store" & vbNewLine & Err.Description, vbCritical, "Error")
    Exit Function
    End With
End Function
   
Private Function ReceiveCredit() As Long
    With rsTemCredit
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblReceivedCredit"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = Val(dtcStaff.BoundText)
        !ReceivedDate = Date
        !ReceivedTime = Now
        !Price = Val(lblNetTotal.Caption)
        !StoreID = UserStoreID
        !ReceivedFromDistributorID = Val(lblSupplierID.Caption)
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        ReceiveCredit = !NewID
        .Close
    End With
End Function

Private Function ReceiveCash() As Long
    With rsTemCredit
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblReceivedCash"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = Val(dtcStaff.BoundText)
        !ReceivedDate = Date
        !ReceivedTime = Now
        !Price = Val(lblNetTotal.Caption)
        !StoreID = UserStoreID
        !ReceivedFromDistributorID = Val(lblSupplierID.Caption)
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        ReceiveCash = !NewID
        .Close
    End With
End Function

Private Function ReceiveCheque() As Long
    With rsTemCredit
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblReceivedCheque"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = Val(dtcStaff.BoundText)
        !ReceivedDate = Date
        !ReceivedTime = Now
        !Price = Val(lblNetTotal.Caption)
        !StoreID = UserStoreID
        !ReceivedFromDistributorID = Val(lblSupplierID.Caption)
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        ReceiveCheque = !NewID
        .Close
    End With
End Function

Private Sub PrintBill()
Dim CsetPrinter As New cSetDfltPrinter
'On Error GoTo eh

    Dim TemResponce As Long
    Dim RetVal As Integer
    CsetPrinter.SetPrinterAsDefault (BillPrinterName)
    RetVal = SelectForm(BillPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1

            Dim i As Integer
            Dim Tab1 As Integer
            Dim Tab2 As Integer
            Dim Tab3 As Integer
            Dim Tab4 As Integer
            Dim Tab5 As Integer
            Dim Tab6 As Integer
            Dim Tab7 As Integer
            Dim Tab8 As Integer
            Dim Tab9 As Integer
            
            Tab1 = 4
            Tab2 = 15
            Tab3 = 36
            Tab4 = 20
            Tab5 = 50
            Tab6 = 55
            Tab7 = 70
            Tab8 = 23
            Tab9 = 65
            With Printer
                '.TrackDefault = False
                '.PaperBin = vbPRBNTractor
                .FontSize = 12
                .Font = "Lucida Console"
                Printer.Print
                Printer.Print Tab(Tab8); "Purchase Cancellation - " & dtcRePayment.Text
                .FontSize = 12
                .Font = "Lucida Console"
                Printer.Print Tab(4); "RUHUNU HOSPITAL (PVT) LTD "
                .FontSize = 10
                .Font = "Lucida Console"
                Printer.Print
                .FontSize = 10
                .Font = "Lucida Console"
                Dim TemString As String
                Printer.Print Tab(Tab1); "GRN No      : "; lblRefillBillID.Caption & " " & TemString; "       Date : "; Format(Date, "dd MM yy"); Tab(Tab6); "Time : "; Time
                Printer.Print Tab(Tab1); "Supplier    : "; lblSupplier.Caption
                Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
                Printer.Print Tab(Tab1); "Item Name"; Tab(Tab3 + 5); "Qty"; Tab(Tab5); Right(Space(12) & "Price", 9); Tab(Tab9); Right(Space(12) & "Value", 13)
                Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
            End With
            Tab1 = 4
            Tab2 = 15
            Tab3 = 36
            Tab4 = 20
            Tab5 = 50
            Tab6 = 55
            Tab7 = 68
            Tab9 = 65
            With GridItem
                For i = 1 To .Rows - 1
                    Printer.Print Tab(Tab1); Left(.TextMatrix(i, 1), 30);
                    Printer.Print Tab(Tab3); Right(Space(10) & (Val((.TextMatrix(i, 5))) + Val((.TextMatrix(i, 7)))), 10);
                    Printer.Print Tab(Tab5); Right(Space(10) & Format(.TextMatrix(i, 10), "0.00"), 10);
                    Printer.Print Tab(Tab7); Right(Space(10) & Format(.TextMatrix(i, 18), "0.00"), 10)
                Next i
    '   0   No
    '   1   Item
    '   2   ItemID
    '   3   PackUnitID
    '   4   IssueUnitID
    '   5   PurchaseQuentity
    '   6   PUnit
    '   7   FreeQuentity
    '   8   PUnit
    '   9   Batch
    '   10  Purchase Price
    '   11  Sales Price
    '   12  Sales Margin
    '   13  BatchID
    '   14
    '   15  IPurchased
    '   16  IFreePurchased
    '   17  IUnitsPerPack
    '   18  Display Price
    '   19  Actual Price
    '   20  DOE
    '   21  RefillID
            
            End With
            With Printer
                Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
                Dim NewTab1 As Integer
                Dim NewTab2 As Integer
                Dim NewTab3 As Integer
                NewTab1 = 42
                NewTab2 = 68
                Printer.Print
                Printer.Print Tab(NewTab1); "Total Return "; Tab(NewTab2); Right((Space(9)) & lblGrossTotal.Caption, 10)
                Printer.Print Tab(NewTab1); "Plus Discount    "; Tab(NewTab2); Right((Space(9)) & lblDiscount.Caption, 10)
                Printer.Print Tab(NewTab1); "Net Return   "; Tab(NewTab2); Right((Space(9)) & lblNetTotal.Caption, 10)
                Printer.Print
                Printer.Print
                Printer.Print Tab(Tab1); "Operate by "; UserName  ' ; Tab(Tab5); "Issued by "; dtcIssueStaff
                Printer.Print Tab(Tab1); "No more returns or cancellations of this GRN is possible"
                .EndDoc
            End With


        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select



    
Exit Sub

eh:
    MsgBox "Printer Error"

End Sub
