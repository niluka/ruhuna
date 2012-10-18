VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDisplayPurchaseReturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Purchase Return"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17055
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
   ScaleHeight     =   9945
   ScaleWidth      =   17055
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDataEntry 
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtValue 
      Height          =   360
      Left            =   14880
      TabIndex        =   45
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtRate 
      Enabled         =   0   'False
      Height          =   360
      Left            =   13920
      TabIndex        =   43
      Top             =   2520
      Width           =   855
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   16080
      TabIndex        =   39
      Top             =   2520
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add"
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
   Begin VB.TextBox txtReturn 
      Height          =   360
      Left            =   6480
      TabIndex        =   38
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   9240
      Width           =   1815
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   375
      Left            =   10800
      TabIndex        =   0
      Top             =   9240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Return"
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
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   7223
      _Version        =   393216
      WordWrap        =   -1  'True
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   14160
      TabIndex        =   1
      Top             =   9240
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSDataListLib.DataCombo dtcRePayment 
      Height          =   360
      Left            =   2400
      TabIndex        =   4
      Top             =   8160
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
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
   Begin VB.Label lblGrossTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   14760
      TabIndex        =   52
      Top             =   660
      Width           =   1935
   End
   Begin VB.Label lblNetTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   14760
      TabIndex        =   51
      Top             =   1620
      Width           =   1935
   End
   Begin VB.Label Label20 
      Caption         =   "Gross Total"
      Height          =   255
      Left            =   12360
      TabIndex        =   50
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label19 
      Caption         =   "Net Total"
      Height          =   255
      Left            =   12360
      TabIndex        =   49
      Top             =   1620
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "Discount"
      Height          =   255
      Left            =   12360
      TabIndex        =   48
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblDiscount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   14760
      TabIndex        =   47
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblDiscountPercent 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11400
      TabIndex        =   46
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "Value"
      Height          =   255
      Left            =   14880
      TabIndex        =   44
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Rate"
      Height          =   255
      Left            =   13920
      TabIndex        =   42
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Free"
      Height          =   255
      Left            =   5040
      TabIndex        =   41
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblFQty 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5040
      TabIndex        =   40
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblIUnit2 
      Height          =   375
      Left            =   7920
      TabIndex        =   37
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "Returned"
      Height          =   255
      Left            =   6480
      TabIndex        =   36
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   15360
      TabIndex        =   35
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblIUnit1 
      Height          =   375
      Left            =   7440
      TabIndex        =   34
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Purchased"
      Height          =   255
      Left            =   3600
      TabIndex        =   33
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3600
      TabIndex        =   32
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblItem 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lblSupplierID 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   29
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblRDiscount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   13200
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
      Caption         =   "Return Discount"
      Height          =   255
      Left            =   10800
      TabIndex        =   23
      Top             =   7860
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Net Return"
      Height          =   255
      Left            =   10800
      TabIndex        =   22
      Top             =   8340
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Return Total"
      Height          =   255
      Left            =   10800
      TabIndex        =   21
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label lblRNetTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   13200
      TabIndex        =   20
      Top             =   8220
      Width           =   1935
   End
   Begin VB.Label lblRGrossTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   375
      Left            =   13200
      TabIndex        =   19
      Top             =   7260
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
      Left            =   13560
      TabIndex        =   12
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Refill Bill ID"
      Height          =   255
      Left            =   12360
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9600
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Time"
      Height          =   255
      Left            =   8640
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Re-Payment Method"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   8160
      Width           =   2055
   End
End
Attribute VB_Name = "frmDisplayPurchaseReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    
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
    GridItem.RowHeight(0) = GridItem.RowHeight(0) * 3
    dtcStaff.BoundText = UserID
End Sub

Private Sub GetValues()
    With rsTemRefillBill
        temSQL = "SELECT tblRefillBill.Date,  tblRefillBill.PaymentMethodID ,tblRefillBill.Time,  tblRefillBill.DiscountPercent, tblRefillBill.RefillBillID, tblDistrubutor.DistributorName, tblDistrubutor.DistributorID, tblStaffReceivedBy.Name, tblStaffCheckedBy.Name, tblRefillBill.Price, tblRefillBill.Discount, tblRefillBill.NetPrice " & _
                    "FROM ((tblRefillBill LEFT JOIN tblDistrubutor ON tblRefillBill.DistributorID = tblDistrubutor.DistributorID) LEFT JOIN tblStaff AS tblStaffCheckedBy ON tblRefillBill.CheckedStaffID = tblStaffCheckedBy.StaffID) LEFT JOIN tblStaff AS tblStaffReceivedBy ON tblRefillBill.StaffID = tblStaffReceivedBy.StaffID " & _
                    "Where (((tblRefillBill.RefillBillID) = " & TxRefillBillID & " ))"
        If .State = 1 Then .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!Date) Then lblDate.Caption = Format(!Date, LongDateFormat)
            If Not IsNull(!price) Then lblGrossTotal.Caption = Format(!price, "#,##0.00")
            If Not IsNull(!Discount) Then lblDiscount.Caption = Format(!Discount, "#,##0.00")
            If Not IsNull(!NetPrice) Then lblNetTotal.Caption = Format(!NetPrice, "#,##0.00")
            If Not IsNull(!RefillBillID) Then lblRefillBillID.Caption = !RefillBillID
            If Not IsNull(!DistributorID) Then lblSupplierID.Caption = !DistributorID
            If Not IsNull(!Time) Then lblTime.Caption = !Time
            If Not IsNull(!DistributorName) Then lblSupplier.Caption = !DistributorName
            If Not IsNull(![tblStaffReceivedBy.Name]) Then lblReceived.Caption = ![tblStaffReceivedBy.Name]
            If Not IsNull(![tblStaffCheckedBy.Name]) Then lblChecked.Caption = ![tblStaffCheckedBy.Name]
            If Not IsNull(!DiscountPercent) Then lblDiscountPercent.Caption = !DiscountPercent
            If Not IsNull(!PaymentMethodID) Then dtcRePayment.BoundText = !PaymentMethodID
        End If
        .Close
    End With
End Sub

Private Sub FillCombos()
    With rsStaff
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblstaff order by listedname"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
        temSQL = "SELECT * from tblpaymentMethod " & _
                    "ORDER BY PaymentMethod"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
                            .ColWidth(i) = 400
                Case 1:     .Text = "Item"
                            .ColWidth(i) = 3800
                Case 5:     .Text = "Purchased"
                            .ColWidth(i) = 1200
                Case 6:     .Text = "Unit"
                            .ColWidth(i) = 1000
                Case 7:     .Text = "Free"
                            .ColWidth(i) = 800
                Case 8:     .Text = "Unit"
                            .ColWidth(i) = 1100
                Case 9:     .Text = "Batch"
                            .ColWidth(i) = 1200
                Case 10:     .Text = "Pruchase Price Unit"
                            .ColWidth(i) = 1100
                Case 11:     .Text = "Slaes Price Per Unit"
                            .ColWidth(i) = 1000
                Case 14:     .Text = "Returned"
                            .ColWidth(i) = 900
                Case 15:     .Text = "Returned Value"
                            .ColWidth(i) = 1200
                Case 19:    .ColWidth(i) = 1200
                            .Text = "Total Pruchase Value"
                Case 20:    .ColWidth(i) = 1700
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
    '   13  BatchID
    '   14  Return AMount
    '   15  Returned Value
    '   16  IFreePurchased
    '   17  IUnitsPerPack
    '   18  Display Price
    '   19  Actual Price
    '   20  DOE
    '   21  RefillID
    EditingData = True
End Sub

Private Sub FillGrid()
    With rsRefillBill
        If .State = 1 Then .Close
        temSQL = "SELECT tblItem.Display, tblRefill.ItemID, tblRefill.RefillID, tblRefill.Price, tblItem.IssueUnitID, tblItem.PackUnitID, tblRefill.Amount, tblRefill.FreeAmount, tblBatch.Batch, tblBatch.BatchID, tblRefill.PPrice, tblRefill.SPrice, tblRefill.PackPPrice, tblRefill.DOE, tblIssueUnit.IssueUnit " & _
                    "FROM ((((tblRefillBill LEFT JOIN tblRefill ON tblRefillBill.RefillBillID = tblRefill.RefillBillID) LEFT JOIN tblItem ON tblRefill.ItemID = tblItem.ItemID) LEFT JOIN tblPackUnit ON tblItem.PackUnitID = tblPackUnit.PackUnitID) LEFT JOIN tblIssueUnit ON tblItem.IssueUnitID = tblIssueUnit.IssueUnitID) LEFT JOIN tblBatch ON tblRefill.BatchID = tblBatch.BatchID " & _
                    "WHERE (((tblRefillBill.RefillBillID)= " & TxRefillBillID & "))"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
                If Not IsNull(!Batch) Then GridItem.TextMatrix(GridItem.Row, 9) = !Batch
                GridItem.TextMatrix(GridItem.Row, 10) = !PPrice
                GridItem.TextMatrix(GridItem.Row, 11) = !SPrice
                'GridItem.TextMatrix(GridItem.Row, 12) = !ItemID
                If Not IsNull(!Batch) Then GridItem.TextMatrix(GridItem.Row, 13) = !BatchID
                'GridItem.TextMatrix(GridItem.Row, 14) = !IssueUnitID
                'GridItem.TextMatrix(GridItem.Row, 15) = !Display
                'GridItem.TextMatrix(GridItem.Row, 16) = !Display
                'GridItem.TextMatrix(GridItem.Row, 17) = !ItemID
                GridItem.TextMatrix(GridItem.Row, 18) = Format(!price, "#,##0.00")
                GridItem.TextMatrix(GridItem.Row, 19) = !price
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
    '   14  Return AMount
    '   15  Returned Value
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
    If IsNumeric(dtcRePayment.BoundText) Then
        MsgBox "Please select a Re-payment Method"
        dtcRePayment.SetFocus
        Exit Sub
    End If
    With rsTemRefillBill
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblRefillBill where refillbillid = " & Val(lblRefillBillID.Caption)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Returned = True
            !ReturnedUserID = UserID
            !ReturnedDate = Date
            !ReturnedTime = Time
            !ReturnedCUserID = Val(dtcChecked.BoundText)
            !RepayPaymentMethodID = Val(dtcRePayment.BoundText)
            If dtcRePayment.Text = "Cash" Then
                !ReceivedCashID = ReceiveCash
            ElseIf dtcRePayment.Text = "Credit" Then
                !ReceivedCreditID = ReceiveCredit
            ElseIf dtcRePayment.Text = "Cheque" Then
                !ReceivedChequeID = ReceiveCheque
            End If
            !ReturnedValue = Val(lblRNetTotal.Caption)
            .Update
        End If
        .Close
    End With
    Dim i As Integer
    For i = 1 To GridItem.Rows - 1
        If Val(GridItem.TextMatrix(i, 14)) <> 0 Then
            With rsTemRefill
                If .State = 1 Then .CancelUpdate
                If .State = 1 Then .Close
                temSQL = "Select * from tblRefill where RefillID = " & GridItem.TextMatrix(i, 21)
                .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !Returned = True
                    !ReturnedUserID = UserID
                    !ReturnedCUserID = UserID
                    !ReturnedDate = Date
                    !ReturnedTime = Time
                    !ReturnedAmount = Val(GridItem.TextMatrix(i, 14))
                    !ReturnedValue = Val(GridItem.TextMatrix(i, 15))
                    If ConsumeStocks(UserStoreID, GridItem.TextMatrix(i, 13), Val(GridItem.TextMatrix(i, 14))) = True Then
                    Else
                    End If
                    .Update
                End If
            End With
        End If
    Next
    If chkPrint.Value = 1 Then PrintBill
    MsgBox "Successfull Cancelled"
    Unload Me
    Exit Sub
eh:
    MsgBox "Error during Cancellation. Please contact Lakmedipro 077 3177874"
    Unload Me
End Sub
   
Private Sub PrintBill()
    Dim CSetPrinter As New cSetDfltPrinter
    
'On Error GoTo eh:

    Dim TemResponce As Long
    Dim RetVal As Integer
    CSetPrinter.SetPrinterAsDefault (BillPrinterName)
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
 '               .TrackDefault = False
'                .PaperBin = vbPRBNTractor
                .FontSize = 12
                .Font = "Lucida Console"
                Printer.Print
                Printer.Print Tab(Tab8 + 10); "Purchase Return - " & dtcRePayment.Text
                .FontSize = 12
                .Font = "Lucida Console"
                Printer.Print Tab(4); "RUHUNU HOSPITAL (PVT) LTD "
                .FontSize = 10
                .Font = "Lucida Console"
                Printer.Print
                .FontSize = 10
                .Font = "Lucida Console"
                Dim TemString As String
                Printer.Print Tab(Tab1); "GRN No      : "; lblRefillBillID.Caption & "-" & TemString; "       Date : "; Format(Date, "dd MM yy"); Tab(Tab6); "Time : "; Time
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
            Tab7 = 70
            Tab9 = 65
            With GridItem
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, 14)) <> 0 Then
                        Printer.Print Tab(Tab1); Left(.TextMatrix(i, 1), 30);
                        Printer.Print Tab(Tab3); Right(Space(10) & (.TextMatrix(i, 14)), 10);
                        Printer.Print Tab(Tab5); Right(Space(12) & Format(.TextMatrix(i, 10), "0.00"), 9);
                        Printer.Print Tab(Tab7); Right(Space(12) & Format(.TextMatrix(i, 15), "0.00"), 8)
                    End If
                Next i
            End With
            With Printer
                Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
                Dim NewTab1 As Integer
                Dim NewTab2 As Integer
                Dim NewTab3 As Integer
                NewTab1 = 40
                NewTab2 = 68
                Printer.Print
                Printer.Print Tab(NewTab1); "Total Return    "; Tab(NewTab2); Right((Space(9)) & lblRGrossTotal.Caption, 10)
                Printer.Print Tab(NewTab1); "Return Discount "; Tab(NewTab2); Right((Space(9)) & lblRDiscount.Caption, 10)
                Printer.Print Tab(NewTab1); "Net Return      "; Tab(NewTab2); Right((Space(9)) & lblRNetTotal.Caption, 10)
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


   
Private Function ConsumeStocks(ByVal IStoreIDValue As Long, ByVal BatchIDValue As Long, ByVal Quentity As Double) As Boolean
    Dim tr As Integer
    On Error GoTo eh
    ConsumeStocks = False
    With rsTemBatch
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblBatchstock where batchid = " & BatchIDValue & " AND StoreID = " & IStoreIDValue & " ORDER BY tblBatchstock.Stock DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
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
        temSQL = "SELECT * FROM tblReceivedCredit"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = dtcStaff.BoundText
        !ReceivedDate = Date
        !ReceivedTime = Time
        !price = Val(lblRNetTotal.Caption)
        !StoreID = UserStoreID
        !ReceivedFromDistributorID = Val(lblSupplierID.Caption)
        .Update
        ReceiveCredit = !ReceivedCreditID
        .Close
    End With
End Function

Private Function ReceiveCash() As Long
    With rsTemCredit
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblReceivedCash"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = dtcStaff.BoundText
        !ReceivedDate = Date
        !ReceivedTime = Time
        !price = Val(lblRNetTotal.Caption)
        !StoreID = UserStoreID
        !ReceivedFromDistributorID = Val(lblSupplierID.Caption)
        .Update
        ReceiveCash = !ReceivedCashID
        .Close
    End With
End Function

Private Function ReceiveCheque() As Long
    With rsTemCredit
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblReceivedCheque"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReceivedSTaffID = dtcStaff.BoundText
        !ReceivedDate = Date
        !ReceivedTime = Time
        !price = Val(lblRNetTotal.Caption)
        !StoreID = UserStoreID
        !ReceivedFromDistributorID = Val(lblSupplierID.Caption)
        .Update
        ReceiveCheque = !ReceivedChequeID
        .Close
    End With
End Function


Private Sub GridItem_DblClick()
    Dim i As Integer
    With GridItem
        If .Rows < 2 Then Exit Sub
        If .Row < 1 Then Exit Sub
        i = .Row
        If Not IsNumeric(.TextMatrix(i, 21)) Then Exit Sub
        NewItem.ID = .TextMatrix(i, 2)
        lblItem.Caption = .TextMatrix(i, 1)
        lblIUnit1.Caption = .TextMatrix(i, 6)
        lblIUnit1.Caption = .TextMatrix(i, 6)
        lblQty.Caption = .TextMatrix(i, 5)
        lblFQty.Caption = .TextMatrix(i, 7)
        lblRow.Caption = i
        txtRate.Text = .TextMatrix(i, 10)
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
    '   13  BatchID
    '   14  Return AMount
    '   15  Returned Value
    '   16  IFreePurchased
    '   17  IUnitsPerPack
    '   18  Display Price
    '   19  Actual Price
    '   20  DOE
    '   21  RefillID
    
End Sub

Private Sub bttnAdd_Click()
    Dim i As Integer
    If Val(txtReturn.Text) > Val(lblQty.Caption) + Val(lblFQty.Caption) Then
        MsgBox "Return can't be more than Purchase"
        txtReturn.SetFocus
        SendKeys "{Home}+{end}"
        Exit Sub
    End If
    If Not IsNumeric(lblRow.Caption) Then
        MsgBox "Nothing to return"
        Exit Sub
    End If
    GridItem.TextMatrix(Val(lblRow.Caption), 14) = Val(txtReturn.Text)
    GridItem.TextMatrix(Val(lblRow.Caption), 15) = Val(txtValue.Text)
    Call CalculateTotals
    Call ClearAddValues
End Sub


Private Sub ClearAddValues()
        lblItem.Caption = Empty
        lblIUnit1.Caption = Empty
        lblIUnit1.Caption = Empty
        lblQty.Caption = Empty
        lblFQty.Caption = Empty
        lblRow.Caption = Empty
        txtRate.Text = Empty
        txtReturn.Text = Empty
        txtValue.Text = Empty
End Sub

Private Sub txtRate_Change()
    Call CalculateReturnValue
End Sub

Private Sub txtReturn_Change()
    Call CalculateReturnValue
End Sub

Private Sub CalculateReturnValue()
    txtValue.Text = Format(Val(txtReturn.Text) * Val(txtRate.Text), "0.00")
End Sub

Private Sub CalculateTotals()
    Dim i As Integer
    Dim ReturnTotal As Double
    Dim ReturnDiscount As Double
    With GridItem
        For i = 1 To .Rows - 1
            ReturnTotal = ReturnTotal + Val(.TextMatrix(i, 15))
        Next
    End With
    lblRGrossTotal.Caption = Format(ReturnTotal, "#,##0.00")
    ReturnDiscount = ReturnTotal * Val(lblDiscountPercent.Caption)
    lblRDiscount.Caption = Format(ReturnDiscount, "#,##0.00")
    lblRNetTotal.Caption = Format(ReturnTotal - ReturnDiscount, "#,##0.00")
End Sub
