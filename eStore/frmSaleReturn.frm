VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSaleReturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Returns"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13785
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
   ScaleHeight     =   8760
   ScaleWidth      =   13785
   WindowState     =   2  'Maximized
   Begin MSDataListLib.DataCombo dtcRepaymentMethod 
      Height          =   360
      Left            =   5880
      TabIndex        =   61
      Top             =   7200
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtItemCost 
      Height          =   375
      Left            =   5520
      TabIndex        =   60
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCostRate 
      Height          =   375
      Left            =   4320
      TabIndex        =   59
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCustomerID 
      Height          =   375
      Left            =   11280
      TabIndex        =   58
      Top             =   4080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   10080
      TabIndex        =   53
      Top             =   7080
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox txtNetReturn 
      Height          =   375
      Left            =   1680
      TabIndex        =   51
      Top             =   8160
      Width           =   2295
   End
   Begin VB.TextBox txtReturnDiscount 
      Height          =   375
      Left            =   1680
      TabIndex        =   49
      Top             =   7680
      Width           =   2295
   End
   Begin VB.TextBox txtRow 
      Height          =   495
      Left            =   4320
      TabIndex        =   48
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRRate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   8640
      TabIndex        =   46
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtRValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   8640
      TabIndex        =   40
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtReturnValue 
      Height          =   375
      Left            =   1680
      TabIndex        =   32
      Top             =   7200
      Width           =   2295
   End
   Begin btButtonEx.ButtonEx bttnUpdate 
      Height          =   495
      Left            =   10080
      TabIndex        =   30
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   255
      Caption         =   "&Update"
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   11400
      TabIndex        =   29
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   495
      Left            =   12360
      TabIndex        =   28
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   255
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
   Begin VB.TextBox txtRQty 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   8640
      TabIndex        =   26
      Top             =   3360
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   4048
      _Version        =   393216
      SelectionMode   =   1
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
   Begin MSDataListLib.DataCombo dtcCheckedStaff 
      Height          =   360
      Left            =   11040
      TabIndex        =   54
      Top             =   2280
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcIssueStaff 
      Height          =   360
      Left            =   11040
      TabIndex        =   55
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label26 
      Caption         =   "Repayment Method"
      Height          =   255
      Left            =   4080
      TabIndex        =   62
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label Label25 
      Caption         =   "Checked By"
      Height          =   255
      Left            =   9360
      TabIndex        =   57
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label24 
      Caption         =   "Issued By"
      Height          =   255
      Left            =   9360
      TabIndex        =   56
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label17 
      Caption         =   "Net Return"
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "Discount"
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label21 
      Caption         =   "Return Rate"
      Height          =   255
      Left            =   6960
      TabIndex        =   47
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label20 
      Caption         =   "Issued Rate"
      Height          =   255
      Left            =   240
      TabIndex        =   45
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblIRate 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   44
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label19 
      Caption         =   "Issued Value"
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lblIValue 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   42
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label11 
      Caption         =   "Return Value"
      Height          =   255
      Left            =   6960
      TabIndex        =   41
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label23 
      Caption         =   "Customer Type"
      Height          =   255
      Left            =   9360
      TabIndex        =   39
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label22 
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   9360
      TabIndex        =   38
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblCustomer 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11040
      TabIndex        =   37
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11040
      TabIndex        =   36
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblBHT 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11040
      TabIndex        =   35
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label18 
      Caption         =   "BHT No."
      Height          =   255
      Left            =   9360
      TabIndex        =   34
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label16 
      Caption         =   "Return Total"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Discount %"
      Height          =   255
      Left            =   4800
      TabIndex        =   31
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblIQty 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label12 
      Caption         =   "Issued Qty"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblItem 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      Top             =   2880
      Width           =   4575
   End
   Begin VB.Label Label10 
      Caption         =   "Item"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblDiscountPercent 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblPaymentMethod 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label lblNetTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblDiscount 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblGTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblBillID 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label lblCheckedBy 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lblUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblTime 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblDate 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Payment Method"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Net Total"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Discount"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Gross Total"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Bill ID"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Checked by"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "User"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Time"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblIUnit1 
      Height          =   375
      Left            =   10560
      TabIndex        =   27
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Return Qty"
      Height          =   255
      Left            =   6960
      TabIndex        =   25
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblIUnit 
      Height          =   375
      Left            =   4440
      TabIndex        =   24
      Top             =   3360
      Width           =   1935
   End
End
Attribute VB_Name = "frmSaleReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsSaleBill As New ADODB.Recordset
    Dim rsSale As New ADODB.Recordset
    Dim rsReturnBill As New ADODB.Recordset
    Dim rsReturn As New ADODB.Recordset
    Dim rsBatchStock As New ADODB.Recordset
    Dim rsIssuedCash As New ADODB.Recordset
    Dim rsIssuedCredit As New ADODB.Recordset
    Dim rsIssuedWoucher As New ADODB.Recordset
    Dim rsIssuedOther As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsViewPaymentMethod As New ADODB.Recordset
    
    Dim temIssueCashID As Long
    Dim temIssueCreditID As Long
    Dim temIssueWoucherID As Long
    Dim temIssueOtherID As Long

    Dim temSql As String
    Dim temReturnBillID As Long
    Dim CsetPrinter As New cSetDfltPrinter
    Dim rsViewStaff As New ADODB.Recordset
    Dim rsCredit As New ADODB.Recordset
    Dim rsTemCustomer As New ADODB.Recordset

Private Sub FillCombos()
    With rsViewStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblstaff order by listedname"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcIssueStaff
        Set .RowSource = rsViewStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With dtcCheckedStaff
        Set .RowSource = rsViewStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With rsViewPaymentMethod
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPaymentMethod order by PaymentMethod"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcRepaymentMethod
        Set .RowSource = rsViewPaymentMethod
        .ListField = "PaymentMethod"
        .BoundColumn = "PaymentMethodID"
    End With
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnAdd_Click()
    If CanAdd = False Then Exit Sub
    Dim i As Integer
    i = Val(txtRow.Text)
    With GridItem
        .TextMatrix(i, 12) = Val(txtRValue.Text)
        .TextMatrix(i, 9) = Val(txtRRate.Text)
        .TextMatrix(i, 11) = Val(txtRQty.Text)
    End With
    Call CalculateReturnValue
    Call ClearAddValues
'   0   No
'   1   Item
'   2   Batch
'   3   Rate
'   4   Amount
'   5   Price
'   6   ItemID
'   7   BatchID
'   8   AMount
'   9   Rate
'   10  IUnit
'   11  Return Qty
'   12  Return Value
'   13  SaleID

End Sub

Private Sub ClearAddValues()
    txtRQty.Text = Empty
    txtRRate.Text = Empty
    txtRValue.Text = Empty
    lblIQty.Caption = Empty
    lblIRate.Caption = Empty
    lblIValue.Caption = Empty
    lblIUnit.Caption = Empty
    lblIUnit1.Caption = Empty
    lblItem.Caption = Empty
End Sub

Private Sub CalculateReturnValue()
    With GridItem
    Dim i As Integer
    Dim RValue As Double
    For i = 1 To .Rows - 1
        RValue = RValue + Val(.TextMatrix(i, 12))
    Next
    End With
    txtReturnValue.Text = Format(RValue, "0.00")
End Sub

Private Function CanAdd() As Boolean
    CanAdd = False
    Dim tr As Integer
    If Val(txtRQty.Text) > Val(lblIQty.Caption) Then
        tr = MsgBox("You can't return more than the issued amount", vbCritical, "Error")
        txtRQty.SetFocus
        SendKeys "{home}+{end}"
        Exit Function
    End If
    If Val(txtRRate.Text) < Val(lblIRate.Caption) Then
        tr = MsgBox("You can't return items at a lower rate than the issued rate", vbCritical, "Error")
        txtRRate.SetFocus
        Exit Function
    End If
    CanAdd = True
End Function

Private Sub bttnUpdate_Click()
    Dim tr As Integer
    Dim TemCost As Double
    Dim TotalCost As Double
    Dim temBilledBHTID As Long
    Dim temBilledOutPatientID     As Long
    Dim temBilledStaffID   As Long
    Dim temBilledUnitID As Long
    
    tr = MsgBox("Are you sure you want to Return this bill?", vbYesNo, "Return")
    
    If tr = vbNo Then
        Exit Sub
    End If
    
    
    If Val(txtReturnValue.Text) <= 0 Then
        tr = MsgBox("There are no items to return", vbCritical, "No Items")
        Exit Sub
    End If
    With rsSaleBill
        If .State = 1 Then .Close
        temSql = "SELECT * from tblSaleBill where SaleBillID = " & TxSaleBillID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Returned = True
            temBilledBHTID = !BilledBHTID
            temBilledOutPatientID = !BilledOutPatientID
            temBilledStaffID = !BilledStaffID
            temBilledUnitID = !BilledUnitID
            !ReturnedUserID = UserID
            !ReturnedDate = Date
            !ReturnedTime = Now
            !ReturnedValue = Val(txtNetReturn.Text)
            .Update
        End If
        .Close
    End With
    With rsReturnBill
        If .State = 1 Then .Close
        temSql = "SELECT * from tblReturnBill"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !SaleBillID = Val(lblBillID.Caption)
        !Date = Date
        !Time = Now
        !StaffID = UserID
        !StoreID = UserStoreID
        !Price = Val(txtReturnValue.Text)
        !Discount = Val(txtReturnDiscount.Text)
        !DiscountPercent = (Val(txtReturnDiscount.Text) * 100) / Val(txtReturnValue.Text)
        !NetPrice = Val(txtNetReturn.Text)
        !BilledBHTID = temBilledBHTID
        !BilledOutPatientID = temBilledOutPatientID
        !BilledStaffID = temBilledStaffID
        !BilledUnitID = temBilledUnitID
        !PaymentMethod = dtcRepaymentMethod.Text
        !PaymentMethodID = Val(dtcRepaymentMethod.BoundText)
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        temReturnBillID = !NewID
        .Close
    End With
    With GridItem
        If rsReturn.State = 1 Then rsReturn.Close
        temSql = "SELECT * from tblReturn"
        
        rsReturn.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        Dim i As Integer
        
        For i = 1 To .Rows - 1
            rsReturn.AddNew
            rsReturn!ReturnBillID = temReturnBillID
            rsReturn!SaleBillID = Val(lblBillID.Caption)
            rsReturn!ItemID = .TextMatrix(i, 6)
            rsReturn!BatchID = .TextMatrix(i, 7)
            rsReturn!StoreID = UserStoreID
            rsReturn!Date = Date
            rsReturn!Time = Now
            rsReturn!StaffID = UserID
            rsReturn!ReturnRate = Val(.TextMatrix(i, 9))
            rsReturn!ReturnAmount = Val(.TextMatrix(i, 11))
            rsReturn!ReturnPrice = Val(.TextMatrix(i, 12))
            TemCost = (Val(.TextMatrix(i, 14)) * Val(.TextMatrix(i, 11))) / Val(.TextMatrix(i, 4))
            rsReturn!Cost = TemCost
            TotalCost = TotalCost + TemCost
            rsReturn.Update
            If rsBatchStock.State = 1 Then rsBatchStock.Close
            temSql = "SELECT * from tblBatchStock where BatchID = " & Val(.TextMatrix(i, 7)) & " AND StoreID = " & UserStoreID
            rsBatchStock.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If rsBatchStock.RecordCount < 1 Then
                rsBatchStock.AddNew
                rsBatchStock!ItemID = Val(.TextMatrix(i, 6))
                rsBatchStock!BatchID = Val(.TextMatrix(i, 7))
                rsBatchStock!Stock = Val(.TextMatrix(i, 11))
            Else
                rsBatchStock!Stock = rsBatchStock!Stock + Val(.TextMatrix(i, 11))
            End If
            rsBatchStock.Update
        Next
    End With
    
    If dtcRepaymentMethod.Text = "Cash" Then
        With rsIssuedCash
            If .State = 1 Then .Close
            temSql = "SELECT * from tblIssuedCash"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !IssuedDate = Date
            !IssuedTime = Now
            !IssuedSTaffID = UserID
            !CheckedStaffID = Val(dtcCheckedStaff.BoundText)
            If lblCustomer.Caption = "Outdoor Customer" Then
                !IssuedToCustomerID = txtCustomerID.Text
            ElseIf lblCustomer.Caption = "Indoor Patient" Then
                !IssuedToBHTID = txtCustomerID.Text
            ElseIf lblCustomer.Caption = "Staff Customer" Then
                !IssuedToSTaffID = txtCustomerID.Text
            ElseIf lblCustomer.Caption = "Hospital Unit" Then
                !IssuedToUnitID = Val(txtCustomerID.Text)
            End If
            !Price = txtNetReturn.Text
            !StoreID = UserStoreID
            !SaleBillID = lblBillID.Caption
            !RefillBillID = temReturnBillID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            temIssueCashID = !NewID
            .Close
        End With
        With rsReturnBill
            temSql = "SELECT * from tblReturnBill where ReturnBillID = " & temReturnBillID
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !IssuedCashID = temIssueCashID
                !NetCost = TotalCost
                .Update
            End If
            .Close
        End With
    ElseIf dtcRepaymentMethod.Text = "Credit" Then
        With rsIssuedCash
            If .State = 1 Then .Close
            temSql = "SELECT * from tblIssuedCredit"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !IssuedDate = Date
            !IssuedTime = Now
            !IssuedSTaffID = UserID
            !CheckedStaffID = Val(dtcCheckedStaff.BoundText)
            If lblCustomer.Caption = "Outdoor Customer" Then
                !IssuedToCustomerID = txtCustomerID.Text
                With rsTemCustomer
                    If .State = 1 Then .Close
                    temSql = "SELECT * from tblPatientMainDetails where patientID = " & txtCustomerID.Text
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !Credit = !Credit + Val(txtNetReturn.Text)
                    End If
                    .Update
                    .Close
                End With
            ElseIf lblCustomer.Caption = "Indoor Patient" Then
                !IssuedToBHTID = txtCustomerID.Text
                With rsTemCustomer
                    If .State = 1 Then .Close
                    temSql = "SELECT * from tblBHT where BHTID = " & txtCustomerID.Text
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !Balance = !Balance + Val(txtNetReturn.Text)
                        .Update
                    End If
                    .Close
                End With
            ElseIf lblCustomer.Caption = "Staff Customer" Then
                !IssuedToSTaffID = txtCustomerID.Text
                With rsTemCustomer
                    If .State = 1 Then .Close
                    temSql = "SELECT * from tblStaff where StaffID = " & txtCustomerID.Text
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !Credit = !Credit + Val(txtNetReturn.Text)
                        .Update
                    End If
                    .Close
                End With
            ElseIf lblCustomer.Caption = "Hospital Unit" Then
                !IssuedToUnitID = Val(txtCustomerID.Text)
            End If
            !Price = txtNetReturn.Text
            !StoreID = UserStoreID
            !SaleBillID = lblBillID.Caption
            !RefillBillID = temReturnBillID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            temIssueCreditID = !NewID
            .Close
        End With
        With rsReturnBill
            temSql = "SELECT * from tblReturnBill where ReturnBillID = " & temReturnBillID
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !IssuedCreditID = temIssueCreditID
                !NetCost = TotalCost
                .Update
            End If
            .Close
        End With
    ElseIf dtcRepaymentMethod.Text = "Credit Card" Then
        With rsIssuedWoucher
            If .State = 1 Then .Close
            temSql = "SELECT * from tblIssuedCreditCard"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !IssuedDate = Date
            !IssuedTime = Now
            !IssuedSTaffID = UserID
            !CheckedStaffID = Val(dtcCheckedStaff.BoundText)
            If lblCustomer.Caption = "Outdoor Customer" Then
                !IssuedToCustomerID = txtCustomerID.Text
            ElseIf lblCustomer.Caption = "Indoor Patient" Then
                !IssuedToBHTID = txtCustomerID.Text
            ElseIf lblCustomer.Caption = "Staff Customer" Then
                !IssuedToSTaffID = txtCustomerID.Text
            ElseIf lblCustomer.Caption = "Hospital Unit" Then
                !IssuedToUnitID = Val(txtCustomerID.Text)
             End If
            !Price = txtNetReturn.Text
            !StoreID = UserStoreID
            !SaleBillID = lblBillID.Caption
            !RefillBillID = temReturnBillID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            temIssueWoucherID = !NewID
            .Close
        End With
        With rsReturnBill
            temSql = "SELECT * from tblReturnBill where ReturnBillID = " & temReturnBillID
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !IssuedCreditCardID = temIssueWoucherID
                !NetCost = TotalCost
                .Update
            End If
            .Close
        End With
    Else
        With rsIssuedOther
            If .State = 1 Then .Close
            temSql = "SELECT * from tblIssuedOther"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !IssuedDate = Date
            !IssuedTime = Now
            !IssuedSTaffID = UserID
            !CheckedStaffID = Val(dtcCheckedStaff.BoundText)
            If lblCustomer.Caption = "Outdoor Customer" Then
                !IssuedToCustomerID = txtCustomerID.Text
            ElseIf lblCustomer.Caption = "Indoor Patient" Then
                !IssuedToBHTID = txtCustomerID.Text
            ElseIf lblCustomer.Caption = "Staff Customer" Then
                !IssuedToSTaffID = txtCustomerID.Text
            ElseIf lblCustomer.Caption = "Hospital Unit" Then
                !IssuedToUnitID = Val(txtCustomerID.Text)
             End If
            !Price = txtNetReturn.Text
            !StoreID = UserStoreID
            !SaleBillID = lblBillID.Caption
            !RefillBillID = temReturnBillID
            .Update
            temSql = "SELECT @@IDENTITY AS NewID"
            .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            temIssueOtherID = !NewID
            .Close
        End With
        With rsReturnBill
            temSql = "SELECT * from tblReturnBill where ReturnBillID = " & temReturnBillID
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !IssuedOtherID = temIssueOtherID
                !NetCost = TotalCost
                .Update
            End If
            .Close
        End With
    
    End If
    
    If chkPrint.Value = 1 Then
        Call SetBillPrinter
        Call SetBillPaper
    End If
    
    MsgBox "Successfully Returned"
    frmBillSearchForReturn.Show
    frmBillSearchForReturn.ZOrder 0
    frmBillSearchForReturn.Top = 0
    frmBillSearchForReturn.Left = 0
    Unload Me
'   0   No
'   1   Item
'   2   Batch
'   3   Rate
'   4   Amount
'   5   Price
'   6   ItemID
'   7   BatchID
'   8   AMount
'   9   Rate
'   10  IUnit
'   11  Return Qty
'   12  Return Value
'   13  SaleID


End Sub

Private Sub Form_Load()
    Call FillCombos
    dtcIssueStaff.BoundText = UserID
    Dim tr As Integer
    If TxSaleBillID = 0 Then
        Unload Me
        Exit Sub
    End If
    With rsSaleBill
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleBill.*, tblPatientMainDetailsOutdoor.FirstName as ODFirstName, tblPatientMainDetailsIndoor.FirstName as IDFirstName, tblBHT.BHT, tblStaffCustomer.Name as SCName, tblPaymentMethod.PaymentMethod, tblStaff.Name as StaffName, tblCStaff.Name as CName, tblStore.Store " & _
                    "FROM ((((((tblStaff AS tblStaffCustomer RIGHT JOIN (tblSaleBill INNER JOIN tblStaff ON tblSaleBill.StaffID = tblStaff.StaffID) ON tblStaffCustomer.StaffID = tblSaleBill.BilledStaffID) LEFT JOIN tblBHT ON tblSaleBill.BilledBHTID = tblBHT.BHTID) LEFT JOIN tblPatientMainDetails AS tblPatientMainDetailsIndoor ON tblBHT.PatientID = tblPatientMainDetailsIndoor.PatientID) LEFT JOIN tblPatientMainDetails AS tblPatientMainDetailsOutdoor ON tblSaleBill.BilledOutPatientID = tblPatientMainDetailsOutdoor.PatientID) LEFT JOIN tblStaff AS tblCStaff ON tblSaleBill.CheckedStaffID = tblCStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblSaleBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStore ON tblSaleBill.BilledUnitID = tblStore.StoreID " & _
                    "WHERE (((tblSaleBill.SaleBillID)=" & TxSaleBillID & ")) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If !Cancelled = True Then
                tr = MsgBox("This bill is already cancelled.", vbCritical, "Cancelled")
                Unload Me
                Exit Sub
            ElseIf !Returned = True Then
                tr = MsgBox("This items in this bill is already Returned.", vbCritical, "Cancelled")
                Unload Me
                Exit Sub
            Else
                lblDate.Caption = Format(!Date, LongDateFormat)
                lblTime.Caption = !Time
                lblUser.Caption = Format(![StaffName], "")
                If Not IsNull(![CName]) Then
                    lblCheckedBy.Caption = ![CName]
                End If
                lblBillID.Caption = !SaleBillID
                lblGTotal.Caption = ![Price]
                lblDiscount.Caption = ![Discount]
                lblDiscountPercent.Caption = ![DiscountPercent]
                lblNetTotal.Caption = ![NetPrice]
                lblPaymentMethod.Caption = ![PaymentMethod]
                dtcRepaymentMethod.BoundText = !PaymentMethodID
                lblBillID.Caption = ![SaleBillID]
                If Not IsNull(![ODFirstName]) Then
                    lblName.Caption = ![ODFirstName]
                    lblCustomer.Caption = "Outdoor Customer"
                    txtCustomerID.Text = !BilledOutPatientID
                ElseIf Not IsNull(![IDFirstName]) Then
                    lblName.Caption = (![IDFirstName])
                    lblCustomer.Caption = "Indoor Patient"
                    lblBHT.Caption = ![BHT]
                    txtCustomerID.Text = !BilledBHTID
                ElseIf Not IsNull(![SCName]) Then
                    lblCustomer.Caption = "Staff Customer"
                    lblName.Caption = ![SCNameName]
                    txtCustomerID.Text = !BilledStaffID
                ElseIf Not IsNull(!Store) Then
                    lblCustomer.Caption = "Hospital Unit"
                    lblName.Caption = ![Store]
                    txtCustomerID.Text = ![BilledUnitID]
                End If
            End If
        Else
            tr = MsgBox("There is no such Bill ID ", vbCritical, "Error")
            Unload Me
            Exit Sub
        End If
        .Close
    End With
    Call FormatItemGrid
    Call FillItemGrid
End Sub


Private Sub FormatItemGrid()
    With GridItem
        .Cols = 15
        .Rows = 1
        Dim i As Integer
        For i = 0 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            Select Case i
                Case 0: .Text = "No."
                        .ColWidth(i) = 400
                Case 1: .Text = "Item"
                        .ColWidth(i) = 3600
                Case 2: .Text = "Batch"
                        .ColWidth(i) = 1000
                Case 3: .Text = "Rate"
                        .ColWidth(i) = 1200
                Case 4: .Text = "Quentity"
                        .ColWidth(i) = 1500
                Case 5: .ColWidth(i) = 1500
                        .Text = "Value"
                Case 11:    .ColWidth(i) = 1200
                            .Text = "Return Qty"
                Case 12:    .ColWidth(i) = 1500
                            .Text = "Return Value"
                Case Else
                        .ColWidth(i) = 1
            End Select
        Next
'   0   No
'   1   Item
'   2   Batch
'   3   Rate
'   4   Amount
'   5   Price
'   6   ItemID
'   7   BatchID
'   8   AMount
'   9   Rate
'   10  IUnit
'   11  Return Qty
'   12  Return Value
'   13  SaleID
'   14  Cost
    End With
End Sub

Private Sub FillItemGrid()
    With rsSale
        If .State = 1 Then .Close
        temSql = "SELECT tblSale.*, tblItem.Display, tblIssueUnit.IssueUnit, tblSaleBill.SaleBillID, tblBatch.Batch " & _
                    "FROM (tblIssueUnit RIGHT JOIN ((tblSaleBill LEFT JOIN tblSale ON tblSaleBill.SaleBillID = tblSale.SaleBillID) LEFT JOIN tblItem ON tblSale.ItemID = tblItem.ItemID) ON tblIssueUnit.IssueUnitID = tblItem.IssueUnitID) LEFT JOIN tblBatch ON tblSale.BatchID = tblBatch.BatchID " & _
                    "WHERE (((tblSaleBill.SaleBillID)=" & TxSaleBillID & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            Dim i As Integer
            .MoveLast
            GridItem.Rows = .RecordCount + 1
            .MoveFirst
            i = 0
            While .EOF = False
                i = i + 1
                GridItem.TextMatrix(i, 0) = i
                If Not IsNull(!Display) Then
                    GridItem.TextMatrix(i, 1) = !Display
                End If
                If Not IsNull(!Batch) Then
                    GridItem.TextMatrix(i, 2) = !Batch
                End If
                GridItem.TextMatrix(i, 3) = Format(!Rate, "0.00")
                GridItem.TextMatrix(i, 4) = !Amount
                GridItem.TextMatrix(i, 5) = Format(!GrossPrice, "0.00")
                GridItem.TextMatrix(i, 6) = !ItemID
                GridItem.TextMatrix(i, 7) = !BatchID
                GridItem.TextMatrix(i, 13) = ![SaleBillID]
                GridItem.TextMatrix(i, 10) = ![IssueUnit]
                GridItem.TextMatrix(i, 14) = ![Cost]
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub CalculateTotals()
    txtReturnDiscount.Text = Format((Val(txtReturnValue.Text) * Val(lblDiscountPercent.Caption) / 100), "0.00")
    txtNetReturn.Text = Format((Val(txtReturnValue.Text) - Val(txtReturnDiscount.Text)), "0.00")
End Sub

Private Sub GridItem_Click()
    With GridItem
        If .Rows < 2 Then Exit Sub
        If .Row < 1 Then Exit Sub
        Dim i As Integer
        i = .Row
            lblIUnit.Caption = .TextMatrix(i, 10)
            lblIUnit1.Caption = .TextMatrix(i, 10)
            lblIValue.Caption = .TextMatrix(i, 5)
            lblIRate.Caption = .TextMatrix(i, 3)
            lblIQty.Caption = .TextMatrix(i, 4)
            txtRow.Text = i
            lblItem.Caption = .TextMatrix(i, 1)
            txtRRate.Text = .TextMatrix(i, 3)
'   0   No
'   1   Item
'   2   Batch
'   3   Rate
'   4   Amount
'   5   Price
'   6   ItemID
'   7   BatchID
'   8   AMount
'   9   R Rate
'   10  IUnit
'   11  Return Qty
'   12  Return Value
'   13  SaleID
    End With
End Sub

Private Sub txtReturnValue_Change()
    Call CalculateTotals
End Sub

Private Sub txtRQty_Change()
    txtRValue.Text = Format((Val(txtRRate.Text) * Val(txtRQty.Text)), "0.00")
End Sub

Private Sub SetBillPrinter()
    CsetPrinter.SetPrinterAsDefault (BillPrinterName)
End Sub

Private Sub SetBillPaper()
    Dim TemResponce As Long
    Dim RetVal As Integer
    RetVal = SelectForm(BillPaperName, Me.hwnd)
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
    If LCase(Left(Trim(HospitalName), 1)) = "m" Then
        
    ElseIf LCase(Left(Trim(HospitalName), 1)) = "r" Then
        RuhunaPrint
    ElseIf LCase(Left(Trim(HospitalName), 1)) = "c" Then
    
    Else
    
    End If
End Sub

Private Sub RuhunaPrint()
On Error GoTo eh:
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

'        .FontSize = 12
'        .Font = "Lucida Console"
'        Printer.Print
'        Printer.Print Tab(Tab8); UserStore & "   -  Goods Return Note"
'        Printer.Print
        .FontSize = 12
        .Font = "Lucida Console"
'        Printer.Print Tab(4); "             RUHUNU HOSPITAL (PVT) LTD "
        Printer.Print Tab(4); "Return Note "
        .FontSize = 10
        .Font = "Lucida Console"
'        Printer.Print Tab(Tab1); "Karapitiya, Galle." & "           Tel: 091-2234059-60, 091-5577113-14"
'        Printer.Print
        Dim TemString As String
        Printer.Print Tab(Tab1); "Issue No        - "; lblBillID.Caption & "-" & TemString; " Issued Date: "; Format(lblDate.Caption, "dd MM yy"); Tab(Tab6); "Issued Time : "; lblTime.Caption
        Printer.Print
        Printer.Print Tab(Tab1); "Return No - "; temReturnBillID & "-" & TemString; " Returned Date: "; Format(Date, "dd MM yy"); Tab(Tab6); "Returned Time : "; Time
        Printer.Print Tab(Tab1); "Patient : "; lblName.Caption; "       BHT : "; lblBHT.Caption
        Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
        Printer.Print Tab(Tab1); "Category"; Tab(Tab2); "Item Name"; Tab(Tab3); "Quentity"; Tab(Tab5); Right(Space(12) & "Price", 9); Tab(Tab9); Right(Space(12) & "Value", 13)
        Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
        Printer.Print
        .FontSize = 10
        .Font = "Lucida Console"
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
        Printer.FontSize = 10
        Printer.Font = "Lucida Console"
'            Printer.Print Tab(Tab1); .TextMatrix(I, 11);
        If Val(.TextMatrix(i, 11)) > 0 Then
                Printer.Print Tab(Tab2); Left(.TextMatrix(i, 1), 20);
                Printer.Print Tab(Tab3); Left(.TextMatrix(i, 11), 24);
                Printer.Print Tab(Tab5); Right(Space(12) & .TextMatrix(i, 9), 9);
                Printer.Print Tab(Tab7); Right(Space(12) & .TextMatrix(i, 12), 8)
        End If
        Next i
    End With
    With Printer
        .Font = 10
        .Font = "Lucida Console"
        Printer.Print
        Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
'        Printer.Print
        .FontSize = 10
        .Font = "Lucida Console"
        Printer.Print Tab(Tab1); "Gross Total"; Tab(Tab4); Right((Space(10)) & (txtReturnValue.Text), 10)
        If Val(txtReturnDiscount.Text) > 0 Then
            Printer.Print Tab(Tab1); "Discount"; Tab(Tab4); Right((Space(10)) & (txtReturnDiscount.Text), 10)
            Printer.Print Tab(Tab1); "Net Total"; Tab(Tab4); Right((Space(10)) & (txtNetReturn.Text), 10)
        End If
'        Printer.Print
        Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
        Printer.Print Tab(Tab1); "Operate by "; UserName; Tab(Tab5)
'        Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
        Printer.Print Tab(Tab1); "Returns are acceptted only once"
'        Printer.Print Tab(Tab1); "--------------------------------------------------------------------------"
        Printer.Print
        Printer.Print
        .EndDoc
    End With
'   0   No
'   1   Item
'   2   Batch
'   3   Rate
'   4   Amount
'   5   Price
'   6   ItemID
'   7   BatchID
'   8   AMount
'   9   R Rate
'   10  IUnit
'   11  Return Qty
'   12  Return Value
'   13  SaleID

Exit Sub

eh:

    MsgBox "printer Error"

End Sub

