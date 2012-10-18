VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmViewSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Bills"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13995
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
   ScaleHeight     =   9570
   ScaleWidth      =   13995
   Begin VB.TextBox txtCustomerID 
      Height          =   375
      Left            =   3600
      TabIndex        =   38
      Top             =   8040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtNetReturn 
      Height          =   375
      Left            =   3600
      TabIndex        =   32
      Top             =   9000
      Width           =   2295
   End
   Begin VB.TextBox txtReturnDiscount 
      Height          =   375
      Left            =   3600
      TabIndex        =   30
      Top             =   8520
      Width           =   2295
   End
   Begin VB.TextBox txtReturnValue 
      Height          =   375
      Left            =   3600
      TabIndex        =   22
      Top             =   8040
      Width           =   2295
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   12480
      TabIndex        =   20
      Top             =   9120
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
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7646
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
      TabIndex        =   34
      Top             =   2160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcIssueStaff 
      Height          =   360
      Left            =   11040
      TabIndex        =   35
      Top             =   1680
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label lblReturn 
      Caption         =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
      Height          =   375
      Left            =   240
      TabIndex        =   40
      Top             =   7560
      Width           =   5535
   End
   Begin VB.Label lblCancelled 
      Caption         =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Top             =   7080
      Width           =   5535
   End
   Begin VB.Label Label20 
      Caption         =   "Issued By"
      Height          =   255
      Left            =   9360
      TabIndex        =   37
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label21 
      Caption         =   "Checked By"
      Height          =   255
      Left            =   9360
      TabIndex        =   36
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblNet 
      Caption         =   "Net Return"
      Height          =   255
      Left            =   2160
      TabIndex        =   33
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label lblDis 
      Caption         =   "Less Discount"
      Height          =   255
      Left            =   2160
      TabIndex        =   31
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label Label23 
      Caption         =   "Customer Type"
      Height          =   255
      Left            =   9360
      TabIndex        =   29
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label22 
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   9360
      TabIndex        =   28
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblCustomer 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11040
      TabIndex        =   27
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11040
      TabIndex        =   26
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblBHT 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11040
      TabIndex        =   25
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label18 
      Caption         =   "BHT No."
      Height          =   255
      Left            =   9360
      TabIndex        =   24
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblGross 
      Caption         =   "Gross Return"
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Discount %"
      Height          =   255
      Left            =   9360
      TabIndex        =   21
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label lblDiscountPercent 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11040
      TabIndex        =   19
      Top             =   8160
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
      Left            =   11040
      TabIndex        =   17
      Top             =   8640
      Width           =   2535
   End
   Begin VB.Label lblDiscount 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11040
      TabIndex        =   16
      Top             =   7680
      Width           =   2535
   End
   Begin VB.Label lblGTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11040
      TabIndex        =   15
      Top             =   7200
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
      Left            =   9360
      TabIndex        =   8
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Discount"
      Height          =   255
      Left            =   9360
      TabIndex        =   7
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Gross Total"
      Height          =   255
      Left            =   9360
      TabIndex        =   6
      Top             =   7200
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
End
Attribute VB_Name = "frmViewSale"
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
    Dim temIssueCreditCardID As Long
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
'    With dtcRepaymentMethod
'        Set .RowSource = rsViewPaymentMethod
'        .ListField = "PaymentMethod"
'        .BoundColumn = "PaymentMethodID"
'    End With
    
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub SetBillPrinter()
    CsetPrinter.SetPrinterAsDefault (BillPrinterName)
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
        temSql = "SELECT tblSaleBill.*, tblPatientMainDetailsOutdoor.FirstName as OFirstName, tblPatientMainDetailsIndoor.FirstName as IFirstName, tblBHT.BHT, tblStaffCustomer.Name as SCName, tblPaymentMethod.PaymentMethod as MyPM, tblStaff.Name, tblCStaff.Name as CName, tblStore.Store " & _
                    "FROM ((((((tblStaff AS tblStaffCustomer RIGHT JOIN (tblSaleBill INNER JOIN tblStaff ON tblSaleBill.StaffID = tblStaff.StaffID) ON tblStaffCustomer.StaffID = tblSaleBill.BilledStaffID) LEFT JOIN tblBHT ON tblSaleBill.BilledBHTID = tblBHT.BHTID) LEFT JOIN tblPatientMainDetails AS tblPatientMainDetailsIndoor ON tblBHT.PatientID = tblPatientMainDetailsIndoor.PatientID) LEFT JOIN tblPatientMainDetails AS tblPatientMainDetailsOutdoor ON tblSaleBill.BilledOutPatientID = tblPatientMainDetailsOutdoor.PatientID) LEFT JOIN tblStaff AS tblCStaff ON tblSaleBill.CheckedStaffID = tblCStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblSaleBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStore ON tblSaleBill.BilledUnitID = tblStore.StoreID " & _
                    "WHERE (((tblSaleBill.SaleBillID)=" & TxSaleBillID & ")) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If !Cancelled = True Then
                lblCancelled.Caption = "This bill was Cancelled "
                lblReturn.Caption = ""
            ElseIf !Returned = True Then
                lblCancelled.Caption = ""
                lblReturn.Caption = "This bill was returned "
            Else
                lblCancelled.Caption = ""
                lblReturn.Caption = "This Bill was NOT cancelled nor returned"
            End If
            lblDate.Caption = Format(!Date, LongDateFormat)
            lblTime.Caption = !Time
            lblUser.Caption = ![Name]
            If Not IsNull(![CName]) Then
                lblCheckedBy.Caption = ![CName]
            End If
            lblBillID.Caption = !SaleBillID
            lblGTotal.Caption = Format(!Price, "0.00")
            lblDiscount.Caption = Format(![Discount], "0.00")
            lblDiscountPercent.Caption = Format(![DiscountPercent], "0.0") & "%"
            lblNetTotal.Caption = Format(![NetPrice], "0.00")
            lblPaymentMethod.Caption = ![MyPM]
            lblBillID.Caption = ![SaleBillID]
            If Not IsNull(![OFirstName]) Then
                lblName.Caption = ![OFirstName]
                lblCustomer.Caption = "Outdoor Customer"
                txtCustomerID.Text = !BilledOutPatientID
            ElseIf Not IsNull(![IFirstName]) Then
                lblName.Caption = (![IFirstName])
                lblCustomer.Caption = "Indoor Patient"
                lblBHT.Caption = ![BHT]
                txtCustomerID.Text = !BilledBHTID
            ElseIf Not IsNull(![SCName]) Then
                lblCustomer.Caption = "Staff Customer"
                lblName.Caption = ![SCName]
                txtCustomerID.Text = !BilledStaffID
            ElseIf Not IsNull(!Store) Then
                lblCustomer.Caption = "Hospital Unit"
                lblName.Caption = ![Store]
                txtCustomerID.Text = ![BilledUnitID]
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
    Call FindReturns
End Sub


Private Sub FindReturns()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "Select * from tblReturnBill where SaleBillID = " & TxSaleBillID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtReturnDiscount.Text = Format(!Discount, "0.00")
            txtReturnValue.Text = Format(!Price, "0.00")
            txtNetReturn.Text = Format(!NetPrice, "0.00")
        Else
            txtReturnDiscount.Text = Empty
            txtReturnValue.Text = Empty
            txtNetReturn.Text = Empty
            txtReturnDiscount.Visible = False
            txtReturnValue.Visible = False
            txtNetReturn.Visible = False
            lblGross.Visible = False
            lblDis.Visible = False
            lblNet.Visible = False
        End If
        .Close
    End With
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
'   15  Cost
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
'On Error GoTo er
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
                GridItem.TextMatrix(i, 9) = Format(!Rate, "0.00")
                GridItem.TextMatrix(i, 4) = !Amount
'                GridItem.TextMatrix(i, 11) = !Amount
                GridItem.TextMatrix(i, 5) = Format(!GrossPrice, "0.00")
                GridItem.TextMatrix(i, 12) = Format(!GrossPrice, "0.00")
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
    With GridItem
        For i = 1 To .Rows - 1
            If rsReturn.State = 1 Then rsReturn.Close
            temSql = "Select * from tblReturn where SaleBillID = " & TxSaleBillID & " And ItemID = " & Val(.TextMatrix(i, 6))
            rsReturn.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If rsReturn.RecordCount > 0 Then
                .TextMatrix(i, 11) = rsReturn!ReturnAmount
                .TextMatrix(i, 12) = rsReturn!ReturnPrice
            End If
            rsReturn.Close
        Next i
    End With
    Exit Sub
    
er:
    MsgBox "Error"
    Exit Sub
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
'        Printer.Print Tab(Tab8); UserStore & "   -  Cancellation Note"
'        Printer.Print
        .FontSize = 12
        .Font = "Lucida Console"
'        Printer.Print Tab(4); "             RUHUNU HOSPITAL (PVT) LTD "
        Printer.Print Tab(4); "Cancellation"
        Printer.Print
        .FontSize = 10
        .Font = "Lucida Console"
'        Printer.Print Tab(Tab1); "Karapitiya, Galle." & "           Tel: 091-2234059-60, 091-5577113-14"
'        Printer.Print
        Dim TemString As String
        Printer.Print Tab(Tab1); "Issue No        - "; lblBillID.Caption & "-" & TemString; " Issued Date: "; Format(lblDate.Caption, "dd MM yy"); Tab(Tab6); "Issued Time : "; lblTime.Caption
        Printer.Print
        Printer.Print Tab(Tab1); "Cancel No - "; temReturnBillID & "-" & TemString; " Cancelled Date: "; Format(Date, "dd MM yy"); Tab(Tab6); "Cancelled Time : "; Time
        Printer.Print Tab(Tab1); "Patient : "; lblName.Caption; "       BHT  : "; lblBHT.Caption
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
            Printer.Print Tab(Tab2); Left(.TextMatrix(i, 1), 20);
            Printer.Print Tab(Tab3); Left(.TextMatrix(i, 11), 24);
            Printer.Print Tab(Tab5); Right(Space(12) & .TextMatrix(i, 9), 9);
            Printer.Print Tab(Tab7); Right(Space(12) & .TextMatrix(i, 12), 8)
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
    MsgBox "Printer Error"


End Sub
