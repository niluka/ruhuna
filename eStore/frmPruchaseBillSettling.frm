VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPruchaseBillSettling 
   Caption         =   "Purchase Bill Settling"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
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
   ScaleHeight     =   9210
   ScaleWidth      =   10515
   Begin VB.TextBox txtComments 
      Height          =   1335
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   6600
      Width           =   5175
   End
   Begin VB.TextBox txtSettled 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox txtToSettle 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      Top             =   7560
      Width           =   2055
   End
   Begin VB.TextBox txtSettling 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   8040
      Width           =   3015
   End
   Begin VB.TextBox txtCellText 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   6600
      Width           =   2055
   End
   Begin btButtonEx.ButtonEx btnUpdate 
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   8520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
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
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9000
      TabIndex        =   21
      Top             =   8760
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
   Begin MSDataListLib.DataCombo dtcSupplier 
      Height          =   360
      Left            =   1800
      TabIndex        =   5
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
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   72613891
      CurrentDate     =   39697
   End
   Begin MSFlexGridLib.MSFlexGrid GridPurchase 
      Height          =   5415
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9551
      _Version        =   393216
      ScrollTrack     =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   72613891
      CurrentDate     =   39697
   End
   Begin MSDataListLib.DataCombo dtcPayment 
      Height          =   360
      Left            =   1800
      TabIndex        =   22
      Top             =   1080
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpPaid 
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   8520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   72613891
      CurrentDate     =   39697
   End
   Begin VB.Label Label10 
      Caption         =   "Paid on"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Payment Comments"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Settling"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   8040
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "To Settle"
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Settled"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   7080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
   Begin VB.Label Label3 
      Caption         =   "Supplier"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Payment Method"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Total Value"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   6600
      Width           =   3015
   End
End
Attribute VB_Name = "frmPruchaseBillSettling"
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
    
    
    Dim rsTemPrice As New ADODB.Recordset
    Dim rsPrice  As New ADODB.Recordset
    Dim temRow As Long
    Dim temCol As Long
    Dim temText As String
    Dim temCellText As String
    Dim temBoxText As String
    

Private Sub btnUpdate_Click()
    Dim temText As String
    GridPurchase.Col = 1
    GridPurchase.Row = 1
    GridPurchase.Col = 10
    GridPurchase.Row = GridPurchase.Rows - 1
    temText = "Are you sure you want to pay Rs. " & txtSettling.Text
    If IsNumeric(dtcSupplier.BoundText) = True Then
        temText = temText & " to " & dtcSupplier.Text
    Else
        temText = temText & " to the respective suppliers?"
    End If
    Dim i As Integer
    Dim n As Integer
    i = MsgBox(temText, vbYesNo)
    If i = vbNo Then
        MsgBox "NOT paid"
        Exit Sub
    Else
        For n = 1 To GridPurchase.Rows - 1
            If Val(GridPurchase.TextMatrix(n, 10)) <> 0 Then
                UpdateSupplierBalance Val(GridPurchase.TextMatrix(n, 12)), Val(GridPurchase.TextMatrix(n, 10)), True, False, False
                UpdateRefillSettle Val(GridPurchase.TextMatrix(n, 11)), Val(GridPurchase.TextMatrix(n, 10)), True, False
            End If
        Next
        MsgBox "Updated successfully"
        Call FillGrid
        Call ClearUpdateValues
    End If
End Sub

Private Sub UpdateSupplierBalance(SupplierID As Long, UpdateValue As Double, AddToSupplier As Boolean, DeductSupplier As Boolean, ResetSupplier As Boolean)
    Dim rsSup As New ADODB.Recordset
    With rsSup
        If .State = 1 Then .Close
        temSql = "Select * from tblDistrubutor where DistributorID = " & SupplierID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If AddToSupplier = True Then
                !Balance = !Balance + UpdateValue
            ElseIf DeductSupplier = True Then
                !Balance = !Balance - UpdateValue
            ElseIf ResetSupplier = True Then
                !Balance = UpdateValue
            End If
            .Update
            .Close
            temSql = "Select * from tblDistributorPayment"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !DistributorID = SupplierID
            !PaymentValue = UpdateValue
            !UserID = UserID
            !PaymentDate = Format(dtpPaid.Value, "dd MMMM yyyy")
            !PaymentTime = Now
            !PaymentComments = txtComments.Text
            .Update
            .Close
        End If
    End With
End Sub

Private Sub UpdateRefillSettle(RefillBillID As Long, UpdateValue As Double, AddToSettle As Boolean, DeductFromSettle As Boolean)
    Dim rsRefill As New ADODB.Recordset
    With rsRefill
        If .State = 1 Then .Close
        temSql = "Select * from tblRefillBill where refillbillid = " & RefillBillID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If IsNull(!SettledValue) = True Then !SettledValue = 0
            If AddToSettle = True Then
                !SettledValue = !SettledValue + UpdateValue
            ElseIf DeductFromSettle = True Then
                !SettledValue = !SettledValue - UpdateValue
            End If
            .Update
        End If
        .Close
    End With
End Sub


Private Sub ClearUpdateValues()
    txtComments.Text = Empty
    dtpPaid.Value = Date
    txtToSettle.Text = Empty
End Sub

Private Sub GridPurchase_Click()
    Dim TemDisPayID As Long
    TemDisPayID = Val(GridPurchase.TextMatrix(GridPurchase.Row, 11))
    If TemDisPayID = 0 Then Exit Sub
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
    
    End With
End Sub

Private Sub GridPurchase_EnterCell()
    txtCellText.Visible = False
    If GridPurchase.Row = 0 Then
    
    ElseIf GridPurchase.CellWidth < 2 Then
    
    ElseIf GridPurchase.Visible = False Then Exit Sub
    
    Else
        temRow = GridPurchase.Row
        temCol = GridPurchase.Col
        temCellText = GridPurchase.TextMatrix(temRow, temCol)
        txtCellText.Top = GridPurchase.Top + GridPurchase.CellTop
        txtCellText.Left = GridPurchase.Left + GridPurchase.CellLeft
        txtCellText.Height = GridPurchase.CellHeight - 60
        txtCellText.Width = GridPurchase.CellWidth
        txtCellText.BackColor = GridPurchase.CellBackColor
        txtCellText.Alignment = GridPurchase.CellAlignment
        txtCellText.Text = temCellText
        txtCellText.Visible = True
        On Error Resume Next
        txtCellText.SetFocus
        SendKeys "{Home}+{end}"
    End If
    
    If GridPurchase.Col = 10 Then
        txtCellText.Locked = False
        Call CalculateSettling
    Else
        txtCellText.Locked = True
    End If

End Sub

Private Sub CalculateSettling()
    Dim i As Integer
    Dim Settling As Double
    For i = 1 To GridPurchase.Rows - 1
        Settling = Settling + Val(GridPurchase.TextMatrix(i, 10))
    Next
    txtSettling.Text = Format(Settling, "0.00")
End Sub

Private Sub GridPurchase_LeaveCell()
    txtCellText.Visible = False
    If GridPurchase.Row = 0 Then
    
    ElseIf GridPurchase.Col = 0 Or GridPurchase.Col = 1 Or GridPurchase.Col = 2 Then
    
    ElseIf GridPurchase.CellWidth < 2 Then
    
    ElseIf GridPurchase.Visible = False Then Exit Sub
    
    Else
    
        temBoxText = txtCellText.Text
        If temBoxText <> temCellText Then
            GridPurchase.TextMatrix(temRow, temCol) = temBoxText
        End If

    End If
    
End Sub

Private Sub GridPurchase_Scroll()
    txtCellText.Visible = False
End Sub

Private Sub txtCellText_KeyDown(KeyCode As Integer, Shift As Integer)
    With GridPurchase
        If KeyCode = vbKeyReturn Then
            If temCol < .Cols - 1 Then
                .Col = temCol + 1
            Else
                .Col = 1
                .Row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyEscape Then
            txtCellText.Text = temText
        ElseIf KeyCode = vbKeyTab Then
            If temCol < .Cols - 1 Then
                .Col = temCol + 1
            Else
                .Col = 1
                .Row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyUp Then
            If temRow > 1 Then
                .Row = temRow - 1
            End If
        ElseIf KeyCode = vbKeyDown Then
            If temRow < .Rows - 1 Then
                .Row = temRow + 1
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If temCol > 1 Then
                .Col = temCol - 1
            End If
        ElseIf KeyCode = vbKeyRight Then
            If temCol < .Cols - 1 Then
                .Col = temCol + 1
            End If
        End If
    End With
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
    dtcPayment.Text = "Credit"
    FillGrid
    Me.WindowState = 2
    Call ClearUpdateValues
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
    
    
    Dim TotalValue As Double
    Dim SettledValue As Double
    Dim ToSettleValue As Double
    Dim SettlingValue As Double
    
    With GridPurchase
        .Clear
        
        .Rows = 1
        .Cols = 13
        
        .Visible = False
        
        .Row = 0
        
        .Col = 0
        .Text = "No"
        .CellAlignment = 4
        
        .Col = 1
        .Text = "Supplier"
        .CellAlignment = 4
        
        .Col = 2
        .Text = "Invoice No."
        .CellAlignment = 4
        
        .Col = 3
        .Text = "Invoice Date"
        .CellAlignment = 4
        
        .Col = 4
        .Text = "GRN Date"
        .CellAlignment = 4
        
        .Col = 5
        .Text = "GRN No"
        .CellAlignment = 4
        
        .Col = 6
        .Text = "Dates Passed"
        .CellAlignment = 4
        
        .Col = 7
        .Text = "Total Payment"
        .CellAlignment = 4
        
        .Col = 8
        .Text = "Settled"
        .CellAlignment = 4
        
        .Col = 9
        .Text = "To Settle"
        .CellAlignment = 4
        
        .Col = 10
        .Text = "Settling"
        .CellAlignment = 4
        
        
        .Col = 11
        .Text = "ID"
        
        
        .ColWidth(0) = 600
        .ColWidth(1) = 3000
        .ColWidth(2) = 1600
        .ColWidth(3) = 1600
        .ColWidth(4) = 1600
        .ColWidth(5) = 1600
        .ColWidth(6) = 1600
        .ColWidth(7) = 1600
        .ColWidth(8) = 1600
        .ColWidth(9) = 1600
        .ColWidth(10) = 1600
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        
        With rsPurchase
            temSql = "SELECT tblDistrubutor.DistributorName, tblRefillBill.ReturnedValue,  tblRefillBill.InvoiceDate, tblRefillBill.Returned, tblRefillBill.Cancelled, tblRefillBill.InvoiceNo, tblRefillBill.SettledValue , tblRefillBill.Date, tblRefillBill.RefillBillID, tblPaymentMethod.PaymentMethod, tblRefillBill.NetPrice, tblRefillBill.DistributorID, tblPaymentMethod.PaymentMethod, tblRefillBill.Date "
            temSql = temSql & "FROM (tblRefillBill LEFT JOIN tblDistrubutor ON tblRefillBill.DistributorID = tblDistrubutor.DistributorID) LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID "
            temSql = temSql & "WHERE tblRefillBill.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
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
                    GridPurchase.Row = i
                    Dim n As Long
                    
                    For n = 0 To GridPurchase.Cols - 1
                        GridPurchase.Col = n
                        If i Mod 2 = 1 Then
                            GridPurchase.CellBackColor = RGB(255, 255, 200)
                        Else
                            GridPurchase.CellBackColor = RGB(255, 255, 30)
                        End If
                    Next
                    GridPurchase.TextMatrix(i, 0) = i
                    If Not IsNull(!DistributorName) Then GridPurchase.TextMatrix(i, 1) = !DistributorName
                    If Not IsNull(!InvoiceNo) Then GridPurchase.TextMatrix(i, 2) = !InvoiceNo
                    If Not IsNull(!InvoiceDate) Then GridPurchase.TextMatrix(i, 3) = !InvoiceDate
                    If Not IsNull(!Date) Then GridPurchase.TextMatrix(i, 4) = !Date
                    If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(i, 5) = !RefillBillID
                    
                    If Not IsNull(!InvoiceDate) Then GridPurchase.TextMatrix(i, 6) = DateDiff("d", !InvoiceDate, Date) & " days"
                    
                    If !Cancelled = True Then
                            If Not IsNull(!NetPrice) Then
                                GridPurchase.TextMatrix(i, 7) = Format(!NetPrice, "#,##0.00")
                            End If
                            GridPurchase.TextMatrix(i, 8) = "Cancelled"
                            GridPurchase.TextMatrix(i, 9) = ""
                    
                    ElseIf !Returned = True Then
                        If IsNull(!NetPrice) = False And IsNull(!ReturnedValue) = False Then
                            GridPurchase.TextMatrix(i, 7) = Format(!NetPrice - !ReturnedValue, "#,##0.00")
                            TotalValue = TotalValue + !NetPrice - !ReturnedValue
                        End If
                        
                        If Not IsNull(!NetPrice) Then
                            If Not IsNull(!SettledValue) Then
                                GridPurchase.TextMatrix(i, 8) = Format(!SettledValue, "#,##0.00")
                                GridPurchase.TextMatrix(i, 9) = Format(!NetPrice - !SettledValue - !ReturnedValue, "#,##0.00")
                                SettledValue = SettledValue + !SettledValue
                                ToSettleValue = ToSettleValue + !NetPrice - !SettledValue - !ReturnedValue
                            Else
                                GridPurchase.TextMatrix(i, 8) = Format(0, "#,##0.00")
                                GridPurchase.TextMatrix(i, 9) = Format(!NetPrice - !ReturnedValue, "#,##0.00")
                                SettledValue = SettledValue + 0
                                ToSettleValue = ToSettleValue + !NetPrice - !ReturnedValue
                            End If
                        End If
                    Else
                        If Not IsNull(!NetPrice) Then
                            GridPurchase.TextMatrix(i, 7) = Format(!NetPrice, "#,##0.00")
                            TotalValue = TotalValue + !NetPrice
                        End If
                        
                        If Not IsNull(!NetPrice) Then
                            If Not IsNull(!SettledValue) Then
                                GridPurchase.TextMatrix(i, 8) = Format(!SettledValue, "#,##0.00")
                                GridPurchase.TextMatrix(i, 9) = Format(!NetPrice - !SettledValue, "#,##0.00")
                                SettledValue = SettledValue + !SettledValue
                                ToSettleValue = ToSettleValue + !NetPrice - !SettledValue
                            Else
                                GridPurchase.TextMatrix(i, 8) = Format(0, "#,##0.00")
                                GridPurchase.TextMatrix(i, 9) = Format(!NetPrice, "#,##0.00")
                                SettledValue = SettledValue + 0
                                ToSettleValue = ToSettleValue + !NetPrice
                            End If
                        End If
                    End If
                    If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(i, 11) = !RefillBillID
                    If Not IsNull(![DistributorID]) Then GridPurchase.TextMatrix(i, 12) = ![DistributorID]
                                        
                    i = i + 1
                    .MoveNext
                Wend
            End If
        End With
        .Visible = True
    End With
    txtTotal.Text = Format(TotalValue, "#,##0.00")
    txtSettled.Text = Format(SettledValue, "#,##0.00")
    txtToSettle.Text = Format(ToSettleValue, "#,##0.00")
    txtSettling.Text = Format(SettlingValue, "#,##0.00")
    
End Sub


Private Sub Form_Resize()
    GridPurchase.Width = Me.Width - 200
End Sub
