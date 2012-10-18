VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHealthSchemeSupplierBillSettling 
   Caption         =   "BHT Bill Settling by Health Scheme Suppliers"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12750
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
   ScaleHeight     =   9450
   ScaleWidth      =   12750
   Begin VB.TextBox txtIncomeBillID 
      Height          =   360
      Left            =   7560
      TabIndex        =   33
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "&Print"
      Height          =   255
      Left            =   6480
      TabIndex        =   31
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   8520
      Width           =   12495
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   6720
         TabIndex        =   28
         Top             =   240
         Width           =   5535
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label29 
         Caption         =   "Paper"
         Height          =   255
         Left            =   6120
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label30 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtComments 
      Height          =   735
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   6240
      Width           =   4695
   End
   Begin VB.TextBox txtSettled 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   10440
      TabIndex        =   19
      Top             =   6720
      Width           =   2055
   End
   Begin VB.TextBox txtToSettle 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   10440
      TabIndex        =   18
      Top             =   7200
      Width           =   2055
   End
   Begin VB.TextBox txtSettling 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   7560
      Width           =   4695
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
      Left            =   10440
      TabIndex        =   9
      Top             =   6240
      Width           =   2055
   End
   Begin btButtonEx.ButtonEx btnUpdate 
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   8040
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
      Left            =   11280
      TabIndex        =   8
      Top             =   8040
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
      Format          =   20643843
      CurrentDate     =   39697
   End
   Begin MSFlexGridLib.MSFlexGrid GridBill 
      Height          =   5055
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   8916
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
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
      Format          =   20643843
      CurrentDate     =   39697
   End
   Begin MSDataListLib.DataCombo dtcPayment 
      Height          =   360
      Left            =   1800
      TabIndex        =   11
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
      Left            =   1680
      TabIndex        =   22
      Top             =   8040
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   20643843
      CurrentDate     =   39697
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   1680
      TabIndex        =   24
      Top             =   7080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   375
      Left            =   6240
      TabIndex        =   32
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "To &Excel"
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
   Begin VB.Label Label11 
      Caption         =   "Payment Method"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Paid on"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Payment Comments"
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Settling"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   7560
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "To Settle"
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Settled"
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   6720
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
      TabIndex        =   13
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Total Value"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   6240
      Width           =   3015
   End
End
Attribute VB_Name = "frmHealthSchemeSupplierBillSettling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsBill As New ADODB.Recordset
    Dim rsViewSupplier As New ADODB.Recordset
    Dim rsViewPayment As New ADODB.Recordset

    Dim temSQL As String
    Dim i As Integer
    Dim CsetPrinter As New cSetDfltPrinter
    Dim NumForms As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean


    Dim rsTemPrice As New ADODB.Recordset
    Dim rsPrice  As New ADODB.Recordset
    Dim temRow As Long
    Dim temCol As Long
    Dim temText As String
    Dim temCellText As String
    Dim temBoxText As String

Private Sub btnExcel_Click()
    GridToExcel GridBill, "Health Scheme Bills"
End Sub

Private Sub btnUpdate_Click()
    Dim temText As String
    Dim DisplayBillID As Long
    Dim MyBHT As New clsBHT
    GridBill.Col = 1
    GridBill.Row = 1
    GridBill.Col = 10
    GridBill.Row = GridBill.Rows - 1
    temText = "Are you sure you want to mark Rs. " & txtSettling.Text & " as Paid"
    If IsNumeric(dtcSupplier.BoundText) = True Then
        temText = temText & " by " & dtcSupplier.Text & "?"
    Else
        temText = temText & " by the respective Companies?"
    End If
    Dim i As Integer
    Dim n As Integer
    i = MsgBox(temText, vbYesNo)
    If i = vbNo Then
        MsgBox "NOT paid"
        Exit Sub
    Else
        For n = 1 To GridBill.Rows - 1
            If Val(GridBill.TextMatrix(n, 10)) <> 0 Then
                UpdateCompanyBalance Val(GridBill.TextMatrix(n, 12)), Val(GridBill.TextMatrix(n, 10)), True, False, False, Val(dtcPayment.BoundText)
                UpdateBHTSettle Val(GridBill.TextMatrix(n, 11)), Val(GridBill.TextMatrix(n, 10)), True, False
                
                DisplayBillID = UpdateIncomeBill(Val(GridBill.TextMatrix(n, 12)), Val(GridBill.TextMatrix(n, 10)), True, False, Val(GridBill.TextMatrix(n, 11)))
                If chkPrint.Value = 1 Then
                    MyBHT.BHTID = Val(GridBill.TextMatrix(n, 11))
                    printBill GridBill.TextMatrix(n, 1), MyBHT.FirstName, MyBHT.BHT, Val(GridBill.TextMatrix(n, 10)), DisplayBillID
                End If
            End If
        Next
        MsgBox "Updated successfully"
        Call fillGrid
        Call ClearUpdateValues
    End If
End Sub


Private Sub UpdateBHTSettle(BHTID As Long, UpdateValue As Double, AddToSettle As Boolean, DeductFromSettle As Boolean)
    Dim rsRefill As New ADODB.Recordset
    With rsRefill
        If .State = 1 Then .Close
        temSQL = "Select * from tblBHT where BHTID = " & BHTID
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If IsNull(!Balance) = True Then !Balance = 0
            If AddToSettle = True Then
                !Balance = !Balance - UpdateValue
            ElseIf DeductFromSettle = True Then
                !Balance = !Balance + UpdateValue
            End If
            If !Balance = 0 Then
                !BillSettled = True
            End If
            .Update
        End If
        .Close
    End With
End Sub


Private Sub dtcSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtcSupplier.Text = Empty
    End If
End Sub

Private Function UpdateIncomeBill(HSSID As Long, UpdateValue As Double, AddToSettle As Boolean, DeductFromSettle As Boolean, BHTID As Long) As Long
    Dim rsRefill As New ADODB.Recordset
    Dim DisplayBillID As Long
    With rsRefill
    
'        If .State = 1 Then .Close
'        temSQL = "Select Count(IncomeBillID) as BillCount from tblIncomeBill where Completed = 1 AND IsHSSPaymentBill = 1"
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            If IsNull(!BillCount) = False Then
'                DisplayBillID = !BillCount + 1
'            Else
'                DisplayBillID = 1
'            End If
'        Else
'            DisplayBillID = 1
'        End If
        
        If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeBill where IncomeBillID = 0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Date = Date
        !Time = Now
        !UserID = UserID
        !StoreID = UserStoreID
        !GrossTotal = UpdateValue
        !Discount = UpdateValue
        !NetTotal = 0
        !IsHSSPaymentBill = True
        !PaidHSSID = HSSID
        !BHTID = BHTID
        If IsNumeric(cmbPaymentMethod.BoundText) = True Then
            !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        Else
            !PaymentMethodID = 1
        End If
        !PaymentComments = txtComments.Text
        !Completed = True
        !CompletedUserID = UserID
        !CompletedDate = Date
        !CompletedTime = Now
'        !DisplayBillID = DisplayBillID
        .Update
        .Close
    
         temSQL = "SELECT @@IDENTITY AS NewID"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        txtIncomeBillID.Text = !NewID
        .Close
   
        DisplayBillID = NewHSSPaymentDisplayBillID(Val(txtIncomeBillID.Text))
    
        If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeBill where IncomeBillID = " & Val(txtIncomeBillID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !DisplayBillID = DisplayBillID
            .Update
        End If
        .Close
    
    
    End With

    UpdateIncomeBill = DisplayBillID


End Function
Private Sub ClearUpdateValues()
    txtComments.Text = Empty
    dtpPaid.Value = Date
    txtToSettle.Text = Empty
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub GridBill_Click()
    Dim TemDisPayID As Long
    TemDisPayID = Val(GridBill.TextMatrix(GridBill.Row, 11))
    If TemDisPayID = 0 Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem

    End With
End Sub

Private Sub GridBill_EnterCell()
    txtCellText.Visible = False
    If GridBill.Row = 0 Then

    ElseIf GridBill.ColWidth(GridBill.Col) < 2 Then

    ElseIf GridBill.Visible = False Then Exit Sub

    Else
        temRow = GridBill.Row
        temCol = GridBill.Col
        temCellText = GridBill.TextMatrix(temRow, temCol)
        txtCellText.Top = GridBill.Top + GridBill.CellTop
        txtCellText.Left = GridBill.Left + GridBill.CellLeft
        txtCellText.Height = GridBill.CellHeight - 60
        txtCellText.Width = GridBill.CellWidth
        txtCellText.BackColor = GridBill.CellBackColor
        txtCellText.Alignment = GridBill.CellAlignment
        txtCellText.Text = temCellText
        txtCellText.Visible = True
        On Error Resume Next
        txtCellText.SetFocus
        SendKeys "{Home}+{end}"
    End If

    If GridBill.Col = 10 Then
        txtCellText.Locked = False
        Call CalculateSettling
    Else
        txtCellText.Locked = True
    End If

End Sub

Private Sub CalculateSettling()
    Dim i As Integer
    Dim Settling As Double
    For i = 1 To GridBill.Rows - 1
        Settling = Settling + Val(GridBill.TextMatrix(i, 10))
    Next
    txtSettling.Text = Format(Settling, "0.00")
End Sub

Private Sub GridBill_LeaveCell()
    txtCellText.Visible = False
    If GridBill.Row = 0 Then

    ElseIf GridBill.Col = 0 Or GridBill.Col = 1 Or GridBill.Col = 2 Then

    ElseIf GridBill.ColWidth(GridBill.Col) < 2 Then

    ElseIf GridBill.Visible = False Then Exit Sub

    Else

        temBoxText = txtCellText.Text
        If temBoxText <> temCellText Then
            GridBill.TextMatrix(temRow, temCol) = temBoxText
        End If

    End If

End Sub

Private Sub GridBill_Scroll()
    txtCellText.Visible = False
End Sub

Private Sub txtCellText_KeyDown(KeyCode As Integer, Shift As Integer)
    With GridBill
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
    fillGrid
End Sub

Private Sub dtcSupplier_Change()
    fillGrid
End Sub

Private Sub dtpFrom_Change()
    fillGrid
End Sub

Private Sub dtpTo_Change()
    fillGrid
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    FillCombos
    dtcPayment.Text = "Credit"
    fillGrid
    Me.WindowState = 2
    Call ClearUpdateValues
    Call PopulatePrinters
    Call GetSettings
End Sub

Private Sub FillCombos()

    With rsViewSupplier
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblHealthSchemeSuppliers order by HealthSchemeSupplierName"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With rsViewPayment
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblPaymentMethod order by PaymentMethod"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcPayment
        Set .RowSource = rsViewPayment
        .ListField = "PaymentMethod"
        .BoundColumn = "PaymentMethodID"
    End With
    With dtcSupplier
        Set .RowSource = rsViewSupplier
        .ListField = "HealthSchemeSupplierName"
        .BoundColumn = "HealthSchemeSupplierID"
    End With
    
    Dim PayMethod As New clsFillCombos
    
    PayMethod.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanReceive", False
    
End Sub

Private Sub fillGrid()
    Dim Total As Double


    Dim TotalValue As Double
    Dim SettledValue As Double
    Dim ToSettleValue As Double
    Dim SettlingValue As Double

    With GridBill
        .Clear

        .Rows = 1
        .Cols = 13

        .Visible = False

        .Row = 0

        .Col = 0
        .Text = "No"
        .CellAlignment = 4

        .Col = 1
        .Text = "Company"
        .CellAlignment = 4

        .Col = 2
        .Text = "BHT No."
        .CellAlignment = 4

        .Col = 3
        .Text = "Admission"
        .CellAlignment = 4

        .Col = 4
        .Text = "Discharge"
        .CellAlignment = 4

        .Col = 5
        .Text = "GRN No"
        .CellAlignment = 4

        .Col = 6
        .Text = "Dates Passed"
        .CellAlignment = 4

        .Col = 7
        .Text = "Total Bill"
        .CellAlignment = 4

        .Col = 8
        .Text = "Paid"
        .CellAlignment = 4

        .Col = 9
        .Text = "To Pay"
        .CellAlignment = 4

        .Col = 10
        .Text = "Paying"
        .CellAlignment = 4
        
        .Col = 11
        .Text = "BHTID"

        .Col = 12
        .Text = "HSSID"


        .ColWidth(0) = 600
        .ColWidth(1) = 3000
        .ColWidth(2) = 1600
        .ColWidth(3) = 1600
        .ColWidth(4) = 1600
        .ColWidth(5) = 0
        .ColWidth(6) = 1600
        .ColWidth(7) = 1600
        .ColWidth(8) = 1600
        .ColWidth(9) = 1600
        .ColWidth(10) = 1600
        .ColWidth(11) = 0
        .ColWidth(12) = 0

        With rsBill
            temSQL = "SELECT tblHealthSchemeSuppliers.HealthSchemeSupplierName, tblHealthSchemeSuppliers.HealthSchemeSupplierID, tblBHT.BHT, tblBHT.DOA, tblBHT.DOD, tblBHT.NetPrice, tblBHT.Balance, tblBHT.BHTID, dbo.tblPatientMainDetails.FirstName "
            temSQL = temSQL & "FROM  dbo.tblBHT LEFT OUTER JOIN dbo.tblPatientMainDetails ON dbo.tblBHT.PatientID = dbo.tblPatientMainDetails.PatientID LEFT OUTER JOIN dbo.tblHealthSchemeSuppliers ON dbo.tblBHT.HealthSchemeSupplierID = dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierID "
            temSQL = temSQL & "WHERE tblBHT.DOD Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
            If IsNumeric(dtcSupplier.BoundText) = True Then
                temSQL = temSQL & " And tblHealthSchemeSuppliers.HealthSchemeSupplierID = " & Val(dtcSupplier.BoundText)
            End If
            If IsNumeric(dtcPayment.BoundText) = True Then
                temSQL = temSQL & " AND tblBHT.PaymentMethodID = " & dtcPayment.BoundText & " "
            End If
            If .State = 1 Then .Close
            .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                .MoveLast
                .MoveFirst
                GridBill.Rows = .RecordCount + 1
                i = 1
                While .EOF = False
                    GridBill.Row = i
                    Dim n As Long

                    For n = 0 To GridBill.Cols - 1
                        GridBill.Col = n
                        If i Mod 2 = 1 Then
                            GridBill.CellBackColor = RGB(255, 255, 200)
                        Else
                            GridBill.CellBackColor = RGB(255, 255, 30)
                        End If
                    Next
                    GridBill.TextMatrix(i, 0) = i
                    If Not IsNull(!HealthSchemeSupplierName) Then GridBill.TextMatrix(i, 1) = !HealthSchemeSupplierName
                    If Not IsNull(!BHT) Then GridBill.TextMatrix(i, 2) = !BHT & " - " & !FirstName
                    If Not IsNull(!DOA) Then GridBill.TextMatrix(i, 3) = Format(!DOA, "dd MMMM yyyy")
                    If Not IsNull(!DOD) Then GridBill.TextMatrix(i, 4) = Format(!DOD, "dd MMMM yyyy")
                    
                    If Not IsNull(!BHTID) Then GridBill.TextMatrix(i, 5) = !BHTID

                    If Not IsNull(!DOD) Then GridBill.TextMatrix(i, 6) = DateDiff("d", !DOD, Date) & " days"

                    GridBill.TextMatrix(i, 7) = Format(!NetPrice, "#,##0.00")
                    TotalValue = TotalValue + !NetPrice
                    
                    GridBill.TextMatrix(i, 8) = Format(!NetPrice - !Balance, "#,##0.00")
                    
                    GridBill.TextMatrix(i, 9) = Format(!Balance, "#,##0.00")
                    
                    SettledValue = SettledValue + !NetPrice - !Balance
                    ToSettleValue = ToSettleValue + !Balance

                    If Not IsNull(!BHTID) Then GridBill.TextMatrix(i, 11) = !BHTID
                    If Not IsNull(![HealthSchemeSupplierID]) Then GridBill.TextMatrix(i, 12) = ![HealthSchemeSupplierID]

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
    GridBill.Width = Me.Width - 200
End Sub

Private Sub txtCellText_LostFocus()
    CalculateSettling
End Sub


Private Sub GetSettings(): On Error Resume Next
    cmbPaymentMethod.BoundText = 5 'Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
    chkPrint.Value = GetSetting(App.EXEName, Me.Name, chkPrint.Name, 1)
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "")
    GetCommonSettings Me
End Sub


Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub PopulatePapers(): On Error Resume Next
    cmbPaper.Clear
    SetPrinter = False
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
'        With FormSize
'            .cx = BillPaperHeight
'            .cy = BillPaperWidth
'        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                cmbPaper.AddItem PtrCtoVbString(.pName)
'                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle): DoEvents
    End If
End Sub

Private Sub cmbPrinter_Change()
    Call PopulatePapers
End Sub

Private Sub cmbPrinter_Click()
    Call PopulatePapers
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveCommonSettings Me
End Sub

Private Sub printBill(CompanyName As String, PatientName As String, BHT As String, PaymentValue As Double, BillNo As Long)
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    
    Dim CenterX As Long
    Dim FieldX As Long
    Dim NoX As Long
    Dim ValueX As Long
    Dim AllLines() As String
    Dim i As Integer
    Dim temY As Long
    Dim n As Long
    
    With MyFOnt
        .Name = DefaultFont.Name
        .Bold = False
        .Italic = False
        .Size = 12
        .Italic = False
        .Underline = False
    End With
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        CenterX = Printer.Width / 2
        NoX = 1440 * 1
        FieldX = 1440 * 1.2
        ValueX = Printer.Width - 1440 * 1
        temY = Printer.CurrentY
        
        
        MyFOnt.Bold = True
        MyFOnt.Size = 13
        PrintingText 0, temY, Printer.Width, 0, HospitalName, CentreAlign, MyFOnt
        
        temY = Printer.CurrentY
        
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText 0, temY, Printer.Width, 0, HospitalDescreption, CentreAlign, MyFOnt
        
        temY = Printer.CurrentY
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText 0, temY, Printer.Width, 0, "Credit Settlement", CentreAlign, MyFOnt
        
        MyFOnt.Size = 11
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Date : " & Format(Date, "dd MM yy"), leftAlign, MyFOnt

        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Time : " & Format(Time, "hh : mm AMPM"), leftAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Company : " & CompanyName, leftAlign, MyFOnt

        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "BHT : " & BHT, leftAlign, MyFOnt
        

        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Patient Name : " & PatientName, leftAlign, MyFOnt
                
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Bill No. : " & BillNo, leftAlign, MyFOnt
        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, "Payment Method : " & cmbPaymentMethod.Text, LeftAlign, MyFOnt
'        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Payment Comments : " & txtComments.Text, leftAlign, MyFOnt
        

        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Bad Detters/Withhalding Tax", leftAlign, MyFOnt
        MyFOnt.Bold = True
        PrintingText FieldX, temY, ValueX, 0, "Rs. " & Format(PaymentValue, "#,##0.00"), rightAlign, MyFOnt
        MyFOnt.Bold = False
        
        Printer.Print
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, ".......", rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Cashier :  " & UserFullName, rightAlign, MyFOnt
        
        Printer.EndDoc
        
    End If
End Sub

Private Sub PrintingText(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, PrintText As String, PrintAlignment As TextAlignment, PrintFont As ReportFont)
    
    If PrintAlignment = leftAlign Then
        Printer.CurrentX = X1
    ElseIf PrintAlignment = rightAlign Then
        Printer.CurrentX = X2 - Printer.TextWidth(PrintText)
    ElseIf PrintAlignment = CentreAlign Then
        Printer.CurrentX = (X1 + X2 / 2) - (Printer.TextWidth(PrintText) / 2)
    Else
        Printer.CurrentX = X1
    End If
    If Y1 <> 0 Then Printer.CurrentY = Y1
    Printer.Font.Name = PrintFont.Name
    Printer.Font.Size = PrintFont.Size
    Printer.Font.Italic = PrintFont.Italic
    Printer.Font.Bold = PrintFont.Bold
    Printer.Font.Underline = PrintFont.Underline
    
    Printer.Print PrintText
End Sub
