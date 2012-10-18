VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSaleReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale Reports"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13275
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
   ScaleHeight     =   8550
   ScaleWidth      =   13275
   Begin VB.Frame Frame6 
      Caption         =   "Bill Details"
      Height          =   1455
      Left            =   6360
      TabIndex        =   26
      Top             =   2160
      Width           =   3015
      Begin VB.CheckBox chkNetTotal 
         Caption         =   "Net Total"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkDiscount 
         Caption         =   "Discount"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkGrossTotal 
         Caption         =   "Gross Total"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Value           =   1  'Checked
         Width           =   2655
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Customer"
      Height          =   1455
      Left            =   3240
      TabIndex        =   22
      Top             =   2160
      Width           =   3015
      Begin VB.CheckBox chkOutPatient 
         Caption         =   "Outpatient"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkInpatient 
         Caption         =   "BHT (Inpatient)"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox chkStaff 
         Caption         =   "Staff Member"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Payment Method"
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   2160
      Width           =   3015
      Begin VB.CheckBox chkCredit 
         Caption         =   "Credit"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CheckBox chkCreditCard 
         Caption         =   "Credit Card"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox chkCheque 
         Caption         =   "Cheque"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox chkCash 
         Caption         =   "Cash"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Value           =   1  'Checked
         Width           =   2655
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridSales 
      Height          =   4095
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnDisplay 
      Height          =   495
      Left            =   9480
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Display"
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
   Begin VB.Frame Frame3 
      Caption         =   "Staff Members"
      Height          =   1815
      Left            =   6360
      TabIndex        =   9
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton optSelectedStaff 
         Caption         =   "Selected Staff Member"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optAllStaff 
         Caption         =   "All Staff Membesr"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo dtcStaff 
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Duration"
      Height          =   1815
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   3015
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   20709379
         CurrentDate     =   29224
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   20709379
         CurrentDate     =   29224
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Department"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin MSDataListLib.DataCombo dtcDepts 
         Height          =   360
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.OptionButton optSelectedDepts 
         Caption         =   "Selected Department"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optAllDepts 
         Caption         =   "All Departments"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   10800
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "&Print"
      Enabled         =   0   'False
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
      Left            =   11880
      TabIndex        =   16
      Top             =   7920
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
End
Attribute VB_Name = "frmSaleReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsSaleBill As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsStore As New ADODB.Recordset
    Dim temSql As String
    Dim FromFormat As Boolean

Private Function CanDisplay() As Boolean
    CanDisplay = False
    Dim tr As Integer
    If chkCash.Value = 0 And chkCredit.Value = 0 And chkCreditCard.Value = 0 And chkCheque.Value = 0 Then
        tr = MsgBox("You have to selecte at least one method of payment", vbCritical, "Payment Method")
        chkCash.SetFocus
        Exit Function
    End If
    If chkStaff.Value = 0 And chkOutPatient.Value = 0 And chkInpatient.Value = 0 Then
        tr = MsgBox("You have to select at least one type of customer", vbCritical, "Customer type")
        chkOutPatient.SetFocus
        Exit Function
    End If
    If optSelectedDepts.Value = True And IsNumeric(dtcDepts.BoundText) = False Then
        tr = MsgBox("You have not selecte the Department", vbCritical, "Department")
        dtcDepts.SetFocus
        Exit Function
    End If
    If optSelectedStaff.Value = True And IsNumeric(dtcStaff.BoundText) = False Then
        tr = MsgBox("You have not selected the staff member", vbCritical, "Staff member")
        dtcStaff.SetFocus
        Exit Function
    End If
    CanDisplay = True
End Function


Private Sub bttnDisplay_Click()
    
    If FromFormat = False Then
        FromFormat = False
        If CanDisplay = False Then Exit Sub
    Else
        FromFormat = False
        Exit Sub
    End If
    
    bttnPrint.Enabled = True
    Dim temSelect As String
    Dim temWhere As String
    Dim temOrderBY As String
    
    temSelect = "SELECT tblSaleBill.Date, tblSaleBill.Time, tblStore.Store, tblStaffUser.Name, tblInPatient.FirstName, tblOutPatient.FirstName, tblStaffCustomer.Name, tblPaymentMethod.PaymentMethod, tblSaleBill.SaleBillID, tblSaleBill.Price, tblSaleBill.Discount, tblSaleBill.NetPrice, tblStaffCustomer.Name " & _
                    "FROM ((((((tblSaleBill LEFT JOIN tblPaymentMethod ON tblSaleBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblStaffUser ON tblSaleBill.StaffID = tblStaffUser.StaffID) LEFT JOIN tblStaff AS tblStaffCustomer ON tblSaleBill.BilledStaffID = tblStaffCustomer.StaffID) LEFT JOIN tblPatientMainDetails AS tblOutPatient ON tblSaleBill.BilledOutPatientID = tblOutPatient.PatientID) LEFT JOIN tblBHT ON tblSaleBill.BilledBHTID = tblBHT.BHTID) LEFT JOIN tblPatientMainDetails AS tblInPatient ON tblBHT.PatientID = tblInPatient.PatientID) LEFT JOIN tblStore ON tblSaleBill.StoreID = tblStore.StoreID "
    temWhere = "WHERE tblSaleBill.Date Between '" & dtpFrom.Value & "' And '" & dtpTo.Value & "' "
    temOrderBY = " Order by SaleBillID"
    
    If optSelectedDepts.Value = True Then
        temWhere = temWhere & " AND tblSaleBill.StoreID=" & dtcDepts.BoundText & " "
    End If
    If optSelectedStaff.Value = True Then
        temWhere = temWhere & " AND tblSaleBill.StaffID=" & dtcStaff.BoundText & " "
    End If
    
    If chkCash.Value = 1 And chkCredit.Value = 0 And chkCreditCard.Value = 0 And chkCheque.Value = 0 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Cash') "
    ElseIf chkCash.Value = 0 And chkCredit.Value = 1 And chkCreditCard.Value = 0 And chkCheque.Value = 0 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Credit') "
    ElseIf chkCash.Value = 0 And chkCredit.Value = 0 And chkCreditCard.Value = 1 And chkCheque.Value = 0 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Credit Card') "
    ElseIf chkCash.Value = 0 And chkCredit.Value = 0 And chkCreditCard.Value = 0 And chkCheque.Value = 1 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Cheque') "
    ElseIf chkCash.Value = 1 And chkCredit.Value = 1 And chkCreditCard.Value = 0 And chkCheque.Value = 0 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Cash' OR tblPaymentMethod.PaymentMethod='Credit') "
    ElseIf chkCash.Value = 1 And chkCredit.Value = 0 And chkCreditCard.Value = 1 And chkCheque.Value = 0 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Cash' OR tblPaymentMethod.PaymentMethod='Credit Card') "
    ElseIf chkCash.Value = 1 And chkCredit.Value = 0 And chkCreditCard.Value = 0 And chkCheque.Value = 1 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Cash' OR tblPaymentMethod.PaymentMethod='Cheque') "
    ElseIf chkCash.Value = 0 And chkCredit.Value = 1 And chkCreditCard.Value = 1 And chkCheque.Value = 0 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Credit' OR tblPaymentMethod.PaymentMethod='Credit Card') "
    ElseIf chkCash.Value = 0 And chkCredit.Value = 1 And chkCreditCard.Value = 0 And chkCheque.Value = 1 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Credit' OR tblPaymentMethod.PaymentMethod='Cheque') "
    ElseIf chkCash.Value = 0 And chkCredit.Value = 0 And chkCreditCard.Value = 1 And chkCheque.Value = 1 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Credit Card' OR tblPaymentMethod.PaymentMethod='Cheque') "
    ElseIf chkCash.Value = 1 And chkCredit.Value = 1 And chkCreditCard.Value = 1 And chkCheque.Value = 0 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Cash' OR tblPaymentMethod.PaymentMethod='Credit Card' OR tblPaymentMethod.PaymentMethod='Credit') "
    ElseIf chkCash.Value = 1 And chkCredit.Value = 1 And chkCreditCard.Value = 0 And chkCheque.Value = 1 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Cash' OR tblPaymentMethod.PaymentMethod='Cheque' OR tblPaymentMethod.PaymentMethod='Credit') "
    ElseIf chkCash.Value = 1 And chkCredit.Value = 0 And chkCreditCard.Value = 1 And chkCheque.Value = 1 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Cash' OR tblPaymentMethod.PaymentMethod='Cheque' OR tblPaymentMethod.PaymentMethod='Credit Card') "
    ElseIf chkCash.Value = 0 And chkCredit.Value = 1 And chkCreditCard.Value = 1 And chkCheque.Value = 1 Then
        temWhere = temWhere & "  AND (tblPaymentMethod.PaymentMethod='Credit Card' OR tblPaymentMethod.PaymentMethod='Cheque' OR tblPaymentMethod.PaymentMethod='Credit') "
    ElseIf chkCash.Value = 1 And chkCredit.Value = 1 And chkCreditCard.Value = 1 And chkCheque.Value = 1 Then
    
    Else
        MsgBox "Error"
        Exit Sub
    End If
    
    If chkOutPatient.Value = 1 And chkInpatient.Value = 1 And chkStaff.Value = 1 Then
        temWhere = temWhere & "  AND (((tblInPatient.FirstName) Is Not Null) or ((tblStaffCustomer.Name) Is Not Null) or ((tblOutPatient.FirstName) Is Not Null)) "
    ElseIf chkOutPatient.Value = 1 And chkInpatient.Value = 1 And chkStaff.Value = 0 Then
        temWhere = temWhere & "  AND (((tblInPatient.FirstName) Is Not Null) or ((tblOutPatient.FirstName) Is Not Null)) "
    ElseIf chkOutPatient.Value = 1 And chkInpatient.Value = 0 And chkStaff.Value = 1 Then
        temWhere = temWhere & "  AND (((tblStaffCustomer.Name) Is Not Null) or ((tblOutPatient.FirstName) Is Not Null)) "
    ElseIf chkOutPatient.Value = 0 And chkInpatient.Value = 1 And chkStaff.Value = 1 Then
        temWhere = temWhere & "  AND (((tblInPatient.FirstName) Is Not Null) or ((tblStaffCustomer.Name) Is Not Null)) "
    ElseIf chkOutPatient.Value = 1 And chkInpatient.Value = 0 And chkStaff.Value = 0 Then
        temWhere = temWhere & "  AND (((tblInPatient.FirstName) Is Null) AND ((tblStaffCustomer.Name) Is Null) AND ((tblOutPatient.FirstName) Is Not Null)) "
    ElseIf chkOutPatient.Value = 0 And chkInpatient.Value = 1 And chkStaff.Value = 0 Then
        temWhere = temWhere & "  AND (((tblInPatient.FirstName) Is Not Null) AND ((tblStaffCustomer.Name) Is Null) AND ((tblOutPatient.FirstName) Is Null)) "
    ElseIf chkOutPatient.Value = 0 And chkInpatient.Value = 0 And chkStaff.Value = 1 Then
        temWhere = temWhere & "  AND (((tblInPatient.FirstName) Is Null) AND ((tblStaffCustomer.Name) Is Not Null) AND ((tblOutPatient.FirstName) Is Null)) "
    ElseIf chkOutPatient.Value = 0 And chkInpatient.Value = 0 And chkStaff.Value = 0 Then
        temWhere = temWhere & "  AND (((tblInPatient.FirstName) Is  Null) AND ((tblStaffCustomer.Name) Is Null) AND ((tblOutPatient.FirstName) Is Null)) "
    End If
    
    If optSelectedDepts.Value = True Then
        temWhere = temWhere & " AND tblSaleBillID.StoreID = " & Val(dtcDepts.BoundText) & " "
    End If
    
    If optSelectedStaff.Value = True Then
        temWhere = temWhere & " StaffID = " & Val(dtcStaff.BoundText) & " "
    End If
    
    Dim i As Integer
    Dim TemGrossTotal As Double
    Dim TemDiscount As Double
    Dim TemNetTotal As Double
    
    
    With rsSaleBill
        If .State = 1 Then .Close
        temSql = temSelect & temWhere & temOrderBY
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                i = i + 1
                GridSales.Rows = GridSales.Rows + 1
                GridSales.TextMatrix(i, 0) = !SaleBillID
                GridSales.TextMatrix(i, 1) = !Date
                GridSales.TextMatrix(i, 2) = !Time
                GridSales.TextMatrix(i, 3) = !Store
                GridSales.TextMatrix(i, 4) = ![tblStaffUser.Name]
                If IsNull(![tblInPatient.FirstName]) = False Then
                    GridSales.TextMatrix(i, 5) = "Inpatient"
                    GridSales.TextMatrix(i, 6) = ![tblInPatient.FirstName]
                ElseIf IsNull(![tblOutPatient.FirstName]) = False Then
                    GridSales.TextMatrix(i, 5) = "Out Patient"
                    GridSales.TextMatrix(i, 6) = ![tblOutPatient.FirstName]
                ElseIf IsNull(![tblStaffCustomer.Name]) = False Then
                    GridSales.TextMatrix(i, 5) = "Staff Customer"
                    GridSales.TextMatrix(i, 6) = ![tblStaffCustomer.Name]
                End If
                GridSales.TextMatrix(i, 7) = ![PaymentMethod]
                GridSales.TextMatrix(i, 8) = ![SaleBillID]
                TemGrossTotal = TemGrossTotal + ![Price]
                GridSales.TextMatrix(i, 9) = Format(![Price], "0.00")
                TemDiscount = TemDiscount + ![Discount]
                GridSales.TextMatrix(i, 10) = Format(![Discount], "0.00")
                TemNetTotal = TemNetTotal + ![NetPrice]
                GridSales.TextMatrix(i, 11) = Format(![NetPrice], "0.00")
                .MoveNext
            Wend
            GridSales.Rows = GridSales.Rows + 1
            i = i + 1
            GridSales.TextMatrix(i, 9) = Format(TemGrossTotal, "0.00")
            GridSales.TextMatrix(i, 10) = Format(TemDiscount, "0.00")
            GridSales.TextMatrix(i, 11) = Format(TemNetTotal, "0.00")
    '   0   ?
    '   1   Date
    '   2   Time
    '   3   Dept
    '   4   User
    '   5   BHT/Outdoor/Staff
    '   6   Customer
    '   7   Payment Mode
    '   8   BillID
    '   9   Gross Total
    '   10   Discount
    '   11  NetPrice
        End If
        .Close
    End With
End Sub

Private Sub chkCash_Click()
    Call FormatWidths
End Sub

Private Sub chkCheque_Click()
    Call FormatWidths
End Sub

Private Sub chkCredit_Click()
    Call FormatWidths
End Sub

Private Sub chkCreditCard_Click()
    Call FormatWidths
End Sub

Private Sub chkDiscount_Click()
    Call FormatWidths
End Sub

Private Sub chkGrossTotal_Click()
    Call FormatWidths
End Sub

Private Sub chkInpatient_Click()
    Call FormatWidths
End Sub

Private Sub chkNetTotal_Click()
    Call FormatWidths
End Sub

Private Sub chkOutPatient_Click()
    Call FormatWidths
End Sub

Private Sub chkStaff_Click()
    Call FormatWidths
End Sub

Private Sub dptSelectedStaff_Click()
    Call FormatWidths
End Sub


Private Sub dtpFrom_Change()
    Call FormatWidths
End Sub

Private Sub dtpTo_Change()
    Call FormatWidths
End Sub

Private Sub Form_Load()
    Call FormatWidths
    Call GetValues
    Call FillCombos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetValues
End Sub

Private Sub GetValues()
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub

Private Sub FormatGrid()
    
    With GridSales
        .Cols = 12
        .Rows = 1
        .Row = 0
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Date"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Time"
        
        .Col = 3
        .CellAlignment = 4
        .Text = "Dept."
        
        .Col = 4
        .CellAlignment = 4
        .Text = "User"
        
        .Col = 5
        .CellAlignment = 4
        .Text = "Catogery"
        
        .Col = 6
        .CellAlignment = 4
        .Text = "Customer"
        
        .Col = 7
        .CellAlignment = 4
        .Text = "Payment"
        
        .Col = 8
        .CellAlignment = 4
        .Text = "Bill ID"
        
        .Col = 9
        .CellAlignment = 4
        .Text = "Gross Total"
        
        .Col = 10
        .CellAlignment = 4
        .Text = "Discount"
        
        .Col = 11
        .CellAlignment = 4
        .Text = "Net Total"
        
        
    End With
    
    '   0   ?
    '   1   Date
    '   2   Time
    '   3   Dept
    '   4   User
    '   5   BHT/Outdoor/Staff
    '   6   Customer
    '   7   Payment Mode
    '   8   BillID
    '   9   Gross Total
    '   10   Discount
    '   11  NetPrice
    
End Sub

Private Sub FillGrid()
'    Dim temSelecet As String
'    Dim temFrom As String
'    Dim temWhere As String
'    Dim temWhereBegninnig As String
'
'    With rsSaleBill
'        If .State = 1 Then .Close
'        temSelecet = "SELECT tblSaleBill.Date, tblSaleBill.Time, tblStore.Store, tblStaffUser.Name, tblInPatient.FirstName, tblOutPatient.FirstName, tblStaffCustomer.Name, tblPaymentMethod.PaymentMethod, tblSaleBill.SaleBillID, tblSaleBill.Price, tblSaleBill.Discount, tblSaleBill.NetPrice "
'        temFrom = "FROM ((((((tblSaleBill LEFT JOIN tblPaymentMethod ON tblSaleBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblStaffUser ON tblSaleBill.StaffID = tblStaffUser.StaffID) LEFT JOIN tblStaff AS tblStaffCustomer ON tblSaleBill.BilledStaffID = tblStaffCustomer.StaffID) LEFT JOIN tblPatientMainDetails AS tblOutPatient ON tblSaleBill.BilledOutPatientID = tblOutPatient.PatientID) LEFT JOIN tblBHT ON tblSaleBill.BilledBHTID = tblBHT.BHTID) LEFT JOIN tblPatientMainDetails AS tblInPatient ON tblBHT.PatientID = tblInPatient.PatientID) LEFT JOIN tblStore ON tblSaleBill.StoreID = tblStore.StoreID "
'        temWhere = " WHERE tblSaleBill.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
'        If optSelectedDepts.Value = True Then
'            temWhere = temWhere & " AND tblSaleBill.StoreID= " & dtcDepts.BoundText & " "
'        End If
'        If optSelectedStaff.Value = True Then
'            temWhere = temWhere & " AND tblSaleBill.StaffID = " & dtcStaff.BoundText & " "
'        End If
'
'
'
'' " AND ((tblSaleBill.StoreID)=1) AND ((tblSaleBill.StaffID)=1));
'
End Sub

Private Sub FormatWidths()


    Call FormatGrid

    Dim temCount As Integer
    With GridSales
        .ColWidth(0) = 1
        If dtpFrom.Value = Date And dtpTo.Value = Date Then
            .ColWidth(1) = 1
        Else
            .ColWidth(1) = 900
        End If
        .ColWidth(2) = 900
        If optAllDepts.Value = True Then
            .ColWidth(3) = 1000
        Else
            .ColWidth(3) = 1
        End If
        If optAllStaff.Value = True Then
            .ColWidth(4) = 1000
        Else
            .ColWidth(4) = 1
        End If
        temCount = 0
        If chkInpatient.Value = 1 Then temCount = temCount + 1
        If chkOutPatient.Value = 1 Then temCount = temCount + 1
        If chkStaff.Value = 1 Then temCount = temCount + 1
        If temCount > 1 Then
            .ColWidth(5) = 1000
        Else
            .ColWidth(5) = 1
        End If
        temCount = 0
        If chkCash.Value = 1 Then temCount = temCount + 1
        If chkCredit.Value = 1 Then temCount = temCount + 1
        If chkCreditCard.Value = 1 Then temCount = temCount + 1
        If chkCheque.Value = 1 Then temCount = temCount + 1
        If temCount > 1 Then
            .ColWidth(7) = 1000
        Else
            .ColWidth(7) = 1
        End If
        temCount = 0
        .ColWidth(8) = 1000
        If chkGrossTotal.Value = 1 Then
            .ColWidth(9) = 1000
        Else
            .ColWidth(9) = 1
        End If
        If chkDiscount.Value = 1 Then
            .ColWidth(10) = 1000
        Else
            .ColWidth(10) = 1
        End If
        If chkNetTotal.Value = 1 Then
            .ColWidth(11) = 1000
        Else
            .ColWidth(11) = 1
        End If
        Dim temColWidth As Long
        .ColWidth(6) = 4000
        
    End With
    
    FromFormat = True
    Call bttnDisplay_Click
    
End Sub

Private Sub SetValues()

End Sub

Private Sub optAllDepts_Click()
    Call FormatWidths
    If optAllDepts.Value = True Then
        dtcDepts.Enabled = False
        dtcDepts.Text = Empty
    Else
        dtcDepts.Enabled = True
    End If
End Sub

Private Sub optAllStaff_Click()
    Call FormatWidths
    If optAllStaff.Value = True Then
        dtcStaff.Enabled = False
        dtcStaff.Text = Empty
    Else
        dtcStaff.Enabled = True
    End If
End Sub

Private Sub optSelectedDepts_Click()
    Call FormatWidths
    If optSelectedDepts.Value = True Then
        dtcDepts.Enabled = True
    Else
        dtcDepts.Enabled = False
        dtcDepts.Text = Empty
    End If
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
    With rsStore
        If .State = 1 Then .Close
        temSql = "SELECT * from tblStore order by store"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcDepts
        Set .RowSource = rsStore
        .ListField = "Store"
        .BoundColumn = "StoreID"
    End With
End Sub

Private Sub optSelectedStaff_Click()
    Call FormatWidths
    If optSelectedDepts.Value = True Then
        dtcStaff.Enabled = True
    Else
        dtcStaff.Enabled = False
        dtcStaff.Text = Empty
    End If
End Sub
