VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDistributorPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distributor Payment"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10755
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
   ScaleHeight     =   8970
   ScaleWidth      =   10755
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10455
      Begin btButtonEx.ButtonEx bttnUpdate 
         Height          =   375
         Left            =   8760
         TabIndex        =   50
         Top             =   7440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Update"
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
      Begin VB.ListBox lstTotal 
         Height          =   1980
         Left            =   8400
         TabIndex        =   47
         Top             =   2160
         Width           =   1335
      End
      Begin btButtonEx.ButtonEx bttnDelete 
         Height          =   375
         Left            =   8400
         TabIndex        =   46
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Delete Invoice"
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
      Begin VB.ListBox lstInvoiceNo 
         Height          =   1980
         ItemData        =   "frmDistributorPayment.frx":0000
         Left            =   7080
         List            =   "frmDistributorPayment.frx":0002
         TabIndex        =   44
         Top             =   2160
         Width           =   1200
      End
      Begin VB.TextBox txtPayAmount 
         Height          =   375
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   4920
         Width           =   5175
      End
      Begin MSFlexGridLib.MSFlexGrid msfInvoice 
         Height          =   2055
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo dtcDistributor 
         Height          =   360
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame framCheque 
         Height          =   2295
         Left            =   240
         TabIndex        =   25
         Top             =   5520
         Visible         =   0   'False
         Width           =   8415
         Begin VB.TextBox txtChequeNo 
            Height          =   375
            Left            =   4320
            TabIndex        =   26
            Top             =   480
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpChequeDate 
            Height          =   375
            Left            =   6720
            TabIndex        =   27
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   80805889
            CurrentDate     =   39537
         End
         Begin MSDataListLib.DataCombo dtcCity 
            Height          =   360
            Left            =   120
            TabIndex        =   28
            Top             =   1680
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcBank 
            Height          =   360
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcAccountNo 
            Height          =   360
            Left            =   120
            TabIndex        =   30
            Top             =   480
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcSignStaffId1 
            Height          =   360
            Left            =   4320
            TabIndex        =   31
            Top             =   1080
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcSignStaffId2 
            Height          =   360
            Left            =   4320
            TabIndex        =   32
            Top             =   1680
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label15 
            Caption         =   "Sign Staff Name 02"
            Height          =   255
            Left            =   4320
            TabIndex        =   39
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "Sign Staff Name 01"
            Height          =   255
            Left            =   4320
            TabIndex        =   38
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Cheqe No"
            Height          =   255
            Left            =   4320
            TabIndex        =   36
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "City"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label19 
            Caption         =   "Bank"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label20 
            Caption         =   "Date"
            Height          =   255
            Left            =   6720
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame framBankSlip 
         Height          =   2295
         Left            =   240
         TabIndex        =   10
         Top             =   5520
         Width           =   8415
         Begin VB.TextBox txtSlipNo 
            Height          =   375
            Left            =   5760
            TabIndex        =   12
            Top             =   480
            Width           =   2415
         End
         Begin MSDataListLib.DataCombo dtcSlipCity 
            Height          =   360
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker dtpSlipDate 
            Height          =   375
            Left            =   4200
            TabIndex        =   13
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   80805889
            CurrentDate     =   39537
         End
         Begin MSDataListLib.DataCombo dtcSlipBank 
            Height          =   360
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcSlipAccountNo 
            Height          =   360
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcSlipCheckStaff 
            Height          =   360
            Left            =   4200
            TabIndex        =   16
            Top             =   1680
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcSlipSingStaff 
            Height          =   360
            Left            =   4200
            TabIndex        =   17
            Top             =   1080
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label9 
            Caption         =   "Bank Name"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Branch Name"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Slip Date"
            Height          =   255
            Left            =   4200
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Check Staff Name"
            Height          =   255
            Left            =   4200
            TabIndex        =   19
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Sing Staff Name"
            Height          =   255
            Left            =   4200
            TabIndex        =   18
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label12 
            Caption         =   "Bank Slip No"
            Height          =   255
            Left            =   5760
            TabIndex        =   21
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.ComboBox cmbPaymode 
         Height          =   360
         Left            =   2520
         TabIndex        =   51
         Top             =   4440
         Width           =   5175
      End
      Begin MSDataListLib.DataCombo dtcPaymode 
         Height          =   360
         Left            =   2520
         TabIndex        =   41
         Top             =   4440
         Visible         =   0   'False
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lblPmode 
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Amount"
         Height          =   255
         Left            =   8400
         TabIndex        =   48
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "Invoice No"
         Height          =   255
         Left            =   7080
         TabIndex        =   45
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Pay Amount"
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label21 
         Caption         =   "Payment Mode"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Invoices"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblDueBalance 
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Total Due Balance"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblbillCountl 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "No Of Invoices"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Distributor Name"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   8880
      TabIndex        =   0
      Top             =   8520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
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
End
Attribute VB_Name = "frmDistributorPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsViewDistributor As New ADODB.Recordset
Dim rsViewPaymode As New ADODB.Recordset
Dim rsTem As New ADODB.Recordset
Dim rsViewAccount As New ADODB.Recordset
Dim rsViewBank As New ADODB.Recordset
Dim rsViewCity As New ADODB.Recordset
Dim rsViewSighStaff As New ADODB.Recordset
Dim TemIssueChequeID As Long
Dim TemIssueSlipID As Long
Dim Temsql As String
Dim I As Integer

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub Filldistributor()
With rsViewDistributor
    If .State = 1 Then .Close
    .Open "Select* From tblDistrubutor Order By DistributorName", cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    Set dtcDistributor.RowSource = rsViewDistributor
    dtcDistributor.BoundColumn = "DistributorID"
    dtcDistributor.ListField = "DistributorName"

End With
End Sub


Private Sub FillPayMode()

With rsViewPaymode
    If .State = 1 Then .Close
    Temsql = "Select * From tblPaymentMethod Order By PaymentMethod"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    Set dtcPaymode.RowSource = rsViewPaymode
    dtcPaymode.BoundColumn = "PaymentMethodID"
    dtcPaymode.ListField = "PaymentMethod"

End With
End Sub

Private Sub bttnDelete_Click()
lstInvoiceNo.Clear
lstTotal.Clear
txtPayAmount.Text = Empty
End Sub

Private Sub bttnUpdate_Click()
If dtcDistributor.BoundText = Empty Then A = MsgBox("Select Distributor Name", vbCritical + vbOKOnly, "Error"): dtcDistributor.SetFocus: SendKeys "{Home}+{End}": Exit Sub
If txtPayAmount.Text = Empty Then A = MsgBox("Add Paying Invoice to List Box", vbCritical + vbOKOnly, "Error"): msfInvoice.SetFocus: SendKeys "{Home}+{End}": Exit Sub
If dtcPaymode.BoundText = Empty Then A = MsgBox("Select Secon Pay Mode", vbCritical + vbOKOnly, "Error"): dtcPaymode.SetFocus: SendKeys "{Home}+{End}": Exit Sub

If CheckPaymentValues = False Then Exit Sub
Call UpdateRefillBill
If cmbPaymode.Text = "Cheque" Then Call SaveChequeDetails
If cmbPaymode.Text = "Slip" Then Call SaveSlipDetails
Call UpdateDistributorBalance
Call ClearVales
End Sub

Private Sub UpdateDistributorBalance()
With rsTem
    If .State = 1 Then .Close
    .Open "Select* From tblDistrubutor Where DistributorID = " & dtcDistributor.BoundText & "", cnnStores, adOpenStatic, adLockOptimistic
    If .RecordCount = 0 Then Exit Sub
    !balance = Val(!balance) - Val(txtPayAmount.Text)
    .Update
    If .State = 1 Then .Close
End With
End Sub

Private Sub UpdateRefillBill()
Dim TemBillID As Long
If lstInvoiceNo.ListCount < 1 Then Exit Sub

For TemBillID = 0 To lstInvoiceNo.ListCount - 1

    With rsTem
    If .State = 1 Then .Close
    Temsql = "Select * from tblRefillBill Where RefillBillID = " & Val(lstInvoiceNo.List(TemBillID)) & " "
    .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic

        If .RecordCount <> 0 Then
        !PaymentMethodID = Val(dtcPaymode.BoundText)
        !PaidStaffID = UserID
        !PaidDate = Date
        !PaidTime = Time
        !PaidPrice = Val(txtPayAmount.Text)
        If Val(!Price) = Val(!PaidPrice) Then
            !FullyPaid = True
        End If
        !ReceivedChequeID = Val(TemIssueChequeID)
        !ReceivedCreditCardID = 0
        !ReceivedSlipID = Val(TemIssueSlipID)
        .Update
        End If
    End With
Next

End Sub

Private Sub ClearVales()
lstInvoiceNo.Clear
lstTotal.Clear
cmbPaymode.Text = Empty
txtPayAmount.Text = Empty
dtcBank.BoundText = Empty
dtcCity.BoundText = Empty
dtcAccountNo.BoundText = Empty
dtcSignStaffId1.BoundText = Empty
dtcSignStaffId2.BoundText = Empty
txtChequeNo.Text = Empty
dtcSlipBank.BoundText = Empty
dtcSlipCity.BoundText = Empty
dtcSlipAccountNo.BoundText = Empty
txtSlipNo.Text = Empty
dtcSlipCheckStaff.BoundText = Empty
dtcSlipSingStaff.BoundText = Empty
End Sub

Private Sub cmbPaymode_Click()
If IsNull(cmbPaymode.Text) = True Then Exit Sub
dtcPaymode.Text = cmbPaymode.Text
End Sub

Private Sub dtcDistributor_Click(Area As Integer)
If IsNumeric(dtcDistributor.BoundText) = False Then Exit Sub
Call FillGrid
End Sub

Private Sub dtcPaymode_Change()
If IsNumeric(dtcPaymode.BoundText) = False Then Exit Sub

    
    If dtcPaymode.BoundText = 7 Then   'Slip
        framCheque.Enabled = False
        framBankSlip.Enabled = True
        framCheque.Visible = False
        framBankSlip.Visible = True
        lblPmode.Caption = dtcPaymode.Text

    ElseIf dtcPaymode.BoundText = 5 Then  'Cheque
        framCheque.Enabled = True
        framBankSlip.Enabled = False
        framCheque.Visible = True
        framBankSlip.Visible = False
        lblPmode.Caption = dtcPaymode.Text
    Else
        framCheque.Enabled = False
        framBankSlip.Enabled = False
        lblPmode.Caption = Empty
        
    End If

End Sub

Private Function CheckPaymentValues() As Boolean
Dim A As Byte
CheckPaymentValues = False

    Select Case dtcPaymode.BoundText
    
    Case 7 'Slip
        If dtcSlipBank.BoundText = Empty Then A = MsgBox("Select Slip Bank Name", vbCritical + vbOKOnly, "Error"): dtcSlipBank.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcSlipCity.BoundText = Empty Then A = MsgBox("Select Slip City Name", vbCritical + vbOKOnly, "Error"): dtcSlipCity.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcSlipAccountNo.BoundText = Empty Then A = MsgBox("Select Slip Account Name", vbCritical + vbOKOnly, "Error"): dtcSlipAccountNo.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If txtSlipNo.Text = Empty Then A = MsgBox("Enter Slip No", vbCritical + vbOKOnly, "Error"): txtSlipNo.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcSlipSingStaff.BoundText = Empty Then A = MsgBox("Select Sing Staff Name", vbCritical + vbOKOnly, "Error"): dtcSlipSingStaff.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcSlipCheckStaff.BoundText = Empty Then A = MsgBox("Select Check Staff Name", vbCritical + vbOKOnly, "Error"): dtcSlipCheckStaff.SetFocus: SendKeys "{Home}+{End}": Exit Function
    Case 5  'Cheque
        If dtcAccountNo.BoundText = Empty Then A = MsgBox("Select Account NO", vbCritical + vbOKOnly, "Error"): dtcAccountNo.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcBank.BoundText = Empty Then A = MsgBox("Select Bank Name", vbCritical + vbOKOnly, "Error"): dtcBank.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcCity.BoundText = Empty Then A = MsgBox("Select City Name", vbCritical + vbOKOnly, "Error"): dtcCity.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If txtChequeNo = Empty Then A = MsgBox("Enter Cheque No", vbCritical + vbOKOnly, "Error"): txtChequeNo.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcSignStaffId1.BoundText = Empty Then A = MsgBox("Select Sing Staff Name", vbCritical + vbOKOnly, "Error"): dtcSignStaffId1.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcSignStaffId2.BoundText = Empty Then A = MsgBox("Select Secon Singd Staff Name", vbCritical + vbOKOnly, "Error"): dtcSignStaffId2.SetFocus: SendKeys "{Home}+{End}": Exit Function
    End Select

CheckPaymentValues = True
End Function

Private Sub Form_Load()
Call FillPayModeCombo
Call Filldistributor
Call FillPayMode
Call FormartGrid
Call FillBankDetails

End Sub

Private Sub FillPayModeCombo()
cmbPaymode.AddItem "Cash"
cmbPaymode.AddItem "Cheque"
cmbPaymode.AddItem "Slip"

End Sub

Private Sub FillBankDetails()
With rsViewAccount
    If .State = 1 Then .Close
    Temsql = "Select * From tblBankAccount Order By AccountNo"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    Set dtcAccountNo.RowSource = rsViewAccount
    dtcAccountNo.BoundColumn = "AccountNoID"
    dtcAccountNo.ListField = "AccountNo"
    Set dtcSlipAccountNo.RowSource = rsViewAccount
    dtcSlipAccountNo.BoundColumn = "AccountNoID"
    dtcSlipAccountNo.ListField = "AccountNo"
   
End With


With rsViewBank
    If .State = 1 Then .Close
    Temsql = "Select * From tblBank Order By Bank"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    Set dtcBank.RowSource = rsViewBank
    dtcBank.BoundColumn = "BankID"
    dtcBank.ListField = "Bank"
    Set dtcSlipBank.RowSource = rsViewBank
    dtcSlipBank.BoundColumn = "BankID"
    dtcSlipBank.ListField = "Bank"
    
End With

With rsViewCity
    If .State = 1 Then .Close
    Temsql = "Select * From tblCity Order By City"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    Set dtcCity.RowSource = rsViewCity
    dtcCity.BoundColumn = "CityId"
    dtcCity.ListField = "City"
    Set dtcSlipCity.RowSource = rsViewCity
    dtcSlipCity.BoundColumn = "CityId"
    dtcSlipCity.ListField = "City"
End With


With rsViewSighStaff
    If .State = 1 Then .Close
    Temsql = "SELECT * from tblstaff order by name"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    
    Set dtcSignStaffId1.RowSource = rsViewSighStaff
    dtcSignStaffId1.ListField = "ListedName"
    dtcSignStaffId1.BoundColumn = "StaffID"
    
    Set dtcSignStaffId2.RowSource = rsViewSighStaff
    dtcSignStaffId2.ListField = "ListedName"
    dtcSlipCheckStaff.BoundColumn = "StaffID"
    
    Set dtcSlipCheckStaff.RowSource = rsViewSighStaff
    dtcSlipCheckStaff.ListField = "ListedName"
    dtcSlipCheckStaff.BoundColumn = "StaffID"
    
    Set dtcSlipSingStaff.RowSource = rsViewSighStaff
    dtcSlipSingStaff.ListField = "ListedName"
    dtcSlipSingStaff.BoundColumn = "StaffID"
    
End With


End Sub


Private Function FormartGrid()
With msfInvoice
    .Clear
    .Cols = 6
    .ColWidth(0) = 500
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + 350)
    
    .Row = 0
    .Col = 0
    .Text = "Id"
    .CellAlignment = 7
    .Col = 1
    .Text = "Invoice No"
    .Col = 2
    .Text = "Date"
    .Col = 3
    .Text = "Total"
    .CellAlignment = 7
    .Col = 4
    .Text = "Paid Amount"
    .Col = 5
    .Text = "Balance"
    .Rows = 1
End With
End Function

Private Function FillGrid()
Dim r As Long
Dim TotalDue As Double
With rsTem
If .State = 1 Then .Close
Temsql = "Select RefillBillID, Date, Price, PaidPrice, Balance,([Price]-[PaidPrice]) as Balance From tblRefillBill Where ((DistributorID = " & dtcDistributor.BoundText & ") and (Fullypaid = False))"
.Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
If .RecordCount = 0 Then: Call FormartGrid: Exit Function
r = 1

Do While .EOF = False
r = r + 1

    With msfInvoice
        .Rows = r
        .Row = r - 1
        .Col = 0
        .Text = r - 1
        .Col = 1
        .Text = rsTem!RefillBillID
        .Col = 2
        .Text = rsTem!Date
        .Col = 3
        .Text = Format(rsTem!Price, "#0.00")
        .Col = 4
        .Text = Format(rsTem!PaidPrice, "#0.00")
        .Col = 5
        .Text = Format(rsTem!balance, "#0.00")
    End With
    TotalDue = Val(TotalDue) + Val(!Price)


.MoveNext
Loop
lblbillCountl.Caption = .RecordCount
lblDueBalance.Caption = Format(TotalDue, "#0.00")
End With
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If rsTem.State = 1 Then rsViewDistributor.Close: Set rsViewDistributor = Nothing
If rsViewDistributor.State = 1 Then rsViewDistributor.Close: Set rsViewTem = Nothing
If rsViewPaymode.State = 1 Then rsViewPaymode.Close: Set rsViewPaymode = Nothing
If rsViewAccount.State = 1 Then rsViewAccount.Close: Set rsViewAccount = Nothing
If rsViewBank.State = 1 Then rsViewBank.Close: Set rsViewBank = Nothing
If rsViewCity.State = 1 Then rsViewCity.Close: Set rsViewCity = Nothing
If rsViewSighStaff.State = 1 Then rsViewSighStaff.Close: Set rsViewSighStaff = Nothing
'If rsviewCreditCard.State = 1 Then rsviewCreditCard.Close: Set rsviewCreditCard = Nothing
'If rsViewSighStaff.State = 1 Then rsViewSighStaff.Close: Set rsViewSighStaff = Nothing

End Sub

Private Sub SaveChequeDetails()
With rsTem
    If .State = 1 Then .Close
    Temsql = "Select * From tblIssueCheque"
    .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic
    
    .AddNew
    !IssueDate = Date
    !IssueTime = Time
    !IssuedStaffID = UserID
    !SignedStaff1ID = Val(dtcSignStaffId1.BoundText)
    !SignedStaff2ID = Val(dtcSignStaffId2.BoundText)
'    If OptionCoustomer.Value = True Then !IssuedToCustomerID = Val(dtcCombo.BoundText)
'    If OptionIndorPatient.Value = True Then !IssuedToBHTID = Val(dtcCombo.BoundText)
'    If optionStaff.Value = True Then !IssuedToStaffID = Val(dtcCombo.BoundText)
    !AccountID = Val(dtcAccountNo.BoundText)
    !BankID = Val(dtcBank.BoundText)
    !BranchID = Val(dtcCity.BoundText)
    !ChequeNo = txtChequeNo.Text
    !ChequeDate = dtpChequeDate.Value
    !Price = Val(txtPayAmount.Text)
    .Update
     TemIssueChequeID = !IssueChequeID
   
End With
End Sub

Private Sub SaveSlipDetails()
With rsTem
    If .State = 1 Then .Close
    Temsql = "Select * From tblIssueSlips"
    .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic
    
    .AddNew
    !IssueDate = Date
    !IssueTime = Time
    !IssuedStaffID = UserID
    !SignedStaff1ID = dtcSlipSingStaff.BoundText
    CheckedStaff1ID = dtcSlipCheckStaff.BoundText
'    If OptionCoustomer.Value = True Then !IssuedToCustomerID = Val(dtcCombo.BoundText)
'    If OptionIndorPatient.Value = True Then !IssuedToBHTID = Val(dtcCombo.BoundText)
'    If optionStaff.Value = True Then !IssuedToStaffID = Val(dtcCombo.BoundText)
    !AccountID = dtcSlipAccountNo.BoundText
    !BankID = dtcSlipBank.BoundText
    !BranchID = dtcSlipCity.BoundText
    !SlipNo = txtSlipNo.Text
    !SlipDate = dtpSlipDate.Value
    !Price = Val(txtPayAmount.Text)
    .Update
    TemIssueSlipID = !IssueChequeID
    
End With
End Sub


Private Sub lstInvoiceNo_Clik()
lstTotal.DataChanged = lstInvoiceNo.DataChanged
End Sub

Private Sub lstTotal_Click()
lstInvoiceNo.DataChanged = lstTotal.DataChanged

End Sub

Private Sub msfInvoice_Click()

If lstInvoiceNo.ListCount = 0 Then

    With msfInvoice
    .Col = 1
    lstInvoiceNo.AddItem .Text
    .Col = 3
    lstTotal.AddItem .Text
    .Col = 3
    txtPayAmount.Text = Val(txtPayAmount.Text) + Val(.Text)
    End With

Else
        With msfInvoice
        .Col = 1
       
            If AlreadyAdded(Val(.Text)) = False Then
                lstInvoiceNo.AddItem .Text
                .Col = 3
                lstTotal.AddItem .Text
                .Col = 3
                txtPayAmount.Text = Format(Val(txtPayAmount.Text) + Val(.Text), "#0.00")
            End If
        End With
End If

End Sub

Private Function AlreadyAdded(ItemID As Long) As Boolean
AlreadyAdded = True

        For I = 0 To lstInvoiceNo.ListCount - 1

        If Val(lstInvoiceNo.List(I)) = ItemID Then Exit Function
        
        Next
        
AlreadyAdded = False
End Function

