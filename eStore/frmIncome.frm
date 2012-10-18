VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmIncome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income Receive"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
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
   ScaleWidth      =   9330
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8895
      Begin btButtonEx.ButtonEx bttnUpdate 
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   6600
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.TextBox txtAmount 
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   3360
         Width           =   4335
      End
      Begin MSDataListLib.DataCombo dtcPayMode 
         Height          =   360
         Left            =   1920
         TabIndex        =   14
         Top             =   2880
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcCombo 
         Height          =   360
         Left            =   1920
         TabIndex        =   12
         Top             =   2400
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select Option"
         Height          =   1455
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   4335
         Begin VB.OptionButton OptionStaff 
            Caption         =   "Staff"
            Height          =   255
            Left            =   2040
            TabIndex        =   10
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton OptionCoustomer 
            Caption         =   "Customer"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton OptionIndorPatient 
            Caption         =   "Indoor Patient"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton OptionDistributor 
            Caption         =   "Distributor"
            Height          =   255
            Left            =   2040
            TabIndex        =   7
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton OptionUnKnown 
            Caption         =   "Unknown"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1695
         End
      End
      Begin MSDataListLib.DataCombo dtcCategory 
         Height          =   360
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame framCreditcard 
         Height          =   2295
         Left            =   240
         TabIndex        =   27
         Top             =   4080
         Width           =   8415
         Begin VB.TextBox txtAuthorizationCode 
            Height          =   375
            Left            =   1680
            TabIndex        =   35
            Top             =   1680
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker dtpCreditCardExpireDate 
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   1680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39537
         End
         Begin VB.TextBox txtCreditCardNo 
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   3975
         End
         Begin MSDataListLib.DataCombo dtcCreditCard 
            Height          =   360
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label12 
            Caption         =   "Authorization Code"
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "Expiry Date"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Credit Card No"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Credit Card Name"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame framCheque 
         Height          =   2295
         Left            =   240
         TabIndex        =   18
         Top             =   4080
         Width           =   8415
         Begin VB.TextBox txtChequeNo 
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   1680
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpChequeDate 
            Height          =   375
            Left            =   2520
            TabIndex        =   25
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39537
         End
         Begin MSDataListLib.DataCombo dtcCity 
            Height          =   360
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcBank 
            Height          =   360
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label8 
            Caption         =   "Date"
            Height          =   255
            Left            =   2520
            TabIndex        =   24
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Cheqe No"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "City"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Bank"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame framBankSlip 
         Height          =   2295
         Left            =   240
         TabIndex        =   37
         Top             =   4080
         Visible         =   0   'False
         Width           =   8415
         Begin VB.TextBox txtSlipNo 
            Height          =   375
            Left            =   1680
            TabIndex        =   39
            Top             =   1680
            Width           =   2415
         End
         Begin MSDataListLib.DataCombo dtcSlipCity 
            Height          =   360
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker dtpSlipDate 
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39537
         End
         Begin MSDataListLib.DataCombo dtcSlipBank 
            Height          =   360
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcSlipCheckStaff 
            Height          =   360
            Left            =   4200
            TabIndex        =   42
            Top             =   480
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label18 
            Caption         =   "Branch Name"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Slip Date"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Bank Slip No"
            Height          =   255
            Left            =   1680
            TabIndex        =   45
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label17 
            Caption         =   "Bank Name"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Check Staff Name"
            Height          =   255
            Left            =   4200
            TabIndex        =   43
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Label lblPmode 
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Amount"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Payment Mode"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Income From"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Income Category"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   7440
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
Attribute VB_Name = "frmIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTem As New ADODB.Recordset
Dim rsViewTem As New ADODB.Recordset
Dim rsViewPaymode As New ADODB.Recordset
Dim rsViewCategory As New ADODB.Recordset
Dim rsViewBank As New ADODB.Recordset
Dim rsViewCity As New ADODB.Recordset
Dim rsviewCreditCard As New ADODB.Recordset
Dim rsViewStaff As New ADODB.Recordset
Dim Temsql As String
Dim TemId As Long

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub FillIncomeCategory()

With rsViewCategory
    If .State = 1 Then .Close
    Temsql = "Select * From tblIncomeCatogery Order By IncomeCatogery"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    Set dtcCategory.RowSource = rsViewCategory
    dtcCategory.BoundColumn = "IncomeCatogeryID"
    dtcCategory.ListField = "IncomeCatogery"

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

Private Sub bttnUpdate_Click()
If CheckValue = False Then Exit Sub
If CheckPaymentValues = False Then Exit Sub
Call SaveIncome
If dtcPaymode.BoundText = 3 Then Call SaveCreditCarddetails
If dtcPaymode.BoundText = 5 Then Call SaveChequeDetails
If dtcPaymode.BoundText = 7 Then Call SaveSlipDetails
Call ClearValues
End Sub

Private Sub SaveSlipDetails()
With rsTem
    If .State = 1 Then .Close
    Temsql = "Select * From tblReceivedSlips"
    .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic


    .AddNew
    !ReceivedDate = Date
    !ReceivedTime = Time
    !ReceivedSTaffID = UserID
    !Price = Val(txtAmount.Text)
    !SlipDate = dtpSlipDate.Value
    !BankID = Val(dtcSlipBank.BoundText)
    !BranchID = Val(dtcSlipCity.BoundText)
    !SlipNo = txtSlipNo.Text
    If OptionCoustomer.Value = True Then !ReceivedBilledOutPatientID = Val(dtcCombo.BoundText)
    If OptionIndorPatient.Value = True Then !ReceivedFromBHTID = Val(dtcCombo.BoundText)
    If optionStaff.Value = True Then !ReceivedFromStaffID = Val(dtcCombo.BoundText)
    If OptionDistributor.Value = True Then !ReceivedFromDistributorID = Val(dtcCombo.BoundText)
    
    .Update

End With

End Sub

Private Sub SaveChequeDetails()
With rsTem
    If .State = 1 Then .Close
    Temsql = "Select * From tblReceiveCheque"
    .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic
    
    .AddNew
    !ReceivedDate = Date
    !ReceivedTime = Time
    !ReceivedSTaffID = UserID
    !Price = Val(txtAmount.Text)
    !ChequeDate = dtpChequeDate.Value
    !BankID = Val(dtcBank.BoundText)
    !BranchID = Val(dtcCity.BoundText)
    !ChequeNo = txtChequeNo.Text
    If OptionCoustomer.Value = True Then !ReceivedBilledOutPatientID = Val(dtcCombo.BoundText)
    'BilledInPatientID
    If OptionIndorPatient.Value = True Then !ReceivedFromBHTID = Val(dtcCombo.BoundText)
    If optionStaff.Value = True Then !ReceivedFromStaffID = Val(dtcCombo.BoundText)
    If OptionDistributor.Value = True Then !ReceivedFromDistributorID = Val(dtcCombo.BoundText)
    .Update
    
End With
End Sub

Private Sub SaveCreditCarddetails()
With rsTem
    If .State = 1 Then .Close
    Temsql = "Select * From tblReceivedCreditCard"
    .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic
    
    .AddNew
    !CreditCardNo = txtCreditCardNo.Text
    !ReceivedSTaffID = UserID
    !CardTypeID = Val(dtcCreditCard.BoundText)
    !AuthrizationCode = txtAuthorizationCode.Text
    !AuthrizationDate = dtpCreditCardExpireDate.Value
    !AuthrizationTime = Time
    !AuthrizationStaffID = UserID
    If OptionCoustomer.Value = True Then !ReceivedBilledOutPatientID = Val(dtcCombo.BoundText)
    'BilledInPatientID
    If OptionIndorPatient.Value = True Then !ReceivedFromBHTID = Val(dtcCombo.BoundText)
    If optionStaff.Value = True Then !ReceivedFromStaffID = Val(dtcCombo.BoundText)
    If OptionDistributor.Value = True Then !ReceivedFromDistributorID = Val(dtcCombo.BoundText)
    .Update
    
End With
End Sub


Private Sub SaveIncome()
With rsTem
    If .State = 1 Then .Close
    Temsql = "Select * From tblIncome"
    .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic

    .AddNew
    !IncomeCatogeryID = Val(dtcCategory.BoundText)
    !Time = Time
    !Date = Date
    !StaffID = UserID
    !StoreID = UserStoreID
    !Price = Val(txtAmount.Text)
    !PaymentModeID = Val(dtcPaymode.BoundText)
    If OptionCoustomer.Value = True Then !BilledOutPatientID = Val(dtcCombo.BoundText)
    If OptionIndorPatient.Value = True Then !BilledBHTID = Val(dtcCombo.BoundText)
    If optionStaff.Value = True Then !BilledStaffID = Val(dtcCombo.BoundText)
    If OptionDistributor.Value = True Then !BilledDistributorID = Val(dtcCombo.BoundText)
    .Update
    
    If .State = 1 Then .Close
End With

End Sub


Private Function CheckValue() As Boolean
Dim A As Byte
CheckValue = False
If dtcCategory.BoundText = Empty Then A = MsgBox("Select Income Category", vbCritical + vbOKOnly, "Error"): dtcCategory.SetFocus: SendKeys "{Home}+{End}": Exit Function
If OptionUnKnown.Value = False And OptionIndorPatient.Value = False And OptionCoustomer.Value = False And OptionDistributor.Value = False And optionStaff.Value = False Then A = MsgBox("Select Income Option", vbCritical + vbOKOnly, "Error"): Exit Function
If OptionIndorPatient.Value = True And dtcCombo.BoundText = Empty Then A = MsgBox("Select Indoor Patient Name ", vbCritical + vbOKOnly, "Error"): dtcCombo.SetFocus: SendKeys "{Home}+{End}": Exit Function
If OptionCoustomer.Value = True And dtcCombo.BoundText = Empty Then A = MsgBox("Select Coustomer Name", vbCritical + vbOKOnly, "Error"): dtcCombo.SetFocus: SendKeys "{Home}+{End}": Exit Function
If OptionDistributor.Value = True And dtcCombo.BoundText = Empty Then A = MsgBox("Select Distributor Name", vbCritical + vbOKOnly, "Error"): dtcCombo.SetFocus: SendKeys "{Home}+{End}": Exit Function
If optionStaff.Value = True And dtcCombo.BoundText = Empty Then A = MsgBox("Select Staff Name", vbCritical + vbOKOnly, "Error"): dtcCombo.SetFocus: SendKeys "{Home}+{End}": Exit Function
If dtcPaymode.BoundText = Empty Then A = MsgBox("Select Payment Mode", vbCritical + vbOKOnly, "Error"): dtcPaymode.SetFocus: SendKeys "{Home}+{End}": Exit Function
If Val(txtAmount.Text) <= 0 Then A = MsgBox("Enter Amount", vbCritical + vbOKOnly, "Error"): txtAmount.SetFocus: SendKeys "{Home}+{End}": Exit Function
CheckValue = True
End Function

Private Function CheckPaymentValues() As Boolean
Dim A As Byte
CheckPaymentValues = False

    Select Case dtcPaymode.BoundText
    
    Case 3 'Credit Card
        If dtcCreditCard.BoundText = Empty Then A = MsgBox("Select Credit Card Name", vbCritical + vbOKOnly, "Error"): dtcCreditCard.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If txtCreditCardNo = Empty Then A = MsgBox("Enter Credit Card No", vbCritical + vbOKOnly, "Error"): txtCreditCardNo.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If txtAuthorizationCode = Empty Then A = MsgBox("Enter Credit Card Aurthorization Code", vbCritical + vbOKOnly, "Error"): txtAuthorizationCode.SetFocus: SendKeys "{Home}+{End}": Exit Function
    Case 5  'Cheque
        If dtcBank.BoundText = Empty Then A = MsgBox("Select Bank Name", vbCritical + vbOKOnly, "Error"): dtcBank.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcCity.BoundText = Empty Then A = MsgBox("Select City Name", vbCritical + vbOKOnly, "Error"): dtcCity.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If txtChequeNo = Empty Then A = MsgBox("Enter Cheque No", vbCritical + vbOKOnly, "Error"): txtChequeNo.SetFocus: SendKeys "{Home}+{End}": Exit Function
    Case 7  'Slip
        If dtcSlipBank.BoundText = Empty Then A = MsgBox("Select Bank Name", vbCritical + vbOKOnly, "Error"): dtcSlipBank.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcSlipCity.BoundText = Empty Then A = MsgBox("Select City Name", vbCritical + vbOKOnly, "Error"): dtcSlipCity.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If txtSlipNo.Text = Empty Then A = MsgBox("Enter Slip No", vbCritical + vbOKOnly, "Error"): txtSlipNo.SetFocus: SendKeys "{Home}+{End}": Exit Function
        If dtcSlipCheckStaff.BoundText = Empty Then A = MsgBox("Select Check Staff Name", vbCritical + vbOKOnly, "Error"): dtcSlipCheckStaff.SetFocus: SendKeys "{Home}+{End}": Exit Function
    End Select

CheckPaymentValues = True
End Function

Private Sub ClearValues()
dtcCategory.BoundText = Empty
OptionUnKnown.Value = False
OptionIndorPatient.Value = False
OptionCoustomer.Value = False
OptionDistributor.Value = False
optionStaff.Value = False
dtcCombo.BoundText = Empty
dtcPaymode.BoundText = Empty
txtAmount.Text = Empty
dtcBank.BoundText = Empty
dtcCity.BoundText = Empty
txtCreditCardNo = Empty
txtAuthorizationCode = Empty
txtChequeNo = Empty
dtcCombo.Enabled = False
txtAmount.Text = Empty
dtcSlipBank.BoundText = Empty
dtcSlipCity.BoundText = Empty
txtSlipNo.Text = Empty
dtcSlipCheckStaff.BoundText = Empty
End Sub

Private Sub dtcPayMode_Change()
If IsNumeric(dtcPaymode.BoundText) = False Then Exit Sub

    If dtcPaymode.BoundText = 3 Then 'Credit Card

        framCheque.Enabled = False
        framCreditcard.Enabled = True
        framCheque.Visible = False
        framCreditcard.Visible = True
        framBankSlip.Visible = False
        framBankSlip.Enabled = False

        lblPmode.Caption = dtcPaymode.Text
        
    ElseIf dtcPaymode.BoundText = 5 Then  'Cheque
    
        framCheque.Enabled = True
        framCreditcard.Enabled = False
        framCheque.Visible = True
        framCreditcard.Visible = False
        framBankSlip.Visible = False
        framBankSlip.Enabled = False

        lblPmode.Caption = dtcPaymode.Text
        
    ElseIf dtcPaymode.BoundText = 7 Then  'Slips
    
        framCheque.Enabled = False
        framCreditcard.Enabled = False
        framCheque.Visible = False
        framCreditcard.Visible = False
        framBankSlip.Visible = True
        framBankSlip.Enabled = True
        lblPmode.Caption = dtcPaymode.Text

    Else
        framCheque.Enabled = False
        framCreditcard.Enabled = False
        lblPmode.Caption = Empty
        framBankSlip.Enabled = False

    End If


End Sub

Private Sub Form_Load()
Call FillPayMode
Call FillIncomeCategory
Call FillBankDetails
Call Fillstaff
FillCreditCardDetails
dtcCombo.Enabled = False
End Sub

Private Sub Fillstaff()

With rsViewStaff
    If .State = 1 Then .Close
    Temsql = "SELECT * from tblstaff order by name"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
       
    Set dtcSlipCheckStaff.RowSource = rsViewStaff
    dtcSlipCheckStaff.ListField = "ListedName"
    dtcSlipCheckStaff.BoundColumn = "StaffID"
    
    
End With
End Sub


Private Sub FillBankDetails()
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
End Sub

Private Sub FillCreditCardDetails()
With rsviewCreditCard
    If .State = 1 Then .Close
    Temsql = "Select * From tblCreditCardType Order By CreditCardType"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    Set dtcCreditCard.RowSource = rsviewCreditCard
    dtcCreditCard.BoundColumn = "CreditCardTypeID"
    dtcCreditCard.ListField = "CreditCardType"
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If rsTem.State = 1 Then rsTem.Close: Set rsTem = Nothing
If rsViewTem.State = 1 Then rsViewTem.Close: Set rsViewTem = Nothing
If rsViewPaymode.State = 1 Then rsViewPaymode.Close: Set rsViewPaymode = Nothing
If rsViewCategory.State = 1 Then rsViewCategory.Close: Set rsViewCategory = Nothing
If rsViewStaff.State = 1 Then rsViewStaff.Close: Set rsViewStaff = Nothing

End Sub

Private Sub OptionCoustomer_Click()
lblName.Caption = "Coustomer"
dtcCombo.Enabled = True
dtcCombo.Text = Empty
     With rsViewTem
    
        If .State = 1 Then .Close
        Temsql = "Select * From tblPatientMainDetails Order By FirstName"
        .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
        Set dtcCombo.RowSource = rsViewTem
        dtcCombo.BoundColumn = "patient_ID"
        dtcCombo.ListField = "firstname"

    End With
    
End Sub

Private Sub OptionDistributor_Click()
lblName.Caption = "Distributor"
dtcCombo.Enabled = True
dtcCombo.Text = Empty

    With rsViewTem
        If .State = 1 Then .Close
        Temsql = "SELECT * from tblDistrubutor order by DistributorName"
        .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
        
        Set dtcCombo.RowSource = rsViewTem
        dtcCombo.ListField = "DistributorName"
        dtcCombo.BoundColumn = "DistributorID"
        
    End With

End Sub

Private Sub OptionIndorPatient_Click()
lblName.Caption = "Indoor Patient"
dtcCombo.Enabled = True
dtcCombo.Text = Empty

    With rsViewTem
        If .State = 1 Then .Close
        Temsql = "SELECT tblBHT.*, tblPatientMainDetails.FirstName FROM tblPatientMainDetails INNER JOIN tblBHT ON tblPatientMainDetails.Patient_ID = tblBHT.PatientID Where (((tblBHT.Discharge) = False))ORDER BY tblPatientMainDetails.FirstName"
        .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
         Set dtcCombo.RowSource = rsViewTem
        dtcCombo.BoundColumn = "BHTID"
        dtcCombo.ListField = "firstname"
 
    End With

End Sub

Private Sub OptionStaff_Click()
lblName.Caption = "Staff"
dtcCombo.Enabled = True
dtcCombo.Text = Empty

    With rsViewTem
        If .State = 1 Then .Close
        Temsql = "SELECT * from tblstaff order by name"
        .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
        Set dtcCombo.RowSource = rsViewTem
        dtcCombo.ListField = "ListedName"
        dtcCombo.BoundColumn = "StaffID"
    End With

'b = CalculateConsumption(5, 1 / 3 / 2008, 31 / 3 / 2008, 100)
'MsgBox b
End Sub

Private Sub OptionUnKnown_Click()
lblName.Caption = "Unknown"
dtcCombo.Enabled = False
End Sub
