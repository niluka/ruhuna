VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExpense 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expense"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
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
   ScaleHeight     =   7380
   ScaleWidth      =   9210
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8895
      Begin btButtonEx.ButtonEx bttnUpdate 
         Height          =   375
         Left            =   6720
         TabIndex        =   15
         Top             =   6000
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
         TabIndex        =   14
         Top             =   3000
         Width           =   4335
      End
      Begin MSDataListLib.DataCombo dtcPayMode 
         Height          =   360
         Left            =   1920
         TabIndex        =   13
         Top             =   2520
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcCombo 
         Height          =   360
         Left            =   1920
         TabIndex        =   11
         Top             =   2040
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select Option"
         Height          =   1095
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   4335
         Begin VB.OptionButton OptionStaff 
            Caption         =   "Staff"
            Height          =   255
            Left            =   2040
            TabIndex        =   9
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton OptionCoustomer 
            Caption         =   "Customer"
            Height          =   255
            Left            =   2040
            TabIndex        =   8
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton OptionIndorPatient 
            Caption         =   "Indoor Patient"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton OptionGeneral 
            Caption         =   "General"
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
      Begin VB.Frame framBankSlip 
         Height          =   2295
         Left            =   360
         TabIndex        =   25
         Top             =   3600
         Width           =   8415
         Begin MSDataListLib.DataCombo dtcSlipCity 
            Height          =   360
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.TextBox txtSlipNo 
            Height          =   375
            Left            =   5760
            TabIndex        =   32
            Top             =   480
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker dtpSlipDate 
            Height          =   375
            Left            =   4200
            TabIndex        =   30
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20643841
            CurrentDate     =   39537
         End
         Begin MSDataListLib.DataCombo dtcSlipBank 
            Height          =   360
            Left            =   120
            TabIndex        =   27
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
            TabIndex        =   42
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
            TabIndex        =   44
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
            TabIndex        =   46
            Top             =   1080
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label5 
            Caption         =   "Sing Staff Name"
            Height          =   255
            Left            =   4200
            TabIndex        =   45
            Top             =   840
            Width           =   2655
         End
         Begin VB.Label Label16 
            Caption         =   "Check Staff Name"
            Height          =   255
            Left            =   4200
            TabIndex        =   43
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Bank Slip No"
            Height          =   255
            Left            =   5760
            TabIndex        =   31
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "Slip Date"
            Height          =   255
            Left            =   4200
            TabIndex        =   29
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Branch Name"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Bank Name"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame framCheque 
         Height          =   2295
         Left            =   360
         TabIndex        =   16
         Top             =   3600
         Visible         =   0   'False
         Width           =   8415
         Begin VB.TextBox txtChequeNo 
            Height          =   375
            Left            =   4320
            TabIndex        =   24
            Top             =   480
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpChequeDate 
            Height          =   375
            Left            =   6720
            TabIndex        =   23
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20643841
            CurrentDate     =   39537
         End
         Begin MSDataListLib.DataCombo dtcCity 
            Height          =   360
            Left            =   120
            TabIndex        =   20
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
            TabIndex        =   18
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
            TabIndex        =   34
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
            TabIndex        =   36
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
            TabIndex        =   38
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
            TabIndex        =   37
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Account No"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Cheqe No"
            Height          =   255
            Left            =   4320
            TabIndex        =   21
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "City"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Bank"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Date"
            Height          =   255
            Left            =   6720
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label Label18 
         Caption         =   "Amount"
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblPmode 
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Payment Mode"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Expense From"
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
      Left            =   6840
      TabIndex        =   0
      Top             =   6840
      Width           =   1695
      _ExtentX        =   2990
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
Attribute VB_Name = "frmExpense"
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
Dim rsViewSighStaff As New ADODB.Recordset
Dim rsViewAccount As New ADODB.Recordset

Dim Temsql As String
Dim TemId As Long

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub FillExpenceCategory()

With rsViewCategory
    If .State = 1 Then .Close
    Temsql = "Select * From tblExpenceCatogery Order By ExpenceCatogery"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    Set dtcCategory.RowSource = rsViewCategory
    dtcCategory.BoundColumn = "ExpenceCatogeryID"
    dtcCategory.ListField = "ExpenceCatogery"

End With
End Sub

Private Sub FillPayMode()

With rsViewPaymode
    If .State = 1 Then .Close
    Temsql = "Select * From tblPaymentMethod Order By PaymentMethod"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    Set dtcPayMode.RowSource = rsViewPaymode
    dtcPayMode.BoundColumn = "PaymentMethodID"
    dtcPayMode.ListField = "PaymentMethod"

End With
End Sub

Private Sub bttnUpdate_Click()
If CheckValue = False Then Exit Sub
If CheckPaymentValues = False Then Exit Sub
Call SaveExpence
If dtcPayMode.BoundText = 5 Then Call SaveChequeDetails
If dtcPayMode.BoundText = 7 Then Call SaveSlipDetails
Call ClearValues
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
    If OptionCoustomer.Value = True Then !IssuedToCustomerID = Val(dtcCombo.BoundText)
    If OptionIndorPatient.Value = True Then !IssuedToBHTID = Val(dtcCombo.BoundText)
    If OptionStaff.Value = True Then !IssuedToStaffID = Val(dtcCombo.BoundText)
    !AccountID = Val(dtcAccountNo.BoundText)
    !ChequeBankID = Val(dtcBank.BoundText)
    !BranchID = Val(dtcCity.BoundText)
    !ChequeNo = txtChequeNo.Text
    !ChequeDate = dtpChequeDate.Value
    !Price = Val(txtAmount.Text)
    .Update
    
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
    If OptionCoustomer.Value = True Then !IssuedToCustomerID = Val(dtcCombo.BoundText)
    If OptionIndorPatient.Value = True Then !IssuedToBHTID = Val(dtcCombo.BoundText)
    If OptionStaff.Value = True Then !IssuedToStaffID = Val(dtcCombo.BoundText)
    !AccountID = dtcSlipAccountNo.BoundText
    !BankID = dtcSlipBank.BoundText
    !BranchID = dtcSlipCity.BoundText
    !SlipNo = txtSlipNo.Text
    !SlipDate = dtpSlipDate.Value
    !Price = txtAmount.Text
    Update
    
End With
End Sub


Private Sub SaveExpence()
With rsTem
    If .State = 1 Then .Close
    Temsql = "Select * From tblExpence"
    .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic

    .AddNew
    !ExpenceCatogeryID = Val(dtcCategory.BoundText)
    !Time = Time
    !Date = Date
    !StaffID = UserID
    !StoreID = UserStoreID
    !Price = Val(txtAmount.Text)
    !PaymentMethodID = Val(dtcPayMode.BoundText)
    If OptionCoustomer.Value = True Then !FromOutPatientID = Val(dtcCombo.BoundText)
    If OptionIndorPatient.Value = True Then !FromBHTID = Val(dtcCombo.BoundText)
    If OptionStaff.Value = True Then !FromStaffID = Val(dtcCombo.BoundText)
    .Update
    
    If .State = 1 Then .Close
End With

End Sub


Private Function CheckValue() As Boolean
Dim A As Byte
CheckValue = False
    If dtcCategory.BoundText = Empty Then A = MsgBox("Select Income Category", vbCritical + vbOKOnly, "Error"): dtcCategory.SetFocus: SendKeys "{Home}+{End}": Exit Function
    If OptionGeneral.Value = False And OptionIndorPatient.Value = False And OptionCoustomer.Value = False And OptionStaff.Value = False Then A = MsgBox("Select Expence Option", vbCritical + vbOKOnly, "Error"): Exit Function
    If OptionIndorPatient.Value = True And dtcCombo.BoundText = Empty Then A = MsgBox("Select Indoor Patient Name ", vbCritical + vbOKOnly, "Error"): dtcCombo.SetFocus: SendKeys "{Home}+{End}": Exit Function
    If OptionCoustomer.Value = True And dtcCombo.BoundText = Empty Then A = MsgBox("Select Coustomer Name", vbCritical + vbOKOnly, "Error"): dtcCombo.SetFocus: SendKeys "{Home}+{End}": Exit Function
    If OptionStaff.Value = True And dtcCombo.BoundText = Empty Then A = MsgBox("Select Staff Name", vbCritical + vbOKOnly, "Error"): dtcCombo.SetFocus: SendKeys "{Home}+{End}": Exit Function
    If dtcPayMode.BoundText = Empty Then A = MsgBox("Select Payment Mode", vbCritical + vbOKOnly, "Error"): dtcPayMode.SetFocus: SendKeys "{Home}+{End}": Exit Function
    If Val(txtAmount.Text) <= 0 Then A = MsgBox("Enter Amount", vbCritical + vbOKOnly, "Error"): txtAmount.SetFocus: SendKeys "{Home}+{End}": Exit Function
CheckValue = True
End Function

Private Function CheckPaymentValues() As Boolean
Dim A As Byte
CheckPaymentValues = False

    Select Case dtcPayMode.BoundText
    
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

Private Sub ClearValues()
    dtcCategory.BoundText = Empty
    OptionGeneral.Value = False
    OptionIndorPatient.Value = False
    OptionCoustomer.Value = False
    OptionStaff.Value = False
    dtcCombo.BoundText = Empty
    dtcPayMode.BoundText = Empty
    txtAmount.Text = Empty
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
    dtcCombo.Enabled = False
End Sub

Private Sub dtcPaymode_Change()
If IsNumeric(dtcPayMode.BoundText) = False Then Exit Sub

    
    If dtcPayMode.BoundText = 7 Then   'Slip
        framCheque.Enabled = False
        framBankSlip.Enabled = True
        framCheque.Visible = False
        framBankSlip.Visible = True
        lblPmode.Caption = dtcPayMode.Text

    ElseIf dtcPayMode.BoundText = 5 Then  'Cheque
        framCheque.Enabled = True
        framBankSlip.Enabled = False
        framCheque.Visible = True
        framBankSlip.Visible = False
        lblPmode.Caption = dtcPayMode.Text
    Else
        framCheque.Enabled = False
        framBankSlip.Enabled = False
        lblPmode.Caption = Empty
    End If
    
End Sub

Private Sub Form_Load()
Call FillPayMode
Call FillExpenceCategory
Call FillBankDetails
Call FillBankAccount
dtcCombo.Enabled = False
End Sub

Private Sub FillBankAccount()
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If rsTem.State = 1 Then rsTem.Close: Set rsTem = Nothing
If rsViewTem.State = 1 Then rsViewTem.Close: Set rsViewTem = Nothing
If rsViewPaymode.State = 1 Then rsViewPaymode.Close: Set rsViewPaymode = Nothing
If rsViewCategory.State = 1 Then rsViewCategory.Close: Set rsViewCategory = Nothing
If rsViewSighStaff.State = 1 Then rsViewSighStaff.Close: Set rsViewSighStaff = Nothing
If rsViewBank.State = 1 Then rsViewBank.Close: Set rsViewBank = Nothing
If rsViewCity.State = 1 Then rsViewCity.Close: Set rsViewCity = Nothing
If rsviewCreditCard.State = 1 Then rsviewCreditCard.Close: Set rsviewCreditCard = Nothing
If rsViewSighStaff.State = 1 Then rsViewSighStaff.Close: Set rsViewSighStaff = Nothing

End Sub

Private Sub OptionCoustomer_Click()
lblName.Caption = "Customer"
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

Private Sub OptionGeneral_Click()
dtcCombo.Enabled = False
lblName.Caption = "General"
End Sub
