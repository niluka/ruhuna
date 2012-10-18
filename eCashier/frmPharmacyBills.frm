VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPharmacyBills 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pharmacy Bill Payments"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
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
   ScaleHeight     =   8325
   ScaleWidth      =   8055
   Begin VB.TextBox txtIncomeBillID 
      Height          =   360
      Left            =   7320
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtTotal 
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   7920
      Width           =   2055
   End
   Begin VB.OptionButton optAll 
      BackColor       =   &H00808000&
      Caption         =   "&All"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtPaymentMethod 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2040
      Width           =   4575
   End
   Begin VB.TextBox txtPayments 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   7440
      Width           =   2055
   End
   Begin VB.ListBox lstBills 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   7695
   End
   Begin VB.CheckBox chkPrint 
      BackColor       =   &H00808000&
      Caption         =   "&Print"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   7440
      Width           =   1455
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16776960
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
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2040
      TabIndex        =   10
      Top             =   1560
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnPay 
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16776960
      Caption         =   "&Pay"
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
   Begin VB.OptionButton optToSettle 
      BackColor       =   &H00808000&
      Caption         =   "&To Settle"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton optSettled 
      BackColor       =   &H00808000&
      Caption         =   "&Settled"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   900
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   75038723
      CurrentDate     =   39963
   End
   Begin MSComCtl2.DTPicker dtpFromTime 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   75038722
      CurrentDate     =   39963
   End
   Begin VB.ListBox lstIDs 
      Height          =   4380
      Left            =   6960
      TabIndex        =   21
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstValues 
      Height          =   4380
      Left            =   7440
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstPaid 
      Height          =   4380
      Left            =   6360
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtpToTime 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   75038722
      CurrentDate     =   39963
   End
   Begin MSDataListLib.DataCombo cmbUser 
      Height          =   360
      Left            =   2040
      TabIndex        =   24
      Top             =   2520
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808000&
      Caption         =   "Setled User"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808000&
      Caption         =   "&To"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      Caption         =   "Bill Value"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808000&
      Caption         =   "Total Value"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      Caption         =   "&Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Payment &Method"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "&From"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "&Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPharmacyBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FirstActi As Boolean
    Dim temSQL As String
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPay_Click()
    Dim rsTem As New ADODB.Recordset
    Dim DisplayBillID As Long
    
    Dim temText As String
    
    Dim tr As Integer
    
    temText = ""
    
    temText = temText & "Pharmmacy Bill ID" & vbTab & " : " & vbTab & lstIDs.List(lstBills.ListIndex) & vbNewLine
    temText = temText & "Bill Amount      " & vbTab & " : " & vbTab & Format(Val(lstValues.List(lstBills.ListIndex)), "#,##0.00") & vbNewLine & vbNewLine
    temText = temText & "Are you sure you want to pay this Bill?"
    
    tr = MsgBox(temText, vbYesNo)
    If tr = vbNo Then Exit Sub
    
    
    
    With rsTem
    
'        If .State = 1 Then .Close
'        temSql = "Select count(IncomeBillID) as BillCOunt from tblIncomeBill where IsPharmacyBill = 1 AND Completed = 1"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
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
        !IsPharmacyBill = True
        !Date = Date
        !Time = Now
        !UserID = UserID
        !StoreID = UserStoreID
        !Completed = True
        !CompletedDate = Date
        !CompletedTime = Now
        !CompletedUserID = UserID
'        !DisplayBillID = DisplayBillID
        !PharmacyBillID = Val(lstIDs.List(lstBills.ListIndex))
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        !PaymentComments = txtPaymentMethod.Text
        !GrossTotal = Val(lstValues.List(lstBills.ListIndex))
        !NetTotal = Val(lstValues.List(lstBills.ListIndex))
        !StoreID = UserStoreID
        .Update
        temSQL = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        txtIncomeBillID.Text = !NewID
        
        DisplayBillID = NewPharmacyDisplayBillID(Val(txtIncomeBillID.Text))
        
        If .State = 1 Then .Close
        
        temSQL = "Select * from tblSaleBill where SaleBillID = " & Val(lstIDs.List(lstBills.ListIndex))
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
            !PaidAtCashier = True
            !PaidAtCashierDate = Date
            !PaidAtCashierTime = Now
            !PaidAtCashierUserID = UserID
            .Update
        End If
        .Close
    
    
        If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeBill where IncomeBillID = " & Val(txtIncomeBillID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !DisplayBillID = DisplayBillID
            .Update
        End If
        .Close
    
    
    End With

    Call FillList

End Sub

Private Sub cmbPaymentMethod_Change()
    Call FillList
End Sub

Private Sub cmbUser_Change()
    Call FillList
End Sub

Private Sub cmbUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbUser.Text = Empty
    End If
End Sub

Private Sub dtpDate_Change()
    Call FillList
End Sub

Private Sub dtpTime_Change()
    Call FillList
End Sub

Private Sub Form_Activate()
    If FirstActi = True Then
        Call GetSettings
        FirstActi = False
    End If
End Sub

Private Sub Form_Load()
    FirstActi = True
    Call GetSettings
    Call FillCombos
    Call FillList
    
    cmbUser.BoundText = UserID
    
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    dtpToTime.Value = CDate(GetSetting(App.EXEName, Me.Name, dtpToTime.Name, "00:00"))
    dtpfromTime.Value = CDate(GetSetting(App.EXEName, Me.Name, dtpfromTime.Name, "00:00"))
    cmbPaymentMethod.BoundText = 1 ' Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
    chkPrint.Value = Val(GetSetting(App.EXEName, Me.Name, chkPrint.Name, 0))
    optSettled.Value = CBool(GetSetting(App.EXEName, Me.Name, optSettled.Name, False))
    optToSettle.Value = CBool(GetSetting(App.EXEName, Me.Name, optToSettle.Name, False))
    optAll.Value = CBool(GetSetting(App.EXEName, Me.Name, optAll.Name, False))
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, dtpfromTime.Name, dtpfromTime.Value
    SaveSetting App.EXEName, Me.Name, dtpToTime.Name, dtpToTime.Value
    SaveSetting App.EXEName, Me.Name, optSettled.Name, optSettled.Value
    SaveSetting App.EXEName, Me.Name, optToSettle.Name, optToSettle.Value
    SaveSetting App.EXEName, Me.Name, optAll.Name, optAll.Value
End Sub

Private Sub FillCombos()
    Dim PayM As New clsFillCombos
    PayM.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanReceive", False
    Dim MyUser As New clsFillCombos
    MyUser.FillBoolCombo cmbUser, "Staff", "Name", "IsAUser", False
End Sub

Private Sub FillList()
    lstIDs.Clear
    lstBills.Clear
    lstValues.Clear
    lstPaid.Clear
    
    lstBills.Visible = False
    
    Dim TotalValue As Double
    Dim rsTem As New ADODB.Recordset
    Dim temText As String
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT * "
        temSQL = temSQL & "From tblSaleBill "
        If optToSettle.Value = True Then
            temSQL = temSQL & "Where (((tblSaleBill.Cancelled) = 0 ) And ((tblSaleBill.PaidAtCashier) = 0 ) And ((tblSaleBill.Date) = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') And ((tblSaleBill.Time) >= '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpfromTime.Value & "') And ((tblSaleBill.Time) <= '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "') And ((tblSaleBill.PaymentMethodID) = " & Val(cmbPaymentMethod.BoundText) & "))"
        ElseIf optSettled.Value = True Then
            If IsNumeric(cmbUser.BoundText) = True Then
                temSQL = temSQL & "Where (((tblSaleBill.Cancelled) = 0 ) And ((tblSaleBill.PaidAtCashier) = 1 )  And ((tblSaleBill.PaidAtCashierUserID) = " & Val(cmbUser.BoundText) & " ) And ((tblSaleBill.Date) = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') And ((tblSaleBill.Time) >= '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpfromTime.Value & "') And ((tblSaleBill.Time) <= '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "') And ((tblSaleBill.PaymentMethodID) = " & Val(cmbPaymentMethod.BoundText) & "))"
            Else
                temSQL = temSQL & "Where (((tblSaleBill.Cancelled) = 0 ) And ((tblSaleBill.PaidAtCashier) = 1 ) And ((tblSaleBill.Date) = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') And ((tblSaleBill.Time) >= '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpfromTime.Value & "') And ((tblSaleBill.Time) <= '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "') And ((tblSaleBill.PaymentMethodID) = " & Val(cmbPaymentMethod.BoundText) & "))"
            End If
        Else
            temSQL = temSQL & "Where (((tblSaleBill.Cancelled) = 0 ) And ((tblSaleBill.Date) = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') And ((tblSaleBill.Time) >= '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpfromTime.Value & "') And ((tblSaleBill.Time) <= '" & dtpToTime.Value & "') And  ((tblSaleBill.PaymentMethodID) = " & Val(cmbPaymentMethod.BoundText) & "))"
        End If
        temSQL = temSQL & "ORDER BY tblSaleBill.SaleBillID"
        
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            temText = Left(!SaleBillID & Space(10), 12)
            temText = temText & vbTab
            temText = temText & Left(Format(!Time, "hh:mm AMPM") & Space(8), 8) & vbTab
            temText = temText & Right(Space(12) & Format(!NetPrice, "0.00"), 12) & vbTab
            If !PaidAtCashier = True Then
                temText = temText & "Paid on " & Format(!PaidAtCashierDate, "dd MMM yyyy")
                lstPaid.AddItem "True"
            Else
                temText = temText & "To Pay"
                lstPaid.AddItem "False"
            End If
            lstBills.AddItem temText
            lstIDs.AddItem !SaleBillID
            TotalValue = TotalValue + !NetPrice
            lstValues.AddItem !NetPrice
            .MoveNext
        Wend
        .Close
    End With
    
    txtTotal.Text = Format(TotalValue, "#,##0.00")
    
    lstBills.Visible = True
    
    If optToSettle.Value = True Then
        btnPay.Enabled = True
    Else
        btnPay.Enabled = False
    End If
    
    If lstBills.ListCount > 0 Then
        lstBills.Selected(0) = True
    End If
End Sub

Private Sub CalculateToPay()
    Dim n As Integer
    Dim ToPay As Double
    For n = 0 To lstIDs.ListCount - 1
        If lstBills.Selected(n) = True Then
            ToPay = ToPay + Val(lstValues.List(n))
        End If
    Next
    txtPayments.Text = Format(ToPay, "0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub lstBills_Click()
    Call CalculateToPay
End Sub

Private Sub lstBills_ItemCheck(Item As Integer)
    If lstPaid.List(Item) = "True" Then lstBills.Selected(Item) = False
End Sub



Private Sub optSettled_Click()
    Call FillList
End Sub

Private Sub optToSettle_Click()
    Call FillList
End Sub
