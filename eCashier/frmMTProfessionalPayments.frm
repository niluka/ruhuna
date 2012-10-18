VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMTProfessionalPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Professional Fee Payments for Medical Tests"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11220
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
   ScaleHeight     =   8250
   ScaleWidth      =   11220
   Begin VB.TextBox txtCount 
      Height          =   375
      Left            =   9240
      TabIndex        =   36
      Top             =   6000
      Width           =   1815
   End
   Begin btButtonEx.ButtonEx btnAll 
      Height          =   375
      Left            =   1200
      TabIndex        =   35
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&All"
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
   Begin VB.TextBox txtTotal 
      Height          =   375
      Left            =   9240
      TabIndex        =   33
      Top             =   5520
      Width           =   1815
   End
   Begin VB.OptionButton optCancelled 
      Caption         =   "Payment Cancelled"
      Height          =   375
      Left            =   8160
      TabIndex        =   31
      Top             =   1560
      Width           =   1935
   End
   Begin VB.OptionButton optPaid 
      Caption         =   "Paid"
      Height          =   375
      Left            =   7080
      TabIndex        =   30
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton optToPay 
      Caption         =   "To Pay"
      Height          =   375
      Left            =   5760
      TabIndex        =   29
      Top             =   1560
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "&Print"
      Height          =   255
      Left            =   4920
      TabIndex        =   19
      Top             =   6480
      Width           =   735
   End
   Begin btButtonEx.ButtonEx btnPay 
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   6480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
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
   Begin VB.TextBox txtPayments 
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   6000
      Width           =   3015
   End
   Begin VB.TextBox txtComments 
      Height          =   375
      Left            =   1800
      TabIndex        =   17
      Top             =   6960
      Width           =   3015
   End
   Begin VB.ListBox lstPayments 
      Height          =   3300
      Left            =   1200
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   2040
      Width           =   9855
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   7560
      TabIndex        =   9
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSpeciality 
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff 
      Height          =   360
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   75104259
      CurrentDate     =   39960
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   75104259
      CurrentDate     =   39960
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   1800
      TabIndex        =   16
      Top             =   6480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   9480
      TabIndex        =   20
      Top             =   7080
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.ListBox lstIDs 
      Height          =   3180
      Left            =   8640
      TabIndex        =   25
      Top             =   2040
      Width           =   375
   End
   Begin VB.ListBox lstPaid 
      Height          =   3180
      Left            =   9240
      TabIndex        =   27
      Top             =   2040
      Width           =   375
   End
   Begin VB.ListBox lstValues 
      Height          =   3180
      Left            =   9240
      TabIndex        =   26
      Top             =   2040
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   7440
      Width           =   10695
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   6600
         TabIndex        =   24
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label17 
         Caption         =   "Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Paper"
         Height          =   255
         Left            =   6000
         TabIndex        =   23
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSDataListLib.DataCombo cmbPtPM 
      Height          =   360
      Left            =   7560
      TabIndex        =   11
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnCancelPayment 
      Height          =   375
      Left            =   5760
      TabIndex        =   32
      Top             =   6960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Cancel Payment"
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   375
      Left            =   5760
      TabIndex        =   38
      Top             =   6000
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Print"
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
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   855
      Left            =   3840
      TabIndex        =   39
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1508
      Appearance      =   3
      Caption         =   "P&rocess"
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
   Begin VB.Label Label10 
      Caption         =   "Count"
      Height          =   255
      Left            =   8280
      TabIndex        =   37
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Total"
      Height          =   255
      Left            =   8280
      TabIndex        =   34
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Payment Method"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Payment Method"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Total Payment"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "From "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Service Category"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Doctor"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Speciality"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMTProfessionalPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim temSql As String
    Dim rsStaff As New ADODB.Recordset
    
    Dim CsetPrinter As New cSetDfltPrinter
    Dim NumForms As Long, i As Long
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
    Dim SuppliedWord As String
    
    
Private Sub btnAll_Click()
    For i = 0 To lstPayments.ListCount - 1
        lstPayments.Selected(i) = True
    Next
End Sub

Private Sub btnCancelPayment_Click()
    Dim n As Integer
    Dim temBillID As Long
    Dim rsTem As New ADODB.Recordset
    
    If IsNumeric(cmbStaff.BoundText) = False Then
        MsgBox "Select to re-pay"
        cmbStaff.SetFocus
        Exit Sub
    End If
    
    If Val(txtPayments.Text) = 0 Then
        MsgBox "Noting to re-pay"
        lstPayments.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Payment Method?"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    
    n = MsgBox("Are you sure you want to re-pay Rs. " & txtPayments.Text & " to " & FullStaffName(Val(cmbStaff.BoundText)), vbYesNo)
    If n = vbNo Then Exit Sub
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblProfessionalPaymentBill"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !StaffID = Val(cmbStaff.BoundText)
        !Date = Date
        !Time = Now
        !UserID = UserID
        !Value = 0 - (Val(txtPayments.Text))
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        !PaymentComments = txtComments.Text
        !IsMedicalTestBill = True
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        temBillID = !NewID
        .Close
    End With
    For n = 0 To lstIDs.ListCount - 1
        If lstPayments.Selected(n) = True Then
            With rsTem
                If .State = 1 Then .Close
                temSql = "Select * from tblProfessionalCharges where ProfessionalChargesID = " & Val(lstIDs.List(n))
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !PaidCancelled = True
                    !PaidCancelledUserID = UserID
                    !PaidCancelledDate = Date
                    !PaidCancelledTime = Now
                    !PaidCancelledDateTime = Now
                    !CancelledProfessionalPaymentBillID = temBillID
                    
                    !Paid = False
                    !PaidFee = 0
                    !PaidUserID = UserID
                    !PaidDate = Date
                    !PaidTime = Now
                    !ProfessionalPaymentBillID = temBillID

                    
                    
                    .Update
                End If
                .Close
            End With
        End If
    Next
    If chkPrint.Value = 1 Then printBill
    
    Call FillList
End Sub

Private Sub btnPay_Click()
    Dim n As Integer
    Dim temBillID As Long
    Dim rsTem As New ADODB.Recordset
    
    If IsNumeric(cmbStaff.BoundText) = False Then
        MsgBox "Select to pay"
        cmbStaff.SetFocus
        Exit Sub
    End If
    
    If Val(txtPayments.Text) = 0 Then
        MsgBox "Noting to pay"
        lstPayments.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Payment Method?"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    
    n = MsgBox("Are you sure you want to pay Rs. " & txtPayments.Text & " to " & FullStaffName(Val(cmbStaff.BoundText)), vbYesNo)
    If n = vbNo Then Exit Sub
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblProfessionalPaymentBill"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !StaffID = Val(cmbStaff.BoundText)
        !Date = Date
        !Time = Now
        !UserID = UserID
        !Value = Val(txtPayments.Text)
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        !PaymentComments = txtComments.Text
        !IsMedicalTestBill = True
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        temBillID = !NewID
        .Close
    End With
    For n = 0 To lstIDs.ListCount - 1
        If lstPayments.Selected(n) = True Then
            With rsTem
                If .State = 1 Then .Close
                temSql = "Select * from tblProfessionalCharges where ProfessionalChargesID = " & Val(lstIDs.List(n))
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !Paid = True
                    !PaidFee = !Fee
                    !PaidUserID = UserID
                    !PaidDate = Date
                    !PaidTime = Now
                    !ProfessionalPaymentBillID = temBillID
                    
                    

                    !PaidCancelled = False
                    !PaidCancelledUserID = UserID
                    !PaidCancelledDate = Date
                    !PaidCancelledTime = Now
                    !PaidCancelledDateTime = Now
                    !CancelledProfessionalPaymentBillID = temBillID

                    
                    .Update
                End If
                .Close
            End With
        End If
    Next
    If chkPrint.Value = 1 Then printBill
    
    Call FillList
End Sub

Private Sub btnPrint_Click()
    Call printBill
End Sub

Private Sub btnProcess_Click()
    Call FillList
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub optCancelled_Click()
    Call FillList
    If optToPay.Value = True Then
        btnPay.Enabled = True
        btnCancelPayment.Enabled = False
    ElseIf optCancelled.Value = True Then
        btnPay.Enabled = False
        btnCancelPayment.Enabled = False
    ElseIf optPaid.Value = True Then
        btnPay.Enabled = False
        btnCancelPayment.Enabled = True
    End If

End Sub

Private Sub optPaid_Click()
    Call FillList
    If optToPay.Value = True Then
        btnPay.Enabled = True
        btnCancelPayment.Enabled = False
    ElseIf optCancelled.Value = True Then
        btnPay.Enabled = False
        btnCancelPayment.Enabled = False
    ElseIf optPaid.Value = True Then
        btnPay.Enabled = False
        btnCancelPayment.Enabled = True
    End If

End Sub

Private Sub optToPay_Click()
    Call FillList
    If optToPay.Value = True Then
        btnPay.Enabled = True
        btnCancelPayment.Enabled = False
    ElseIf optCancelled.Value = True Then
        btnPay.Enabled = False
        btnCancelPayment.Enabled = False
    ElseIf optPaid.Value = True Then
        btnPay.Enabled = False
        btnCancelPayment.Enabled = True
    End If
End Sub

Private Sub cmbCategory_Change()
'    Call FillList
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub cmbPtPM_Change()
'    Call FillList
End Sub

Private Sub cmbSpeciality_Change()
    With rsStaff
        If .State = 1 Then .Close
        If IsNumeric(cmbSpeciality.BoundText) = True Then
            temSql = "SELECT tblStaff.Name AS NameWithTitle, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID where SpecialityID = " & Val(cmbSpeciality.BoundText) & " ORDER BY Name"
        Else
            temSql = "SELECT tblStaff.Name AS NameWithTitle, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID Order BY Name"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbStaff
        Set .RowSource = rsStaff
        .ListField = "NameWithTitle"
        .BoundColumn = "StaffID"
        .Text = Empty
    End With
End Sub

Private Sub cmbSpeciality_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbStaff.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbSpeciality.Text = Empty
    End If
End Sub

Private Sub cmbStaff_Change()
'    Call FillList
End Sub

Private Sub cmbStaff_Click(Area As Integer)
'    Call FillList
End Sub

Private Sub cmbStaff_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnPay.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbStaff.Text = Empty
    End If
End Sub

Private Sub dtpFrom_Change()
'    Call FillList
End Sub


Private Sub dtpTo_Change()
'    Call FillList
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
    Call GetSettings
'    Call FillList
    
End Sub

Private Sub printBill()
    Dim temBillPoints As MyBillPoints
    
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    
    Dim temText As String
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
        .Size = 9
        .Italic = False
        .Underline = False
    End With
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        temBillPoints = PrintThisBill("", cmbPaymentMethod.Text, FullStaffName(Val(cmbStaff.BoundText)), Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "Payment Voucher", "")
        
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        
        temY = temBillPoints.DY
        
        Printer.CurrentY = temY
        
        Printer.Print
        
        temText = Left("Date" & Space(10), 12)
        temText = temText & vbTab & vbTab
        temText = temText & Left("BillID" & Space(7), 7) & vbTab
        temText = temText & Right(Space(12) & "Fee", 12) & vbTab

        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, temText, leftAlign, MyFOnt
        
        For i = 0 To lstPayments.ListCount - 1
            If lstPayments.Selected(i) = True Then
                temY = Printer.CurrentY
                PrintingText FieldX, temY, ValueX, 0, lstPayments.List(i), leftAlign, MyFOnt
            End If
        Next
        
'        For i = 1 To gridService.Rows - 1
'            temY = Printer.CurrentY
'            n = i
'            PrintingText FieldX, temY, NoX, 0, CStr(n), RightAlign, MyFOnt
'            PrintingText FieldX, temY, ValueX, 0, gridService.TextMatrix(i, 2), LeftAlign, MyFOnt
'            PrintingText FieldX, temY, ValueX, 0, gridService.TextMatrix(i, 4), RightAlign, MyFOnt
'        Next
        
        Printer.Print
        
        temY = Printer.CurrentY
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText FieldX, temY, ValueX, 0, "Total", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtPayments.Text, rightAlign, MyFOnt
        MyFOnt.Bold = False
        MyFOnt.Size = 9
        
        
        Printer.Print
                
        temY = temBillPoints.CY - Printer.TextHeight("...")
                
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, ".........................", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, ".........................", rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Cashier :  " & UserFullName, leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, FullStaffName(Val(cmbStaff.BoundText)), rightAlign, MyFOnt
        
        
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


Private Sub FillCombos()
    Dim SC As New clsFillCombos
    SC.FillAnyCombo cmbCategory, "ServiceCategory", True
    Dim PayMethod As New clsFillCombos
    PayMethod.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanPay", False
    Dim Speciality As New clsFillCombos
    Speciality.FillAnyCombo cmbSpeciality, "Speciality", False
    Dim PtPayMethod As New clsFillCombos
    PtPayMethod.FillAnyCombo cmbPtPM, "PaymentMethod", False
    
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpFrom.Value = DateSerial(Year(Date), 1, 1)
    dtpTo.Value = Date
    cmbPaymentMethod.BoundText = 1 ' Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
    On Error Resume Next
    chkPrint.Value = GetSetting(App.EXEName, Me.Name, chkPrint.Name, "1")
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "")
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
End Sub


Private Sub FillList()
    
    txtPayments.Text = Empty
    txtTotal.Text = Empty
    txtCount.Text = Empty
    
    lstIDs.Clear
    lstPayments.Clear
    lstValues.Clear
    lstPaid.Clear
    
    Dim temTotal As Double
    Dim temCount As Double
    Dim rsTem As New ADODB.Recordset
    Dim temText As String
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblProfessionalCharges.*, tblServiceCategory.ServiceCategory, tblServiceSubCategory.ServiceSubCategory, tblProfessionalCharges.PatientServiceCharge, tblIncomeBill.IncomeBillID, tblIncomeBill.DisplayBillID "
        temSql = temSql & "FROM (((tblPatientService RIGHT JOIN tblProfessionalCharges ON tblPatientService.PatientServiceID = tblProfessionalCharges.PatientServiceID) LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID) LEFT JOIN tblIncomeBill ON tblPatientService.MedicalTestBillID = tblIncomeBill.IncomeBillID "
        temSql = temSql & "Where tblIncomeBill.Cancelled =  0   AND tblProfessionalCharges.Cancelled =  0   AND tblProfessionalCharges.Fee > 0 "
        If optToPay.Value = True Then
            temSql = temSql & " And Paid = 0 "
        ElseIf optCancelled.Value = True Then
            temSql = temSql & " And Paid = 0 And PaidCancelled = 1 "
        ElseIf optPaid.Value = True Then
            temSql = temSql & " And Paid = 1 and PaidCancelled = 0 "
        End If
        
        If IsNumeric(cmbStaff.BoundText) = True Then
            temSql = temSql & " AND tblProfessionalCharges.StaffID = " & Val(cmbStaff.BoundText) & "  "
        End If
        
        temSql = temSql & " AND ForMedicalTestBillID <> 0 "
        If IsNumeric(cmbCategory.BoundText) = True Then
            temSql = temSql & " AND tblServiceCategory.ServiceCategoryID = " & Val(cmbCategory.BoundText)
        End If
        
        temSql = temSql & " AND tblProfessionalCharges.Date between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' order by tblProfessionalCharges.Date, tblProfessionalCharges.Time"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            temText = Left(Format(![Date], "dd MMM yyyy") & Space(10), 12)
            temText = temText & vbTab
            temText = temText & Left(Format(![DisplayBillID], "0") & Space(7), 7)
            temText = temText & vbTab
            temText = temText & Right(Space(12) & Format(![Fee], "0.00"), 12) & vbTab
            If IsNumeric(cmbStaff.BoundText) = False Then
                temText = temText & Left(FullStaffName(!StaffID) & Space(27), 27) & vbTab
            End If
            
            ', tblIncomeBill.DisplayBillID,
            If !Paid = True Then
                temText = temText & "Paid on " & Format(!PaidDate, "dd MMM yyyy")
                lstPaid.AddItem "True"
            Else
                temText = temText & "To Pay"
                lstPaid.AddItem "False"
            End If
            lstPayments.AddItem temText
            lstIDs.AddItem !ProfessionalChargesID
            lstValues.AddItem ![Fee]
            temTotal = temTotal + ![Fee]
            temCount = temCount + 1
            .MoveNext
        Wend
        .Close
    End With
    txtTotal.Text = Format(temTotal, "0.00")
    txtCount.Text = temCount
End Sub

Private Sub CalculateToPay()
    Dim n As Integer
    Dim ToPay As Double
    For n = 0 To lstIDs.ListCount - 1
        If lstPayments.Selected(n) = True Then
            ToPay = ToPay + Val(lstValues.List(n))
        End If
    Next
    txtPayments.Text = Format(ToPay, "0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub lstPayments_Click()
    Call CalculateToPay
End Sub

Private Sub lstPayments_ItemCheck(Item As Integer)
    If optPaid.Value = False Then
        If lstPaid.List(Item) = "True" Then lstPayments.Selected(Item) = False
    End If
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

