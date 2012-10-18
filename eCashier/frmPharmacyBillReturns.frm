VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPharmacyBillReturns 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pharmacy Bill Returns"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10155
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
   ScaleHeight     =   7830
   ScaleWidth      =   10155
   Begin TabDlg.SSTab SSTab1 
      Height          =   2055
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   3625
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Re-Payment"
      TabPicture(0)   =   "frmPharmacyBillReturns.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtPayments"
      Tab(0).Control(1)=   "txtPaymentMethod"
      Tab(0).Control(2)=   "cmbPaymentMethod"
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(5)=   "Label5"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Printing"
      TabPicture(1)   =   "frmPharmacyBillReturns.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label30"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label29"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmbPrinter"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmbPaper"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtPayments 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72960
         TabIndex        =   19
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtPaymentMethod 
         Height          =   375
         Left            =   -72960
         TabIndex        =   18
         Top             =   1440
         Width           =   3615
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   930
         Width           =   4695
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   4695
      End
      Begin MSDataListLib.DataCombo cmbPaymentMethod 
         Height          =   360
         Left            =   -72960
         TabIndex        =   20
         Top             =   960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Re-payment Method"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Comments"
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Value"
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Paper"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   930
         Width           =   1815
      End
      Begin VB.Label Label30 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
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
      Height          =   3900
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   9855
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "&Print"
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   6960
      Width           =   1335
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      Top             =   7320
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
   Begin btButtonEx.ButtonEx btnPay 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   7320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Return Bill Repay"
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
      Caption         =   "&Return Bills"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton optSettled 
      Caption         =   "&Returned && Paid Bills"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
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
      Format          =   67829763
      CurrentDate     =   39963
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67829762
      CurrentDate     =   39963
   End
   Begin VB.ListBox lstIDs 
      Height          =   4860
      Left            =   6720
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstValues 
      Height          =   4860
      Left            =   6360
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstPaid 
      Height          =   4860
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPharmacyBillReturns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FirstActi As Boolean
    Dim temSql As String
    
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
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPay_Click()
    Dim temIncomeBillID As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where PharmacyBillID = " & Val(lstIDs.List(lstBills.ListIndex))
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            temIncomeBillID = !IncomeBillID
        End If
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeReturnBill where IncomeReturnBillID = 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReturnDate = Date
        !ReturnTime = Now
        !ReturnUserID = UserID
        !ReturnValue = Val(txtPayments.Text)
        !IncomeBillID = temIncomeBillID
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        .Update
        .Close
        temSql = "Select * from tblSaleBill where SaleBillID = " & Val(lstIDs.List(lstBills.ListIndex))
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
            !PaidReturnedAtCashier = True
            !PaidReturnedAtCashierDate = Date
            !PaidReturnedAtCashierTime = Now
            !PaidReturnedAtCashierUserID = UserID
            !PaidReturnedValue = Val(lstValues.List(lstBills.ListIndex))
        End If
        .Update
        .Close
    
    End With
    If chkPrint.Value = 1 Then printBill
    Call FillList

End Sub

Private Sub cmbPaymentMethod_Change()
    Call FillList
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
    Call PopulatePrinters
    Call PopulatePapers
    Call FillCombos
    Call GetSettings
    Call FillList
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    dtpTime.Value = CDate(GetSetting(App.EXEName, Me.Name, dtpTime.Name, "00:00"))
    cmbPaymentMethod.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
    chkPrint.Value = Val(GetSetting(App.EXEName, Me.Name, chkPrint.Name, 0))
    optSettled.Value = CBool(GetSetting(App.EXEName, Me.Name, optSettled.Name, False))
    optToSettle.Value = CBool(GetSetting(App.EXEName, Me.Name, optToSettle.Name, False))
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, dtpTime.Name, dtpTime.Value
    SaveSetting App.EXEName, Me.Name, optSettled.Name, optSettled.Value
    SaveSetting App.EXEName, Me.Name, optToSettle.Name, optToSettle.Value
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    
End Sub

Private Sub FillCombos()
    Dim PayM As New clsFillCombos
    PayM.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanReceive", False
End Sub

Private Sub FillList()
    lstIDs.Clear
    lstBills.Clear
    lstValues.Clear
    lstPaid.Clear
    
    lstBills.Visible = False
    
    Dim rsTem As New ADODB.Recordset
    Dim temText As String
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleBill.* "
        temSql = temSql & "From tblSaleBill "
        If optToSettle.Value = True Then
            temSql = temSql & "Where (((tblSaleBill.PaidAtCashier) = 1) And ((tblSaleBill.PaidReturnedAtCashier) = 0 ) And ((tblSaleBill.Date) = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') And ((tblSaleBill.Time) > '" & dtpTime.Value & "') And ((tblSaleBill.PaymentMethodID) = " & Val(cmbPaymentMethod.BoundText) & ") AND ((tblSaleBill.Returned)= 1 ))"
        Else
            temSql = temSql & "Where (((tblSaleBill.PaidAtCashier) = 1) And ((tblSaleBill.PaidReturnedAtCashier) = 1 ) And ((tblSaleBill.Date) = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') And ((tblSaleBill.Time) > '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTime.Value & "') And ((tblSaleBill.PaymentMethodID) = " & Val(cmbPaymentMethod.BoundText) & ") AND ((tblSaleBill.Returned)= 1 ))"
        End If
        temSql = temSql & "ORDER BY tblSaleBill.SaleBillID"
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            temText = Left(!SaleBillID & Space(10), 12)
            temText = temText & vbTab
            temText = temText & Left(!ReturnedTime & Space(7), 7) & vbTab
            temText = temText & Right(Space(12) & Format(!ReturnedValue, "0.00"), 12) & vbTab
            If !PaidAtCashier = True And !PaidReturnedAtCashier = False Then
                temText = temText & "Paid on " & Format(!PaidAtCashierDate, "dd MMM yyyy")
                lstPaid.AddItem "True"
            ElseIf !PaidAtCashier = True And !PaidReturnedAtCashier = True Then
                temText = temText & "Repaid on " & Format(!PaidReturnedAtCashierDate, "dd MMM yyyy")
                lstPaid.AddItem "False"
            End If
            lstBills.AddItem temText
            lstIDs.AddItem !SaleBillID
            lstValues.AddItem !ReturnedValue
            .MoveNext
        Wend
        .Close
    End With
    
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

Private Sub printBill()
    Dim temBillPoints As MyBillPoints
    
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
        .Size = 9
        .Italic = False
        .Underline = False
    End With
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        temBillPoints = PrintThisBill(CStr(TxSaleBillID), cmbPaymentMethod.Text, "", Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "Pharmacy Bill Returns", "")
        
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        
        temY = temBillPoints.DY
        
        Printer.CurrentY = temY
                
        Printer.Print
        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Re-Payment Method : " & cmbPaymentMethod.Text, leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Bill No : " & lstIDs.List(lstBills.ListIndex), leftAlign, MyFOnt
        
        Printer.Print
        
        temY = Printer.CurrentY
        Printer.FontBold = True
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText FieldX, temY, ValueX, 0, "Total", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtPayments.Text, rightAlign, MyFOnt
        Printer.FontBold = False
        MyFOnt.Size = 9
        
        Printer.Print
        Printer.Print
        
        temY = temBillPoints.CY
        PrintingText FieldX, temY, ValueX, 0, "Customer:", leftAlign, MyFOnt
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



