VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGSBPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Green Sheet Payments"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
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
   ScaleHeight     =   6555
   ScaleWidth      =   6270
   Begin VB.TextBox txtBillID 
      Height          =   375
      Left            =   4800
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   6015
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   960
         TabIndex        =   18
         Top             =   720
         Width           =   4935
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label29 
         Caption         =   "Paper"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label30 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtPayment 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   3840
      Width           =   2655
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2040
      TabIndex        =   11
      Top             =   3360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2880
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   75038722
      CurrentDate     =   39962
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   2400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   75038723
      CurrentDate     =   39962
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtDetails 
      Height          =   1215
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   4095
   End
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Top             =   5880
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
   Begin btButtonEx.ButtonEx btnPay 
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin VB.Label Label6 
      Caption         =   "&Time"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Payment &Method"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "&Value"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Date"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&Green Sheet No."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "&Details"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmGSBPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Dim MyBHT As New clsBHT
    Dim rsBHT As New ADODB.Recordset
    
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
        Dim DisplayBillID As Long

    
Private Sub FillCombos()
    With rsBHT
        If .State = 1 Then .Close
        temSQL = "Select * from tblBHT where IsGSB = 1 order by BHT"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
    Dim Pay As New clsFillCombos
    Pay.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanReceive", False
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    dtpTime.Value = Time
    cmbPaymentMethod.BoundText = 1 '  Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
    chkPrint.Value = GetSetting(App.EXEName, Me.Name, chkPrint.Name, 1)
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "")
End Sub
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPay_Click()
    If IsNumeric(cmbBHT.BoundText) = False Then
        MsgBox "BHT?"
        cmbBHT.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Payment Method?"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    If Val(txtValue.Text) = 0 Then
        MsgBox "Value?"
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        
        
        
'        If .State = 1 Then .Close
'        temSql = "Select Count(IncomeBillID) as BillCount from tblIncomeBill where Completed = 1 and IsGSBill = 1"
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
        !BHTID = Val(cmbBHT.BoundText)
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        !Date = dtpDate.Value
        !Time = dtpTime.Value
        !UserID = UserID
        !StoreID = UserStoreID
        !GrossTotal = Val(txtValue.Text)
        !Completed = True
        !CompletedDate = Date
        !CompletedTime = Now
        !CompletedUserID = UserID
        !IsGSBill = True
        !NetTotal = Val(txtValue.Text)
'        !DisplayBillID = DisplayBillID
        .Update
        temSQL = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        txtBillID.Text = !NewID
        .Close
        
        DisplayBillID = NewGSBDisplayBillID(Val(txtBillID.Text))

        If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeBill where IncomeBillID = " & Val(txtBillID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !DisplayBillID = DisplayBillID
            .Update
        End If
        .Close

        
    End With
    
    If MyBHT.BHTID <> Val(cmbBHT.BoundText) Then
        MyBHT.BHTID = Val(cmbBHT.BoundText)
    End If
    If MyBHT.Discharge = True Then
        UpdateBHTBalance MyBHT.BHTID, Val(txtValue.Text), False, True, False
    End If
    
    If chkPrint.Value = 1 Then printBill
    
    Call ClearValues
    
    
    cmbBHT.SetFocus
End Sub

Private Sub ClearValues()
    txtPayment.Text = Empty
    cmbBHT.Text = Empty
    txtValue.Text = Empty
    txtDetails.Text = Empty
    dtpDate.Value = Date
    dtpTime.Value = Time
    txtBillID.Text = Empty
End Sub

Private Sub cmbBHT_Change()
    If IsNumeric(cmbBHT.BoundText) = False Then Exit Sub
    MyBHT.BHTID = Val(cmbBHT.BoundText)
    Call displayDetails
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
    Call PopulatePapers
    Call GetSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub displayDetails(): On Error Resume Next
    Dim temText As String
    temText = "Name : " & MyBHT.FirstName & vbNewLine
    temText = temText & "Green Sheet No. : " & MyBHT.BHT & vbNewLine
    temText = temText & "Age : " & MyBHT.AgeInWords & vbNewLine
'    temText = temText & "Referred By : " & MyBHT.ReferringDoctor
    txtDetails.Text = temText
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
        temBillPoints = PrintThisBill(CStr(DisplayBillID), cmbPaymentMethod.Text, MyBHT.FirstName, Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "Green Sheet Payments", "")
        
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        temY = temBillPoints.DY
        
        Printer.CurrentY = temY
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Bill No : " & txtBillID.Text, leftAlign, MyFOnt
        
        
        Printer.Print
        
        AllLines = SeperateLines(txtDetails.Text)
        For i = 0 To UBound(AllLines) - 1
            PrintingText FieldX, temY, ValueX, 0, AllLines(i), leftAlign, MyFOnt
            temY = Printer.CurrentY
        Next
        
        
        

        Printer.Print
        
        temY = Printer.CurrentY
        Printer.FontBold = False
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText FieldX, temY, ValueX, 0, "Total", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, "Rs. " & txtValue.Text, rightAlign, MyFOnt
        Printer.FontBold = False
        MyFOnt.Bold = False
        MyFOnt.Size = 9
        
        Printer.Print
        Printer.Print
        
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, ".......", rightAlign, MyFOnt
        
        temY = temBillPoints.CY
        
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


Private Sub txtValue_LostFocus()
    txtValue.Text = Format(Val(txtValue.Text), "0.00")
End Sub
