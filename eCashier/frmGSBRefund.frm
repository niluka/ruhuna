VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmGSBRefund 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Green Sheet Bill Refunds"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
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
   ScaleHeight     =   8415
   ScaleWidth      =   8010
   Begin VB.TextBox txtBillAmount 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txtToRefund 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4440
      Width           =   2055
   End
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin VB.TextBox txtPreviousRefunds 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox txtReturnID 
      Height          =   495
      Left            =   5760
      TabIndex        =   27
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPayments 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox txtDetails 
      Height          =   2175
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   720
      Width           =   5415
   End
   Begin VB.TextBox txtRefundingValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   5880
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1575
      Left            =   120
      TabIndex        =   26
      Top             =   6720
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   2778
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Payments"
      TabPicture(0)   =   "frmGSBRefund.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(2)=   "cmbPaymentMethod"
      Tab(0).Control(3)=   "txtPaymentMethod"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Printing"
      TabPicture(1)   =   "frmGSBRefund.frx":001C
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
      Begin VB.TextBox txtPaymentMethod 
         Height          =   375
         Left            =   -72960
         TabIndex        =   20
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   840
         TabIndex        =   24
         Top             =   930
         Width           =   4215
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   840
         TabIndex        =   22
         Top             =   480
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo cmbPaymentMethod 
         Height          =   360
         Left            =   -72960
         TabIndex        =   18
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label11 
         Caption         =   "Re-payment Co&mments"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "Re-payment &Method"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "Paper"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   930
         Width           =   1815
      End
      Begin VB.Label Label30 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1815
      End
   End
   Begin btButtonEx.ButtonEx btnRefund 
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Refund"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   7920
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
   Begin VB.CheckBox chkPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   6240
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label7 
      Caption         =   "Bill Amount"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "To Refund"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Previous Refunds"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Total Payments"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "GSB No."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Details"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Refunding Value"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   1815
   End
End
Attribute VB_Name = "frmGSBRefund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
    Dim MyBHT As New clsBHT
    Dim rsBHT As New ADODB.Recordset

Private Sub btnClose_Click()
    Unload Me
End Sub


Private Sub printBill()
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
        PrintingText 0, temY, Printer.Width, 0, "", CentreAlign, MyFOnt
'                PrintingText 0, temY, Printer.Width, 0, HospitalName, CentreAlign, MyFOnt
        
                
        temY = Printer.CurrentY

        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText 0, temY, Printer.Width, 0, "", CentreAlign, MyFOnt
        
        Printer.Print
        Printer.Print
        Printer.Print
        
        temY = Printer.CurrentY
        
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText 0, temY, Printer.Width, 0, "Green Sheet Payment Refund Woucher", CentreAlign, MyFOnt

        MyFOnt.Size = 11


        Printer.Print

        Printer.Print

        temY = Printer.CurrentY


        AllLines = SeperateLines(txtDetails.Text)
        For i = 0 To UBound(AllLines) - 1
            PrintingText FieldX, temY, ValueX, 0, AllLines(i), leftAlign, MyFOnt
            temY = Printer.CurrentY
        Next


'        PrintingText FieldX, temY, ValueX, 0, "Re-Payment Method : " & cmbPaymentMethod.Text, LeftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Return Receipt No : " & txtReturnID.Text, leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Repayment Date : " & Format(Date, "dd MMMM yyyy"), leftAlign, MyFOnt
        
        Printer.Print

        temY = Printer.CurrentY
        Printer.FontBold = True
        PrintingText FieldX, temY, ValueX, 0, "Total Payments", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtPayments.Text, rightAlign, MyFOnt
        Printer.FontBold = False

        temY = Printer.CurrentY
        PrintingText FieldX, temY, FieldX, 0, "Bill Amount", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtBillAmount.Text, rightAlign, MyFOnt

        temY = Printer.CurrentY
        PrintingText FieldX, temY, FieldX, 0, "Previous Refunds", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtPreviousRefunds.Text, rightAlign, MyFOnt


'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, FieldX, 0, "Previous Payments", LeftAlign, MyFOnt
'        PrintingText FieldX, temY, ValueX, 0, txtPreviousRefunds.Text, rightAlign, MyFOnt

        Printer.Print

        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Total Return", leftAlign, MyFOnt
        MyFOnt.Bold = True
        PrintingText FieldX, temY, ValueX, 0, txtRefundingValue.Text, rightAlign, MyFOnt
        MyFOnt.Bold = False

        Printer.Print
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText ValueX, temY, ValueX, 0, ".........", rightAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, ".........", leftAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText ValueX, temY, ValueX, 0, "Patient", rightAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, "Cashier : " & UserFullName, leftAlign, MyFOnt
    
        Printer.Print

        temY = Printer.CurrentY
        PrintingText ValueX, temY, ValueX, 0, ".........", rightAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, ".........", leftAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText ValueX, temY, ValueX, 0, "Accountant", rightAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, "Manager", leftAlign, MyFOnt


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

Private Sub btnProcess_Click()
    If IsNumeric(cmbBHT.BoundText) = False Then Exit Sub
    MyBHT.BHTID = Val(cmbBHT.BoundText)
    Call DisplayDetails
    txtPayments.Text = Format(FFillPayments, "0.00")
    txtPreviousRefunds.Text = Format(FillPreviousRefunds, "0.00")
    Unload frmGSBSummery
    frmGSBSummery.Show
    frmGSBSummery.cmbBHT.BoundText = Val(cmbBHT.BoundText)
    frmGSBSummery.btnProcess_Click
    txtBillAmount.Text = frmGSBSummery.lblNetChargeF.Caption
    txtToRefund.Text = frmGSBSummery.lblBalanceF.Caption
    txtPayments.Text = frmGSBSummery.lblPaymentsF.Caption
End Sub

Private Sub btnRefund_Click()
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Please select a Re-Payment Method"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    If Val(txtPayments.Text) < Val(txtRefundingValue.Text) + Val(txtPreviousRefunds.Text) Then
        MsgBox "You can't pay more than the paid value"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeReturnBill"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ReturnDate = Date
        !ReturnTime = Now
        !ReturnUserID = UserID
        !ReturnValue = Val(txtRefundingValue.Text)
        !BHTID = Val(cmbBHT.BoundText)
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        If Trim(txtPayments.Text) <> "" Then
            !PaymentComments = txtPayments.Text
        End If
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        txtReturnID.Text = !NewID
        .Close
    End With
    
    If MyBHT.BHTID <> Val(cmbBHT.BoundText) Then
        MyBHT.BHTID = Val(cmbBHT.BoundText)
    End If
    If MyBHT.Discharge = True Then
        UpdateBHTBalance MyBHT.BHTID, Val(txtRefundingValue.Text), True, False, False
    End If
    
    
    MsgBox "Successfully returned"
    printBill 'If chkPrint.Value = 1 Then Call printBill
    Call ClearSearchValues
    Call ClearBillValues
    cmbBHT.SetFocus
End Sub

Private Sub ClearBillValues()
    txtDetails.Text = Empty
    txtPaymentMethod.Text = Empty
    txtRefundingValue.Text = Empty
    txtPreviousRefunds.Text = Empty
    txtReturnID.Text = Empty
    txtPayments.Text = Empty
    cmbPaymentMethod.BoundText = 1
End Sub

Private Sub ClearSearchValues()
    cmbBHT.Text = Empty
    txtRefundingValue.Text = Empty
    txtPreviousRefunds.Text = Empty
    txtPayments.Text = Empty
End Sub




Private Function FillPreviousRefunds()
    FillPreviousRefunds = 0
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue FROM tblIncomeReturnBill WHERE (((tblIncomeReturnBill.BHTID)=" & Val(cmbBHT.BoundText) & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfReturnValue) = False Then
                FillPreviousRefunds = !SumOfReturnValue
            End If
        End If
        .Close
    End With
End Function

Private Sub DisplayDetails(): On Error Resume Next
    Dim temText As String
    Dim r As Long
    temText = "Patient Name : " & MyBHT.FirstName & vbNewLine
'    temText = temText & "Guardian : " & MyBHT.GuardianName & vbNewLine
'    temText = temText & "Address : " & MyBHT.PtAddress & vbNewLine
    temText = temText & "Green Sheet No. : " & MyBHT.BHT & vbNewLine
'    temText = temText & "Age : " & MyBHT.AgeInWords & vbNewLine
'    temText = temText & "Payment Method : " & MyBHT.PaymentMethod
    
    If MyBHT.HealthSchemeSupplier <> "" Then
        temText = temText & " (" & MyBHT.HealthSchemeSupplier & ")" & vbNewLine
    Else
        temText = temText & vbNewLine
    End If
    If MyBHT.Comments <> "" Then
        temText = temText & MyBHT.Comments & vbNewLine
    End If

    txtDetails.Text = temText
End Sub

Private Function FFillPayments()
    Dim TotalPayments As Double
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where Completed = 1 AND IsInwardPaymentBill = 1 AND Cancelled = 0  AND BHTID = " & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            TotalPayments = TotalPayments + !NetTotal
            .MoveNext
        Wend
        .Close
    End With
    FFillPayments = TotalPayments
End Function

Private Sub Form_Load()
    Call PopulatePrinters
    Call PopulatePapers
    Call FillCombos
    Call GetSettings
End Sub

Private Sub GetSettings(): On Error Resume Next
    cmbPaymentMethod.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
    chkPrint.Value = GetSetting(App.EXEName, Me.Name, chkPrint.Name, 1)
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
End Sub

Private Sub FillCombos()
    Dim PayM As New clsFillCombos
    PayM.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanPay", False
    With rsBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where IsGSB = 1 order by BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
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





Private Sub txtToRefund_Change()
    txtRefundingValue.Text = txtToRefund.Text
End Sub
