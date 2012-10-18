VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmGSBBillCancellation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancellation of Green Sheet Bill Payment"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
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
   ScaleHeight     =   6720
   ScaleWidth      =   6255
   Begin VB.TextBox txtDisplayBillID 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   240
      Width           =   3615
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   1080
      TabIndex        =   24
      Top             =   6210
      Width           =   4695
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   1080
      TabIndex        =   23
      Top             =   5760
      Width           =   4695
   End
   Begin VB.TextBox txtPaymentMethod 
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   4560
      Width           =   3615
   End
   Begin VB.TextBox txtRemarks 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1680
      Width           =   3615
   End
   Begin btButtonEx.ButtonEx btnCancelBill 
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   5040
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Cancel Payment"
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
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   3615
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox txtDate 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox txtGrossTotal 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3600
      Width           =   3615
   End
   Begin VB.TextBox txtPM 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txtPt 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox txtBillID 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   3615
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   4560
      TabIndex        =   15
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   1320
      TabIndex        =   18
      Top             =   5040
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2280
      TabIndex        =   20
      Top             =   4080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label29 
      Caption         =   "Paper"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   6210
      Width           =   1815
   End
   Begin VB.Label Label30 
      Caption         =   "Printer"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Re-payment Co&mments"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Re-payment &Method"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "User"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Time"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Gross Total"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Payment "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Patient"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Bill No"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmGSBBillCancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim FirstActi As Boolean

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

Private Sub btnCancelBill_Click()
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Please select a Re-Payment Method"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    If ProfessionalFeePaidGSB(Val(txtBillID.Text)) = True Then
        MsgBox "You can't cancel this bill as Professional Payments are already done"
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where IncomeBIllID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Cancelled = True
            !CancelledDate = Date
            !CancelledTime = Now
            !CancelledUserID = UserID
            !CancelledValue = Val(txtGrossTotal.Text)
            !cancelledPaymentMethodID = Val(cmbPaymentMethod.BoundText)
            !CancelledPaymentComments = txtPaymentMethod.Text
            .Update
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblProfessionalCharges where ForBHTID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            While .EOF = False
                !Cancelled = True
                !CancelledDate = Date
                !CancelledTime = Now
                !CancelledUserID = UserID
                .Update
                .MoveNext
            Wend
        End If
        .Close
    End With
    
    If MyBHT.Discharge = True Then
        UpdateBHTBalance MyBHT.BHTID, Val(txtGrossTotal.Text), True, False, False
    End If
    
    
    printBill 'If chkPrint.Value = 1 Then Call printBill
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
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
            
        temBillPoints = PrintThisBill(txtDisplayBillID.Text, cmbPaymentMethod.Text, MyBHT.FirstName, Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "GSB Cancellation", "")
        
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        temY = temBillPoints.DY
        
        
        Printer.CurrentY = temY
        
        Printer.Print

        Printer.Print
        
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Total", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtGrossTotal.Text, rightAlign, MyFOnt
        MyFOnt.Bold = False
        MyFOnt.Size = 9
        
        Printer.Print
        Printer.Print
        
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, ".......", rightAlign, MyFOnt
        
        temY = temBillPoints.CY - 1440
        PrintingText FieldX, temY, ValueX, 0, "Customer", leftAlign, MyFOnt
        
        temY = temBillPoints.CY - 1440
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

Private Sub Form_Activate()
    If FirstActi = True Then
        Call FormatGrid
        Call BillDetails
        Call FillGrid
        FirstActi = False
    End If
End Sub

Private Sub BillDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeBill.*, tblIncomeBill.IncomeBillID as IBIncomeBillID, tblIncomeBill.DisplayBillID, tblIncomeBill.CompletedTime as IBCompletedTime ,  tblIncomeBill.CompletedDate as IBCompletedDate , tblPatientMainDetails.FirstName, tblPaymentMethod.PaymentMethod as MyPM, tblIncomeBill.NetTotal, tblBookedUser.Name as BName, tblCancelledUser.Name as CName, tblIncomeBill.Cancelled as IBCancelled, tblIncomeBill.CancelledDate as IBCancelledDate, tblIncomeBill.CancelledTime as IBCancelledTime, tblRefundMethod.PaymentMethod as MyRPM " & _
                    "FROM ((((tblIncomeBill LEFT JOIN tblPaymentMethod ON tblIncomeBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblBookedUser ON tblIncomeBill.CompletedUserID = tblBookedUser.StaffID) LEFT JOIN tblStaff AS tblCancelledUser ON tblIncomeBill.CancelledUserID = tblCancelledUser.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundMethod ON tblIncomeBill.CancelledPaymentMethodID = tblRefundMethod.PaymentMethodID) LEFT JOIN tblPatientMainDetails ON tblIncomeBill.PatientID = tblPatientMainDetails.PatientID " & _
                    "WHERE tblIncomeBill.IncomeBillID = " & Val(txtBillID.Text)
        temSql = temSql & " Order by DisplayBillID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtTime.Text = Format(!IBCompletedTime, "hh:mm AMPM")
            txtDate.Text = Format(!IBCOmpletedDate, "dd MMMM yyyy")
            txtPt.Text = Format(!FirstName, "")
            txtPM.Text = !MyPM
            cmbPaymentMethod.Text = !MyPM
            txtGrossTotal.Text = Format(!GrossTotal, "0.00")
            txtUser.Text = !BName
            If !IBCancelled = True Then
                txtRemarks.Text = "Cancelled at " & !IBCancelledTime & " on " & ![IBCancelledDate] & " by " & ![CName] & "(" & ![MyRPM] & ")"
                btnCancelBill.Enabled = False
            End If
            txtDisplayBillID.Text = Format(!DisplayBillID, "0")
        End If
        MyBHT.BHTID = !BHTID
        .Close
    End With
End Sub

Private Sub Form_Load()
    FirstActi = True
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
    GetCommonSettings Me
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    SaveCommonSettings Me
End Sub

Private Sub FillCombos()
    Dim PayM As New clsFillCombos
    PayM.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanPay", False

End Sub

Private Sub FillGrid()
End Sub

Private Sub FormatGrid()
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


