VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMTRefund 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medical Test Refunds"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
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
   ScaleHeight     =   7455
   ScaleWidth      =   6990
   Begin VB.TextBox txtIncomeBillID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5400
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid gridService 
      Height          =   1695
      Left            =   5760
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   2990
      _Version        =   393216
   End
   Begin VB.TextBox txtDetail 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtReturnID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4200
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRefundingValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox txtRemarks 
      Height          =   1695
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox txtPreviousRefunds 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtBillValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   2055
   End
   Begin btButtonEx.ButtonEx btnSearch 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Search"
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
   Begin VB.TextBox txtBillID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   5640
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   2778
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Payments"
      TabPicture(0)   =   "frmMTRefund.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtPaymentMethod"
      Tab(0).Control(1)=   "cmbPaymentMethod"
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(3)=   "Label12"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Printing"
      TabPicture(1)   =   "frmMTRefund.frx":001C
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
         TabIndex        =   14
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   840
         TabIndex        =   13
         Top             =   930
         Width           =   4215
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   840
         TabIndex        =   12
         Top             =   480
         Width           =   4215
      End
      Begin MSDataListLib.DataCombo cmbPaymentMethod 
         Height          =   360
         Left            =   -72960
         TabIndex        =   15
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
         TabIndex        =   18
         Top             =   480
         Width           =   1815
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
   Begin btButtonEx.ButtonEx btnRefund 
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "&Refund"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      Top             =   6840
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
      TabIndex        =   22
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Refunding Value"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Remarks"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Previous Refunds"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Value"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Bill ID"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMTRefund"
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
        .Name = "Arial"
        .Bold = False
        .Italic = False
        .Size = 9
        .Italic = False
        .Underline = False
    End With

    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        temBillPoints = PrintThisBill(txtBillID.Text, txtDetail.Text, cmbPaymentMethod.Text, Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "Medical Test Refund", "")
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        
        temY = temBillPoints.DY
        
        Printer.CurrentY = temY

        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, txtDetail.Text, leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Re-Payment Method : " & cmbPaymentMethod.Text, leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Bill No : " & txtBillID.Text, leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Return Receipt No : " & txtReturnID.Text, leftAlign, MyFOnt
        Printer.Print
        
        Printer.Print

        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Total Return", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtRefundingValue.Text, rightAlign, MyFOnt


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

Private Sub btnRefund_Click()
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Please select a Re-Payment Method"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    If Val(txtPreviousRefunds.Text) + Val(txtRefundingValue.Text) > Val(txtBillValue.Text) Then
        MsgBox "You can't pay more than the bill value"
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
        !IncomeBillID = Val(txtIncomeBillID.Text)
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        txtReturnID.Text = !NewID
        .Close
    End With
    MsgBox "Successfully returned"
    printBill 'If chkPrint.Value = 1 Then Call printBill
    Call ClearSearchValues
    Call ClearBillValues
    txtBillID.SetFocus
End Sub

Private Sub FillGrid()
    Call FormatGrid
    Dim rsTem As New ADODB.Recordset
    Dim TotalCharge As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientService.PatientServiceID, tblPatientService.ServiceDate, tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory, tblPatientService.Comments, tblPatientService.Charge " & _
                    "FROM (tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID " & _
                    "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.MedicalTestBillID)<> 0)  AND ((tblPatientService.MedicalTestBillID)=" & Val(txtBillID.Text) & ")) " & _
                    "ORDER BY tblPatientService.PatientServiceID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridService.Rows = gridService.Rows + 1
            gridService.Row = gridService.Rows - 1
            gridService.Col = 0
            gridService.Text = !PatientServiceID
            gridService.Col = 1
            gridService.Text = !ServiceDate
            gridService.Col = 2
            If IsNull(!ServiceSubcategory) = True Then
                gridService.Text = !ServiceCategory
            Else
                gridService.Text = !ServiceCategory & " - " & !ServiceSubcategory
            End If
            gridService.Col = 3
            gridService.Text = !Comments
            gridService.Col = 4
            gridService.Text = Format(!Charge, "0.00")
            TotalCharge = TotalCharge + !Charge
            .MoveNext
        Wend
    End With
End Sub

Private Sub FormatGrid()
    '   0   ID
    '   1   Date
    '   2   Service
    '   3   Comments
    '   4   Charges
    With gridService
        .Cols = 5
        .Rows = 1
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        .ColWidth(0) = 0
        
        
        .Col = 1
        .ColWidth(1) = 0
        .Text = "Date"
        
        .Col = 2
        .ColWidth(2) = 3500
        .Text = "Service"
        
        .Col = 3
        .ColWidth(3) = 3000
        .Text = "Comments "
        
        .Col = 4
        .ColWidth(4) = 1200
        .Text = "Charge"
    End With
End Sub


Private Sub ClearBillValues()
    txtBillValue.Text = Empty
    txtPaymentMethod.Text = Empty
    txtRefundingValue.Text = Empty
    txtRemarks.Text = Empty
    txtBillID.Text = Empty
    txtPreviousRefunds.Text = Empty
    txtReturnID.Text = Empty
End Sub

Private Sub ClearSearchValues()
    txtBillValue.Text = Empty
    txtRefundingValue.Text = Empty
    txtRemarks.Text = Empty
    txtPreviousRefunds.Text = Empty
End Sub


Private Sub btnSearch_Click()
    Dim rsTem As New ADODB.Recordset
    Dim temText As String
    
    Call ClearSearchValues
    Call FormatGrid
    Call FillGrid
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue FROM tblIncomeReturnBill WHERE (((tblIncomeReturnBill.IncomeBillID)=" & Val(txtBillID.Text) & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfReturnValue) = False Then
                txtPreviousRefunds.Text = Format(!SumOfReturnValue, "0.00")
            End If
        End If
        .Close
    End With
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeBill.*, tblBHTPatientMainDetails.FirstName as BHTFirstName, tblBHT.BHT, tblIncomeBill.CompletedDate as IBCompletedDate,  tblIncomeBill.CompletedTime as IBCompletedTime, tblPatientMainDetails.FirstName as MedicalTestFirstName, tblPaymentMethod.PaymentMethod as MyPM, tblIncomeBill.NetTotal, tblBookedUser.Name as BName, tblCancelledUser.Name as CName, tblIncomeBill.Cancelled as IBCancelled, tblIncomeBill.CancelledDate as IBCancelledDate, tblIncomeBill.CancelledTime as IBCancelledTime, tblRefundMethod.PaymentMethod as MyRPM, tblIncomeBill.ReturnedValue as IBReturnValue " & _
                    "FROM ((((((tblIncomeBill LEFT JOIN tblPaymentMethod ON tblIncomeBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblBookedUser ON tblIncomeBill.UserID = tblBookedUser.StaffID) LEFT JOIN tblStaff AS tblCancelledUser ON tblIncomeBill.CancelledUserID = tblCancelledUser.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundMethod ON tblIncomeBill.CancelledPaymentMethodID = tblRefundMethod.PaymentMethodID) LEFT JOIN tblPatientMainDetails ON tblIncomeBill.PatientID = tblPatientMainDetails.PatientID) LEFT JOIN tblBHT ON tblIncomeBill.BHTID = tblBHT.BHTID) LEFT JOIN tblPatientMainDetails AS tblBHTPatientMainDetails ON tblBHT.PatientID = tblBHTPatientMainDetails.PatientID " & _
                    "WHERE IsMedicalTestBill = 1 AND tblIncomeBill.DisplayBillID = " & Val(txtBillID.Text)
        temSql = temSql & " Order by IncomeBillID Desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtIncomeBillID.Text = !IncomeBillID
            txtBillValue.Text = Format(![NetTotal], "0.00")
            If IsNull(![MedicalTestFirstName]) = False Then
                temText = "Patient : " & ![MedicalTestFirstName]
                txtDetail.Text = temText
            ElseIf IsNull(![BHT]) = False Then
                If IsNull(![BHTFirstName]) = False Then
                    temText = "Patient : " & ![BHTFirstName]
                End If
                txtDetail.Text = temText & " - " & "BHT : " & ![[tblBHT.BHT]]
                temText = temText & vbNewLine & "BHT : " & ![[tblBHT.BHT]]
            End If
            
            temText = temText & vbNewLine & "Bill Date : " & Format(![IBCOmpletedDate], "dd MMMM yyyy") & " "
            temText = temText & vbNewLine & "Bill Time : " & Format(![IBCompletedTime], "HH : MM (AMPM)") & " "
            temText = temText & vbNewLine & "Booked User : " & ![BName]
            temText = temText & vbNewLine & "Payment Method : " & ![MyPM]
            If !IBCancelled = True Then
                temText = temText = temText & vbNewLine & "Cancelled at " & !IBCancelledTime & " on " & ![IBCancelledDate] & " by " & ![CName] & "(" & ![MyRPM] & ")"
                btnRefund.Enabled = False
            Else
                If Val(txtPreviousRefunds.Text) >= Val(txtBillValue.Text) Then
                    btnRefund.Enabled = False
                Else
                    btnRefund.Enabled = True
                End If
            End If
            txtRemarks.Text = temText
        End If
        .Close
    End With

End Sub

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




