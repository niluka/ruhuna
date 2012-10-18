VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMTBillCancellation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancellation of Medical Test Bills"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
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
   ScaleHeight     =   8175
   ScaleWidth      =   9825
   Begin TabDlg.SSTab SSTab1 
      Height          =   1575
      Left            =   120
      TabIndex        =   23
      Top             =   6480
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   2778
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Payments"
      TabPicture(0)   =   "frmMedicalTestBillCancellation.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmbPaymentMethod"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtPaymentMethod"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Printing"
      TabPicture(1)   =   "frmMedicalTestBillCancellation.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label29"
      Tab(1).Control(1)=   "Label30"
      Tab(1).Control(2)=   "cmbPaper"
      Tab(1).Control(3)=   "cmbPrinter"
      Tab(1).ControlCount=   4
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   -74160
         TabIndex        =   29
         Top             =   480
         Width           =   4695
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   -74160
         TabIndex        =   28
         Top             =   930
         Width           =   4695
      End
      Begin VB.TextBox txtPaymentMethod 
         Height          =   375
         Left            =   2280
         TabIndex        =   24
         Top             =   960
         Width           =   4095
      End
      Begin MSDataListLib.DataCombo cmbPaymentMethod 
         Height          =   360
         Left            =   2280
         TabIndex        =   25
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label30 
         Caption         =   "Printer"
         Height          =   255
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label29 
         Caption         =   "Paper"
         Height          =   255
         Left            =   -74880
         TabIndex        =   30
         Top             =   930
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Re-payment &Method"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Re-payment Co&mments"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.TextBox txtRemarks 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1560
      Width           =   8295
   End
   Begin btButtonEx.ButtonEx btnCancelBill 
      Height          =   495
      Left            =   6720
      TabIndex        =   19
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Cancel Bill"
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
   Begin MSFlexGridLib.MSFlexGrid gridService 
      Height          =   4455
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7858
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtDate 
      Height          =   375
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtNetTotal 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtGrossTotal 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtPM 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtPt 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtBillID 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8160
      TabIndex        =   20
      Top             =   7560
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
      Left            =   6720
      TabIndex        =   32
      Top             =   7200
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "User"
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Time"
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Date"
      Height          =   255
      Left            =   6600
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Net Total"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Discount"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Gross Total"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Payment "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Patient"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Bill No"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMTBillCancellation"
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


Private Sub btnCancelBill_Click()
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Please select a Re-Payment Method"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If

    Dim rstem As New ADODB.Recordset
    With rstem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where IncomeBIllID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Cancelled = True
            !CancelledDate = Date
            !CancelledTime = Time
            !CancelledUserID = UserID
            !CancelledValue = Val(txtNetTotal.Text)
            !cancelledPaymentMethodID = Val(cmbPaymentMethod.BoundText)
            !CancelledPaymentComments = txtPaymentMethod.Text
            .Update
        End If
        .Close
    End With
    With rstem
        If .State = 1 Then .Close
        temSql = "Select * from tblProfessionalCharges where ForOPDBIllID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            While .EOF = False
                !Cancelled = True
                !CancelledDate = Date
                !CancelledTime = Time
                !CancelledUserID = UserID
                .Update
                .MoveNext
            Wend
        End If
        .Close
    End With
    
    If chkPrint.Value = 1 Then Call printBill
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


Private Sub printBill()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    
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
        
        
        MyFOnt.Bold = True
        MyFOnt.Size = 13
        PrintingText 0, 1440 * 1, Printer.Width, 0, HospitalName, CentreAlign, MyFOnt
        
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText 0, 1440 * 1, Printer.Width, 0, HospitalDescreption, CentreAlign, MyFOnt
        
        
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText 0, 0, Printer.Width, 0, "OPD Services Bill Cancellation", CentreAlign, MyFOnt
        
        MyFOnt.Size = 11
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Patient : " & txtPt.Text, LeftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Re-Payment Method : " & cmbPaymentMethod.Text, LeftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Bill No : " & txtBillID.Text, LeftAlign, MyFOnt
        
        Printer.Print
        
        
        For i = 1 To gridService.Rows - 1
            temY = Printer.CurrentY
            n = i
            PrintingText FieldX, temY, NoX, 0, CStr(n), RightAlign, MyFOnt
            PrintingText FieldX, temY, ValueX, 0, gridService.TextMatrix(i, 2), LeftAlign, MyFOnt
            PrintingText FieldX, temY, ValueX, 0, gridService.TextMatrix(i, 4), RightAlign, MyFOnt
        Next
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Total", LeftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtNetTotal.Text, RightAlign, MyFOnt
        Printer.Print
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, ".......", RightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Cashier :  " & UserName, LeftAlign, MyFOnt
        
        Printer.EndDoc
        
    End If
End Sub

Private Sub PrintingText(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, PrintText As String, PrintAlignment As TextAlignment, PrintFont As ReportFont)
    
    If PrintAlignment = LeftAlign Then
        Printer.CurrentX = X1
    ElseIf PrintAlignment = RightAlign Then
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
        Call FillGrid
        Call BillDetails
        FirstActi = False
    End If
End Sub

Private Sub BillDetails()
    Dim rstem As New ADODB.Recordset
    With rstem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeBill.*, tblIncomeBill.CompletedTime, tblPatientMainDetails.FirstName, tblPaymentMethod.PaymentMethod, tblIncomeBill.NetTotal, tblBookedUser.Name, tblCancelledUser.Name, tblIncomeBill.Cancelled, tblIncomeBill.CancelledDate, tblIncomeBill.CancelledTime, tblCancelledUser.Name, tblRefundMethod.PaymentMethod " & _
                    "FROM ((((tblIncomeBill LEFT JOIN tblPaymentMethod ON tblIncomeBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblBookedUser ON tblIncomeBill.CompletedUserID = tblBookedUser.StaffID) LEFT JOIN tblStaff AS tblCancelledUser ON tblIncomeBill.CancelledUserID = tblCancelledUser.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundMethod ON tblIncomeBill.CancelledPaymentMethodID = tblRefundMethod.PaymentMethodID) LEFT JOIN tblPatientMainDetails ON tblIncomeBill.PatientID = tblPatientMainDetails.PatientID " & _
                    "WHERE tblIncomeBill.IncomeBillID = " & Val(txtBillID.Text)
        temSql = temSql & " Order by IncomeBillID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtTime.Text = ![tblIncomeBill.CompletedTime]
            txtDate.Text = Format(!Date, "dd MMMM yyyy")
            txtPt.Text = Format(!FirstName, "")
            txtPM.Text = ![tblPaymentMethod.PaymentMethod]
            txtNetTotal.Text = Format(![tblIncomeBill.NetTotal], "0.00")
            txtGrossTotal.Text = Format(!GrossTotal, "0.00")
            txtDiscount.Text = Format(!Discount, "0.00")
            txtUser.Text = ![tblBookedUser.Name]
            If ![tblIncomeBill.Cancelled] = True Then
                txtRemarks.Text = "Cancelled at " & ![tblIncomeBill.CancelledTime] & " on " & ![tblIncomeBill.CancelledDate] & " by " & ![tblCancelledUser.Name] & "(" & ![tblRefundMethod.PaymentMethod] & ")"
                btnCancelBill.Enabled = False
            End If
        End If
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

Private Sub GetSettings()
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

Private Sub FillGrid()
    Call FormatGrid
    Dim rstem As New ADODB.Recordset
    Dim rsT As New ADODB.Recordset
    Dim TotalCharge As Double
    With rstem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientService.PatientServiceID, tblPatientService.ServiceDate, tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory, tblPatientService.Comments, tblPatientService.HospitalCharge " & _
                    "FROM (tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID " & _
                    "WHERE (((tblPatientService.Deleted)=False) AND ((tblPatientService.OPDBillID)<> 0)  AND ((tblPatientService.OPDBillID)=" & Val(txtBillID.Text) & ")) " & _
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
            If IsNull(!ServiceSubCategory) = True Then
                gridService.Text = !ServiceCategory
            Else
                gridService.Text = !ServiceCategory & " - " & !ServiceSubCategory
            End If
            gridService.Col = 3
            gridService.Text = !Comments
            gridService.Col = 4
            gridService.Text = Format(!HospitalCharge, "0.00")
            TotalCharge = TotalCharge + !HospitalCharge
            
            
            If rsT.State = 1 Then rsT.Close
            temSql = "SELECT tblStaff.Name, tblProfessionalCharges.* " & _
                        "FROM (tblPatientService LEFT JOIN tblProfessionalCharges ON tblPatientService.PatientServiceID = tblProfessionalCharges.PatientServiceID) LEFT JOIN tblStaff ON tblProfessionalCharges.StaffID = tblStaff.StaffID " & _
                        "WHERE (((tblProfessionalCharges.ForOPDBillID)=" & Val(txtBillID.Text) & "))"
            rsT.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            While rsT.EOF = False
                gridService.Rows = gridService.Rows + 1
                gridService.Row = gridService.Rows - 1
                gridService.Col = 0
                gridService.Col = 1
                gridService.Text = rsT!Date
                gridService.Col = 2
                If IsNull(rsT!Name) = True Then
                    gridService.Text = ""
                Else
                    gridService.Text = rsT!Name
                End If
                gridService.Col = 4
                gridService.Text = Format(rsT!Fee, "0.00")
                TotalCharge = TotalCharge + rsT!Fee
                rsT.MoveNext
            Wend
                
            
            
            
            
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
        .Text = "Hospital Fee"
        
        .Col = 4
        .ColWidth(4) = 1200
        .Text = "Professional Fee"
        
        .Col = 4
        .ColWidth(4) = 1200
        .Text = "Total"
        
        
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

Private Sub PopulatePapers()
    cmbPaper.Clear
    SetPrinter = False
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
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
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub cmbPrinter_Change()
    Call PopulatePapers
End Sub

Private Sub cmbPrinter_Click()
    Call PopulatePapers
End Sub


