VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmOPDBillReprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprint of OPD Bills"
   ClientHeight    =   7575
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
   ScaleHeight     =   7575
   ScaleWidth      =   9825
   Begin VB.TextBox txtDisplayBillID 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   120
      Width           =   1815
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   840
      TabIndex        =   24
      Top             =   7050
      Width           =   4695
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   840
      TabIndex        =   23
      Top             =   6600
      Width           =   4695
   End
   Begin VB.TextBox txtRemarks 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1560
      Width           =   8295
   End
   Begin btButtonEx.ButtonEx btnReprint 
      Height          =   495
      Left            =   6840
      TabIndex        =   19
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Reprint"
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
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   6480
      Width           =   1815
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8280
      TabIndex        =   20
      Top             =   6960
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
   Begin VB.Label Label29 
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   7050
      Width           =   1815
   End
   Begin VB.Label Label30 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Width           =   1815
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
Attribute VB_Name = "frmOPDBillReprint"
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
        .Name = DefaultFont.Name   ' "Arial"
        .Bold = False
        .Italic = False
        .Size = 9
        .Italic = False
        .Underline = False
    End With
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        temBillPoints = PrintThisBill(txtDisplayBillID.Text, txtPM.Text, txtPt.Text, Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "OPD Services Bill (Reprint)", "")
        
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        
        temY = temBillPoints.DY
        
        
        
        For i = 1 To gridService.Rows - 1
            n = i
            'PrintingText FieldX, temY, NoX, 0, CStr(n), rightAlign, MyFOnt
            PrintingText FieldX, temY, ValueX, 0, gridService.TextMatrix(i, 2), leftAlign, MyFOnt
            PrintingText FieldX, temY, ValueX, 0, gridService.TextMatrix(i, 4), rightAlign, MyFOnt
        Next
        
        Printer.Print
        
        temY = Printer.CurrentY
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText FieldX, temY, ValueX, 0, "Total", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtNetTotal.Text, rightAlign, MyFOnt
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

Private Sub btnReprint_Click()
    Call printBill
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
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeBill.*,  tblIncomeBill.CompletedTime as IBCompletedTime, tblIncomeBill.CompletedDate as IBCompletedDate, tblPatientMainDetails.FirstName, tblPaymentMethod.PaymentMethod as MyPM, tblIncomeBill.NetTotal, tblBookedUser.Name as BName, tblCancelledUser.Name as CName, tblIncomeBill.Cancelled  as IBCancelled, tblIncomeBill.CancelledDate as IBCancelledDate, tblIncomeBill.CancelledTime as IBCancelledTime, tblRefundMethod.PaymentMethod as MyRPM " & _
                    "FROM ((((tblIncomeBill LEFT JOIN tblPaymentMethod ON tblIncomeBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblBookedUser ON tblIncomeBill.CompletedUserID = tblBookedUser.StaffID) LEFT JOIN tblStaff AS tblCancelledUser ON tblIncomeBill.CancelledUserID = tblCancelledUser.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundMethod ON tblIncomeBill.CancelledPaymentMethodID = tblRefundMethod.PaymentMethodID) LEFT JOIN tblPatientMainDetails ON tblIncomeBill.PatientID = tblPatientMainDetails.PatientID " & _
                    "WHERE tblIncomeBill.IncomeBillID = " & Val(txtBillID.Text)
        temSql = temSql & " ORDER by tblIncomeBill.DisplayBillID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtTime.Text = ![IBCompletedTime]
            txtDate.Text = Format(!IBCOmpletedDate, "dd MMMM yyyy")
            txtPt.Text = Format(!FirstName, "")
            txtPM.Text = ![MyPM]
            txtNetTotal.Text = Format(![NetTotal], "0.00")
            txtGrossTotal.Text = Format(!GrossTotal, "0.00")
            txtDiscount.Text = Format(!Discount, "0.00")
            txtUser.Text = ![BName]
            If ![IBCancelled] = True Then
                txtRemarks.Text = "Cancelled at " & ![IBCancelledTime] & " on " & ![IBCancelledDate] & " by " & ![CName] & "(" & ![MyRPM] & ")"
                btnReprint.Enabled = False
            End If
            If IsNull(!DisplayBillID) = False Then txtDisplayBillID.Text = !DisplayBillID
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

Private Sub GetSettings(): On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
    GetCommonSettings Me
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    SaveCommonSettings Me
End Sub

Private Sub FillCombos()
End Sub

Private Sub FillGrid()
    Call FormatGrid
    Dim rsTem As New ADODB.Recordset
    Dim TotalCharge As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientService.PatientServiceID, tblPatientService.ServiceDate, tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory, tblPatientService.Comments, tblPatientService.Charge " & _
                    "FROM (tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID " & _
                    "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.OPDBillID)<> 0)  AND ((tblPatientService.OPDBillID)=" & Val(txtBillID.Text) & ")) " & _
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


