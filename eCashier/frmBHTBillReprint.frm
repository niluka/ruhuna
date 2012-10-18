VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmBHTBillReprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprint of Inward Payments"
   ClientHeight    =   6600
   ClientLeft      =   2490
   ClientTop       =   -2445
   ClientWidth     =   4905
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
   ScaleHeight     =   6600
   ScaleWidth      =   4905
   Visible         =   0   'False
   Begin VB.TextBox txtDisplayBillID 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   120
      Width           =   3255
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   840
      TabIndex        =   19
      Top             =   5490
      Width           =   3735
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   840
      TabIndex        =   18
      Top             =   5040
      Width           =   3735
   End
   Begin VB.TextBox txtRemarks 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2520
      Width           =   3255
   End
   Begin btButtonEx.ButtonEx btnReprint 
      Height          =   495
      Left            =   1800
      TabIndex        =   14
      Top             =   6000
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
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox txtDate 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3000
      Width           =   3255
   End
   Begin VB.TextBox txtGrossTotal 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4440
      Width           =   3255
   End
   Begin VB.TextBox txtPM 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox txtPt 
      Height          =   1335
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox txtBillID 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   6000
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
      TabIndex        =   21
      Top             =   5490
      Width           =   1815
   End
   Begin VB.Label Label30 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "User"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Time"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Gross Total"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Payment "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
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
Attribute VB_Name = "frmBHTBillReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim FirstActi As Boolean

    Dim MyBHT As New clsBHT


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
        .Name = DefaultFont.Name
        .Bold = False
        .Italic = False
        .Size = 9
        .Italic = False
        .Underline = False
    End With
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        temBillPoints = PrintThisBill(txtDisplayBillID.Text, txtPM.Text, MyBHT.FirstName, Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "BHT Payments - Reprint", "")
        
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        temY = temBillPoints.DY
        
        
        Printer.CurrentY = temY
        Printer.Print
        temY = Printer.CurrentY
        
        AllLines = SeperateLines(txtPt.Text)
        For i = 0 To UBound(AllLines) - 1
            PrintingText FieldX, temY, ValueX, 0, AllLines(i), leftAlign, MyFOnt
            temY = Printer.CurrentY
        Next
        
        
        Printer.Print
        
        temY = Printer.CurrentY
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText FieldX, temY, ValueX, 0, "Total", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, "Rs. " & txtGrossTotal.Text, rightAlign, MyFOnt
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
        temSql = "SELECT tblIncomeBill.*, tblIncomeBill.IncomeBillID as IBIncomeBillID, tblIncomeBill.DisplayBillID, tblIncomeBill.CompletedTime as IBCompletedTime ,  tblIncomeBill.CompletedDate as IBCompletedDate , tblPatientMainDetails.FirstName, tblPaymentMethod.PaymentMethod as MyPM, tblIncomeBill.NetTotal, tblBookedUser.Name as BName, tblCancelledUser.Name as CName, tblIncomeBill.Cancelled as IBCancelled, tblIncomeBill.CancelledDate as IBCancelledDate, tblIncomeBill.CancelledTime as IBCancelledTime, tblRefundMethod.PaymentMethod as MyRPM " & _
                    "FROM ((((tblIncomeBill LEFT JOIN tblPaymentMethod ON tblIncomeBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblBookedUser ON tblIncomeBill.CompletedUserID = tblBookedUser.StaffID) LEFT JOIN tblStaff AS tblCancelledUser ON tblIncomeBill.CancelledUserID = tblCancelledUser.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundMethod ON tblIncomeBill.CancelledPaymentMethodID = tblRefundMethod.PaymentMethodID) LEFT JOIN tblPatientMainDetails ON tblIncomeBill.PatientID = tblPatientMainDetails.PatientID " & _
                    "WHERE tblIncomeBill.IncomeBillID = " & Val(txtBillID.Text)
        temSql = temSql & " Order by DisplayBillID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            MyBHT.BHTID = !BHTID
            txtTime.Text = Format(!IBCompletedTime, "hh:mm AMPM")
            txtDate.Text = Format(!IBCOmpletedDate, "dd MMMM yyyy")
            DisplayDetails
            txtPM.Text = !MyPM
            txtGrossTotal.Text = Format(!GrossTotal, "0.00")
            txtUser.Text = !BName
            If !IBCancelled = True Then
                txtRemarks.Text = "Cancelled at " & !IBCancelledTime & " on " & ![IBCancelledDate] & " by " & ![CName] & "(" & ![MyRPM] & ")"
                btnReprint.Enabled = False
            End If
            If IsNull(!DisplayBillID) = False Then txtDisplayBillID.Text = !DisplayBillID
        End If
        .Close
    End With
End Sub

Private Sub DisplayDetails(): On Error Resume Next
    Dim temText As String
    temText = "Name : " & MyBHT.FirstName & vbNewLine
    temText = temText & "BHT : " & MyBHT.BHT & vbNewLine
    temText = temText & "Age : " & MyBHT.AgeInWords & vbNewLine
    temText = temText & "Admitted : " & Format(MyBHT.DOA, "dd MMMM yyyy") & " at " & Format(MyBHT.TOA, "HH:MM AMPM") & vbNewLine
    If MyBHT.Discharge = True Then
        temText = temText & "Discharged :" & Format(MyBHT.DOD, "dd MMMM yyyy") & " at " & Format(MyBHT.TOD, "HH:MM AMPM")
    Else
        'temText = temText & "Not yet discharged"
    End If
    txtPt.Text = temText
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


