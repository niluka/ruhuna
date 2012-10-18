VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLabBillList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lab Bill List"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10995
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
   ScaleHeight     =   8985
   ScaleWidth      =   10995
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   8520
      Width           =   4695
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   8040
      Width           =   4695
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2040
      TabIndex        =   7
      Top             =   1080
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
      Left            =   9600
      TabIndex        =   4
      Top             =   8040
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin MSFlexGridLib.MSFlexGrid gridBill 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11033
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
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
      Format          =   77856771
      CurrentDate     =   40006
   End
   Begin MSDataListLib.DataCombo cmbUser 
      Height          =   360
      Left            =   2040
      TabIndex        =   8
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnToExcel 
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   7200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&To Excel"
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
      BackColor       =   &H00C0C0FF&
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Payment Method"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "User"
      Height          =   255
      Left            =   120
      TabIndex        =   5
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
Attribute VB_Name = "frmLabBillList"
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
    dtpDate.Value = Date
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    SaveCommonSettings Me
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
    GetCommonSettings Me
End Sub

Public Sub FormatGrid()
    With gridBill
        .Rows = 1
        .Cols = 9
        
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        .ColWidth(0) = 0
        
        
        .Col = 1
        .Text = "Bill ID"
        
        
        .Col = 2
        .Text = "Time"
        
        .Col = 3
        .Text = "Patient"
        
        .Col = 4
        .Text = "Payment"
        
        .Col = 5
        .Text = "Prof."
        
        .Col = 6
        .Text = "Hos."
        
        .Col = 7
        .Text = "TOTAL"
        
        .Col = 8
        .Text = "Remarks"
        
        .ColWidth(0) = 0
        
    End With
End Sub

Public Sub FillGrid()


    Dim Total As Double
    Dim H As Double
    Dim P As Double

    Dim temH As Double
    Dim Temp As Double


    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeBill.IncomeBillID,  tblIncomeBill.DisplayBillID, tblIncomeBill.CompletedTime, tblPatientMainDetails.FirstName, tblPaymentMethod.PaymentMethod as MyPM, tblIncomeBill.NetTotal, tblBookedUser.Name as BName, tblCancelledUser.Name as CName, tblIncomeBill.Cancelled, tblIncomeBill.CancelledDate, tblIncomeBill.CancelledTime, tblRefundMethod.PaymentMethod as MyRPM " & _
                    "FROM ((((tblIncomeBill LEFT JOIN tblPaymentMethod ON tblIncomeBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblBookedUser ON tblIncomeBill.CompletedUserID = tblBookedUser.StaffID) LEFT JOIN tblStaff AS tblCancelledUser ON tblIncomeBill.CancelledUserID = tblCancelledUser.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundMethod ON tblIncomeBill.CancelledPaymentMethodID = tblRefundMethod.PaymentMethodID) LEFT JOIN tblPatientMainDetails ON tblIncomeBill.PatientID = tblPatientMainDetails.PatientID " & _
                    "WHERE tblIncomeBill.Completed = 1  AND tblIncomeBill.IsLabBill = 1  AND tblIncomeBill.CompletedDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbUser.BoundText) = True Then temSql = temSql & " AND tblIncomeBill.CompletedUserID = " & Val(cmbUser.BoundText)
        If IsNumeric(cmbPaymentMethod.BoundText) = True Then temSql = temSql & " And tblIncomeBill.PaymentMethodID = " & Val(cmbPaymentMethod.BoundText)
        temSql = temSql & " Order by DisplayBillID"
        
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridBill.Rows = gridBill.Rows + 1
            gridBill.Row = gridBill.Rows - 1
            
            gridBill.Col = 0
            gridBill.Text = !IncomeBillID
            
            gridBill.Col = 1
            gridBill.Text = Format(!DisplayBillID)
            
            gridBill.Col = 2
            gridBill.Text = Format(!CompletedTime, "hh:mm AMPM")
            
            gridBill.Col = 3
            gridBill.Text = Format(!FirstName, "")
            
            gridBill.Col = 4
            gridBill.Text = ![MyPM]
            
            Temp = BillPFee(!IncomeBillID)
            gridBill.Col = 5
            gridBill.Text = Format(Temp, "0.00")
            
            temH = BillHFee(!IncomeBillID)
            gridBill.Col = 6
            gridBill.Text = Format(temH, "0.00")
            
            gridBill.Col = 7
            gridBill.Text = Format(!NetTotal, "0.00")
            
            
            gridBill.Col = 8
            If ![Cancelled] = True Then
                gridBill.Text = "Cancelled at " & ![CancelledTime] & " on " & ![CancelledDate] & " by " & ![CName] & "(" & ![MyRPM] & ")"
            Else
                Total = Total + !NetTotal
                H = H + temH
                P = P + Temp
            End If
        
            .MoveNext
        Wend
        .Close
    End With

    gridBill.Rows = gridBill.Rows + 1
    gridBill.Row = gridBill.Rows - 1
    
   
    gridBill.Col = 1
    gridBill.Text = "Total"
    
    gridBill.Col = 5
    gridBill.Text = Format(P, "0.00")
    
    gridBill.Col = 6
    gridBill.Text = Format(H, "0.00")
    
    gridBill.Col = 7
    gridBill.Text = Format(Total, "0.00")


End Sub

Private Function BillPFee(BillID As Long) As Double
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT SUM(ProfessionalCharge) AS P FROM dbo.tblPatientService WHERE     (Deleted = 0) AND (LabBillID = " & BillID & ")"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!P) = False Then
                BillPFee = !P
            Else
                BillPFee = 0
            End If
        Else
            BillPFee = 0
        End If
        .Close
    End With
End Function

Private Function BillHFee(BillID As Long) As Double
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT SUM(HospitalCharge) AS H FROM dbo.tblPatientService WHERE     (Deleted = 0) AND (LabBillID = " & BillID & ")"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!H) = False Then
                BillHFee = !H
            Else
                BillHFee = 0
            End If
        Else
            BillHFee = 0
        End If
        .Close
    End With
End Function

Private Sub FillCombos()
    Dim Staff As New clsFillCombos
    Staff.FillSpecificFieldBoolCombo cmbUser, "Staff", "Name", "Name", "IsAUser", False
    Dim PM As New clsFillCombos
    PM.FillSpecificFieldBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "PaymentMethod", "CanReceive", False
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    
    Dim temBillPoints As MyBillPoints
    Dim ThisReportFormat As PrintReport
    Dim MyFOnt As ReportFont

    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        GetPrintDefaults ThisReportFormat
        GridPrint gridBill, ThisReportFormat, "Lab Bill List", Format(dtpDate.Value, "dd MMMM yyyy")
        Printer.EndDoc
    End If
    
    
End Sub

Private Sub btnToExcel_Click()
    GridToExcel gridBill, "Lab Bills", Format(dtpDate.Value, "dd MMMM yyyy") & " - " & cmbUser.Text & " - " & cmbPaymentMethod.Text

End Sub

Private Sub cmbPaymentMethod_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbPaymentMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnPrint.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbPaymentMethod.Text = Empty
    End If
End Sub

Private Sub cmbUser_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPaymentMethod.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbUser.Text = Empty
    End If
End Sub

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbUser.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtpDate.Value = Date
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
    Call GetSettings
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub gridBill_DblClick()
    Dim temBillID As Long
    temBillID = Val(gridBill.TextMatrix(gridBill.Row, 0))
    If temBillID <> 0 Then
        frmLabBillCancellation.txtBillID.Text = temBillID
        frmLabBillCancellation.Show
        frmLabBillCancellation.ZOrder 0
        frmLabBillCancellation.Top = 0
        frmLabBillCancellation.Left = 0
    End If
End Sub
