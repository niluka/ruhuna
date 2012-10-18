VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAgentBills 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agent Payments"
   ClientHeight    =   8625
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
   ScaleHeight     =   8625
   ScaleWidth      =   10995
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   7200
      Width           =   4695
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   7650
      Width           =   4695
   End
   Begin VB.TextBox txtBillID 
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
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
      Left            =   9480
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
      Left            =   8160
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
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8705
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67895299
      CurrentDate     =   40006
   End
   Begin MSDataListLib.DataCombo cmbAgent 
      Height          =   360
      Left            =   2040
      TabIndex        =   8
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnReprint 
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   8160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx btnToExcel 
      Height          =   495
      Left            =   6840
      TabIndex        =   15
      Top             =   8040
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67895299
      CurrentDate     =   40006
   End
   Begin VB.Label Label6 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Payment Method"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Agent"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAgentBills"
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

    
Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    SaveCommonSettings Me
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpFrom.Value = Date
    dtpTo.Value = Date
    cmbPaymentMethod.BoundText = 7 ' Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
    GetCommonSettings Me
End Sub

Private Sub FormatGrid()
    With gridBill
        .Rows = 1
        .Cols = 7
        
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        
        .Col = 1
        .Text = "Bill ID"
        
        .Col = 2
        .Text = "Time"
        
        .Col = 3
        .Text = "Agent"
        
        .Col = 4
        .Text = "Payment"
        
        .Col = 5
        .Text = "Value"
        
        .Col = 6
        .Text = "Remarks"
        
        .ColWidth(0) = 0
    End With
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeBill.IncomeBillID,  tblIncomeBill.DisplayBillID, tblIncomeBill.CompletedTime, tblAgent.Agent, tblPaymentMethod.PaymentMethod as MyPM, tblIncomeBill.NetTotal, tblBookedUser.Name as BName, tblCancelledUser.Name as CName, tblIncomeBill.Cancelled, tblIncomeBill.CancelledDate, tblIncomeBill.CancelledTime, tblRefundMethod.PaymentMethod as MyRPM " & _
                    "FROM ((((tblIncomeBill LEFT JOIN tblPaymentMethod ON tblIncomeBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblBookedUser ON tblIncomeBill.CompletedUserID = tblBookedUser.StaffID) LEFT JOIN tblStaff AS tblCancelledUser ON tblIncomeBill.CancelledUserID = tblCancelledUser.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundMethod ON tblIncomeBill.CancelledPaymentMethodID = tblRefundMethod.PaymentMethodID) LEFT JOIN tblAgent ON tblIncomeBill.AgentID = tblAgent.AgentID " & _
                    "WHERE tblIncomeBill.Completed = 1  AND tblIncomeBill.IsAgentBill = 1  AND tblIncomeBill.CompletedDate  between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  "
        If IsNumeric(cmbAgent.BoundText) = True Then
            temSql = temSql & " AND tblIncomeBill.AgentID = " & Val(cmbAgent.BoundText)
            
        End If
        If IsNumeric(cmbPaymentMethod.BoundText) = True Then temSql = temSql & " And tblIncomeBill.PaymentMethodID = " & Val(cmbPaymentMethod.BoundText)
        temSql = temSql & " Order by DisplayBillID"
        
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridBill.Rows = gridBill.Rows + 1
            gridBill.Row = gridBill.Rows - 1
            
            gridBill.Col = 0
            gridBill.Text = !IncomeBillID
            
            gridBill.Col = 1
            gridBill.Text = Format(!DisplayBillID, "0")
            
            gridBill.Col = 2
            gridBill.Text = Format(!CompletedTime, "hh:mm AMPM")
            
            gridBill.Col = 3
            gridBill.Text = Format(!Agent, "")
            
            gridBill.Col = 4
            gridBill.Text = ![MyPM]
            
            gridBill.Col = 5
            gridBill.Text = Format(!NetTotal, "0.00")
            
            
            gridBill.Col = 6
            If ![Cancelled] = True Then
                gridBill.Text = "Cancelled at " & ![CancelledTime] & " on " & ![CancelledDate] & " by " & ![CName] & "(" & ![MyRPM] & ")"
            End If
        
            .MoveNext
        Wend
        .Close
    End With
    
    
    
End Sub

Private Sub FillCombos()
    Dim Staff As New clsFillCombos
    Staff.FillAnyCombo cmbAgent, "Agent", True
    Dim PM As New clsFillCombos
    PM.FillSpecificFieldBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "PaymentMethod", "CanReceive", False
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
        temBillPoints = PrintThisBill(txtBillID.Text, cmbPaymentMethod.Text, cmbAgent.Text, Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "Agent Payments", "")
        
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        
        temY = temBillPoints.DY
        
        Printer.CurrentY = temY
                
        Printer.Print
        
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Agent : " & gridBill.TextMatrix(gridBill.Row, 2), leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Payment Method : " & cmbPaymentMethod.Text, leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Bill No : " & gridBill.TextMatrix(gridBill.Row, 1), leftAlign, MyFOnt
        
        Printer.Print
        
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Value", leftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, gridBill.TextMatrix(gridBill.Row, 4), rightAlign, MyFOnt
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, "Remarks", leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, ValueX, 0, gridBill.TextMatrix(gridBill.Row, 5), rightAlign, MyFOnt
        
        
        Printer.Print
        Printer.Print
        
        temY = temBillPoints.CY
        PrintingText FieldX, temY, ValueX, 0, "Cashier :  " & UserFullName, rightAlign, MyFOnt
        
        Printer.EndDoc
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


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    
    GridPrint gridBill, ThisReportFormat
    Printer.EndDoc
    
End Sub

Private Sub btnReprint_Click()
    Call printBill
End Sub

Private Sub btnToExcel_Click()
    GridToExcel gridBill, "Agent Payments"
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

Private Sub cmbAgent_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbAgent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPaymentMethod.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbAgent.Text = Empty
    End If
End Sub

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbAgent.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtpFrom.Value = Date
        dtpTo.Value = Date
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
    Call PopulatePapers
    cmbPrinter_Click
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
        frmOPDBillCancellation.txtBillID.Text = temBillID
        frmOPDBillCancellation.Show
        frmOPDBillCancellation.ZOrder 0
        frmOPDBillCancellation.Top = 0
        frmOPDBillCancellation.Left = 0
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



