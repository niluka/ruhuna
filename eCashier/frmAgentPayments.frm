VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAgentPayments 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agent Payments"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
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
   ScaleHeight     =   4620
   ScaleWidth      =   6945
   Begin VB.TextBox txtBillID 
      Height          =   495
      Left            =   600
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPaymentMethod 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   6495
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   960
         TabIndex        =   16
         Top             =   720
         Width           =   4695
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   960
         TabIndex        =   14
         Top             =   210
         Width           =   4695
      End
      Begin VB.Label Label29 
         BackColor       =   &H0080C0FF&
         Caption         =   "Paper"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label30 
         BackColor       =   &H0080C0FF&
         Caption         =   "Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   210
         Width           =   1815
      End
   End
   Begin VB.CheckBox chkPrint 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Print"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   4575
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   8438015
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
   Begin MSDataListLib.DataCombo cmbCode 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnUpdate 
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   8438015
      Caption         =   "&Update"
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
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbAgent 
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Agent"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080C0FF&
      Caption         =   "Payment Co&mments"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080C0FF&
      Caption         =   "Payment &Method"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Value"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Code"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAgentPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim DisplayBillID As Long
    
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

    Dim rsSecession As New ADODB.Recordset


Private Sub ClearBillValues()
    
    cmbAgent.Text = Empty
    cmbCode.Text = Empty
    txtValue.Text = Empty
    txtBillID.Text = Empty
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnUpdate_Click()
    If IsNumeric(cmbAgent.BoundText) = False Then
        cmbAgent.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        
'        If .State = 1 Then .Close
'        temSql = "Select count(IncomeBillID) as BillCount from tblIncomeBill where Completed = 1 and IsAgentBill = 1"
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
        temSql = "Select * from tblIncomeBill where IncomeBillID = 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Date = Date
        !Time = Now
        !UserID = UserID
        !Completed = True
        !CompletedDate = Date
        !CompletedTime = Now
        !CompletedUserID = UserID
        !AgentID = Val(cmbAgent.BoundText)
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        !PaymentComments = txtPaymentMethod.Text
        !IsAgentBill = True
        !GrossTotal = Val(txtValue.Text)
        !NetTotal = Val(txtValue.Text)
        !StoreID = UserStoreID
'        !DisplayBillID = DisplayBillID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        txtBillID.Text = !NewID
        .Close
    
        DisplayBillID = NewAgentDisplayBillID(Val(txtBillID.Text))
        
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where IncomeBillID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !DisplayBillID = DisplayBillID
            .Update
        End If
        .Close
    
    End With
    If chkPrint.Value = 1 Then printBill
    MsgBox "Agent Bill ID : " & vbTab & txtBillID.Text
    
    Call ClearBillValues
    cmbCode.SetFocus
    
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
        temBillPoints = PrintThisBill(CStr(DisplayBillID), cmbPaymentMethod.Text, cmbAgent.Text, Format(Date, "dd MM yyyy"), Format(Time, "hh:mm AMPM"), "Agent Payments", "")
        
        
        CenterX = temBillPoints.CenterX
'        NoX = temBillPoints.VX
        FieldX = temBillPoints.DX
        ValueX = temBillPoints.VX
        
        temY = temBillPoints.DY
        
        Printer.CurrentY = temY
                
        Printer.Print
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, FieldX, 0, "Agent : " & cmbAgent.Text, leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, FieldX, 0, "Payment Method : " & cmbPaymentMethod.Text, leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, FieldX, 0, "Bill No : " & DisplayBillID, leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, FieldX, 0, "Value : " & txtValue.Text, leftAlign, MyFOnt
        temY = Printer.CurrentY
        PrintingText FieldX, temY, FieldX, 0, "Date : " & Format(Date, "dd MM yy") & vbTab & vbTab & Format(Time, "HH:MM AMPM"), leftAlign, MyFOnt
        
        
        Printer.Print
       
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, FieldX, 0, "Total ", leftAlign, MyFOnt
        PrintingText ValueX, temY, ValueX, 0, txtValue.Text, rightAlign, MyFOnt
        
        
        Printer.Print
        Printer.Print
        
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

Private Sub cmbAgent_Change()
    cmbCode.BoundText = cmbAgent.BoundText
End Sub

Private Sub cmbAgent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtValue.SetFocus
    End If
End Sub

Private Sub cmbCode_Change()
    cmbAgent.BoundText = cmbCode.BoundText
End Sub

Private Sub cmbCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbAgent.SetFocus
    End If
End Sub

Private Sub cmbPaymentMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPaymentMethod.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call PopulatePrinters
    Call PopulatePapers
    Call FillCombos
    Call GetSettings
End Sub

Private Sub GetSettings(): On Error Resume Next
    cmbPaymentMethod.BoundText = 1 ' Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
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
    Dim Cat As New clsFillCombos
    Cat.FillAnyCombo cmbAgent, "Agent", True
    Dim BHT As New clsFillCombos
    BHT.FillSpecificIDField cmbCode, "Agent", "AgentID", "Code", True
    Dim PM As New clsFillCombos
    PM.FillSpecificFieldBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "PaymentMethod", "CanReceive", False
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

Private Sub txtPaymentMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnUpdate_Click
    End If
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPaymentMethod.SetFocus
    End If
End Sub

Private Sub txtValue_LostFocus()
    txtValue.Text = Format(Val(txtValue.Text), "0.00")
End Sub
