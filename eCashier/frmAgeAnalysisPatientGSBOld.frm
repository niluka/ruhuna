VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAgeAnalysisPatientGSBOld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Age Analysis of Outstandings of Health Scheme Suppliers for BHT Patients"
   ClientHeight    =   9090
   ClientLeft      =   2295
   ClientTop       =   -2445
   ClientWidth     =   10680
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
   ScaleHeight     =   9090
   ScaleWidth      =   10680
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   7800
      Visible         =   0   'False
      Width           =   5655
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   3975
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label18 
         Caption         =   "Paper"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9120
      TabIndex        =   3
      Top             =   8520
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
   Begin MSFlexGridLib.MSFlexGrid gridBalance 
      Height          =   6855
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12091
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   1920
      TabIndex        =   9
      Top             =   1080
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnToExcel 
      Height          =   495
      Left            =   7800
      TabIndex        =   11
      Top             =   8520
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
   Begin VB.Label Label42 
      Caption         =   "Paid as"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblSubtopic 
      Alignment       =   2  'Center
      Caption         =   "Subtopic"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   10095
   End
   Begin VB.Label lblTopic 
      Alignment       =   2  'Center
      Caption         =   "TOPIC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "frmAgeAnalysisPatientGSBOld"
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
    Dim SuppliedWord As String

Private Sub btnToExcel_Click()

    GridToExcel gridBalance, "Age Analysis of Patient Balance - " & cmbPaymentMethod.Text, "Date - " & Format(Date, LongDateFormat) & " - Time " & Format(Time, "HH MM AMPM")

End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub cmbPaymentMethod_Click(Area As Integer)
    Call FormatGrid
    Call FillList
End Sub

Private Sub cmbPaymentMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbPaymentMethod.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call PopulatePrinters
    Call FillCombos
    Call GetSettings
    Call FillList
End Sub

Private Sub FillCombos()
    Dim PM As New clsFillCombos
    PM.FillAnyCombo cmbPaymentMethod, "PaymentMethod", False
End Sub

Private Sub printBill()
'    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text) : doevents
'    Dim MyFOnt As ReportFont
'
'    PrinterName = Printer.DeviceName
'    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
'        ClosePrinter (PrinterHandle) : doevents
'    End If
'    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text) : doevents
'
'    Dim temText As String
'    Dim CenterX As Long
'    Dim FieldX As Long
'    Dim NoX As Long
'    Dim ValueX As Long
'    Dim AllLines() As String
'    Dim i As Integer
'    Dim temY As Long
'    Dim n As Long
'
''    With MyFont
''        .Name =  DefaultFont.Name
''        .Bold = False
''        .Italic = False
''        .Size = 12
''        .Italic = False
''        .Underline = False
''    End With
'
'    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
'        CenterX = Printer.Width / 2
'        NoX = 1440 * 1
'        FieldX = 1440 * 1.2
'        ValueX = Printer.Width - 1440 * 1
'        temY = Printer.CurrentY

''        MyFont.Bold = False
''        MyFont.Size = 13
'        PrintingText 0, temY, Printer.Width, 0, HospitalName, CentreAlign, MyFOnt
'        temY = Printer.CurrentY

''        MyFont.Bold = False
''        MyFont.Size = 12
'        PrintingText 0, temy, Printer.Width, 0, "Payment Voucher", CentreAlign, MyFont
'
''        MyFont.Size = 11
'
'        Printer.Print
'
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, "Doctor : " & cmbStaff.Text, LeftAlign, MyFont
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, "Payment Method : " & cmbPaymentMethod.Text, LeftAlign, MyFont
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, "Date : " & Format(Date, "dd MM yy"), LeftAlign, MyFont
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, "Time : " & Format(Time, "HH : MM AMPM"), LeftAlign, MyFont
'
'        Printer.Print
'
'        temText = Left("Date" & Space(10), 12)
'        temText = temText & vbTab & vbTab
'        temText = temText & Left("GSB" & Space(7), 7) & vbTab
'        temText = temText & Right(Space(12) & "Fee", 12) & vbTab
'
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, temText, LeftAlign, MyFont
'
'        For i = 0 To lstPayments.ListCount - 1
'            If lstPayments.Selected(i) = True Then
'                temY = Printer.CurrentY
'                PrintingText FieldX, temY, ValueX, 0, lstPayments.List(i), LeftAlign, MyFont
'            End If
'        Next
'
''        For i = 1 To gridService.Rows - 1
''            temY = Printer.CurrentY
''            n = i
''            PrintingText FieldX, temY, NoX, 0, CStr(n), RightAlign, MyFont
''            PrintingText FieldX, temY, ValueX, 0, gridService.TextMatrix(i, 2), LeftAlign, MyFont
''            PrintingText FieldX, temY, ValueX, 0, gridService.TextMatrix(i, 4), RightAlign, MyFont
''        Next
'
'        Printer.Print
'
'        temY = Printer.CurrentY
'        Printer.FontBold = True
'        PrintingText FieldX, temY, ValueX, 0, "Total", LeftAlign, MyFOnt
'        PrintingText FieldX, temY, ValueX, 0, txtPayments.Text, rightAlign, MyFOnt
'        Printer.FontBold = False
'
'
'
'        Printer.Print
'
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, ".........................", LeftAlign, MyFont
'        PrintingText FieldX, temY, ValueX, 0, ".........................", RightAlign, MyFont
'
'        temY = Printer.CurrentY
'        PrintingText FieldX, temY, ValueX, 0, "Cashier :  " & userfullname, rightAlign, MyFOnt
'        PrintingText FieldX, temY, ValueX, 0, cmbStaff.Text, RightAlign, MyFont
'
'
'        Printer.EndDoc
'
'    End If
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

Private Sub GetSettings(): On Error Resume Next
    GetCommonSettings Me
    lblTopic.Caption = "Age Analysis"
    lblSubtopic.Caption = ""
    cmbPaymentMethod.BoundText = GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1)
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "")
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
End Sub


Private Sub FormatGrid()

            gridBalance.Cols = 7
            gridBalance.Rows = 1
            gridBalance.Row = 0
            
            gridBalance.ColWidth(0) = 0
'            gridBalance.ColWidth(1) = 2600
'            gridBalance.ColWidth(2) = 1200
'            gridBalance.ColWidth(3) = 1200
'            gridBalance.ColWidth(4) = 1200
'            gridBalance.ColWidth(5) = 1200
'            gridBalance.ColWidth(6) = 1600
            
            
            gridBalance.Col = 0
            'gridBalance.Text = "HealthSchemeSupplierID"
            gridBalance.Col = 1
            gridBalance.Text = "Health Scheme Supplier"
            gridBalance.Col = 2
            gridBalance.Text = "< 30 d"
            gridBalance.Col = 3
            gridBalance.Text = "30 - 60"
            gridBalance.Col = 4
            gridBalance.Text = "61 - 90"
            gridBalance.Col = 5
            gridBalance.Text = ">90"
            gridBalance.Col = 6
            gridBalance.Text = "Total"
End Sub

Private Sub FillList()
    Screen.MousePointer = vbHourglass
    Dim rsHS As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
    
    Dim temText As String
    Dim Below30dBalance As Double
    Dim Between30And60Balance As Double
    Dim Between60And90Balance As Double
    Dim Above90Balance As Double
    Dim Below30dBalanceTotal As Double
    Dim Between30And60BalanceTotal As Double
    Dim Between60And90BalanceTotal As Double
    Dim Above90BalanceTotal As Double
    Dim DaysPassed As Long
    
    With rsTem
        If IsNumeric(cmbPaymentMethod.BoundText) = True Then
            temSql = "SELECT tblPatientMainDetails.PatientID, tblPatientMainDetails.FirstName, tblBHT.* " & _
                        "FROM tblBHT LEFT JOIN tblPatientMainDetails ON tblBHT.PatientID = tblPatientMainDetails.PatientID " & _
                        "WHERE (((tblBHT.Discharge)=1) AND ((tblBHT.IsGSB)=True) AND ((tblBHT.PaymentMethodID)=" & Val(cmbPaymentMethod.BoundText) & ") AND ((tblBHT.Balance)>0))" & _
                        "ORDER BY tblPatientMainDetails.FirstName"
        Else
            temSql = "SELECT tblPatientMainDetails.PatientID, tblPatientMainDetails.FirstName, tblBHT.* " & _
                        "FROM tblBHT LEFT JOIN tblPatientMainDetails ON tblBHT.PatientID = tblPatientMainDetails.PatientID " & _
                        "WHERE (((tblBHT.Discharge)=1) AND ((tblBHT.IsGSB)=True) AND ((tblBHT.Balance)>0))" & _
                        "ORDER BY tblPatientMainDetails.FirstName"
        End If
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        While .EOF = False
            Below30dBalance = 0
            Between30And60Balance = 0
            Between60And90Balance = 0
            Above90Balance = 0
            DaysPassed = DateDiff("d", !DOD, Date)
            Select Case DaysPassed:
                Case Is < 30:
                                Below30dBalance = Below30dBalance + !Balance
                                Below30dBalanceTotal = Below30dBalanceTotal + !Balance
                
                Case 30 To 60:
                                Between30And60Balance = Between30And60Balance + !Balance
                                Between30And60BalanceTotal = Between30And60BalanceTotal + !Balance
                
                Case 61 To 90:
                                Between60And90Balance = Between60And90Balance + !Balance
                                Between60And90BalanceTotal = Between60And90BalanceTotal + !Balance
                
                Case Is > 90:
                                Above90Balance = Above90Balance + !Balance
                                Above90BalanceTotal = Above90BalanceTotal + !Balance
                Case Else
                    
                    MsgBox "Error"
            
            End Select
            
            
            
            gridBalance.Rows = gridBalance.Rows + 1
            gridBalance.Row = gridBalance.Rows - 1
            gridBalance.Col = 0
            'gridBalance.Text = ![tblBHT.PatientID]
            gridBalance.Col = 1
            gridBalance.Text = !FirstName
            gridBalance.Col = 2
            gridBalance.Text = Format(Below30dBalance, "#,##0.00")
            gridBalance.Col = 3
            gridBalance.Text = Format(Between30And60Balance, "#,##0.00")
            gridBalance.Col = 4
            gridBalance.Text = Format(Between60And90Balance, "#,##0.00")
            gridBalance.Col = 5
            gridBalance.Text = Format(Above90Balance, "#,##0.00")
            gridBalance.Col = 6
            gridBalance.Text = Format(Below30dBalance + Between30And60Balance + Above90Balance + Between60And90Balance, "#,##0.00")
            
            
            
            .MoveNext
        Wend
        .Close
    End With
    
    
    gridBalance.Rows = gridBalance.Rows + 1
    gridBalance.Row = gridBalance.Rows - 1
    gridBalance.Col = 0
    gridBalance.Text = Empty
    gridBalance.Col = 1
    gridBalance.Text = "Total"
    gridBalance.Col = 2
    gridBalance.Text = Format(Below30dBalanceTotal, "#,##0.00")
    gridBalance.Col = 3
    gridBalance.Text = Format(Between30And60BalanceTotal, "#,##0.00")
    gridBalance.Col = 4
    gridBalance.Text = Format(Between60And90BalanceTotal, "#,##0.00")
    gridBalance.Col = 5
    gridBalance.Text = Format(Above90BalanceTotal, "#,##0.00")
    gridBalance.Col = 6
    gridBalance.Text = Format(Below30dBalanceTotal + Between30And60BalanceTotal + Above90BalanceTotal + Between60And90BalanceTotal, "#,##0.00")
    
    Screen.MousePointer = vbDefault
    
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
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                cmbPaper.AddItem PtrCtoVbString(.pName)
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


