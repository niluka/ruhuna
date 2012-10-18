VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCurrentCompanyBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Current Company Balance"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12075
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
   ScaleHeight     =   8445
   ScaleWidth      =   12075
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   6960
      Width           =   5175
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   4215
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label17 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Paper"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   9120
      TabIndex        =   0
      Top             =   7320
      Width           =   2775
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   10680
      TabIndex        =   1
      Top             =   7800
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
      Left            =   9360
      TabIndex        =   2
      Top             =   7800
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
   Begin MSFlexGridLib.MSFlexGrid gridService 
      Height          =   6375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11245
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   77856771
      CurrentDate     =   40028
   End
   Begin VB.Label Label1 
      Caption         =   "Today"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Total Value"
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   7320
      Width           =   2775
   End
End
Attribute VB_Name = "frmCurrentCompanyBalance"
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

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        GridPrint gridService, ThisReportFormat, "Company Balance", "On " & Format(dtpTo.Value, "dd MMMM yyyy")
        Printer.EndDoc
    End If
    
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
    Call FormatGrid
    Call GetSettings
    Call FillGrid
End Sub


Private Sub FillCombos()

End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpTo.Value = Date
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "")
    Dim i As Integer
    With gridService
        For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(App.EXEName, Me.Name, i, .ColWidth(i)))
        Next
    End With
    
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    Dim i As Integer
    With gridService
        For i = 0 To .Cols - 1
            SaveSetting App.EXEName, Me.Name, i, .ColWidth(i)
        Next
    End With
End Sub

Private Sub FormatGrid()
    With gridService
        .Clear
    
        .Rows = 1
        .Cols = 3
        
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        
        .Col = 1
        .Text = "Company"
        
        .Col = 2
        .Text = "Balance"
        
    End With
End Sub

Private Sub FillGrid()
    txtTotal.Text = Empty
    
    Dim temTotal As Double
    Dim temCount As Double
    
    Dim rsTem As New ADODB.Recordset
    
    Dim temText As String
    
    With rsTem
        If .State = 1 Then .Close
        
        temSql = "SELECT tblHealthSchemeSuppliers.*  FROM tblHealthSchemeSuppliers ORDER BY tblHealthSchemeSuppliers.HealthSchemeSupplierName"
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        gridService.Visible = False
        
        While .EOF = False
            gridService.Rows = gridService.Rows + 1
            gridService.Row = gridService.Rows - 1
            
            gridService.Col = 0
            gridService.Text = !HealthSchemeSupplierID
            
            gridService.Col = 1
            gridService.Text = !HealthSchemeSupplierName
            
            gridService.Col = 2
            gridService.Text = Format(!Balance, "0.00")
            
            temTotal = temTotal + ![Balance]

            .MoveNext
        Wend
        .Close
    End With
    
    gridService.Visible = True
    
    txtTotal.Text = Format(temTotal, "0.00")
    
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




