VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmExpenceDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expence Details"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14055
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
   ScaleHeight     =   9345
   ScaleWidth      =   14055
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   11160
      TabIndex        =   18
      Top             =   8280
      Width           =   2775
   End
   Begin VB.TextBox txtHTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   11160
      TabIndex        =   17
      Top             =   7800
      Width           =   2775
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   12720
      TabIndex        =   15
      Top             =   8760
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
      Left            =   11400
      TabIndex        =   14
      Top             =   8760
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
      Height          =   6015
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   10610
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   8520
      Width           =   10215
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Width           =   4215
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label18 
         Caption         =   "Paper"
         Height          =   255
         Left            =   5280
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSDataListLib.DataCombo cmbCat 
      Height          =   360
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   77791235
      CurrentDate     =   40028
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   77791235
      CurrentDate     =   40028
   End
   Begin MSDataListLib.DataCombo cmbSC 
      Height          =   360
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   7800
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.Label Label7 
      Caption         =   "Total Count"
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   8280
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "Total Expences"
      Height          =   255
      Left            =   8640
      TabIndex        =   16
      Top             =   7800
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Subcategory"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmExpenceDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim SCC As New clsFillCombos
    Dim SCS As New clsFillCombos

    Dim temSql As String
    Dim rsStaff As New ADODB.Recordset
    
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
    
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridService, Me.Caption, "From " & Format(dtpFrom.Value, "dd mmmm yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy") & vbTab & cmbCat.Text & vbTab & cmbSC.Text
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    GridPrint gridService, ThisReportFormat, Me.Caption, "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    Printer.EndDoc

End Sub

Private Sub cmbCat_Change()
    Call FormatGrid
    Call FillGrid
    SCC.FillSpecificIDFielterField cmbSC, "ServiceSubcategory", "ServiceCategoryID", Val(cmbCat.BoundText), "ServiceSubcategory", True
    cmbSC.Text = Empty
End Sub

Private Sub cmbCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbSC.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbCat.Text = Empty
    End If
End Sub

Private Sub cmbS_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbSC_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbSC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbSC.Text = Empty
    End If
End Sub

Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
    Call FormatGrid
    Call GetSettings
    Call FillGrid
End Sub

Private Sub printBill()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents: DoEvents
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        
        
        Printer.EndDoc
        
    End If
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCat, "ServiceCategory", "ServiceCategory", "ForExpence", True
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpFrom.Value = Date
    dtpTo.Value = Date
    cmbCat.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbCat.Name, 1))
    cmbSC.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbSC.Name, 1))
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
    SaveSetting App.EXEName, Me.Name, cmbSC.Name, cmbSC.BoundText
    SaveSetting App.EXEName, Me.Name, cmbCat.Name, cmbCat.BoundText
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
        .Cols = 6
        
        .Row = 0
        
        .Col = 0
        .Text = "Date"
        
        .Col = 1
        .Text = "Time"
                
        
        .Col = 2
        .Text = "Category"
        
        .Col = 3
        .Text = "Subcategory"
        
        .Col = 4
        .Text = "Comments "
        
        .Col = 5
        .Text = "Value"
        
    End With
End Sub

Private Sub FillGrid()
    txtHTotal.Text = Empty
    txtCount.Text = Empty
    
    Dim temHTotal As Double
    Dim temPTotal As Double
    Dim temCount As Double
    
    Dim rsTem As New ADODB.Recordset
    
    Dim temText As String
    
    With rsTem
        If .State = 1 Then .Close
        
        temSql = "SELECT tblPatientService.PatientServiceID, tblPatientService.Charge, tblPatientService.Comments, tblIncomeBill.CompletedTime, tblIncomeBill.CompletedDate, tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory " & _
                    "FROM tblIncomeBill RIGHT JOIN ((tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID) ON tblIncomeBill.IncomeBillID = tblPatientService.ExpenceBillID " & _
                    "WHERE tblIncomeBill.Cancelled = 0  AND tblIncomeBill.Completed = 1 AND tblIncomeBill.CompletedDate Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbCat.BoundText) = True Then
            temSql = temSql & "  AND tblPatientService.ServiceCategoryID = " & cmbCat.BoundText & " "
        End If
        
        If IsNumeric(cmbSC.BoundText) = True Then
            temSql = temSql & "  AND tblPatientService.ServiceSubCategoryID = " & cmbSC.BoundText & " "
        End If
        temSql = temSql & " AND tblPatientService.ExpenceBillID <> 0 "
        temSql = temSql & " ORDER BY tblPatientService.PatientServiceID"
        
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        gridService.Visible = False
        
        While .EOF = False
            gridService.Rows = gridService.Rows + 1
            gridService.Row = gridService.Rows - 1
            
            gridService.Col = 0
            gridService.Text = Format(!CompletedDate, "dd MMMM yyyy")
            
            gridService.Col = 1
            gridService.Text = Format(!CompletedTime, "MM HH AMPM")
            
            gridService.Col = 2
            gridService.Text = !ServiceCategory
            
            gridService.Col = 3
            gridService.Text = Format(!ServiceSubcategory, "") & Space(15)
            
            
            gridService.Col = 4
            gridService.Text = Format(!Comments, "0.00")
            
            gridService.Col = 5
            gridService.Text = Format(!Charge, "0.00")
            temHTotal = temPTotal + Format(!Charge, "0.00")
            
            .MoveNext
        Wend
        .Close
    End With
    
    gridService.Rows = gridService.Rows + 1
    gridService.Row = gridService.Rows - 1

    gridService.Col = 0
    gridService.Text = "Totals"

    gridService.Col = 5
    gridService.Text = Format(temHTotal, "0.00")

    
    gridService.Visible = True
    txtHTotal.Text = Format(temHTotal, "0.00")
    txtCount.Text = temCount
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


Private Sub optBHT_Click()
    Call FormatGrid
    Call FillGrid

End Sub

Private Sub optLab_Click()
    Call FormatGrid
    Call FillGrid

End Sub

Private Sub optOPD_Click()
    Call FormatGrid
    Call FillGrid

End Sub

