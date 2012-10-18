VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmServiceCategoryList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Services"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9990
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
   ScaleHeight     =   8970
   ScaleWidth      =   9990
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   7080
      TabIndex        =   21
      Top             =   7920
      Width           =   2775
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   7080
      TabIndex        =   20
      Top             =   7440
      Width           =   2775
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8640
      TabIndex        =   17
      Top             =   8400
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
      Left            =   7320
      TabIndex        =   16
      Top             =   8400
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
      Height          =   5175
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9128
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   7440
      Width           =   5175
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   720
         TabIndex        =   12
         Top             =   720
         Width           =   4215
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label18 
         Caption         =   "Paper"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSDataListLib.DataCombo cmbCat 
      Height          =   360
      Left            =   1440
      TabIndex        =   7
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
      TabIndex        =   5
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   15990787
      CurrentDate     =   40028
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   15990787
      CurrentDate     =   40028
   End
   Begin MSDataListLib.DataCombo cmbSC 
      Height          =   360
      Left            =   1440
      TabIndex        =   8
      Top             =   1200
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbS 
      Height          =   360
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label7 
      Caption         =   "Total Count"
      Height          =   375
      Left            =   5640
      TabIndex        =   19
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Total Value"
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   7440
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Secession"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
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
Attribute VB_Name = "frmServiceCategoryList"
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

Private Sub cmbCat_Change()
    Call FormatGrid
    Call FillGrid
    SCC.FillSpecificIDFielterField cmbSC, "ServiceSubcategory", "ServiceCategoryID", Val(cmbCat.BoundText), "ServiceSubcategory", True
    SCS.FillSpecificIDFielterField cmbS, "ServiceSecession", "ServiceCategoryID", Val(cmbCat.BoundText), "ServiceSecession", True
    cmbSC.Text = Empty
    cmbS.Text = Empty
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

Private Sub cmbS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnPrint.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbS.Text = Empty
    End If
End Sub

Private Sub cmbSC_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbSC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbS.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
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
    Call GetSettings
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub printBill()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        
        
        Printer.EndDoc
        
    End If
End Sub

Private Sub FillCombos()
    Dim SC As New clsFillCombos
    SC.FillAnyCombo cmbCat, "ServiceCategory", True
'    Dim PayMethod As New clsFillCombos
'    PayMethod.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanPay", False
'    Dim SC As New clsFillCombos
'    SC.FillAnyCombo cmbSpeciality, "ServiceSubCategory", True
'    Dim PtPayMethod As New clsFillCombos
'    PtPayMethod.FillAnyCombo cmbPtPM, "PaymentMethod", False
    
End Sub

Private Sub GetSettings()
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
    cmbCat.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbCat.Name, 1))
    cmbSC.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbSC.Name, 1))
    cmbS.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbS.Name, 1))
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "")
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbSC.Name, cmbSC.BoundText
    SaveSetting App.EXEName, Me.Name, cmbS.Name, cmbS.BoundText
    SaveSetting App.EXEName, Me.Name, cmbCat.Name, cmbCat.BoundText
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
End Sub

Private Sub FormatGrid()
    With gridService
        .Clear
    
        .Rows = 1
        .Cols = 7
        
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        
        .Col = 1
        .Text = "Date"
        
        .Col = 2
        .Text = "Time"
        
        .Col = 3
        .Text = "Category"
        
        .Col = 4
        .Text = "Subcategory"
        
        .Col = 5
        .Text = "Secession"
        
        .Col = 6
        .Text = "Fee"
        
    
    End With
End Sub

Private Sub FillGrid()
    txtTotal.Text = Empty
    txtCount.Text = Empty
    
    Dim temTotal As Double
    Dim temCount As Double
    
    Dim rstem As New ADODB.Recordset
    
    Dim temText As String
    
    With rstem
        If .State = 1 Then .Close
        
        temSql = "SELECT tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory, tblServiceSecession.ServiceSecession, tblPatientService.* " & _
                    "FROM ((tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID) LEFT JOIN tblServiceSecession ON tblPatientService.SecessionID = tblServiceSecession.ServiceSecessionID " & _
                    "WHERE tblPatientService.ServiceDate Between #" & Format(dtpFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpTo.Value, "dd MMMM yyyy") & "# "
        temSql = temSql & " AND tblPatientService.Deleted = False "
        
        If IsNumeric(cmbCat.BoundText) = True Then
            temSql = temSql & "  AND tblPatientService.ServiceCategoryID = " & cmbCat.BoundText & " "
        End If
        
        If IsNumeric(cmbSC.BoundText) = True Then
            temSql = temSql & "  AND tblPatientService.ServiceSubCategoryID = " & cmbSC.BoundText & " "
        End If
        
        If IsNumeric(cmbS.BoundText) = True Then
            temSql = temSql & "AND tblPatientService.ServiceSubcategoryID = " & cmbS.BoundText & " "
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        gridService.Visible = False
        
        While .EOF = False
            gridService.Rows = gridService.Rows + 1
            gridService.Row = gridService.Rows - 1
            
            gridService.Col = 0
            gridService.Text = !PatientServiceID
            
            gridService.Col = 1
            gridService.Text = Format(!ServiceDate, "dd MMM yy")
            
            gridService.Col = 2
            gridService.Text = Format(!ServiceTime, "hh:mm AMPM")
            
            gridService.Col = 3
            gridService.Text = Format(!ServiceCategory, "")
            
            gridService.Col = 4
            gridService.Text = Format(!ServiceSubCategory, "")
            
            gridService.Col = 5
            gridService.Text = Format(!ServiceSecession, "")
            
            gridService.Col = 6
            gridService.Text = Format(!Charge, "")
            
            temTotal = temTotal + ![Charge]
            temCount = temCount + 1
            .MoveNext
        Wend
        .Close
    End With
    
    gridService.Visible = True
    
    txtTotal.Text = Format(temTotal, "0.00")
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


