VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMTCreditLetters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Settle Letter for Medical Test Bills"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13185
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
   ScaleHeight     =   9300
   ScaleWidth      =   13185
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11880
      TabIndex        =   8
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
      Left            =   10560
      TabIndex        =   7
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
   Begin MSDataListLib.DataCombo cmdHSS 
      Height          =   360
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid gridMT 
      Height          =   7095
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   12515
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   77987843
      CurrentDate     =   40078
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   77987843
      CurrentDate     =   40078
   End
   Begin btButtonEx.ButtonEx btnToExcel 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   8760
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
   Begin VB.Label Label3 
      Caption         =   "Company"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
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
Attribute VB_Name = "frmMTCreditLetters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsHSS As New ADODB.Recordset
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
    Dim MyTotal As Double
    Dim MyHSS As New clsHSS
    Dim AllLines() As String
    Dim myworkbook As Excel.Workbook
    Dim myworksheet As Excel.Worksheet
    Dim mychart As Excel.Chart
    Dim TemPath As String
    Dim FSys As New Scripting.FileSystemObject
    TemPath = FSys.GetParentFolderName(Database) & "\LetterMT.xls"
    If FSys.FileExists(TemPath) = False Then
        MsgBox "The Excel file is not located in the database folder"
        Exit Sub
    End If
    Set myworkbook = GetObject(TemPath)
    Set myworksheet = myworkbook.Worksheets.Item(1)
    MyHSS.HSSID = Val(cmdHSS.BoundText)
    
    myworksheet.Cells(8, 7) = Format(Date, "dd MMMM yyyy")
    myworksheet.Cells(16, 3) = "Invoice for the Period from " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    
    AllLines = SeperateLines(MyHSS.Name & vbNewLine & MyHSS.Address)
    For i = 0 To UBound(AllLines) - 1
        myworksheet.Cells(10 + i, 1) = AllLines(i)
    Next
    
    With gridMT
        For i = 1 To .Rows - 1
            myworksheet.Cells(19 + i, 1) = .TextMatrix(i, 0)
            myworksheet.Cells(19 + i, 3) = .TextMatrix(i, 1)
            myworksheet.Cells(19 + i, 4) = .TextMatrix(i, 2)
            myworksheet.Cells(19 + i, 5) = .TextMatrix(i, 3)
            myworksheet.Cells(19 + i, 6) = .TextMatrix(i, 4)
            myworksheet.Cells(19 + i, 7) = .TextMatrix(i, 5)
            MyTotal = MyTotal + Val(.TextMatrix(i, 5))
        Next
    End With
    
    myworksheet.Cells(20 + i, 7) = Format(MyTotal, "#,##0.00")
    
    myworksheet.Cells(24 + i, 1) = "Thanking you,"
    myworksheet.Cells(25 + i, 1) = "Yours Faithfully"
    myworksheet.Cells(27 + i, 1) = "Accountant"
    myworksheet.Cells(28 + i, 1) = "Ruhunu Hospital (Pvt) Ltd."
    
    
    If FSys.FileExists(FSys.GetParentFolderName(Database) & "\CreditLetterMT " & MyHSS.Name & Format(Date, " dd MM yyyy") & ".xls") = True Then
        FSys.DeleteFile (FSys.GetParentFolderName(Database) & "\CreditLetterMT " & MyHSS.Name & Format(Date, " dd MM yyyy") & ".xls"), True
    End If
    
    myworkbook.SaveAs FSys.GetParentFolderName(Database) & "\CreditLetterMT " & MyHSS.Name & Format(Date, " dd MM yyyy") & ".xls"

    
    
    
    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
    
    
    If SelectForm(ReportPaperName, Me.hwnd) = 1 Then
        myworksheet.PrintOut
    End If

End Sub

Private Sub btnToExcel_Click()
    GridToExcel gridMT, "Medical Test Bills", Format(dtpFrom.Value, "dd MMMM yyyy") & vbTab & Format(dtpTo.Value, "dd MMMM yyyy") & vbTab & cmdHSS.Text
End Sub

Private Sub cmdHSS_Change()
    Call FormatGrid
    Call FillGrid
End Sub


Private Sub dtpTo_LostFocus()
    Call FormatGrid
    Call FillGrid

End Sub

Private Sub dtpFrom_LostFocus()
    Call FormatGrid
    Call FillGrid

End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call GetSettings
    Call FillCombos
    Call FillGrid
End Sub

Private Sub GetSettings(): On Error Resume Next
    GetCommonSettings Me
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
End Sub

Private Sub FillCombos()
    With rsHSS
        If .State = 1 Then .Close
        temSql = "SELECT tblHealthSchemeSuppliers.HealthSchemeSupplierName, tblHealthSchemeSuppliers.HealthSchemeSupplierID " & _
                    "From tblHealthSchemeSuppliers " & _
                    "ORDER BY tblHealthSchemeSuppliers.HealthSchemeSupplierName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmdHSS
        Set .RowSource = rsHSS
        .ListField = "HealthSchemeSupplierName"
        .BoundColumn = "HealthSchemeSupplierID"
    End With
End Sub

Private Sub FillGrid()
    Dim rsMTBill As New ADODB.Recordset
    Dim rsDate As New ADODB.Recordset
    Dim rsIx As New ADODB.Recordset

    If rsDate.State = 1 Then rsDate.Close
    temSql = "SELECT tblIncomeBill.CompletedDate " & _
                "From tblIncomeBill " & _
                "GROUP BY tblIncomeBill.IsMedicalTestBill, tblIncomeBill.Completed, tblIncomeBill.Cancelled, tblIncomeBill.CompletedDate, tblIncomeBill.HSSID " & _
                "HAVING (((tblIncomeBill.IsMedicalTestBill)=1) AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.CompletedDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND  ((tblIncomeBill.HSSID)=1)) " & _
                "ORDER BY tblIncomeBill.CompletedDate"
    rsDate.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    While rsDate.EOF = False
        gridMT.Rows = gridMT.Rows + 1
        gridMT.Row = gridMT.Rows - 1
        gridMT.Col = 0
        gridMT.Text = Format(rsDate!CompletedDate, "DD MMMM yyyy")
    
        If rsMTBill.State = 1 Then rsMTBill.Close
        temSql = "SELECT tblIncomeBill.*, tblPatientMainDetails.FirstName " & _
                    "FROM tblIncomeBill LEFT JOIN tblPatientMainDetails ON tblIncomeBill.PatientID = tblPatientMainDetails.PatientID " & _
                    "WHERE (((tblIncomeBill.IsMedicalTestBill)=1) AND ((tblIncomeBill.CompletedDate)='" & Format(rsDate!CompletedDate, "dd MMMM yyyy") & "') AND  ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.HSSID)=" & Val(cmdHSS.BoundText) & ")) "
        rsMTBill.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While rsMTBill.EOF = False
            If gridMT.TextMatrix(gridMT.Rows - 1, 2) = "" Then
            
            Else
                gridMT.Rows = gridMT.Rows + 1
                gridMT.Row = gridMT.Rows - 1
            End If
            
            gridMT.Col = 1
            gridMT.Text = Format(rsMTBill!PaymentComments, "")
            gridMT.Col = 3
            gridMT.Text = Format(rsMTBill!FirstName, "")
            gridMT.Col = 2
            gridMT.Text = Format(rsMTBill!DisplayBillID, "0")
            
            If rsIx.State = 1 Then rsIx.Close
            temSql = "SELECT tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory, tblPatientService.Charge  " & _
                        "FROM ((tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID) LEFT JOIN tblServiceSecession ON tblPatientService.SecessionID = tblServiceSecession.ServiceSecessionID " & _
                        "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.MedicalTestBillID)<> 0)  AND ((tblPatientService.MedicalTestBillID)=" & Format(rsMTBill!IncomeBillID, "0") & ")) " & _
                        "ORDER BY tblPatientService.PatientServiceID"
            rsIx.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            While rsIx.EOF = False
                If gridMT.TextMatrix(gridMT.Rows - 1, 4) = "" Then
                
                Else
                    gridMT.Rows = gridMT.Rows + 1
                    gridMT.Row = gridMT.Rows - 1
                End If
                If Format(rsIx!ServiceSubcategory, "") <> "" Then
                    gridMT.Col = 4
                    gridMT.Text = Format(rsIx!ServiceSubcategory, "")
                Else
                    gridMT.Col = 4
                    gridMT.Text = Format(rsIx!ServiceCategory, "")
                End If
                gridMT.Col = 5
                gridMT.Text = rsIx!Charge
                rsIx.MoveNext
            Wend
            
            rsMTBill.MoveNext
        Wend
        
        rsDate.MoveNext
    Wend
        
        
End Sub

Private Sub FormatGrid()
    With gridMT
        .Clear
        .Rows = 1
        .Cols = 6
        
        .Col = 0
        .Text = "Date"
        
        .Col = 1
        .Text = "Proposal No."
        
        .Col = 2
        .Text = "Receipt No."
        
        .Col = 3
        .Text = "Name"
        
        .Col = 4
        .Text = "Test"
        
        .Col = 5
        .Text = "Amount"
        
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub
