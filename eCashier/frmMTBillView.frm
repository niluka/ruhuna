VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMTBillsView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medical Service Bills"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   FillColor       =   &H0080FFFF&
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
   ScaleHeight     =   8565
   ScaleWidth      =   8565
   Begin VB.TextBox txtDisplayBillID 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   1080
      Width           =   4575
   End
   Begin MSDataListLib.DataCombo cmbHSS 
      Height          =   360
      Left            =   4200
      TabIndex        =   13
      Top             =   6240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.CheckBox chkForeigner 
      Caption         =   "&Foreigner"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtPaymentMethod 
      Height          =   855
      Left            =   4200
      TabIndex        =   7
      Top             =   6720
      Width           =   4095
   End
   Begin VB.TextBox txtMedicalTestBillID 
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   16777088
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
   Begin MSFlexGridLib.MSFlexGrid gridService 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5741
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbPatient 
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
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   4200
      TabIndex        =   5
      Top             =   6240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   1560
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   78446595
      CurrentDate     =   39956
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   2040
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   78446594
      CurrentDate     =   39956
   End
   Begin VB.Label Label17 
      Caption         =   "&Serial No"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Time"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblHSS 
      Caption         =   "Credit Company"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Payment Co&mments"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF8080&
      Caption         =   "Payment &Method"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   5880
      Width           =   4095
   End
   Begin VB.Label Label6 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Pa&tient"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Bill Date && Time"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmMTBillsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsSC As New ADODB.Recordset
    Dim temSql As String
    Dim rsSPC As New ADODB.Recordset
    Dim rsStaff() As New ADODB.Recordset
    Dim PSCCount As Long
    Dim FirstActi As Boolean
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

    Dim rsSecession As New ADODB.Recordset


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call GetSettings
End Sub

Private Sub GetSettings(): On Error Resume Next
'    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
'    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
    GetCommonSettings Me
End Sub

Private Sub SaveSettings()
'    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
'    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    SaveCommonSettings Me
End Sub

Private Sub FillCombos()
    Dim BHT As New clsFillCombos
    ''BHT.FillSpecificIDField cmbPatient, "PatientMainDetails", "PatientID", "FirstName", False
    Dim PayM As New clsFillCombos
    PayM.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanReceive", False
    With rsHSS
        If .State = 1 Then .Close
        temSql = "SELECT tblHealthSchemeSuppliers.* FROM tblHealthSchemeSuppliers ORDER BY tblHealthSchemeSuppliers.HealthSchemeSupplierName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbHSS
        Set .RowSource = rsHSS
        .ListField = "HealthSchemeSupplierName"
        .BoundColumn = "HealthSchemeSupplierID"
    End With
End Sub

Private Sub FillGrid()
    Call FormatGrid
    Dim rsTem As New ADODB.Recordset
    Dim TotalCharge As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientService.PatientServiceID, tblPatientService.ServiceDate, tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory, tblPatientService.Comments, tblPatientService.Charge, tblServiceSecession.ServiceSecession, tblPatientService.SerialNo " & _
                    "FROM ((tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID) LEFT JOIN tblServiceSecession ON tblPatientService.SecessionID = tblServiceSecession.ServiceSecessionID " & _
                    "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.MedicalTestBillID)<> 0)  AND ((tblPatientService.MedicalTestBillID)=" & Val(txtMedicalTestBillID.Text) & ")) " & _
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
                gridService.Text = !ServiceSubcategory   '  !ServiceCategory & " - " & !ServiceSubCategory
            End If
            gridService.Col = 3
            gridService.Text = !Comments
            gridService.Col = 4
            gridService.Text = Format(!Charge, "0.00")
            TotalCharge = TotalCharge + !Charge
            
            gridService.Col = 5
            gridService.Text = Format(!ServiceSecession, "")
            gridService.Col = 6
            gridService.Text = Format(!SerialNo, "")
            
            
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
    '   5   Secession
    '   6   Serial
    
    With gridService
        .Cols = 7
        .Rows = 1
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        .ColWidth(0) = 0
        
        
        .Col = 1
        .ColWidth(1) = 0
        .Text = "Date"
        
        .Col = 2
        .ColWidth(2) = 2500
        .Text = "Service"
        
        .Col = 3
        .ColWidth(3) = 2500
        .Text = "Comments "
        
        .Col = 4
        .ColWidth(4) = 1200
        .Text = "Charge"
        
        .Col = 5
        .ColWidth(4) = 1200
        .Text = "Secession"
        
        .Col = 6
        .ColWidth(4) = 1200
        .Text = "Serial"
        
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        
        
    End With
    lblTotal.Caption = "0.00"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub DisplayBillDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where IncomeBillID = " & Val(txtMedicalTestBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtDisplayBillID.Text = !DisplayBillID
            dtpDate.Value = !CompletedDate
            dtpTime.Value = !CompletedTime
            cmbPatient.BoundText = !PatientID
            cmbPaymentMethod.BoundText = !PaymentMethodID
            txtPaymentMethod.Text = !PaymentComments
            lblTotal.Caption = Format(!GrossTotal, "0.00")
            cmbHSS.BoundText = !HSSID
        Else
            MsgBox "Error. Bill NOT added. Please reenter"
            Unload Me
            Exit Sub
        End If
        .Close
    End With
End Sub

'Private Sub PopulatePrinters()
'    Dim MyPrinter As Printer
'    For Each MyPrinter In Printers
'        cmbPrinter.AddItem MyPrinter.DeviceName
'    Next
'End Sub
'
'Private Sub PopulatePapers(): on error resume next
'    cmbPaper.Clear
'    SetPrinter = False
'    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text) : doevents
'    PrinterName = Printer.DeviceName
'    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
''        With FormSize
''            .cx = BillPaperHeight
''            .cy = BillPaperWidth
''        End With
'        ReDim aFI1(1)
'        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
'        ReDim Temp(BytesNeeded)
'        ReDim aFI1(BytesNeeded / Len(FI1))
'        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
'        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
'        For i = 0 To NumForms - 1
'            With aFI1(i)
'                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
'                'ComboBillPrinterPapers.AddItem FormItem
'                cmbPaper.AddItem PtrCtoVbString(.pName)
''                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
'            End With
'        Next i
'        ClosePrinter (PrinterHandle) : doevents
'    End If
'End Sub
'
'Private Sub cmbPrinter_Change()
'    Call PopulatePapers
'End Sub
'
'Private Sub cmbPrinter_Click()
'    Call PopulatePapers
'End Sub

Private Sub txtMedicalTestBillID_Change()
    Call FormatGrid
    Call FillGrid
    Call DisplayBillDetails
End Sub
