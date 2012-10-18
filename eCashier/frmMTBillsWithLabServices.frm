VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMTBillsWithLabServices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Summery Details"
   ClientHeight    =   8985
   ClientLeft      =   3930
   ClientTop       =   -180
   ClientWidth     =   11145
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
   ScaleWidth      =   11145
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   3120
      TabIndex        =   9
      Top             =   8040
      Width           =   2415
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   3120
      TabIndex        =   7
      Top             =   8520
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid gridBill 
      Height          =   5895
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10398
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9840
      TabIndex        =   1
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
      Left            =   8520
      TabIndex        =   0
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   78118915
      CurrentDate     =   39969
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   7200
      TabIndex        =   12
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Excel"
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
      Left            =   6360
      TabIndex        =   13
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   78118915
      CurrentDate     =   39969
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   1800
      TabIndex        =   15
      Top             =   1560
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Category"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "To"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "From"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Total Value"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   8520
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Total Count"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   8040
      Width           =   2775
   End
   Begin VB.Label lblTopic 
      Alignment       =   2  'Center
      Caption         =   "Topic"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10815
   End
   Begin VB.Label lblSubtopic 
      Alignment       =   2  'Center
      Caption         =   "Topic"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   10815
   End
End
Attribute VB_Name = "frmMTBillsWithLabServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsTem As New ADODB.Recordset
    Dim MyGrid As Grid
    Dim MyGridRow() As GridRow
    Dim MyGridCell() As GridCell
    Dim i As Integer
    Dim n As Integer
    Dim TotalValue As Double
    Dim TotalCount As Long

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridBill, "Medical Test Bills with Lab Services", "From : " & Format(dtpFrom.Value, "dd MMMM yyyy") & " To : " & Format(dtpTo.Value, "dd MMMM yyyy")
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    
    
    With ThisReportFormat
        .ColSpace = 500
        
        .SubTopicFontSize = 8
    
    End With
    
    GridPrint gridBill, ThisReportFormat, "Medical Test Bills with Lab Services", "From : " & Format(dtpFrom.Value, "dd MMMM yyyy") & " To : " & Format(dtpTo.Value, "dd MMMM yyyy")
    Printer.EndDoc
    
End Sub


Private Sub CatIncome()
    TotalValue = 0
    TotalCount = 0
    Dim PreviousBillNo As Long
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(cmbCategory.BoundText) = False Then
            temSql = "SELECT dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierName, dbo.tblPatientMainDetails.FirstName, dbo.tblIncomeBill.IncomeBillID, " & _
                          "dbo.tblIncomeBill.DisplayBillID, dbo.tblIncomeBill.Cancelled, dbo.tblIncomeBill.NetTotal, dbo.tblIncomeBill.PaymentComments, dbo.tblIncomeBill.CancelledDate, dbo.tblIncomeBill.CancelledTime " & _
                            "FROM         dbo.tblIncomeBill LEFT OUTER JOIN " & _
                            "dbo.tblPatientMainDetails ON dbo.tblIncomeBill.PatientID = dbo.tblPatientMainDetails.PatientID LEFT OUTER JOIN " & _
                            "dbo.tblHealthSchemeSuppliers ON dbo.tblIncomeBill.HSSID = dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierID "
            temSql = temSql & "where tblIncomeBill.CompletedDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' AND tblIncomeBill.IsMedicalTestBill = 1 AND tblIncomeBill.Completed = 1 "
            temSql = temSql & " ORDER by tblIncomeBill.DisplayBillID"
        Else
            temSql = "SELECT dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierName, dbo.tblPatientMainDetails.FirstName, dbo.tblIncomeBill.IncomeBillID, " & _
                          "dbo.tblIncomeBill.DisplayBillID, dbo.tblIncomeBill.Cancelled, dbo.tblIncomeBill.NetTotal, dbo.tblIncomeBill.PaymentComments, dbo.tblIncomeBill.CancelledDate, dbo.tblIncomeBill.CancelledTime " & _
                            "FROM         dbo.tblIncomeBill LEFT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.MedicalTestBillID  LEFT OUTER JOIN dbo.tblPatientMainDetails ON dbo.tblIncomeBill.PatientID = dbo.tblPatientMainDetails.PatientID LEFT OUTER JOIN dbo.tblHealthSchemeSuppliers ON dbo.tblIncomeBill.HSSID = dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierID "
            temSql = temSql & "where tblIncomeBill.CompletedDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' AND tblIncomeBill.IsMedicalTestBill = 1 AND tblIncomeBill.Completed = 1 AND dbo.tblPatientService.ServiceCategoryID = " & Val(cmbCategory.BoundText) & " "
            temSql = temSql & " ORDER by tblIncomeBill.DisplayBillID"
        End If
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 0 To .RecordCount - 1
                If PreviousBillNo <> !DisplayBillID Then
                    PreviousBillNo = !DisplayBillID
                    gridBill.Rows = gridBill.Rows + 1
                    gridBill.Row = gridBill.Rows - 1
                    
                    gridBill.Col = 0
                    gridBill.Text = ""  '![IncomeBillID]
                    
                    gridBill.Col = 0
                    gridBill.Text = ![DisplayBillID]
                    
                    gridBill.Col = 1
                    gridBill.CellAlignment = 1
                    gridBill.Text = !HealthSchemeSupplierName
                    
                    gridBill.Col = 2
                    gridBill.CellAlignment = 1
                    gridBill.Text = !FirstName
                    
                    
                    gridBill.Col = 3
                    gridBill.Text = Format(!NetTotal, "0.00")
                    
                    gridBill.Col = 4
                    If !Cancelled = True Then
                    gridBill.Text = "Cancelled on " & Format(!CancelledDate, "dd MM yy") & " at " & Format(!CancelledTime, "mm:hh AMPM")
                    Else
                    gridBill.Text = ""
                    TotalValue = TotalValue + !NetTotal
                    TotalCount = TotalCount + 1
                    End If
                    
                    gridBill.Col = 5
                    gridBill.Text = Format(!PaymentComments, "")
                               
                End If
                
                .MoveNext
            
            Next i
            
            
        End If
        .Close
    End With
End Sub


Private Sub CatReturn()
    
    TotalValue = 0
    TotalCount = 0
    With rsTem
        If .State = 1 Then .Close
        
            temSql = "SELECT tblIncomeReturnBill.*, tblIncomeBill.*,  tblIncomeReturnBill.IncomeBillID as RIncomeBillID, tblIncomeBill.IncomeBillID, tblBHT.BHT, tblOPDPatient.FirstName as OPDFirstName, tblBHTPatient.FirstName as BHTFirstName " & _
                        "FROM ((tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID) LEFT JOIN tblPatientMainDetails AS tblOPDPatient ON tblIncomeBill.PatientID = tblOPDPatient.PatientID) LEFT JOIN (tblBHT LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID) ON tblIncomeReturnBill.BHTID = tblBHT.BHTID "
            temSql = temSql & "Where tblIncomeReturnBill.ReturnDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' AND tblIncomeBill.IsMedicalTestBill = 1 AND tblIncomeReturnBill.Cancelled = 0 "
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            TotalCount = .RecordCount
            ReDim MyGridRow(.RecordCount)
            For i = 0 To .RecordCount - 1
                MyGridCell(0).Col = 0
'                If IsNull(![RIncomeBillID]) = True Then
'                    MyGridCell(0).Text = "BHT/GSB"
'                Else
'                    MyGridCell(0).Text = ![RIncomeBillID]
'                End If
'                MyGridCell(1).Col = 1
'
'                If IsNull(![DisplayBillID]) = True Then
'                    MyGridCell(1).Text = "BHT/GSB"
'                Else
'                    MyGridCell(1).Text = ![DisplayBillID]
'                End If
'                MyGridCell(2).Col = 1
'                MyGridCell(2).CellAlignment = 1
'                MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "")
'
'
'                If !PharmacyBillID <> 0 Then
'                    MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "") & " " & Format(![Agent], "") & " " & !PharmacyBillID
'                Else
'                    MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "") & " " & Format(![Agent], "")
'                End If
'
'
'
'                MyGridCell(3).Col = 2
'                MyGridCell(3).Text = Format(!NetTotal, "0.00")
'                MyGridCell(4).Col = 3
'                MyGridCell(4).Text = Format(!CompletedTime, "HH:MM AMPM")
'                MyGridCell(5).Col = 4
'                MyGridCell(5).Text = "Returned on " & Format(!ReturnDate, "dd MMMM yyyy") & " at " & !ReturnTime
'                MyGridCell(6).Col = 5
'                MyGridCell(6).Text = "Returned Value " & Format(!ReturnValue, "0.00")
'                MyGridRow(i).RowCells = MyGridCell
'                TotalValue = TotalValue + !ReturnValue
'                MyGridRow(i).RowCells = MyGridCell
                .MoveNext
            Next i
            
            With gridBill
                .Rows = UBound(MyGridRow) + 1
                For i = 1 To .Rows - 1
                    For n = 0 To UBound(MyGridRow(i - 1).RowCells) - 1
                        .TextMatrix(i, n) = MyGridRow(i - 1).RowCells(n).Text
                    Next n
                Next i
            End With
        End If
        .Close
    End With
End Sub

Private Sub CatCancellation()
    TotalValue = 0
    TotalCount = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierName, dbo.tblPatientMainDetails.FirstName, dbo.tblIncomeBill.IncomeBillID, " & _
                      "dbo.tblIncomeBill.DisplayBillID, dbo.tblIncomeBill.Cancelled, dbo.tblIncomeBill.NetTotal, dbo.tblIncomeBill.PaymentComments, dbo.tblIncomeBill.CancelledDate, dbo.tblIncomeBill.CancelledTime " & _
                        "FROM         dbo.tblIncomeBill LEFT OUTER JOIN " & _
                        "dbo.tblPatientMainDetails ON dbo.tblIncomeBill.PatientID = dbo.tblPatientMainDetails.PatientID LEFT OUTER JOIN " & _
                        "dbo.tblHealthSchemeSuppliers ON dbo.tblIncomeBill.HSSID = dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierID "

        
        temSql = temSql & "where tblIncomeBill.Cancelled = 1 AND tblIncomeBill.CompletedDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' AND tblIncomeBill.IsMedicalTestBill = 1 AND tblIncomeBill.Completed = 1 "
        temSql = temSql & " ORDER by tblIncomeBill.DisplayBillID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            TotalCount = TotalCount - .RecordCount
            For i = 0 To .RecordCount - 1
                
                gridBill.Rows = gridBill.Rows + 1
                gridBill.Row = gridBill.Rows - 1
                
                gridBill.Col = 0
                gridBill.Text = ![IncomeBillID]
                
                gridBill.Col = 1
                gridBill.Text = ![DisplayBillID]
               
                gridBill.Col = 2
                gridBill.CellAlignment = 1
                gridBill.Text = !HealthSchemeSupplierName
                
                gridBill.Col = 3
                gridBill.CellAlignment = 1
                gridBill.Text = !FirstName
                
                
                gridBill.Col = 4
                gridBill.Text = Format(!NetTotal, "0.00")
                
                gridBill.Col = 5
                gridBill.Text = "Cancellation"
                
                gridBill.Col = 6
                gridBill.Text = Format(!PaymentComments, "")
                
                TotalValue = TotalValue - !NetTotal
                .MoveNext
            
            Next i
            
            
        End If
        .Close
    End With
    
End Sub

Private Sub cmbCategory_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbCategory.Text = Empty
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
    Call FormatGrid
    Call GetSettings
    Call FillGrid
End Sub

Private Sub FormatGrid()
    With gridBill
        .Cols = 6

        .Rows = 1

        .Row = 0

        .Col = 0
        .Text = "ID"

        .Col = 0
        .Text = "Bill ID"

        .Col = 1
        .Text = "Company"

        .Col = 2
        .Text = "Patient"

        .Col = 3
        .Text = "Value"

        .Col = 4
        .Text = "Remarks"

        .Col = 5
        .Text = "Comments"
        
        .ColWidth(0) = 1000
        
    End With
End Sub

Private Sub FillGrid()
    Dim TemString As String
    
    TotalValue = 0
    
    CatIncome
'    CatCancellation
'    CatReturn
    
    With gridBill
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 1
        .Text = "Total"
        .Col = 3
        .Text = Format(TotalValue, "0.00")
        
    End With
    
    
    txtValue.Text = Format(TotalValue, "0.00")
    txtCount.Text = TotalCount
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCategory, "ServiceCategory", "ServiceCategory", "ForMT", True

End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpFrom.Value = Date ' GetSetting(App.EXEName, Me.Name, dtpFrom.Name, Date)
    dtpTo.Value = Date 'GetSetting(App.EXEName, Me.Name, dtpTo.Name, Date)
    GetCommonSettings Me
    lblTopic.Caption = HospitalName
    lblSubtopic.Caption = "Medical Test Bills with Lab Services - From : " & Format(dtpFrom.Value, "dd MMMM yyyy") & " To : " & Format(dtpTo.Value, "dd MMMM yyyy")
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

