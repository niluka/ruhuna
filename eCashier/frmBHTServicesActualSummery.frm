VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBHTServicesActualSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BHT Services Actual Summery"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11070
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
   ScaleHeight     =   8160
   ScaleWidth      =   11070
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9720
      TabIndex        =   6
      Top             =   7440
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
   Begin MSFlexGridLib.MSFlexGrid gridService 
      Height          =   5175
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9128
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSC 
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   7440
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
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   78970883
      CurrentDate     =   40182
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   78970883
      CurrentDate     =   40182
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   7080
      TabIndex        =   13
      Top             =   7440
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
   Begin VB.Label Label12 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "To"
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   120
      Width           =   735
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
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   6960
      Width           =   3615
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
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label lblSC 
      Caption         =   "Service Subcategory"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Service Category"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmBHTServicesActualSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsSC As New ADODB.Recordset
    Dim temSql As String
    Dim rsSPC As New ADODB.Recordset
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridService, "BHT Services Actual Summery", "From " & Format(dtpFromDate.Value, "dd MMMM yyyy") & " to " & Format(dtpToDate.Value, "dd MMMM yyyy")
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    GridPrint gridService, ThisReportFormat, "BHT Services Actual Summery", "From " & Format(dtpFromDate.Value, "dd MMMM yyyy") & " to " & Format(dtpToDate.Value, "dd MMMM yyyy")
    Printer.EndDoc
End Sub

Private Sub cmbCategory_Change()
    If IsNumeric(cmbCategory.BoundText) = False Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsSC
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSubCategory where   Deleted = 0 AND ForBHT = 1 AND ServiceCategoryID = " & Val(cmbCategory.BoundText) & " Order By ServiceSubCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbSC
        Set .RowSource = rsSC
        .ListField = "ServiceSubcategory"
        .BoundColumn = "ServiceSubcategoryID"
        .Text = Empty
    End With
    Call FormatGrid
    Call FillGrid
    
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If cmbSC.Visible = False Then
            SendKeys "{home}+{end}"
        Else
            cmbSC.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbCategory.Text = Empty
    End If
End Sub


Private Sub cmbSC_Change()
    Call FormatGrid
    Call FillGrid

End Sub

Private Sub cmbSC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
    ElseIf KeyCode = vbKeyEscape Then
        cmbSC.Text = Empty
    End If
    
End Sub

Private Sub dtpFromDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpToDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub


Private Sub Form_Load()
    Call GetSettings
    Call FillCombos
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
    GetCommonSettings Me
End Sub


Private Sub SaveSettings()
    SaveCommonSettings Me
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCategory, "ServiceCategory", "ServiceCategory", "ForBHT", True
End Sub

Private Sub FillGrid()
    Screen.MousePointer = vbHourglass
    gridService.Visible = False
    
    Call FormatGrid
    Dim rsTem As New ADODB.Recordset
    Dim TotalCharge As Double
    With rsTem
        If .State = 1 Then .Close
'        temSql = "SELECT tblPatientService.PatientServiceID, tblPatientService.ServiceDate, tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory, tblPatientService.Comments, tblPatientService.Charge " & _
'                    "FROM (tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID " & _
'                    "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.BHTID)=" & Val(cmbBHT.BoundText) & ")) " & _
'                    "ORDER BY tblPatientService.PatientServiceID"
        
        temSql = "SELECT dbo.tblBHT.BHT, dbo.tblPatientService.BHTID, dbo.tblPatientMainDetails.FirstName, dbo.tblPatientService.Charge, dbo.tblPatientService.Comments, dbo.tblPaymentMethod.PaymentMethod, dbo.tblPatientService.ServiceDate "
        
        temSql = temSql & "FROM dbo.tblPatientMainDetails RIGHT OUTER JOIN dbo.tblBHT ON dbo.tblPatientMainDetails.PatientID = dbo.tblBHT.PatientID RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblBHT.BHTID = dbo.tblPatientService.BHTID LEFT OUTER JOIN dbo.tblPaymentMethod ON dbo.tblBHT.PaymentMethodID = dbo.tblPaymentMethod.PaymentMethodID "
        
        temSql = temSql & "WHERE dbo.tblPatientService.Deleted = 0 AND dbo.tblBHT.BHTID <> 0 AND dbo.tblBHT.IsBHT =1 AND dbo.tblPatientService.ServiceDate BETWEEN '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' "
        
        If IsNumeric(cmbCategory.BoundText) = True Then
            temSql = temSql & " AND dbo.tblPatientService.ServiceCategoryID = " & Val(cmbCategory.BoundText) & " "
        End If
        If IsNumeric(cmbSC.BoundText) = True Then
            temSql = temSql & " AND dbo.tblPatientService.ServiceSubcategoryID = " & Val(cmbSC.BoundText) & " "
        End If
        
        temSql = temSql & ""
        
        temSql = temSql & ""
        
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridService.Rows = gridService.Rows + 1
            gridService.Row = gridService.Rows - 1
            
            gridService.Col = 0
            gridService.Text = Format(!BHT, "")
            
            gridService.Col = 1
            gridService.Text = !BHTID
            
            gridService.Col = 2
            gridService.Text = Format(!FirstName, "")
                        
            gridService.Col = 3
            gridService.Text = Format(!Charge, "0.00")
            TotalCharge = TotalCharge + !Charge
            
            gridService.Col = 4
            gridService.Text = !Comments
            
            gridService.Col = 5
            gridService.Text = Format(!PaymentMethod, "0.00")
            
            gridService.Col = 6
            gridService.Text = Format(!ServiceDate, "dd MMMM yyyy")
            
            
            
            .MoveNext
        Wend
    End With
    lblTotal.Caption = Format(TotalCharge, "0.00")

    Screen.MousePointer = vbDefault
    gridService.Visible = True


End Sub

Private Sub FormatGrid()
    With gridService
        .Cols = 7
        .Rows = 1
        .Row = 0
        
'        .ColWidth(0) = 1600
        
        .Col = 0
        .Text = "BHT"
        
        .Col = 1
        .Text = "Final Bill No"
        
        .Col = 2
        .Text = "Patient"
        
        .Col = 3
        .Text = "Amount"
        
        .Col = 4
        .Text = "Comment"
        
        .Col = 5
        .Text = "Paid As"
        
        .Col = 6
        .Text = "Date"
        
    End With
    lblTotal.Caption = "0.00"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSettings
End Sub
