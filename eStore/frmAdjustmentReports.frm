VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdjustmentReports 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjustment Report"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
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
   ScaleHeight     =   7170
   ScaleWidth      =   8745
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   6480
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   6480
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
   Begin MSFlexGridLib.MSFlexGrid gridA 
      Height          =   4695
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8281
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbICat 
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67633155
      CurrentDate     =   39904
   End
   Begin MSComCtl2.DTPicker dtpto 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   720
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67633155
      CurrentDate     =   39904
   End
   Begin btButtonEx.ButtonEx btnFill 
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Fill"
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
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Top             =   6480
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
   Begin VB.Label Label3 
      Caption         =   "Item Category"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAdjustmentReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsICat As New ADODB.Recordset
    Dim AppExcel As Excel.Application
    
    Dim myworkbook As Excel.Workbook
    
    Dim myWorkSheet1 As Excel.Worksheet
    Dim myWorkSheet2 As Excel.Worksheet
    Dim myWorkSheet3 As Excel.Worksheet
    
    Dim tempath As String
    Dim FSys As New Scripting.FileSystemObject

    Dim temTopic As String
    Dim temSubTopic As String

    
Private Sub btnExcel_Click()
    Screen.MousePointer = vbHourglass
    frmPleaseWait.Show
    
    Dim i As Integer
    
    Dim TemRangeAddress As String
    
    Set AppExcel = CreateObject("Excel.Application")
    Set myworkbook = AppExcel.Workbooks.Add
    Set myWorkSheet1 = AppExcel.Worksheets(1)
    Set myWorkSheet2 = AppExcel.Worksheets(2)
    Set myWorkSheet3 = AppExcel.Worksheets(3)
    
    
    temTopic = HospitalName
    temSubTopic = "Stock Adjustment - " & cmbICat.Text
    
    myWorkSheet1.Cells(1, 1).Value = temTopic
    myWorkSheet1.Cells(2, 1).Value = temSubTopic
    
    myWorkSheet1.Cells(3, 1) = "Item"
    myWorkSheet1.Cells(3, 2) = "Batch"
    myWorkSheet1.Cells(3, 3) = "Adjustment"
    
    i = 3
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.Display, tblBatch.Batch, tblAdjustment.Amount " & _
                    "FROM (tblAdjustment LEFT JOIN tblItem ON tblAdjustment.ItemID = tblItem.ItemID) LEFT JOIN tblBatch ON tblAdjustment.BatchID = tblBatch.BatchID " & _
                    "WHERE (((tblAdjustment.Date) Between #" & Format(dtpFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpto.Value, "dd MMMM yyyy") & "#) AND ((tblItem.ItemCategoryID)=" & Val(cmbICat.BoundText) & ")) " & _
                    "ORDER BY tblItem.Display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                i = i + 1
                If Not IsNull(!Display) Then myWorkSheet1.Cells(i, 1) = !Display
                If Not IsNull(!Batch) Then myWorkSheet1.Cells(i, 2) = !Batch
                If Not IsNull(!Amount) Then myWorkSheet1.Cells(i, 3) = !Amount
                .MoveNext
            Wend
        End If
        .Close
    End With
    
    ' Column Headings
    TemRangeAddress = ("A3:C3")
    myWorkSheet1.Range(TemRangeAddress).Font.Bold = True
    myWorkSheet1.Range(TemRangeAddress).Font.Size = 13
    myWorkSheet1.Range(TemRangeAddress).BorderAround 13, xlMedium, xlColorIndexAutomatic
    myWorkSheet1.Range(TemRangeAddress).Orientation = xlTickLabelOrientationUpward
    myWorkSheet1.Range(TemRangeAddress).HorizontalAlignment = xlHAlignCenter
    
    ' Row Headings
    TemRangeAddress = ("A4:A" & (i + 1))
    myWorkSheet1.Range(TemRangeAddress).Font.Bold = True
    myWorkSheet1.Range(TemRangeAddress).Font.Size = 13
    myWorkSheet1.Range(TemRangeAddress).BorderAround 13, xlMedium, xlColorIndexAutomatic
    myWorkSheet1.Columns(1).AutoFit
    
    
    'Column Totals
    TemRangeAddress = ("A" & i + 1 & ":" & "E" & i + 1)
    myWorkSheet1.Range(TemRangeAddress).Font.Bold = True
    myWorkSheet1.Range(TemRangeAddress).Font.Size = 13
    myWorkSheet1.Range(TemRangeAddress).BorderAround 11, xlMedium, xlColorIndexAutomatic
    
    ' Numeric Values
    TemRangeAddress = ("E2:" & "E" & i + 2)
    myWorkSheet1.Range(TemRangeAddress).NumberFormat = "#,##0.00"
    
    Dim TemString As String
    TemString = ""
    While FSys.FileExists(App.Path & "\" & temTopic & " " & temSubTopic & TemString & ".xls") = True
        TemString = TemString & "1"
    Wend
    
    DoEvents
    
    myWorkSheet1.Activate
    myworkbook.SaveAs (App.Path & "\" & temTopic & " " & temSubTopic & TemString & ".xls")
    ExcelFilePath = App.Path & "\" & temTopic & " " & temSubTopic & TemString & ".xls"
    
    AppExcel.ActiveWorkbook.Close
    AppExcel.Quit

    Set myWorkSheet1 = Nothing
    Set myWorkSheet2 = Nothing
    Set myWorkSheet3 = Nothing
    Set myworkbook = Nothing
    Set AppExcel = Nothing
    
    
    Screen.MousePointer = vbDefault
    Unload frmPleaseWait
    frmChart.Show
    

End Sub

Private Sub btnFill_Click()
    Call FormatGrid
    Call FillGrid
    
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    dtpFrom.Value = Date
    dtpto.Value = Date
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.Display, tblBatch.Batch, tblAdjustment.Amount " & _
                    "FROM (tblAdjustment LEFT JOIN tblItem ON tblAdjustment.ItemID = tblItem.ItemID) LEFT JOIN tblBatch ON tblAdjustment.BatchID = tblBatch.BatchID " & _
                    "WHERE (((tblAdjustment.Date) Between #" & Format(dtpFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpto.Value, "dd MMMM yyyy") & "#) AND ((tblItem.ItemCategoryID)=" & Val(cmbICat.BoundText) & ")) " & _
                    "ORDER BY tblItem.Display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                gridA.Rows = gridA.Rows + 1
                gridA.Row = gridA.Rows - 1
                gridA.Col = 0
                gridA.Text = gridA.Row
                
                gridA.Col = 1
                gridA.Text = !Display
                
                gridA.Col = 2
                gridA.Text = !Batch
                
                gridA.Col = 3
                gridA.Text = !Amount
                
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub FormatGrid()
    With gridA
        .Clear
        
        .Cols = 5
        .Rows = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "No."
        
        .Col = 1
        .Text = "Item"
        
        .Col = 2
        .Text = "Batch"
        
        .Col = 3
        .Text = "Adjustment"
        
        
        .ColWidth(0) = 600
        .ColWidth(1) = 3500
        .ColWidth(2) = 2000
        .ColWidth(3) = 1600
        
    End With
End Sub


Private Sub FillCombos()
    With rsICat
        If .State = 1 Then .Close
        temSql = "Select * from tblItemCategory order by ItemCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbICat
        Set .RowSource = rsICat
        .ListField = "ItemCategory"
        .BoundColumn = "ItemCategoryID"
    End With
End Sub
