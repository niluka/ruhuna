VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSalesChart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Chart"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "SaveChart"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   8520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SaveLine"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   8520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SavePie"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   8520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9720
      TabIndex        =   3
      Top             =   7320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnPieChart 
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   7320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Pie "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnLineChart 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   7320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Line Chart"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnBarChart 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   7320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Bar Chart"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OLE oleLine 
      BackStyle       =   0  'Transparent
      Class           =   "Excel.Sheet.8"
      Enabled         =   0   'False
      Height          =   6615
      Left            =   360
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\ePharmacyAssistant\ExcelSheet\MontthlySalesLineChart.xls"
      TabIndex        =   9
      Top             =   480
      Width           =   11055
   End
   Begin VB.OLE oleBar 
      BackStyle       =   0  'Transparent
      Class           =   "Excel.Sheet.8"
      Enabled         =   0   'False
      Height          =   6615
      Left            =   360
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\ePharmacyAssistant\ExcelSheet\MontthlySalesBarChart.xls"
      TabIndex        =   8
      Top             =   480
      Width           =   11055
   End
   Begin VB.OLE olePie 
      BackStyle       =   0  'Transparent
      Class           =   "Excel.Sheet.8"
      Enabled         =   0   'False
      Height          =   6615
      Left            =   360
      SizeMode        =   1  'Stretch
      SourceDoc       =   "C:\ePharmacyAssistant\ExcelSheet\MontthlySalesChart.xls"
      TabIndex        =   7
      Top             =   480
      Width           =   11055
   End
End
Attribute VB_Name = "frmSalesChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrCon As String
Dim ExCon As ADODB.Connection
Dim ExRs As ADODB.Recordset

Dim DateFrom As Date
Dim DateTo As Date
Dim TemJanuarySales As Double
Dim TemFebruarySales As Double
Dim TemMarchrSales As Double
Dim TemApirlSales As Double
Dim TemMaySales As Double
Dim TemJuneSales As Double
Dim TemJulySales As Double
Dim TemAuguestSales As Double
Dim TemSeptemberSales As Double
Dim TemOctoberSales As Double
Dim TemNovemberSales As Double
Dim TemDecemberSales As Double


Private Sub ExcelBarchartSheetUpdate()

StrCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & App.Path & "\ExcelSheet\MonthlySalesBarChart.xls;" & _
"Extended Properties=Excel 8.0"
      
Set ExCon = New ADODB.Connection

ExCon.Open StrCon
    
Set ExRs = New ADODB.Recordset
ExRs.CursorLocation = adUseClient
ExRs.Open "Select * from [Sheet1$]", ExCon, adOpenStatic, adLockOptimistic

With ExRs
   .Fields("January").Value = TemJanuarySales
   .Fields("February").Value = TemFebruarySales
   .Fields("March").Value = TemMarchrSales
   .Fields("April").Value = TemApirlSales
   .Fields("May").Value = TemMaySales
   .Fields("June").Value = TemJuneSales
   .Fields("July").Value = TemJulySales
   .Fields("August").Value = TemAuguestSales
   .Fields("September").Value = TemSeptemberSales
   .Fields("October").Value = TemOctoberSales
   .Fields("November").Value = TemNovemberSales
   .Fields("December").Value = TemDecemberSales
   
   .Update
End With

If ExRs.State = 1 Then ExRs.Close: Set ExRs = Nothing
If ExCon.State = 1 Then ExCon.Close: Set ExCon = Nothing

End Sub

Private Sub ExcelPiechartSheetUpdate()

StrCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & App.Path & "\ExcelSheet\MonthlySalesPieChart.xls;" & _
"Extended Properties=Excel 8.0"
      
Set ExCon = New ADODB.Connection

ExCon.Open StrCon
    
Set ExRs = New ADODB.Recordset
ExRs.CursorLocation = adUseClient
ExRs.Open "Select * from [Sheet1$]", ExCon, adOpenStatic, adLockOptimistic

With ExRs
   .Fields("January").Value = TemJanuarySales
   .Fields("February").Value = TemFebruarySales
   .Fields("March").Value = TemMarchrSales
   .Fields("April").Value = TemApirlSales
   .Fields("May").Value = TemMaySales
   .Fields("June").Value = TemJuneSales
   .Fields("July").Value = TemJulySales
   .Fields("August").Value = TemAuguestSales
   .Fields("September").Value = TemSeptemberSales
   .Fields("October").Value = TemOctoberSales
   .Fields("November").Value = TemNovemberSales
   .Fields("December").Value = TemDecemberSales
   
   .Update
End With

If ExRs.State = 1 Then ExRs.Close: Set ExRs = Nothing
If ExCon.State = 1 Then ExCon.Close: Set ExCon = Nothing

End Sub

Private Sub ExcelLinechartSheetUpdate()

StrCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & App.Path & "\ExcelSheet\MonthlySalesLineChart.xls;" & _
"Extended Properties=Excel 8.0"
      
Set ExCon = New ADODB.Connection

ExCon.Open StrCon
    
Set ExRs = New ADODB.Recordset
ExRs.CursorLocation = adUseClient
ExRs.Open "Select * from [Sheet1$]", ExCon, adOpenStatic, adLockOptimistic

With ExRs
   .Fields("January").Value = TemJanuarySales
   .Fields("February").Value = TemFebruarySales
   .Fields("March").Value = TemMarchrSales
   .Fields("April").Value = TemApirlSales
   .Fields("May").Value = TemMaySales
   .Fields("June").Value = TemJuneSales
   .Fields("July").Value = TemJulySales
   .Fields("August").Value = TemAuguestSales
   .Fields("September").Value = TemSeptemberSales
   .Fields("October").Value = TemOctoberSales
   .Fields("November").Value = TemNovemberSales
   .Fields("December").Value = TemDecemberSales
   
   .Update
End With

If ExRs.State = 1 Then ExRs.Close: Set ExRs = Nothing
If ExCon.State = 1 Then ExCon.Close: Set ExCon = Nothing

End Sub

Private Sub bttnBarChart_Click()
oleLine.Visible = False
olePie.Visible = False
oleBar.Visible = True
'Call OpenBarChart
oleBar.CreateLink (App.Path & "\ExcelSheet\MontthlySalesBarChart.xls")
oleBar.Update
'Call SaveBarchart
End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub CalculateJanuarySales()
DateFrom = DateSerial(Year(Date), 1, 1)
DateTo = DateSerial(Year(Date), 1, 31)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemJanuarySales = TemJanuarySales + !Netttotal
    .MoveNext
    Loop
    
    If .State = 1 Then .Close
    
End With
End Sub

Private Sub bttnLineChart_Click()
olePie.Visible = False
oleBar.Visible = False
oleLine.Visible = True
'Call OpenLinechart
oleLine.CreateLink (App.Path & "\ExcelSheet\MontthlySalesLineChart.xls")
oleLine.Update
'Call SaveLinechart
End Sub

Private Sub bttnPieChart_Click()
olePie.Visible = True
oleLine.Visible = False
oleBar.Visible = False
'Call OpenPieChart
olePie.CreateLink (App.Path & "\ExcelSheet\MontthlySalesPieChart.xls")
Call olePie.Update
'Call SavePieechart
End Sub

Private Sub Command1_Click()
SavePieechart
End Sub

Private Sub CalculateFebruarySales()

DateFrom = DateSerial(Year(Date), 2, 1)

Dim TemDate As Date
TemDate = DateSerial(Year(Date), 1, 31)
TemDate = Day(DateAdd("m", 1, TemDate))
DateTo = DateSerial(Year(Date), 2, TemDate)


With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemFebruarySales = TemFebruarySales + !Netttotal
    .MoveNext
    Loop
    
    If .State = 1 Then .Close
    
End With
End Sub

Private Sub CalculateMarchSales()
DateFrom = DateSerial(Year(Date), 3, 1)
DateTo = DateSerial(Year(Date), 3, 31)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemMarchrSales = TemMarchrSales + !Netttotal
    .MoveNext
    Loop
    
    If .State = 1 Then .Close

End With

End Sub

Private Sub CalculateAprilSales()
DateFrom = DateSerial(Year(Date), 4, 1)
DateTo = DateSerial(Year(Date), 4, 30)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemApirlSales = TemApirlSales + !Netttotal
    .MoveNext
    Loop
    
    If .State = 1 Then .Close

End With

End Sub

Private Sub CalculateMaySales()
DateFrom = DateSerial(Year(Date), 5, 1)
DateTo = DateSerial(Year(Date), 5, 31)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemMaySales = TemMaySales + !Netttotal
    .MoveNext
    Loop
    
    If .State = 1 Then .Close

End With

End Sub

Private Sub CalculateJuneSales()
DateFrom = DateSerial(Year(Date), 6, 1)
DateTo = DateSerial(Year(Date), 6, 30)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemJuneSales = TemJuneSales + !Netttotal
    .MoveNext
    Loop
    
    If .State = 1 Then .Close

End With

End Sub

Private Sub CalculateJulySales()
DateFrom = DateSerial(Year(Date), 7, 1)
DateTo = DateSerial(Year(Date), 7, 31)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemJulySales = TemJulySales + !Netttotal
    .MoveNext
    Loop
    
    If .State = 1 Then .Close

End With

End Sub

Private Sub CalculateAugustSales()
DateFrom = DateSerial(Year(Date), 8, 1)
DateTo = DateSerial(Year(Date), 8, 31)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemAuguestSales = TemAuguestSales + !Netttotal
    .MoveNext
    Loop
    
    If .State = 1 Then .Close

End With

End Sub

Private Sub CalculateSeptemberSales()
DateFrom = DateSerial(Year(Date), 9, 1)
DateTo = DateSerial(Year(Date), 9, 30)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemSeptemberSales = TemSeptemberSales + !Netttotal
    .MoveNext
    Loop
    
    If .State = 1 Then .Close

End With

End Sub

Private Sub CalculateOctomerSales()
DateFrom = DateSerial(Year(Date), 10, 1)
DateTo = DateSerial(Year(Date), 10, 31)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    TemOctoberSales = TemOctoberSales + !Netttotal
    .MoveNext
    Loop
    
    If .State = 1 Then .Close

End With

End Sub

Private Sub CalculateNovemberSales()
DateFrom = DateSerial(Year(Date), 11, 1)
DateTo = DateSerial(Year(Date), 11, 30)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
     TemNovemberSales = TemNovemberSales + !Netttotal
    .MoveNext
    Loop
   
    If .State = 1 Then .Close
    
End With
End Sub

Private Sub CalculateDecemberSales()
DateFrom = DateSerial(Year(Date), 12, 1)
DateTo = DateSerial(Year(Date), 12, 31)

With DataEnvironment1.rscmmdBill

    If .State = 1 Then .Close
'    .Source = "Select RetailBill.* From RetailBill Where (BillDate >= #" & DateFrom & "# and BillDate <= #" & DateTo & "#)"
    .Source = "Select RetailBill.* From RetailBill Where (BillDate Between #" & DateFrom & "# and  #" & DateTo & "#)"
    .Open
    
    If .RecordCount = 0 Then Exit Sub
    
    Do While .EOF = False
    
    TemDecemberSales = TemDecemberSales + !Netttotal
    .MoveNext
    Loop

    If .State = 1 Then .Close
    
End With
End Sub

Private Sub SaveBarchart()
Dim FNum As Integer

    On Error GoTo Cancel
    
    FNum = FreeFile
    Open App.Path & "/ExcelSheet/Bar.Lmp" For Binary As #FNum
    oleBar.SaveToFile (FNum)
    Close #FNum
    Exit Sub
    
Cancel:
    MsgBox "COuld not save file"
    Close #FNum

End Sub

Private Sub SaveLinechart()
Dim FNum As Integer

    On Error GoTo Cancel
    
    FNum = FreeFile
    Open App.Path & "/ExcelSheet/Line.Lmp" For Binary As #FNum
    oleLine.SaveToFile (FNum)
    Close #FNum
    Exit Sub
    
Cancel:
    MsgBox "COuld not save file"
    Close #FNum

End Sub

Private Sub SavePieechart()
Dim FNum As Integer

    On Error GoTo Cancel
    
    FNum = FreeFile
    Open App.Path & "/ExcelSheet/Pie.Lmp" For Binary As #FNum
    olePie.SaveToFile (FNum)
    Close #FNum
    Exit Sub
    
Cancel:
    MsgBox "COuld not save file"
    Close #FNum

End Sub

Private Sub OpenLinechart()
Dim FNum As Integer

    On Error GoTo Cancel
    
    FNum = FreeFile
    Open App.Path & "/ExcelSheet/Line.Lmp" For Binary As #FNum
    oleLine.SaveToFile (FNum)
    Close #FNum
    Exit Sub
    
Cancel:
    MsgBox "COuld not save file"
    Close #FNum

End Sub

Private Sub OpenBarChart()
Dim FNum As Integer

    On Error GoTo Cancel
    
    FNum = FreeFile
    Open App.Path & "/ExcelSheet/Bar.Lmp" For Binary As #FNum
    oleBar.ReadFromFile (FNum)
    Close #FNum
    Exit Sub
    
Cancel:
    MsgBox "Could not load file"
    Close #FNum
End Sub

Private Sub OpenPieChart()
Dim FNum As Integer

    On Error GoTo Cancel
    
    FNum = FreeFile
    Open App.Path & "/ExcelSheet/Pie.Lmp" For Binary As #FNum
    olePie.ReadFromFile (FNum)
    Close #FNum
    Exit Sub
    
Cancel:
    MsgBox "Could not load file"
    Close #FNum
End Sub

Private Sub Command2_Click()
SaveLinechart
End Sub

Private Sub Command3_Click()
SaveBarchart
End Sub

Private Sub Form_Load()
'Call CalculateJanuarySales
'Call CalculateFebruarySales
'Call CalculateMarchSales
'Call CalculateAprilSales
'Call CalculateMaySales
'Call CalculateJuneSales
'Call CalculateJulySales
'Call CalculateAugustSales
'Call CalculateSeptemberSales
'Call CalculateNovemberSales
'Call CalculateDecemberSales
'Call ExcelBarchartSheetUpdate
'Call ExcelPiechartSheetUpdate
'Call ExcelLinechartSheetUpdate
'oleLine.Visible = False
'olePie.Visible = False
'oleBar.Visible = True
'oleBar.CreateLink (App.Path & "\ExcelSheet\MontthlySalesBarChart.xls")
'oleBar.Update
'oleBar.SourceDoc = App.Path & "\ExcelSheet\MontthlySalesBarChart.xls"
'olePie.SourceDoc = App.Path & "\ExcelSheet\MontthlySalesLineChart.xls"
'oleLine.SourceDoc = App.Path & "\ExcelSheet\MontthlySalesPieChart.xls"
End Sub

