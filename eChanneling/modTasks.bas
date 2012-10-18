Attribute VB_Name = "modTasks"
Option Explicit
    Dim temSQL As String


Public Sub NullToSero()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblPatientFacility where Personalrefund is Null "
        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
        While .EOF = False
            !Personalrefund = 0
            .Update
            .MoveNext
        Wend
        .Close
    End With


    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblPatientFacility where  institutionrefund is Null "
        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
        While .EOF = False
            !institutionrefund = 0
            .Update
            .MoveNext
        Wend
        .Close
    End With





End Sub



Public Sub SaveCommonSettings(MyForm As Form)
    Dim MyCtrl As Control
    Dim i As Integer
    For Each MyCtrl In MyForm.Controls
        If TypeOf MyCtrl Is MSFlexGrid Then
            For i = 0 To MyCtrl.Cols - 1
                SaveSetting App.EXEName, MyForm.Name & MyCtrl.Name, i, MyCtrl.ColWidth(i)
            Next
        End If
    Next
    SaveSetting App.EXEName, MyForm.Name, "Top", MyForm.Top
    SaveSetting App.EXEName, MyForm.Name, "Left", MyForm.Left
End Sub

Public Sub GetCommonSettings(MyForm As Form)
    Dim MyCtrl As Control
    Dim i As Integer
    For Each MyCtrl In MyForm.Controls
        If TypeOf MyCtrl Is MSFlexGrid Then
            For i = 0 To MyCtrl.Cols - 1
                MyCtrl.ColWidth(i) = GetSetting(App.EXEName, MyForm.Name & MyCtrl.Name, i, MyCtrl.ColWidth(i))
                MyCtrl.AllowUserResizing = flexResizeColumns
            Next
        End If
    Next
    MyForm.Top = GetSetting(App.EXEName, MyForm.Name, "Top", MyForm.Top)
    MyForm.Left = GetSetting(App.EXEName, MyForm.Name, "Left", MyForm.Left)
End Sub

Public Sub GridToExcel(ExportGrid As MSFlexGrid, Optional Topic As String, Optional Subtopic As String)
    If ExportGrid.Rows <= 1 Then
        MsgBox "Noting to Export"
        Exit Sub
    End If
    
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myWorkSheet1 As Excel.Worksheet
    Dim temRow As Integer
    Dim temCol As Integer
    
    Set AppExcel = CreateObject("Excel.Application")
    Set myworkbook = AppExcel.Workbooks.Add
    Set myWorkSheet1 = AppExcel.Worksheets(1)
    
    myWorkSheet1.Cells(1, 1) = Topic
    myWorkSheet1.Cells(2, 1) = Subtopic
    
    For temRow = 0 To ExportGrid.Rows - 1
        For temCol = 0 To ExportGrid.Cols - 1
            myWorkSheet1.Cells(temRow + 3, temCol + 1) = ExportGrid.TextMatrix(temRow, temCol)
        Next
    Next temRow
    
    myworkbook.SaveAs (App.Path & "\" & Topic & ".xls")
    myworkbook.Save
    myworkbook.Close
    
    ShellExecute 0&, "open", App.Path & "\" & Topic & ".xls", "", "", vbMaximizedFocus
    
End Sub
