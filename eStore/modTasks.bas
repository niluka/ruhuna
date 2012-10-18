Attribute VB_Name = "modTasks"
Option Explicit


Public Function SortItemCollection(col As Collection, strPropertyName, Optional blnCompareNumeric As Boolean = False) As Collection
    Dim colNew As Collection
    Dim objCurrent As Object
    Dim objCompare As Object
    Dim lngCompareIndex As Long
    Dim strCurrent As String
    Dim strCompare As String
    Dim blnGreaterValueFound As Boolean

    'make a copy of the collection, ripping through it one item
    'at a time, adding to new collection in right order...
    
    Set colNew = New Collection
    
    For Each objCurrent In col
    
        'get value of current item...
        strCurrent = CallByName(objCurrent, strPropertyName, VbGet)
        
        'setup for compare loop
        blnGreaterValueFound = False
        lngCompareIndex = 0
        
        For Each objCompare In colNew
            lngCompareIndex = lngCompareIndex + 1
            
            strCompare = CallByName(objCompare, strPropertyName, VbGet)
            
            'optimization - instead of doing this for every iteration, have 2 different loops...
            If blnCompareNumeric = True Then
                'this means we are looking for a numeric sort order...
                
                If Val(strCurrent) < Val(strCompare) Then
                    'found an item in compare collection that is greater...
                    'add it to the new collection...
                    blnGreaterValueFound = True
                    colNew.Add objCurrent, , lngCompareIndex
                    Exit For
                End If
                
            Else
                'this means we are looking for a string sort...
                
                If strCurrent < strCompare Then
                    'found an item in compare collection that is greater...
                    'add it to the new collection...
                    blnGreaterValueFound = True
                    colNew.Add objCurrent, , lngCompareIndex
                    Exit For
                End If
            
            End If
        Next
        
        'if we didn't find something bigger, just add it to the end of the new collection...
        If blnGreaterValueFound = False Then
            colNew.Add objCurrent
        End If
              
    Next

    'return the new collection...
    Set SortItemCollection = colNew
    Set colNew = Nothing

End Function


Public Sub SaveCommonSettings(MyForm As Form)
    Dim MyCtrl As Control
    Dim I As Integer
    For Each MyCtrl In MyForm.Controls
        If TypeOf MyCtrl Is MSFlexGrid Then
            For I = 0 To MyCtrl.Cols - 1
                SaveSetting App.EXEName, MyForm.Name & MyCtrl.Name, I, MyCtrl.ColWidth(I)
            Next
        End If
    Next
    SaveSetting App.EXEName, MyForm.Name, "Top", MyForm.Top
    SaveSetting App.EXEName, MyForm.Name, "Left", MyForm.Left
End Sub

Public Sub GetCommonSettings(MyForm As Form)
    Dim MyCtrl As Control
    Dim I As Integer
    For Each MyCtrl In MyForm.Controls
        If TypeOf MyCtrl Is MSFlexGrid Then
            For I = 0 To MyCtrl.Cols - 1
                MyCtrl.ColWidth(I) = GetSetting(App.EXEName, MyForm.Name & MyCtrl.Name, I, MyCtrl.ColWidth(I))
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
    
    myworkbook.SaveAs (App.Path & "\" & Topic & " - " & Subtopic & ".xls")
    myworkbook.Save
    myworkbook.Close
    
    ShellExecute 0&, "open", App.Path & "\" & Topic & " - " & Subtopic & ".xls", "", "", vbMaximizedFocus
End Sub
