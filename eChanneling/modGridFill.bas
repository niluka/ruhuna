Attribute VB_Name = "modGridFill"
Option Explicit
    Dim temSQL As String
    
Public Sub FillAnyGrid(InputSql As String, InputGrid As MSFlexGrid)
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = InputSql
        .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
        
        InputGrid.Clear
        
        InputGrid.Rows = 1
        InputGrid.Cols = .Fields.Count
        
        InputGrid.Row = 0
                    
        For i = 0 To .Fields.Count - 1
            InputGrid.col = i
            InputGrid.Text = .Fields(i).Name
        Next i
        
        While .EOF = False
            InputGrid.Rows = InputGrid.Rows + 1
            InputGrid.Row = InputGrid.Rows - 1
            For i = 0 To .Fields.Count - 1
                InputGrid.col = i
                If IsNull(.Fields(i).Value) = False Then
                    InputGrid.Text = .Fields(i).Value
                End If
            Next i
            .MoveNext
        Wend
        .Close
    End With
End Sub





Public Function FillTotalGrid(InputSql As String, InputGrid As MSFlexGrid, TotalNameCol As Integer, TotalCols() As Integer, OmitRepeatCols() As Integer) As Double()
    Dim rsTem As New ADODB.Recordset
    Dim colTotal() As Double
    Dim previousValue() As String
    Dim i As Integer
    Dim col As Integer
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = InputSql
        .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
        
        InputGrid.Clear
        
        InputGrid.Rows = 1
        InputGrid.Cols = .Fields.Count
        
        ReDim colTotal(.Fields.Count)
        ReDim previousValue(.Fields.Count)
        
        InputGrid.Row = 0
                    
        For i = 0 To .Fields.Count - 1
            InputGrid.col = i
            InputGrid.Text = .Fields(i).Name
        Next i
        
        While .EOF = False
            InputGrid.Rows = InputGrid.Rows + 1
            InputGrid.Row = InputGrid.Rows - 1
            For i = 0 To .Fields.Count - 1
                InputGrid.col = i
                If UBound(OmitRepeatCols) > 0 Then
                    For col = 0 To UBound(OmitRepeatCols) - 1
                        If OmitRepeatCols(col) = i Then
                            If previousValue(i) <> .Fields(i).Value Then
                                previousValue(i) = .Fields(i).Value
                                If Not IsNull(.Fields(i).Value) Then
                                    InputGrid.Text = .Fields(i).Value
                                End If
                            End If
                        Else
                            If Not IsNull(.Fields(i).Value) Then
                                InputGrid.Text = .Fields(i).Value
                            End If
                        End If
                    Next
                Else
                    If Not IsNull(.Fields(i).Value) Then
                        InputGrid.Text = .Fields(i).Value
                    End If
                End If
                For col = 0 To UBound(TotalCols) - 1
                    If TotalCols(col) = i Then
                        If IsNull((.Fields(i).Value)) = False Then
                            colTotal(i) = colTotal(i) + Val(.Fields(i).Value)
                        End If
                    End If
                Next
            Next i
            .MoveNext
        Wend
        .Close
    End With
    
    If UBound(TotalCols) > 0 Then
        InputGrid.Rows = InputGrid.Rows + 1
        InputGrid.Row = InputGrid.Rows - 1
        InputGrid.col = TotalNameCol
        InputGrid.Text = "Total"
        For i = 0 To InputGrid.Cols - 1
            InputGrid.col = i
            For col = 0 To UBound(TotalCols) - 1
                If TotalCols(col) = i Then
                    InputGrid.Text = colTotal(i)
                End If
            Next
        Next i
    End If
    FillTotalGrid = colTotal
End Function

Public Sub ReplaceGridCOlText(grid As MSFlexGrid, col As Integer, FromStr As String, ToStr As String)
    Dim i As Integer
    With grid
        For i = 0 To .Rows - 1
            If .TextMatrix(i, col) = FromStr Then
                .TextMatrix(i, col) = ToStr
            End If
        Next
    End With
End Sub
