Attribute VB_Name = "modGridFill"
Option Explicit
    Dim temSql As String
    
Public Sub FillAnyGrid(InputSql As String, InputGrid As MSFlexGrid)
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    
    With rsTem
        If .State = 1 Then .Close
        temSql = InputSql
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
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


