Attribute VB_Name = "modStringFunctions"
Public Function SeperateLines(InputText As String) As String()
    Dim temSeperateLines() As String
    Dim LineBrakCount As Long
    Dim LineBrakeLocation As Long
    Dim i As Integer
    
    Dim TemString As String
    TemString = InputText
    
    LineBrakCount = 0
    InputText = TemString

    While InStr(InputText, vbNewLine) > 0
        LineBrakeLocation = InStr(InputText, vbNewLine)
        InputText = Right(InputText, Len(InputText) - LineBrakeLocation - 1)
        LineBrakCount = LineBrakCount + 1
    Wend
    

    i = 0
    
    InputText = TemString
    ReDim temSeperateLines(LineBrakCount + 1)
    While InStr(InputText, vbNewLine) > 0
        LineBrakeLocation = InStr(InputText, vbNewLine)
        temSeperateLines(i) = Left(InputText, LineBrakeLocation - 1)
        InputText = Right(InputText, Len(InputText) - LineBrakeLocation - 1)
        i = i + 1
    Wend
    temSeperateLines(i) = InputText
    SeperateLines = temSeperateLines
End Function
