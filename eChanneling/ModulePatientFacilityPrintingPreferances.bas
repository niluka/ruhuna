Attribute VB_Name = "ModulePatientFacilityPrintingPreferances"
Option Explicit
Dim RemainingNotes As String
Dim CurrentNotes As String
Dim PresentCurrentNotes As String
Dim RemainingCurrentNotes As String

Dim RemainingNotesLength As Long
Dim CurrentNotesLength As Long
Dim PresentCurrentNotesLength As Long
Dim RemainingCurrentNotesLength As Long

Dim TextTopY As Long
Dim TextBotY As Long
Dim TextLX As Long
Dim TextRX As Long

Dim PaperTopY As Long
Dim PaperBotY As Long
Dim PaperLeftX As Long
Dim PaperRightX As Long

Dim TextXMargin As Single
Dim TextYMargin As Single
Dim InbetweenX As Single
'Dim InbetweenY As Single




' ************************************** Preferances


' ***************** Printing
Public PreferanceLeftMargin As Long
Public PreferanceRightMargin As Long
Public PreferanceUpperMargin As Long
Public PreferanceLowerMargin As Long
Public PreferanceInbetweenSpace As Long
Public PreferanceHeaderFontName As String
Public PreferanceHeaderFontSize As Long
Public PreferanceHeaderFontBold As Boolean
Public PreferanceHeaderFontItalic As Boolean
Public PreferanceBodyFontName As String
Public PreferanceBodyFontSize As Long
Public PreferanceBodyFontBold As Boolean
Public PreferanceBodyFontItalic As Boolean

Public PreferanceLX As Double
Public PreferanceRX As Double
Public PreferanceTY As Double
Public PreferanceBY As Double
Public PreferancechkInstitutionName As Boolean
Public PreferancetxtInstitutionName As String
Public PreferanceInstitutionNameFontName As String
Public PreferanceInstitutionNameFontSize As Double
Public PreferanceInstitutionNameFontBold As Boolean
Public PreferanceInstitutionNameFontItalic As Boolean
Public PreferancechkInstitutionAddress As Boolean
Public PreferancetxtInstitutionAddress As String
Public PreferanceInstitutionAddressFontName As String
Public PreferanceInstitutionAddressFontSize As Double
Public PreferanceInstitutionAddressFontBold As Boolean
Public PreferanceInstitutionAddressFontItalic As Boolean
Public PreferancechkInstitutionContact As Boolean
Public PreferancetxtInstitutionContact As String
Public PreferanceInstitutionContactFontName As String
Public PreferanceInstitutionContactFontSize As Double
Public PreferanceInstitutionContactFontBold  As Boolean
Public PreferanceInstitutionContactFontItalic As Boolean
Public PreferanceChkMessage As Boolean
Public PreferancetxtMessage As String
Public PreferanceMessageFontName As String
Public PreferanceMessageFontSize As Double
Public PreferanceMessageFontBold As Boolean
Public PreferanceMessageFontItalic As Boolean

Public PreferanceInstitutionNameLX As Double
Public PreferanceInstitutionNameRX As Double
Public PreferanceInstitutionNameTY As Double
Public PreferanceInstitutionNameBY As Double
Public PreferanceInstitutionAddressLX As Double
Public PreferanceInstitutionAddressRX As Double
Public PreferanceInstitutionAddressTY As Double
Public PreferanceInstitutionAddressBY As Double
Public PreferanceInstitutionContactLX As Double
Public PreferanceInstitutionContactRX As Double
Public PreferanceInstitutionContactTY As Double
Public PreferanceInstitutionContactBY As Double
Public PreferanceMessageLX As Double
Public PreferanceMessageRX As Double
Public PreferanceMessageTY As Double
Public PreferanceMessageBY As Double


Public PreferancechkPatientName As Boolean
Public PreferancechkPatientAge As Boolean
Public PreferancechkPatientID As Boolean
Public PreferancechkPatientSex As Boolean
Public PreferancechkLblPatientName As Boolean
Public PreferancechkLblPatientAge As Boolean
Public PreferancechkLblPatientID As Boolean
Public PreferancechkLblPatientSex As Boolean
Public PreferancetxtLblPatientName As String
Public PreferancetxtLblPatientAge As String
Public PreferancetxtLblPatientID As String
Public PreferancetxtLblPatientSex As String

Public PreferanceLblPatientNameLX As Double
Public PreferanceLblPatientAgeLX As Double
Public PreferanceLblPatientIDLX As Double
Public PreferanceLblPatientSexLX As Double
Public PreferancePatientNameLX As Double
Public PreferancePatientAgeLX As Double
Public PreferancePatientIDLX As Double
Public PreferancePatientSexLX As Double

Public PreferanceLblPatientNameTY As Double
Public PreferanceLblPatientAgeTY As Double
Public PreferanceLblPatientIDTY As Double
Public PreferanceLblPatientSexTY As Double
Public PreferancePatientNameTY As Double
Public PreferancePatientAgeTY As Double
Public PreferancePatientIDTY As Double
Public PreferancePatientSexTY As Double

Public PreferanceLabelFontName As String
Public PreferanceLabelFontSize As Long
Public PreferanceLabelFontBold As Boolean
Public PreferanceLabelFontItalic As Boolean
Public PreferanceTextFontName As String
Public PreferanceTextFontSize As Long
Public PreferanceTextFontBold As Boolean
Public PreferanceTextFontItalic As Boolean
Public PreferanceChkLblResults As Boolean
Public PreferanceChkLblComments As Boolean
Public PreferanceTxtLblResults As String
Public PreferanceTxtLblComments As String
Public PreferanceChkResults As Boolean
Public PreferanceChkComments As Boolean

Public LblResultsLX As Double
Public LblResultsTY As Double
Public LblCommentsLX As Double
Public LblCommentsTY As Double

Public ResultsLX As Double
Public ResultsRX As Double
Public ResultsTY As Double
Public ResultsBY As Double
Public CommentsLX As Double
Public CommentsRX As Double
Public CommentsTY As Double
Public CommentsBY As Double

Public TopicFontName As String
Public TopicFontSize As Long
Public TopicFontBold As Boolean
Public TopicFontItalic As Boolean
Public ValueFontName As String
Public ValueFontSize As Long
Public ValueFontBold As Boolean
Public ValueFontItalic As Boolean

Public InbetweenY As Double
Public HLineY1 As Double
Public HLineY2 As Double
Public HLineY3 As Double
Public HLineY4 As Double
Public chkHLine3 As Boolean
Public chkHLine2 As Boolean
Public chkHLine4 As Boolean


Public Sub PrintingResults()
InbetweenY = Printer.ScaleWidth / 100
PaperLeftX = Printer.ScaleWidth * PreferanceLX
PaperRightX = Printer.ScaleWidth * PreferanceRX
PaperTopY = Printer.ScaleHeight * PreferanceTY
PaperBotY = Printer.ScaleHeight * PreferanceBY
'Printer.Line (PaperLeftX, PaperTopY)-(PaperRightX, PaperBotY), , B
If PreferancechkInstitutionName = True Then PrintName
If PreferancechkInstitutionAddress = True Then PrintAddress
If PreferanceChkMessage = True Then PrintMessage
If PreferancechkLblPatientName = True Then PrintLblPatientName
If PreferancechkLblPatientAge = True Then PrintLblPatientAge
If PreferancechkLblPatientSex = True Then PrintLblPatientSex
If PreferancechkLblPatientID = True Then PrintLblPatientID

End Sub

Private Sub PrintName()
TextRX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferanceInstitutionNameRX)
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferanceInstitutionNameLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferanceInstitutionNameTY)
TextBotY = PaperTopY + ((PaperBotY - PaperTopY) * PreferanceInstitutionNameBY)
'Printer.Line (TextLX, TextTopY)-(TextRX, TextBotY), , B
Printer.FontName = PreferanceInstitutionNameFontName
Printer.FontSize = PreferanceInstitutionNameFontSize
Printer.FontBold = PreferanceInstitutionNameFontBold
Printer.FontItalic = PreferanceInstitutionNameFontItalic
Printer.CurrentY = TextTopY ' + Printer.TextHeight(PreferancetxtInstitutionName)
AddNotes (PreferancetxtInstitutionName)
End Sub

Private Sub PrintAddress()
TextRX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferanceInstitutionAddressRX)
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferanceInstitutionAddressLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferanceInstitutionAddressTY)
TextBotY = PaperTopY + ((PaperBotY - PaperTopY) * PreferanceInstitutionAddressBY)
'Printer.Line (TextLX, TextTopY)-(TextRX, TextBotY), , B
Printer.FontName = PreferanceInstitutionAddressFontName
Printer.FontSize = PreferanceInstitutionAddressFontSize
Printer.FontBold = PreferanceInstitutionAddressFontBold
Printer.FontItalic = PreferanceInstitutionAddressFontItalic
Printer.CurrentY = TextTopY '+ Printer.TextHeight(PreferancetxtInstitutionName)
AddNotes (PreferancetxtInstitutionAddress)
End Sub



Private Sub PrintMessage()
TextRX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferanceMessageRX)
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferanceMessageLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferanceMessageTY)
TextBotY = PaperTopY + ((PaperBotY - PaperTopY) * PreferanceMessageBY)
'Printer.Line (TextLX, TextTopY)-(TextRX, TextBotY), , B
Printer.FontName = PreferanceMessageFontName
Printer.FontSize = PreferanceMessageFontSize
Printer.FontBold = PreferanceMessageFontBold
Printer.FontItalic = PreferanceMessageFontItalic
Printer.CurrentY = TextTopY '+ Printer.TextHeight(PreferancetxtMessage)
AddNotes (PreferancetxtMessage)
End Sub


Private Sub PrintLblPatientName()
TextRX = PaperRightX
TextBotY = PaperBotY
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferanceLblPatientNameLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferanceLblPatientNameTY)
Printer.FontName = PreferanceLabelFontName
Printer.FontSize = PreferanceLabelFontSize
Printer.FontBold = PreferanceLabelFontBold
Printer.FontItalic = PreferanceLabelFontItalic
Printer.CurrentY = TextTopY
AddNotes (PreferancetxtLblPatientName)
End Sub
Private Sub PrintLblPatientAge()
TextRX = PaperRightX
TextBotY = PaperBotY
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferanceLblPatientAgeLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferanceLblPatientAgeTY)
Printer.FontName = PreferanceLabelFontName
Printer.FontSize = PreferanceLabelFontSize
Printer.FontBold = PreferanceLabelFontBold
Printer.FontItalic = PreferanceLabelFontItalic
Printer.CurrentY = TextTopY
AddNotes (PreferancetxtLblPatientAge)
End Sub
Private Sub PrintLblPatientSex()
TextRX = PaperRightX
TextBotY = PaperBotY
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferanceLblPatientSexLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferanceLblPatientSexTY)
Printer.FontName = PreferanceLabelFontName
Printer.FontSize = PreferanceLabelFontSize
Printer.FontBold = PreferanceLabelFontBold
Printer.FontItalic = PreferanceLabelFontItalic
Printer.CurrentY = TextTopY
AddNotes (PreferancetxtLblPatientSex)
End Sub
Private Sub PrintLblPatientID()
TextRX = PaperRightX
TextBotY = PaperBotY
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferanceLblPatientIDLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferanceLblPatientIDTY)
Printer.FontName = PreferanceLabelFontName
Printer.FontSize = PreferanceLabelFontSize
Printer.FontBold = PreferanceLabelFontBold
Printer.FontItalic = PreferanceLabelFontItalic
Printer.CurrentY = TextTopY
AddNotes (PreferancetxtLblPatientID)
End Sub


Public Sub PrintPatientName(ByVal PrintingText As String)
TextRX = PaperRightX
TextBotY = PaperBotY
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferancePatientNameLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferancePatientNameTY)
Printer.FontName = PreferanceLabelFontName
Printer.FontSize = PreferanceLabelFontSize
Printer.FontBold = PreferanceLabelFontBold
Printer.FontItalic = PreferanceLabelFontItalic
Printer.CurrentY = TextTopY
AddNotes (PrintingText)
End Sub
Public Sub PrintPatientAge(ByVal PrintingText As String)
TextRX = PaperRightX
TextBotY = PaperBotY
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferancePatientAgeLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferancePatientAgeTY)
Printer.FontName = PreferanceLabelFontName
Printer.FontSize = PreferanceLabelFontSize
Printer.FontBold = PreferanceLabelFontBold
Printer.FontItalic = PreferanceLabelFontItalic
Printer.CurrentY = TextTopY
AddNotes (PrintingText)
End Sub
Public Sub PrintPatientSex(ByVal PrintingText As String)
TextRX = PaperRightX
TextBotY = PaperBotY
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferancePatientSexLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferancePatientSexTY)
Printer.FontName = PreferanceLabelFontName
Printer.FontSize = PreferanceLabelFontSize
Printer.FontBold = PreferanceLabelFontBold
Printer.FontItalic = PreferanceLabelFontItalic
Printer.CurrentY = TextTopY
AddNotes (PrintingText)
End Sub
Public Sub PrintPatientID(ByVal PrintingText As String)
TextRX = PaperRightX
TextBotY = PaperBotY
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * PreferancePatientIDLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * PreferancePatientIDTY)
Printer.FontName = PreferanceLabelFontName
Printer.FontSize = PreferanceLabelFontSize
Printer.FontBold = PreferanceLabelFontBold
Printer.FontItalic = PreferanceLabelFontItalic
Printer.CurrentY = TextTopY
AddNotes (PrintingText)
End Sub
Public Sub PrintLblResults(ByVal PrintingText As String)
TextRX = PaperRightX
TextBotY = PaperBotY
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * LblResultsLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * LblResultsTY)
Printer.FontName = TopicFontName
Printer.FontSize = TopicFontSize
Printer.FontBold = TopicFontBold
Printer.FontItalic = TopicFontItalic
Printer.CurrentY = TextTopY
AddNotes (PrintingText)
End Sub

Public Sub PrintLblComments(ByVal PrintingText As String)
TextRX = PaperRightX
TextBotY = PaperBotY
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * LblCommentsLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * LblCommentsTY)
Printer.FontName = TopicFontName
Printer.FontSize = TopicFontSize
Printer.FontBold = TopicFontBold
Printer.FontItalic = TopicFontItalic
Printer.CurrentY = TextTopY
AddNotes (PrintingText)
End Sub

Public Sub PrintResultstList(ByVal PrintingText As String)
TextRX = PaperLeftX + ((PaperRightX - PaperLeftX) * ResultsRX)
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * ResultsLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * ResultsTY)
TextBotY = PaperTopY + ((PaperBotY - PaperTopY) * ResultsBY)
'Printer.Line (TextLX, TextTopY)-(TextRX, TextBotY), , B
Printer.FontName = ValueFontName
Printer.FontSize = ValueFontSize
Printer.FontBold = ValueFontBold
Printer.FontItalic = ValueFontItalic
Printer.CurrentY = TextTopY '+ Printer.TextHeight(PreferancetxtMessage)
AddNotes (PrintingText)
End Sub


Public Sub PrintComments(ByVal PrintingText As String)
TextRX = PaperLeftX + ((PaperRightX - PaperLeftX) * CommentsRX)
TextLX = PaperLeftX + ((PaperRightX - PaperLeftX) * CommentsLX)
TextTopY = PaperTopY + ((PaperBotY - PaperTopY) * CommentsTY)
TextBotY = PaperTopY + ((PaperBotY - PaperTopY) * CommentsBY)
'Printer.Line (TextLX, TextTopY)-(TextRX, TextBotY), , B
Printer.FontName = ValueFontName
Printer.FontSize = ValueFontSize
Printer.FontBold = ValueFontBold
Printer.FontItalic = ValueFontItalic
Printer.CurrentY = TextTopY '+ Printer.TextHeight(PreferancetxtMessage)
AddNotes (PrintingText)
End Sub



Public Sub PrintLines()
Printer.Line (PaperLeftX, HLineY1 * (PaperBotY - PaperTopY))-(PaperRightX, HLineY1 * (PaperBotY - PaperTopY))
If chkHLine2 = True Then Printer.Line (PaperLeftX, HLineY2 * (PaperBotY - PaperTopY))-(PaperRightX, HLineY2 * (PaperBotY - PaperTopY))
If chkHLine3 = True Then Printer.Line (PaperLeftX, HLineY3 * (PaperBotY - PaperTopY))-(PaperRightX, HLineY3 * (PaperBotY - PaperTopY))
If chkHLine4 = True Then Printer.Line (PaperLeftX, HLineY4 * (PaperBotY - PaperTopY))-(PaperRightX, HLineY4 * (PaperBotY - PaperTopY))
End Sub


Private Sub AddNotes(SendingText)
With Printer
RemainingNotes = SendingText
While InStr(RemainingNotes, Chr(13) & Chr(10)) > 0
    CurrentNotesLength = InStr(RemainingNotes, Chr(13) & Chr(10))
    CurrentNotes = Left(RemainingNotes, CurrentNotesLength - 1)
'    If Trim(CurrentNotes) <> "" Then PrintCurrentNotes
    
    PrintCurrentNotes
    
    RemainingNotesLength = Len(RemainingNotes) - (CurrentNotesLength + 1)
    RemainingNotes = Right(RemainingNotes, RemainingNotesLength)
Wend
'If Trim(RemainingNotes) <> "" And Trim(RemainingNotes) <> Chr(13) & Chr(10) Then
    CurrentNotes = Right(RemainingNotes, (Len(RemainingNotes) - 0))
    PrintCurrentNotes
'End If
.CurrentY = .CurrentY + InbetweenY
TextBotY = .CurrentY
End With
End Sub

Private Sub PrintCurrentNotes()
With Printer
If .TextWidth(CurrentNotes) > TextRX - (TextLX) Then
    BreakCurrentNotes
Else
    .CurrentX = TextLX  ' ((TextRX - TextLX) / 2) - (.TextWidth(CurrentNotes) / 2)
    
    
'    Printer.Print CurrentNotes
    
        PrintingWithSuperscript (CurrentNotes)

    
End If
End With
End Sub

Private Sub BreakCurrentNotes()
If InStr(CurrentNotes, " ") < 1 Then
    BreakNoSpaces
Else
    BreakSpaces
End If
End Sub

Private Sub BreakSpaces()
With Printer
Dim SearchPosition
RemainingCurrentNotes = CurrentNotes
SearchPosition = Len(RemainingCurrentNotes)
While InStr(RemainingCurrentNotes, " ") > 0
RemainingCurrentNotesLength = Len(RemainingCurrentNotes)
If SearchPosition = 0 Then SearchPosition = 1
PresentCurrentNotesLength = InStrRev(RemainingCurrentNotes, " ", SearchPosition)
PresentCurrentNotes = Left(RemainingCurrentNotes, PresentCurrentNotesLength)
If Printer.TextWidth(PresentCurrentNotes) < TextRX - (TextLX) Then
    .CurrentX = TextLX '((TextRX - TextLX) / 2) - (.TextWidth(PresentCurrentNotes) / 2)
    
    
'    Printer.Print PresentCurrentNotes
    PrintingWithSuperscript (PresentCurrentNotes)
    
    RemainingCurrentNotes = Right(RemainingCurrentNotes, RemainingCurrentNotesLength - PresentCurrentNotesLength)
    SearchPosition = Len(RemainingCurrentNotes)
Else
    SearchPosition = SearchPosition - 1
End If
Wend
If Printer.TextWidth(RemainingCurrentNotes) < TextRX - (TextLX) Then
    .CurrentX = TextLX  ' ((TextRX - TextLX) / 2) - (.TextWidth(RemainingCurrentNotes) / 2)
    
    
    'Printer.Print RemainingCurrentNotes

' ******************************

    PrintingWithSuperscript (RemainingCurrentNotes)

' ************************


Else
    BreakNoSpaces
End If
End With
End Sub

Private Sub BreakNoSpaces()
Dim A
Dim TextLength
Dim TemText

With Printer
TextLength = 0
.CurrentX = TextLX

For A = 1 To Len(CurrentNotes)
TemText = Mid(CurrentNotes, A, 1)
TextLength = TextLength + Printer.TextWidth(TemText)
If TextLength >= TextRX - (TextLX) Then
    Printer.Print TemText
    TextLength = 0
    .CurrentX = TextLX
Else
    Printer.Print TemText;
End If

Next
End With
End Sub


Private Sub PrintingWithSuperscript(ByVal SuppliedForSuperscript As String)


Dim SuperscriptPresent As Boolean
Dim SuperscriptLocation As Long
Dim BeforeSuperscriptText As String
Dim AfterSuperscriptText As String
Dim SuperScriptText As String
'Dim SuppliedForSuperscript As String
Dim SpaceLocation As Long
Dim SendingX As Long
Dim SendingY As Long
Dim SuperscriptElevation As Double

SendingX = Printer.CurrentX
SendingY = Printer.CurrentY

SuppliedForSuperscript = SuppliedForSuperscript

SuperscriptLocation = InStr(1, SuppliedForSuperscript, "^")

SuperscriptElevation = 0.2

If SuperscriptLocation = 0 Then
    Printer.CurrentX = SendingX
    Printer.CurrentY = SendingY
    Printer.Print SuppliedForSuperscript
Else
    If SuperscriptLocation > 0 Then BeforeSuperscriptText = Left(SuppliedForSuperscript, SuperscriptLocation - 1)
    AfterSuperscriptText = Right(SuppliedForSuperscript, (Len(SuppliedForSuperscript) - SuperscriptLocation))
    AfterSuperscriptText = Trim(AfterSuperscriptText)
    If InStr(1, AfterSuperscriptText, " ") = 0 Then
        Printer.CurrentX = SendingX
        Printer.CurrentY = SendingY
        Printer.Print BeforeSuperscriptText
        Printer.CurrentX = SendingX + Printer.TextWidth(BeforeSuperscriptText)
        Printer.CurrentY = SendingY - (Printer.TextHeight(BeforeSuperscriptText) * SuperscriptElevation)
        Printer.FontSize = Printer.FontSize - 2
        Printer.Print AfterSuperscriptText
        Printer.FontSize = Printer.FontSize + 2
    Else
        SpaceLocation = InStr(1, AfterSuperscriptText, " ")
        SuperScriptText = Left(AfterSuperscriptText, SpaceLocation)
        AfterSuperscriptText = Right(AfterSuperscriptText, (Len(AfterSuperscriptText) - SpaceLocation))
        
        Printer.CurrentX = SendingX
        Printer.CurrentY = SendingY
        
        Printer.Print BeforeSuperscriptText
        
        Printer.CurrentX = SendingX + Printer.TextWidth(BeforeSuperscriptText)
        Printer.CurrentY = SendingY - (Printer.TextHeight(BeforeSuperscriptText) * SuperscriptElevation)
        Printer.FontSize = Printer.FontSize - 2
        Printer.Print SuperScriptText
        Printer.FontSize = Printer.FontSize + 2
        Printer.CurrentY = SendingY
        Printer.CurrentX = SendingX + Printer.TextWidth(BeforeSuperscriptText & SuperScriptText)
        Printer.Print AfterSuperscriptText
        
    End If
    
    
End If

Printer.CurrentY = SendingY
Printer.Print


End Sub



