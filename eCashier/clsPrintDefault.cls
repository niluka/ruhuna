VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Private MyPrintDefaultsV As lmpItem
    
Public Property Get MyPrintDefaults() As lmpItem
    Call GetMyPrintDefaults
    MyPrintDefaults = MyPrintDefaultsV
End Property

Private Sub GetMyPrintDefaults()
    Dim MyPrinter As Printer
    For Each MyPrinter In VB.Printers
        If Printer.DeviceName = MyPrinter.DeviceName Then
            MyPrintDefaults.Printer = MyPrinter.DeviceName
        End If
    Next

'    With MyPrintDefaults
'        .TopMargin = 1440
'        .LeftMargin = 1440
'        .RightMargin = 1440
'        .BottomMargin = 1440
'        .ColFontName = "Verdana"
'        .ColFontSize = 10
'        .ColFontBold = False
'        .ColFontItalic = False
'        .ColFontUnderline = False
'        .ColFontStrikeThrough = False
'        .ColAlignment = Left
'        .TopFontName = "Verdana"
'        .TopFontSize = 11
'        .TopFontBold = True
'        .TopFontItalic = False
'        .TopFontUnderline = False
'        .TopFontStrikeThrough = False
'        .TopAlignment = Left
'        .Paper = "A4"
'        .TopicFontName = "Verdana"
'        .TopicFontSize = 13
'        .TopicFontBold = True
'        .TopicFontItalic = False
'        .TopicFontUnderline = False
'        .TopicFontStrikeThrough = False
'        .TopAlignment = Centre
'        .SubTopicFontName = "Verdana"
'        .SubTopicFontSize = 12
'        .SubTopicFontBold = True
'        .SubTopicFontItalic = False
'        .SubTopicFontUnderline = False
'        .SubTopicFontStrikeThrough = False
'        .SubTopicAlignment = Centre
'        .Topic = HospitalName
'        .Subtopic = HospitalDescreption
'        .Header = ""
'        .Footer = ""
'        .HeaderFontName = "Verdana"
'        .HeaderFontSize = 9
'        .HeaderFontBold = False
'        .HeaderFontItalic = False
'        .HeaderFontUnderline = False
'        .HeaderX = 1440
'        .HeaderY = 1440 / 2
'        .HeaderAlignment = Left
'        .FooterFontName = "Verdana"
'        .FooterFontSize = 9
'        .FooterFontBold = False
'        .FooterFontItalic = False
'        .FooterFontUnderline = False
'        .FooterX = 1440
'        .FooterY = Printer.Height - (1440 / 2)
'        .FooterAlignment = Right
'        .PageNoX = Printer.Width - 1440
'        .PageNoY = Printer.Height - (1440 / 3 * 1)
'        .PageNoFontName = "Verdana"
'        .PageNoFontSize = 9
'        .PageNoFontBold = False
'        .PageNoFontItalic = False
'        .PageNoFontUnderline = False
'        .PageNoAlignment = Centre
'    End With
    GetPrintDefaults = MyPrintDefaults
End Sub
