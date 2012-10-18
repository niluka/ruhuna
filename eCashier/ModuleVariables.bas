Attribute VB_Name = "ModuleVariables"
Option Explicit

'API
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
    Public Const LB_SETTABSTOPS = &H192
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'Programing Variables
Public Type ReportFont
    Name As String
    Size As Long
    Bold As Boolean
    Italic As Boolean
    Strikethrough As Boolean
    Underline As Boolean
End Type

Public Type GridCell
        Col As Long
        Row As Long
        Text As String
        CellAlignment As Single
        CellBackColor As OLE_COLOR
        CellFontBold As Boolean
        CellFontItalic As Boolean
        CellFontName As String
        CellFontSize As Single
        CellFontStrikeThrough As Boolean
        CellFontUnderline As BookmarkEnum
        CellFontWidth As Long
        CellForeColor As OLE_COLOR
        CellHeight As Long
        CellLeft As Long
        CellPicture As IPictureDisp
        CellPictureAlignment As Integer
        CellTextStyle As TextStyleSettings
        CellTop As Long
        CellWidth As Long
End Type

Public Type GridRow
        RowCells() As GridCell
        Row As Long
        RowData() As Long
        RowHeight As Long
        RowHeightMin As Long
        RowIsVisible As Boolean
End Type

Public Type Grid
        GridRows() As GridRow
        Appearance As AppearanceSettings
        BackColor As OLE_COLOR
        BackColorBkg As OLE_COLOR
        BackColorFixed As OLE_COLOR
        BackColorSel As OLE_COLOR
        BorderStyle As BorderStyleSettings
        Container As Object
        Enabled As Boolean
        FillStyle As FillStyleSettings
        FixedAlignment() As Integer
        Font As IFontDisp
        FontWidth As Single
        ForeColor As OLE_COLOR
        ForeColorFixed As OLE_COLOR
        ForeColorSel As OLE_COLOR
        FormatString As String
        GridColor As OLE_COLOR
        GridColorFixed As OLE_COLOR
        Gridlines As GridLineSettings
        GridLinesFixed As GridLineSettings
        GridLineWidth As Long
        Height As Long
        HighLight As HighLightSettings
        GridName As String
        GridWidth As Long
        WordWrap As Boolean
End Type

Public DefaultFont As ReportFont





Public Enum ControlType
    TextBox = 1
    ComboBox = 2
    ListBox = 3
    DataCombo = 4
    DataList = 5
    Grid = 6
    Button = 7
    MenuItem = 8
    CheckBox = 9
    OptionButton = 10
    DateTimePicker = 11
    Label = 12
    SSTab = 13
    Unknown = 100
End Enum



' database Variables
Public Database As String
Public DatabasePath As String
Public cnnStores As New ADODB.Connection


' User Variables
Public UserFullName As String
Public UserName As String
Public UserID As Long
Public UserAuthority As Long
Public UserStoreID As Long
Public UserStore As String


' Data Exchange Between forms
Public OrderBillID As Long
Public TxSaleBillID As Long
Public TxRefillBillID As Long

' Store Variables
Public HospitalName As String
Public HospitalDescreption As String
Public HospitalAddress As String
Public LongAd As String
Public ShortAd As String

Public LabName As String
Public LabDescreption As String
Public LabAddress As String

Public RadiologyName As String
Public RadiologyDescreption As String
Public RadiologyAddress As String



' Printing Preferances
Public PrintingOnBlankPaper As Boolean
Public PrintingOnPrintedPaper As Boolean
Public BillPrinterName As String
Public BillPaperName As String
Public PrescreptionPrinterName As String
Public PrescreptionPaperName As String
Public ReportPrinterName As String
Public ReportPaperName As String
Public BillPaperHeight As Long
Public BillPaperWidth As Long
Public ReportPaperWidth As Long
Public ReportPaperHeight As Long

' Program Preferances
Public HighRate As Integer

