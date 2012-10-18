Attribute VB_Name = "ModuleVariables"
Option Explicit

'API
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Const LB_SETTABSTOPS = &H192
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Type ReportFont
    Name As String
    Size As Long
    Bold As Boolean
    Italic As Boolean
    Strikethrough As Boolean
    Underline As Boolean
End Type

Public DefaultFont As ReportFont


' database Variables
Public Database As String
Public DatabasePath As String
Public cnnStores As New ADODB.Connection


' User Variables
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

