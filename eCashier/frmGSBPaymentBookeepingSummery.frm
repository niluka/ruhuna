VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGSBPaymentBookeepingSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSB Payment Bookeeping Summmery"
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10845
   ScaleWidth      =   13230
   Begin MSFlexGridLib.MSFlexGrid gridSummery 
      Height          =   9495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   16748
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   22085635
      CurrentDate     =   39960
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   22085635
      CurrentDate     =   39960
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   10200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "To &Excel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   10200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11880
      TabIndex        =   7
      Top             =   10200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Process"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmGSBPaymentBookeepingSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim CsetPrinter As New cSetDfltPrinter
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    
    
    Dim temBHT As String
    Dim temBHTID As Long
    Dim temDOD As Date
    Dim temDOA As Date
    Dim temPt As String
    Dim temPM As String
    Dim temCC As String
    Dim temFBID As Long
    Dim temTotal As Double
    

Private Sub FillGrid()
    Screen.MousePointer = vbHourglass
    gridSummery.Visible = False
    
    temTotal = 0
    Call FillDirectPayments
    Call FillCompanyPayments
    Call FillRefunds
    
    Screen.MousePointer = vbDefault
    gridSummery.Visible = True

End Sub

Private Sub FillDirectPayments()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT     TOP 100 PERCENT dbo.tblIncomeBill.IncomeBillID, dbo.tblIncomeBill.DisplayBillID, dbo.tblIncomeBill.NetTotal, dbo.tblIncomeBill.BHTID, dbo.tblPaymentMethod.PaymentMethod , dbo.tblIncomeBill.CompletedDate, dbo.tblIncomeBill.CompletedTime, dbo.tblIncomeBill.PaymentComments " & _
                    "FROM         dbo.tblIncomeBill LEFT OUTER JOIN dbo.tblPaymentMethod ON dbo.tblIncomeBill.PaymentMethodID = dbo.tblPaymentMethod.PaymentMethodID " & _
                    "WHERE     (dbo.tblIncomeBill.IsGSBill = 1) AND (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.CompletedDate BETWEEN CONVERT(DATETIME, '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "', 102) AND CONVERT(DATETIME, '" & Format(dtpTo.Value, "dd MMMM yyyy") & "', 102)) " & _
                    "ORDER BY dbo.tblIncomeBill.CompletedDate, dbo.tblIncomeBill.CompletedTime"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            NewRow !BHTID, !CompletedDate, !DisplayBillID, "Direct Payment", !PaymentMethod, Format(!PaymentComments, ""), !NetTotal
            temTotal = temTotal + !NetTotal
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub FillCompanyPayments()
'    Dim rsTem As New ADODB.Recordset
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "SELECT     TOP 100 PERCENT dbo.tblIncomeBill.IncomeBillID, dbo.tblIncomeBill.DisplayBillID, dbo.tblIncomeBill.NetTotal, dbo.tblIncomeBill.BHTID, dbo.tblPaymentMethod.PaymentMethod , dbo.tblIncomeBill.CompletedDate, dbo.tblIncomeBill.CompletedTime, dbo.tblIncomeBill.PaymentComments " & _
'                    "FROM         dbo.tblIncomeBill LEFT OUTER JOIN dbo.tblPaymentMethod ON dbo.tblIncomeBill.PaymentMethodID = dbo.tblPaymentMethod.PaymentMethodID " & _
'                    "WHERE     (dbo.tblIncomeBill.IsHSSPaymentBill = 1) AND (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.CompletedDate BETWEEN CONVERT(DATETIME, '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "', 102) AND CONVERT(DATETIME, '" & Format(dtpTo.Value, "dd MMMM yyyy") & "', 102)) " & _
'                    "ORDER BY dbo.tblIncomeBill.CompletedDate, dbo.tblIncomeBill.CompletedTime"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While .EOF = False
'            NewRow !BHTID, !CompletedDate, !DisplayBillID, "Company Payment", !PaymentMethod, Format(!PaymentComments, ""), !NetTotal
'            temTotal = temTotal + !NetTotal
'            .MoveNext
'        Wend
'        .Close
'    End With

End Sub

Private Sub FillRefunds()
    Dim rsTem As New ADODB.Recordset
    Dim temMyBHT As New clsBHT
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT     TOP 100 PERCENT dbo.tblIncomeReturnBill.IncomeReturnBillID,  dbo.tblIncomeReturnBill.ReturnValue, dbo.tblIncomeReturnBill.BHTID, dbo.tblPaymentMethod.PaymentMethod , dbo.tblIncomeReturnBill.ReturnDate, dbo.tblIncomeReturnBill.ReturnTime, dbo.tblIncomeReturnBill.PaymentComments " & _
                    "FROM         dbo.tblIncomeReturnBill LEFT OUTER JOIN dbo.tblPaymentMethod ON dbo.tblIncomeReturnBill.PaymentMethodID = dbo.tblPaymentMethod.PaymentMethodID " & _
                    "WHERE   dbo.tblIncomeReturnBill.BHTID <> 0 AND   (dbo.tblIncomeReturnBill.Cancelled = 0) AND (dbo.tblIncomeReturnBill.ReturnDate BETWEEN CONVERT(DATETIME, '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "', 102) AND CONVERT(DATETIME, '" & Format(dtpTo.Value, "dd MMMM yyyy") & "', 102)) " & _
                    "ORDER BY dbo.tblIncomeReturnBill.ReturnDate, dbo.tblIncomeReturnBill.ReturnTime"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            temMyBHT.BHTID = !BHTID
            If temMyBHT.IsGSB = True Then
                NewRow !BHTID, !ReturnDate, !IncomeReturnBillID, "Refund", !PaymentMethod, Format(!PaymentComments, ""), 0 - !ReturnValue
                temTotal = temTotal - !ReturnValue
            End If
            .MoveNext
        Wend
        .Close
    End With

End Sub

Private Sub NewRow(BHTID As Long, BillDate As String, ReceiptNo As String, Descreption As String, PaidAs As String, Details As String, Value As Double)
    
    Dim temMyBHT As New clsBHT
    
    temMyBHT.BHTID = BHTID
    
    temBHTID = BHTID
    temBHT = temMyBHT.BHT
    temPM = temMyBHT.PaymentMethod
    temCC = temMyBHT.HealthSchemeSupplier
    temFBID = temMyBHT.BHTID
    temPt = temMyBHT.FirstName
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = temBHT
    gridSummery.Col = 1
    gridSummery.Text = temBHTID
    gridSummery.Col = 2
    gridSummery.Text = temPt
    gridSummery.Col = 3
    gridSummery.Text = temCC
    
    gridSummery.Col = 4
    gridSummery.Text = Format(BillDate, "dd MMM yy")
    gridSummery.Col = 5
    gridSummery.Text = ReceiptNo
    gridSummery.Col = 6
    gridSummery.Text = Descreption
    gridSummery.Col = 7
    gridSummery.Text = PaidAs
    gridSummery.Col = 8
    gridSummery.Text = Details
    gridSummery.Col = 9
    gridSummery.Text = Format(Value, "#,##0.00")

End Sub

Private Sub FormatGrid()
    With gridSummery
        .Clear
        
        .Cols = 10
        .Rows = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "GSB"
        
        .Col = 1
        .Text = "Final Bill No"
        
        .Col = 2
        .Text = "Patient"
        
        .Col = 3
        .Text = "Company"
        
        .Col = 4
        .Text = "Date"
        
        .Col = 5
        .Text = "Receipt No."
        
        .Col = 6
        .Text = "Descreption"
        
        .Col = 7
        .Text = "Paid as"
        
        .Col = 8
        .Text = "Details"
        
        .Col = 9
        .Text = "Value"
        
    End With
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpFrom.Value = Date
    dtpTo.Value = Date
    GetCommonSettings Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridSummery, "Book Keeping Summery For GS Bills", "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")

End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    ThisReportFormat.ReportPrintOrientation = Landscape
    
    
    
    GetPrintDefaults ThisReportFormat
    
    With ThisReportFormat
        
        .LeftMargin = 0
        .ColSpace = 70
        
        .TopicFontSize = 11
        .TopicFontName = "Tahoma"
        
        .SubTopicFontSize = 10
        .SubTopicFontName = "Tahoma"
        
        .HeaderFontName = "Tahoma"
        .HeaderFontSize = 8
        .HeaderFontBold = False
        .HeaderFontUnderline = False
        
        .ColTopicFontName = "Tahoma"
        .ColTopicFontSize = 8
        .ColTopicFontBold = False
        .ColTopicFontUnderline = False
        
        .ColFontSize = 7
        .ColFontName = "Tahoma"
        
    End With
    
    
    
    GridPrint gridSummery, ThisReportFormat, "Book Keeping Summery For GSBs", "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    Printer.EndDoc

End Sub

Private Sub btnProcess_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call GetSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSettings
End Sub
