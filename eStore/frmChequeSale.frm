VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmChequeSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque Sale"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8640
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   7200
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   66715651
      CurrentDate     =   29224
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   66715651
      CurrentDate     =   29224
   End
   Begin MSFlexGridLib.MSFlexGrid GridSales 
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Total Cost"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label lblTotalCost 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Total Cheque Collection"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmChequeSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsSale As New ADODB.Recordset
    Dim CSetPrinter As New cSetDfltPrinter
    Dim rsReport As New ADODB.Recordset

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    Dim TemResponce As Long
    Dim RetVal As Integer
    
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
        Case FORM_SELECTED   ' 1
            Call SelectPrint
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
    End Select
End Sub

Private Sub SelectPrint()
    With rsReport
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleBill.SaleBillID, tblSaleBill.NetCost, tblSaleBill.NetPrice, tblSaleBill.Time, tblSaleBill.Date, tblStaff.Name, tblSaleBill.NetPrice " & _
                    "FROM (tblSaleBill LEFT JOIN tblStaff ON tblSaleBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblSaleBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID " & _
                    "WHERE (((tblSaleBill.Date) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblPaymentMethod.PaymentMethod)='Cheque'))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtrCashSale
        Set .DataSource = rsReport
        .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
        .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
        .Sections("Section4").Controls.Item("lblTopic").Caption = "Total Cheque Sale"
        .Sections("Section4").Controls.Item("lblSubTopic").Caption = "From " & dtpFrom.Value & " to " & dtpTo.Value
        .Show
    End With
End Sub


Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FormatGrid()
    With GridSales
        .Clear
        
        .Rows = 1
        .Cols = 5
        
        .Row = 0
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Bill ID"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Date"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Time"
        
        .Col = 3
        .CellAlignment = 4
        .Text = "Staff"
        
        .Col = 4
        .CellAlignment = 4
        .Text = "Amount"
        
    .ColWidth(0) = 800
    .ColWidth(1) = 1600
    .ColWidth(2) = 1600
    .ColWidth(3) = 2200
    .ColWidth(4) = 1800
    
    End With
    
    
    lblTotal.Caption = "0.00"
    
End Sub

Private Sub FillGrid()
    With rsSale
        If .State = 1 Then .Close
        temSql = "SELECT tblSaleBill.SaleBillID, tblSaleBill.NetCost, tblSaleBill.Date, tblSaleBill.Time, tblStaff.Name, tblSaleBill.NetPrice " & _
                    "FROM (tblSaleBill LEFT JOIN tblStaff ON tblSaleBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblSaleBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID " & _
                    "WHERE (((tblSaleBill.Date) Between '" & dtpFrom.Value & "' And '" & dtpTo.Value & "') AND ((tblPaymentMethod.PaymentMethod)='Cheque'))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            GridSales.Rows = .RecordCount + 1
            .MoveFirst
            Dim i As Integer
            Dim TCash As Double
            Dim TCost As Double
            While .EOF = False
                i = i + 1
                GridSales.TextMatrix(i, 0) = !SaleBillID
                GridSales.TextMatrix(i, 1) = !Date
                GridSales.TextMatrix(i, 2) = !Time
                GridSales.TextMatrix(i, 3) = !Name
                GridSales.TextMatrix(i, 4) = Format(!NetPrice, "0.00")
                
                TCash = TCash + !NetPrice
                If Not IsNull(!NetCost) Then
                    TCost = TCost + !NetCost
                End If
                .MoveNext
            Wend
        End If
    End With
    lblTotal.Caption = Format(TCash, "0.00")
    lblTotalCost.Caption = Format(TCost, "0.00")
End Sub

