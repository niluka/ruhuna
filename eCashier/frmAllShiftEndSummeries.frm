VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAllShiftEndSummeries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Shift End Summeries"
   ClientHeight    =   8400
   ClientLeft      =   855
   ClientTop       =   -2445
   ClientWidth     =   13185
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
   ScaleHeight     =   8400
   ScaleWidth      =   13185
   Begin MSFlexGridLib.MSFlexGrid gridSummery 
      Height          =   5175
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9128
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11880
      TabIndex        =   8
      Top             =   7800
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   10560
      TabIndex        =   7
      Top             =   7800
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
      ShowFocus       =   0
   End
   Begin VB.Frame Frame1 
      Height          =   7575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   12975
      Begin btButtonEx.ButtonEx btnProcess 
         Height          =   375
         Left            =   5760
         TabIndex        =   11
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.ComboBox cmbType 
         Height          =   360
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   1200
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   78053379
         CurrentDate     =   39969
      End
      Begin VB.Label lblBillTimes 
         AutoSize        =   -1  'True
         Caption         =   "Bill Details"
         Height          =   240
         Left            =   5880
         TabIndex        =   9
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label Label8 
         Caption         =   "Bill Type"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblSubtopic 
         Alignment       =   2  'Center
         Caption         =   "Topic"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   7455
      End
      Begin VB.Label lblTopic 
         Alignment       =   2  'Center
         Caption         =   "Topic"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   7455
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   1200
         Width           =   1455
      End
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   9240
      TabIndex        =   12
      Top             =   7800
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
End
Attribute VB_Name = "frmAllShiftEndSummeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsTem As New ADODB.Recordset
    Dim temUserID As Long
    
    Private Type MyTotal
        CashTotal As Double
        ChequeTotal As Double
        CreditTotal As Double
        SlipsTotal As Double
        CardTotal As Double
        GrandTotal As Double
    End Type
    

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
        .HeaderFontSize = 10
        .HeaderFontBold = False
        .HeaderFontUnderline = False
        
        .ColTopicFontName = "Tahoma"
        .ColTopicFontSize = 10
        .ColTopicFontBold = False
        .ColTopicFontUnderline = False
        
        .ColFontSize = 10
        .ColFontName = "Tahoma"
        
    End With
    
    
    
    GridPrint gridSummery, ThisReportFormat, "All Shiftend summeries", "On " & Format(dtpDate.Value, "dd MMMM yyyy")
    Printer.EndDoc

End Sub

Private Sub btnExcel_Click()
    GridToExcel gridSummery, "All Shiftend summeries", "On " & Format(dtpDate.Value, "dd MMMM yyyy")

End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnProcess_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Function FillDetails() As MyTotal

    Dim TemString As String

    Dim CashTotal As Double
    Dim CreditTotal As Double
    Dim CardTotal As Double
    Dim ChequeTotal As Double
    Dim SlipsTotal As Double
    Dim GrandTotal As Double

    Dim CashTotalA As Double
    Dim CreditTotalA As Double
    Dim CardTotalA As Double
    Dim ChequeTotalA As Double
    Dim SlipsTotalA As Double
    Dim GrandTotalA As Double


    Dim CashTotalC As Double
    Dim CreditTotalC As Double
    Dim CardTotalC As Double
    Dim ChequeTotalC As Double
    Dim SlipsTotalC As Double
    Dim GrandTotalC As Double

    Dim CashTotalR As Double
    Dim CreditTotalR As Double
    Dim CardTotalR As Double
    Dim ChequeTotalR As Double
    Dim SlipsTotalR As Double
    Dim GrandTotalR As Double

    Dim CashTotalP As Double
    Dim CreditTotalP As Double
    Dim CardTotalP As Double
    Dim ChequeTotalP As Double
    Dim SlipsTotalP As Double
    Dim GrandTotalP As Double

    
    Select Case cmbType.Text
        Case "OPD Bills": TemString = "OPD"
        Case "Lab Bills": TemString = "Lab"
        Case "Pharmacy Bills": TemString = "Pharmacy"
        Case "Inward Bills": TemString = "InwardPayment"
        Case "Medical Test Bills": TemString = "MedicalTest"
        Case "Green Sheet Bills": TemString = "GS"
        Case "All Bills": TemString = "All"
        Case "Agent Bills": TemString = "Agent"
        Case "Expence Bills": TemString = "Expence"
        Case "Roentgents Bills": TemString = "R"
        Case "Health Scheme Supplier Payments": TemString = "HSSPayment"
        Case Else:  Exit Function
    End Select
    
    
    CashTotal = CatIncome(TemString, 1)
    CreditTotal = CatIncome(TemString, 4)
    ChequeTotal = CatIncome(TemString, 5)
    SlipsTotal = CatIncome(TemString, 7)
    CardTotal = CatIncome(TemString, 3)
    
    CashTotalC = CatCancellation(TemString, 1)
    CreditTotalC = CatCancellation(TemString, 4)
    ChequeTotalC = CatCancellation(TemString, 5)
    SlipsTotalC = CatCancellation(TemString, 7)
    CardTotalC = CatCancellation(TemString, 3)
    
    CashTotalR = CatReturn(TemString, 1)
    CreditTotalR = CatReturn(TemString, 4)
    ChequeTotalR = CatReturn(TemString, 5)
    SlipsTotalR = CatReturn(TemString, 7)
    CardTotalR = CatReturn(TemString, 3)
    
    CashTotalP = CatPay(TemString, 1)
    CreditTotalP = CatPay(TemString, 4)
    ChequeTotalP = CatPay(TemString, 5)
    SlipsTotalP = CatPay(TemString, 7)
    CardTotalP = CatPay(TemString, 3)
    
'    lblCashTotal.Caption = Format(CashTotal, "#,##0.00")
'    lblCreditTotal.Caption = Format(CreditTotal, "#,##0.00")
'    lblChequeTotal.Caption = Format(ChequeTotal, "#,##0.00")
'    lblSlipsTotal.Caption = Format(SlipsTotal, "#,##0.00")
'    lblCardTotal.Caption = Format(CardTotal, "#,##0.00")
'
'    lblCashReturn.Caption = Format(CashTotalR, "#,##0.00")
'    lblCreditReturn.Caption = Format(CreditTotalR, "#,##0.00")
'    lblChequeReturn.Caption = Format(ChequeTotalR, "#,##0.00")
'    lblSlipsReturn.Caption = Format(SlipsTotalR, "#,##0.00")
'    lblCardReturn.Caption = Format(CardTotalR, "#,##0.00")
'
'    lblCashCancellation.Caption = Format(CashTotalC, "#,##0.00")
'    lblCreditCancellation.Caption = Format(CreditTotalC, "#,##0.00")
'    lblChequeCancellation.Caption = Format(ChequeTotalC, "#,##0.00")
'    lblSlipsCancellation.Caption = Format(SlipsTotalC, "#,##0.00")
'    lblCardCancellation.Caption = Format(CardTotalC, "#,##0.00")
'
'    lblCashPaid.Caption = Format(CashTotalP, "#,##0.00")
'    lblCreditPaid.Caption = Format(CreditTotalP, "#,##0.00")
'    lblChequePaid.Caption = Format(ChequeTotalP, "#,##0.00")
'    lblSlipsPaid.Caption = Format(SlipsTotalP, "#,##0.00")
'    lblCardPaid.Caption = Format(CardTotalP, "#,##0.00")
'
'
    FillDetails.CashTotal = Format(CashTotal - (CashTotalC + CashTotalR + CashTotalP), "#,##0.00")
    FillDetails.CreditTotal = Format(CreditTotal - (CreditTotalC + CreditTotalR + CreditTotalP), "#,##0.00")
    FillDetails.ChequeTotal = Format(ChequeTotal - (ChequeTotalC + ChequeTotalR + ChequeTotalP), "#,##0.00")
    FillDetails.SlipsTotal = Format(SlipsTotal - (SlipsTotalC + SlipsTotalR + SlipsTotalP), "#,##0.00")
    FillDetails.CardTotal = Format(CardTotal - (CardTotalC + CardTotalR + CardTotalP), "#,##0.00")
    
    
    GrandTotal = CashTotal + CreditTotal + ChequeTotal + SlipsTotal + CardTotal
    GrandTotalR = CashTotalR + CreditTotalR + ChequeTotalR + SlipsTotalR + CardTotalR
    GrandTotalC = CashTotalC + CreditTotalC + ChequeTotalC + SlipsTotalC + CardTotalC
    GrandTotalP = CashTotalP + CreditTotalP + ChequeTotalP + SlipsTotalP + CardTotalP
    
'    lblGrandTotal.Caption = Format(GrandTotal, "#,##0.00")
'    lblReturnTotal.Caption = Format(GrandTotalR, "#,##0.00")
'    lblCancellationTotal.Caption = Format(GrandTotalC, "#,##0.00")
'    lblPaidTotal.Caption = Format(GrandTotalP, "#,##0.00")
'
'    lblGrandTotalN.Caption = Format(GrandTotal - (GrandTotalR + GrandTotalC + GrandTotalP), "#,##0.00")

    FillDetails.GrandTotal = Format(GrandTotal - (GrandTotalR + GrandTotalC + GrandTotalP), "#,##0.00")

'    lblBillTimes.Caption = "Bills From " & MinTime & " to " & MaxTime

End Function

Private Function CatIncome(ByVal IncomeCategory As String, PaymentMethodID As Long) As Double
    CatIncome = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeBill.NetTotal) AS SumOfNetTotal From tblIncomeBill "
        If IncomeCategory = "All" Then
            temSql = temSql & "Where (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.CompletedUserID)=" & temUserID & ") AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "Where (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "')  AND ((tblIncomeBill.CompletedUserID)=" & temUserID & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1)AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfNetTotal) = False Then
                CatIncome = !SumOfNetTotal
            End If
        End If
        .Close
    End With
End Function

Private Function CatReturn(ByVal IncomeCategory As String, PaymentMethodID As Long) As Double
    CatReturn = 0
    If IncomeCategory = "InwardPayment" Or IncomeCategory = "GS" Then
        CatReturn = CatBHTGSBReturn(IncomeCategory, PaymentMethodID)
        Exit Function
    End If
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue " & _
                    "FROM tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID "
        If IncomeCategory <> "All" Then
            temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblIncomeReturnBill.ReturnUserID)=" & temUserID & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        Else
            temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblIncomeReturnBill.ReturnUserID)=" & temUserID & ") AND  ((tblIncomeReturnBill.Cancelled)=0))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfReturnValue) = False Then
                CatReturn = !SumOfReturnValue
            End If
        End If
        .Close
    End With
End Function


Private Function CatBHTGSBReturn(ByVal IncomeCategory As String, PaymentMethodID As Long) As Double
    CatBHTGSBReturn = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue " & _
                    "FROM tblIncomeReturnBill INNER JOIN tblBHT ON tblIncomeReturnBill.BHTID = tblBHT.BHTID "
        If IncomeCategory = "InwardPayment" Then
            temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblIncomeReturnBill.ReturnUserID)=" & temUserID & ") AND  ((tblBHT.IsBHT)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        ElseIf IncomeCategory = "GS" Then
            temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblIncomeReturnBill.ReturnUserID)=" & temUserID & ") AND  ((tblBHT.IsGSB)=1) AND  ((tblIncomeReturnBill.Cancelled)=0))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfReturnValue) = False Then
                CatBHTGSBReturn = !SumOfReturnValue
            End If
        End If
        .Close
    End With
End Function

Private Function CatCancellation(ByVal IncomeCategory As String, PaymentMethodID As Long) As Double
    CatCancellation = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeBill.NetTotal) AS SumOfNetTotal " & _
                    "From tblIncomeBill "
        If IncomeCategory <> "All" Then
            temSql = temSql & "Where (((tblIncomeBill.CancelledDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "')  AND ((tblIncomeBill.CancelledUserID)=" & temUserID & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeBill.Cancelled)=1) AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "Where (((tblIncomeBill.CancelledDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Cancelled)=1)  AND ((tblIncomeBill.CancelledUserID)=" & temUserID & ") AND  ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfNetTotal) = False Then
                CatCancellation = !SumOfNetTotal
            End If
        End If
        .Close
    End With
End Function

Private Function CatPay(ByVal IncomeCategory As String, PaymentMethodID As Long) As Double
    CatPay = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblProfessionalPaymentBill.Value) AS SumOfValue " & _
                    "FROM tblProfessionalPaymentBill "
        If IncomeCategory <> "All" Then
            temSql = temSql & "Where (((tblProfessionalPaymentBill.Date)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblProfessionalPaymentBill.Is" & IncomeCategory & "Bill)=1) AND ((tblProfessionalPaymentBill.UserID)=" & temUserID & ") AND ((tblProfessionalPaymentBill.PaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "Where (((tblProfessionalPaymentBill.Date)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblProfessionalPaymentBill.UserID)=" & temUserID & ") AND ((tblProfessionalPaymentBill.PaymentMethodID)=" & PaymentMethodID & "))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfValue) = False Then
                CatPay = !SumOfValue
            End If
        End If
        .Close
    End With
End Function

Private Sub Form_Load()
    Call FormatGrid
    Call GetSettings
    cmbType.AddItem "Inward Bills"
    cmbType.AddItem "Green Sheet Bills"
    cmbType.AddItem "OPD Bills"
    cmbType.AddItem "Roentgents Bills"
    cmbType.AddItem "Lab Bills"
    cmbType.AddItem "Pharmacy Bills"
    cmbType.AddItem "Medical Test Bills"
    cmbType.AddItem "Agent Bills"
    cmbType.AddItem "Expence Bills"
    cmbType.AddItem "Health Scheme Supplier Payments"
    cmbType.AddItem "Health Screening Test Bills"
    cmbType.AddItem "All Bills"
    cmbType.Text = "All Bills"
    Call FormatGrid
    Call FillDetails
    
End Sub

Private Sub FormatGrid()
    With gridSummery
        .Clear
        .Rows = 1
        .Cols = 8
        
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        
        .Col = 1
        .Text = "User"
        
        .Col = 2
        .Text = "Cash"
        
        .Col = 3
        .Text = "Credit"
        .Col = 4
        .Text = "Cheque"
        .Col = 6
        .Text = "Slips"
        .Col = 5
        .Text = "Card"
        .Col = 7
        .Text = "Total"
        
        
        
    End With
    
    
'    GrandTotal = CashTotal + CreditTotal + ChequeTotal + SlipsTotal + CardTotal
'    GrandTotalR = CashTotalR + CreditTotalR + ChequeTotalR + SlipsTotalR + CardTotalR
'    GrandTotalC = CashTotalC + CreditTotalC + ChequeTotalC + SlipsTotalC + CardTotalC
'    GrandTotalP = CashTotalP + CreditTotalP + ChequeTotalP + SlipsTotalP + CardTotalP
    
    
End Sub


Private Sub FillGrid()

    Dim AllUserNetTotal As Double
    Dim AllUserCashTotal As Double
    Dim AllUserCreditTotal As Double
    Dim AllUserCardTotal As Double
    Dim AllUserChequeTotal As Double
    Dim AllUserSlipsTotal As Double

    Screen.MousePointer = vbHourglass
    
    Dim rsTem As New ADODB.Recordset
    Dim temNetTotal As MyTotal
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where Deleted = 0 and IsAUser = 1 order By Name"
        
        temUserID = 0
        AllUserNetTotal = 0
        AllUserCashTotal = 0
        AllUserCreditTotal = 0
        AllUserChequeTotal = 0
        AllUserCardTotal = 0
        AllUserSlipsTotal = 0
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            temUserID = !StaffID
            temNetTotal = FillDetails
            If temNetTotal.CashTotal <> 0 Or temNetTotal.CreditTotal <> 0 Or temNetTotal.CardTotal <> 0 Or temNetTotal.ChequeTotal <> 0 Or temNetTotal.SlipsTotal <> 0 Or temNetTotal.GrandTotal <> 0 Then
            
                AllUserNetTotal = AllUserNetTotal + temNetTotal.GrandTotal
                AllUserCashTotal = AllUserCashTotal + temNetTotal.CashTotal
                AllUserCreditTotal = AllUserCreditTotal + temNetTotal.CreditTotal
                AllUserChequeTotal = AllUserChequeTotal + temNetTotal.ChequeTotal
                AllUserCardTotal = AllUserCardTotal + temNetTotal.CardTotal
                AllUserSlipsTotal = AllUserSlipsTotal + temNetTotal.SlipsTotal
            
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                
                gridSummery.Col = 0
                gridSummery.Text = !StaffID
                
                gridSummery.Col = 1
                gridSummery.Text = !Name
                
                gridSummery.Col = 2
                gridSummery.Text = Format(temNetTotal.CashTotal, "#,##0.00")
                
                gridSummery.Col = 3
                gridSummery.Text = Format(temNetTotal.CreditTotal, "#,##0.00")
                gridSummery.Col = 4
                gridSummery.Text = Format(temNetTotal.ChequeTotal, "#,##0.00")
                gridSummery.Col = 5
                gridSummery.Text = Format(temNetTotal.CardTotal, "#,##0.00")
                gridSummery.Col = 6
                gridSummery.Text = Format(temNetTotal.SlipsTotal, "#,##0.00")
                gridSummery.Col = 7
                gridSummery.Text = Format(temNetTotal.GrandTotal, "#,##0.00")
                
            
            End If
            .MoveNext
        Wend
        .Close
    End With
    gridSummery.ColWidth(0) = 0
    

    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    
    gridSummery.Col = 0
    gridSummery.Text = 0
    
    gridSummery.Col = 1
    gridSummery.Text = "Total"
    
    gridSummery.Col = 2
    gridSummery.Text = Format(AllUserCashTotal, "#,##0.00")
    gridSummery.Col = 3
    gridSummery.Text = Format(AllUserCreditTotal, "#,##0.00")
    gridSummery.Col = 4
    gridSummery.Text = Format(AllUserChequeTotal, "#,##0.00")
    gridSummery.Col = 6
    gridSummery.Text = Format(AllUserSlipsTotal, "#,##0.00")
    gridSummery.Col = 5
    gridSummery.Text = Format(AllUserCardTotal, "#,##0.00")
    gridSummery.Col = 7
    gridSummery.Text = Format(AllUserNetTotal, "#,##0.00")

    
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Function MinTime() As Date
    Dim IncomeCategory As String
    IncomeCategory = "All"
    MinTime = Time
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT min(tblIncomeBill.CompletedTime) AS MinTime From tblIncomeBill "
        If IncomeCategory = "All" Then
            temSql = temSql & "WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.CompletedUserID)=" & temUserID & ") AND ((tblIncomeBill.Completed)=1))"
        Else
            temSql = temSql & "WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "')  AND ((tblIncomeBill.CompletedUserID)=" & temUserID & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1)AND ((tblIncomeBill.Completed)=1))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MinTime) = False Then
                If MinTime > !MinTime Then
                    MinTime = !MinTime
                End If
            End If
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Min(tblIncomeReturnBill.ReturnTime) AS MinTime " & _
                    "FROM tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID "
        If IncomeCategory <> "All" Then
            temSql = temSql & "WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.ReturnUserID)=" & temUserID & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        Else
            temSql = temSql & "WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.ReturnUserID)=" & temUserID & ") AND  ((tblIncomeReturnBill.Cancelled)=0))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MinTime) = False Then
                If MinTime > !MinTime Then
                    MinTime = !MinTime
                End If
            End If
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Min(tblProfessionalPaymentBill.Time) AS MinTime " & _
                    "FROM tblProfessionalPaymentBill "
        If IncomeCategory <> "All" Then
            temSql = temSql & "WHERE (((tblProfessionalPaymentBill.Date)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblProfessionalPaymentBill.Is" & IncomeCategory & "Bill)=1) AND ((tblProfessionalPaymentBill.UserID)=" & temUserID & "))"
        Else
            temSql = temSql & "WHERE (((tblProfessionalPaymentBill.Date)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblProfessionalPaymentBill.UserID)=" & temUserID & "))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MinTime) = False Then
                If MinTime > !MinTime Then
                    MinTime = !MinTime
                End If
            End If
        End If
        .Close
    End With

End Function

Private Function MaxTime() As Date
    Dim IncomeCategory As String
    IncomeCategory = "All"
    MaxTime = Time
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Max(tblIncomeBill.CompletedTime) AS MaxTime From tblIncomeBill "
        If IncomeCategory = "All" Then
            temSql = temSql & "WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.CompletedUserID)=" & temUserID & ") AND ((tblIncomeBill.Completed)=1))"
        Else
            temSql = temSql & "WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "')  AND ((tblIncomeBill.CompletedUserID)=" & temUserID & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1)AND ((tblIncomeBill.Completed)=1))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MaxTime) = False Then
                If MaxTime < !MaxTime Then
                    MaxTime = !MaxTime
                End If
            End If
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Max(tblIncomeReturnBill.ReturnTime) AS MaxTime " & _
                    "FROM tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID "
        If IncomeCategory <> "All" Then
            temSql = temSql & "WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.ReturnUserID)=" & temUserID & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        Else
            temSql = temSql & "WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.ReturnUserID)=" & temUserID & ") AND  ((tblIncomeReturnBill.Cancelled)=0))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MaxTime) = False Then
                If MaxTime < !MaxTime Then
                    MaxTime = !MaxTime
                End If
            End If
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Max(tblProfessionalPaymentBill.Time) AS MaxTime " & _
                    "FROM tblProfessionalPaymentBill "
        If IncomeCategory <> "All" Then
            temSql = temSql & "WHERE (((tblProfessionalPaymentBill.Date)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblProfessionalPaymentBill.Is" & IncomeCategory & "Bill)=1) AND ((tblProfessionalPaymentBill.UserID)=" & temUserID & "))"
        Else
            temSql = temSql & "WHERE (((tblProfessionalPaymentBill.Date)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblProfessionalPaymentBill.UserID)=" & temUserID & "))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!MaxTime) = False Then
                If MaxTime < !MaxTime Then
                    MaxTime = !MaxTime
                End If
            End If
        End If
        .Close
    End With

End Function

Private Sub SaveSettings()
    SaveCommonSettings Me
End Sub


Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
'    lblPrintDetails.Caption = "Printed on " & Format(Date, "dd MMMM yyyy") & " at " & Time
    lblTopic = HospitalName
    lblSubtopic.Caption = "Shift End Summery"
    GetCommonSettings Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSettings
End Sub

Private Sub gridSummery_DblClick()
    Dim temID As Long
    With gridSummery
        temID = Val(.TextMatrix(.Row, 0))
        If temID <> 0 Then
            Unload frmShiftEndSummery
            
            frmShiftEndSummery.Show
            frmShiftEndSummery.ZOrder 0
            frmShiftEndSummery.cmbUser.BoundText = temID
            frmShiftEndSummery.dtpDate.Value = dtpDate.Value
            frmShiftEndSummery.FillDetails
            
        End If
    End With

End Sub
