VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDayEndSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Day End Summery"
   ClientHeight    =   8385
   ClientLeft      =   855
   ClientTop       =   -2445
   ClientWidth     =   14685
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
   ScaleHeight     =   8385
   ScaleWidth      =   14685
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   7800
      Width           =   4695
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   7800
      Width           =   4695
   End
   Begin VB.ComboBox cmbType 
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   13320
      TabIndex        =   28
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
      Left            =   12000
      TabIndex        =   27
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
   End
   Begin VB.Frame Frame1 
      Height          =   7575
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   14415
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
         Format          =   78643203
         CurrentDate     =   39969
      End
      Begin VB.Label lblPrintDetails 
         AutoSize        =   -1  'True
         Caption         =   "Bill Details"
         Height          =   240
         Left            =   120
         TabIndex        =   50
         Top             =   7200
         Width           =   870
      End
      Begin VB.Label lblCashPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   48
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Label lblCreditPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3750
         TabIndex        =   47
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblCardPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5580
         TabIndex        =   46
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblChequePaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   7410
         TabIndex        =   45
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblSlipsPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   9240
         TabIndex        =   44
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblPaidTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   43
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Bill Type"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblCashTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   42
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label lblCreditTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3750
         TabIndex        =   41
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label lblCardTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5580
         TabIndex        =   40
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label lblChequeTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7410
         TabIndex        =   39
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label lblSlipsTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9240
         TabIndex        =   38
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label lblGrandTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   37
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Net Income"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Cancellations"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label lblCashCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   34
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblCreditCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3750
         TabIndex        =   33
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblCardCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5580
         TabIndex        =   32
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblChequeCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   7410
         TabIndex        =   31
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblSlipsCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   9240
         TabIndex        =   30
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblCancellationTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   29
         Top             =   4080
         Width           =   1695
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
         TabIndex        =   26
         Top             =   600
         Width           =   13935
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
         TabIndex        =   25
         Top             =   240
         Width           =   13935
      End
      Begin VB.Label Label42 
         Caption         =   "Total"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label lblGrandTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   23
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblSlipsTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   9240
         TabIndex        =   22
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblChequeTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   7410
         TabIndex        =   21
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblCreditTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3750
         TabIndex        =   19
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblCashTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   18
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblReturnTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11520
         TabIndex        =   17
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label lblSlipsReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   9240
         TabIndex        =   16
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label lblChequeReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   7410
         TabIndex        =   15
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label lblCreditReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3750
         TabIndex        =   13
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label lblCashReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
         Height          =   255
         Left            =   12120
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Returns"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Slips"
         Height          =   255
         Left            =   10200
         TabIndex        =   9
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Cheque"
         Height          =   255
         Left            =   8370
         TabIndex        =   8
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Card"
         Height          =   255
         Left            =   6180
         TabIndex        =   7
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit"
         Height          =   255
         Left            =   4710
         TabIndex        =   6
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cash"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblCardTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5580
         TabIndex        =   20
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblCardReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   5580
         TabIndex        =   14
         Top             =   3720
         Width           =   1695
      End
   End
   Begin VB.Label Label29 
      Caption         =   "Paper"
      Height          =   255
      Left            =   5760
      TabIndex        =   54
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label30 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   53
      Top             =   7800
      Width           =   1815
   End
End
Attribute VB_Name = "frmDayEndSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsTem As New ADODB.Recordset
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

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
    
    
        lblPrintDetails.Caption = "Printed on " & Format(Date, "dd MMMM yyyy") & " at " & Time
    
    
        Dim MyControl As Control
        
        For Each MyControl In Controls
            If Left(MyControl.Name, 3) = "lbl" Or Left(MyControl.Name, 3) = "Lab" Then
                If MyControl.Alignment = 0 Then
                    Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
                ElseIf MyControl.Alignment = 1 Then
                    Printer.CurrentX = (MyControl.Left / Frame1.Width * Printer.Width) + (MyControl.Width / Frame1.Width * Printer.Width) - Printer.TextWidth(MyControl.Caption)
                ElseIf MyControl.Alignment = 2 Then
                    Printer.CurrentX = (MyControl.Left / Frame1.Width * Printer.Width) + ((MyControl.Width / Frame1.Width * Printer.Width) / 2) - (Printer.TextWidth(MyControl.Caption) / 2)
                End If
                Printer.Font.Size = MyControl.Font.Size
                Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
                Printer.Print MyControl.Caption
            ElseIf Left(MyControl.Name, 3) = "cmb" Then
                Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
                Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
                Printer.Print MyControl.Text
            ElseIf Left(MyControl.Name, 3) = "dtp" Then
                Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
                Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
                Printer.Print Format(MyControl.Value, "dd MMMM yyyy")
            
            End If
        Next
        
        Printer.EndDoc
    End If
End Sub

Private Sub cmbType_Change()
'    lblBillType.Caption = cmbType.Text
    Call FillDetails
End Sub

Private Sub cmbType_Click()
'    lblBillType.Caption = cmbType.Text
    Call FillDetails
End Sub

Private Sub dtpDate_Change()
'    format(dtpDate.value, "dd MMMM yyyy") = Format(dtpDate.Value, "dd MMMM yyyy")
    Call FillDetails
End Sub



Private Sub FillDetails()
    Dim TemString As String

    Dim CashTotal As Double
    Dim CreditTotal As Double
    Dim CardTotal As Double
    Dim ChequeTotal As Double
    Dim SlipsTotal As Double
    Dim GrandTotal As Double

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
        Case Else:  Exit Sub
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
    
    lblCashTotal.Caption = Format(CashTotal, "#,##0.00")
    lblCreditTotal.Caption = Format(CreditTotal, "#,##0.00")
    lblChequeTotal.Caption = Format(ChequeTotal, "#,##0.00")
    lblSlipsTotal.Caption = Format(SlipsTotal, "#,##0.00")
    lblCardTotal.Caption = Format(CardTotal, "#,##0.00")
    
    lblCashReturn.Caption = Format(CashTotalR, "#,##0.00")
    lblCreditReturn.Caption = Format(CreditTotalR, "#,##0.00")
    lblChequeReturn.Caption = Format(ChequeTotalR, "#,##0.00")
    lblSlipsReturn.Caption = Format(SlipsTotalR, "#,##0.00")
    lblCardReturn.Caption = Format(CardTotalR, "#,##0.00")
    
    lblCashCancellation.Caption = Format(CashTotalC, "#,##0.00")
    lblCreditCancellation.Caption = Format(CreditTotalC, "#,##0.00")
    lblChequeCancellation.Caption = Format(ChequeTotalC, "#,##0.00")
    lblSlipsCancellation.Caption = Format(SlipsTotalC, "#,##0.00")
    lblCardCancellation.Caption = Format(CardTotalC, "#,##0.00")
    
    lblCashPaid.Caption = Format(CashTotalP, "#,##0.00")
    lblCreditPaid.Caption = Format(CreditTotalP, "#,##0.00")
    lblChequePaid.Caption = Format(ChequeTotalP, "#,##0.00")
    lblSlipsPaid.Caption = Format(SlipsTotalP, "#,##0.00")
    lblCardPaid.Caption = Format(CardTotalP, "#,##0.00")
   
    
    lblCashTotalN.Caption = Format(CashTotal - (CashTotalC + CashTotalR + CashTotalP), "#,##0.00")
    lblCreditTotalN.Caption = Format(CreditTotal - (CreditTotalC + CreditTotalR + CreditTotalP), "#,##0.00")
    lblChequeTotalN.Caption = Format(ChequeTotal - (ChequeTotalC + ChequeTotalR + ChequeTotalP), "#,##0.00")
    lblSlipsTotalN.Caption = Format(SlipsTotal - (SlipsTotalC + SlipsTotalR + SlipsTotalP), "#,##0.00")
    lblCardTotalN.Caption = Format(CardTotal - (CardTotalC + CardTotalR + CardTotalP), "#,##0.00")
    
    
    GrandTotal = CashTotal + CreditTotal + ChequeTotal + SlipsTotal + CardTotal
    GrandTotalR = CashTotalR + CreditTotalR + ChequeTotalR + SlipsTotalR + CardTotalR
    GrandTotalC = CashTotalC + CreditTotalC + ChequeTotalC + SlipsTotalC + CardTotalC
    GrandTotalP = CashTotalP + CreditTotalP + ChequeTotalP + SlipsTotalP + CardTotalP
    
    lblGrandTotal.Caption = Format(GrandTotal, "#,##0.00")
    lblReturnTotal.Caption = Format(GrandTotalR, "#,##0.00")
    lblCancellationTotal.Caption = Format(GrandTotalC, "#,##0.00")
    lblPaidTotal.Caption = Format(GrandTotalP, "#,##0.00")
    
    lblGrandTotalN.Caption = Format(GrandTotal - (GrandTotalR + GrandTotalC + GrandTotalP), "#,##0.00")


End Sub



Private Function CatIncome(ByVal IncomeCategory As String, PaymentMethodID As Long) As Double
    CatIncome = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeBill.NetTotal) AS SumOfNetTotal From tblIncomeBill "
        If IncomeCategory = "All" Then
            temSql = temSql & "WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1)AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
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
            temSql = temSql & "WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ") AND ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        Else
            temSql = temSql & "WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ") AND ((tblIncomeReturnBill.Cancelled)=0))"
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
            temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblBHT.IsBHT)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        ElseIf IncomeCategory = "GS" Then
            temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblBHT.IsGSB)=1) AND  ((tblIncomeReturnBill.Cancelled)=0))"
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
            temSql = temSql & "WHERE (((tblIncomeBill.CancelledDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeBill.Cancelled)=1) AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "WHERE (((tblIncomeBill.CancelledDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Cancelled)=1) AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
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
            temSql = temSql & "WHERE (((tblProfessionalPaymentBill.Date)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblProfessionalPaymentBill.Is" & IncomeCategory & "Bill)=1) AND ((tblProfessionalPaymentBill.PaymentMethodID)=" & PaymentMethodID & "))"
        
        Else
            temSql = temSql & "WHERE (((tblProfessionalPaymentBill.Date)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblProfessionalPaymentBill.PaymentMethodID)=" & PaymentMethodID & "))"
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
    Call PopulatePrinters
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
    Call FillDetails
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    lblTopic.Caption = HospitalName
    lblSubtopic.Caption = "Day End Summery - " & Format(dtpDate.Value, "dd MMMM yyyy")
    lblPrintDetails.Caption = "Printed on " & Format(Date, "dd MMMM yyyy") & " at " & Time
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
    GetCommonSettings Me
    
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    
    SaveCommonSettings Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSettings
End Sub

Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub PopulatePapers(): On Error Resume Next
    cmbPaper.Clear
    SetPrinter = False
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
'        With FormSize
'            .cx = BillPaperHeight
'            .cy = BillPaperWidth
'        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                cmbPaper.AddItem PtrCtoVbString(.pName)
'                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle): DoEvents
    End If
End Sub

Private Sub cmbPrinter_Change()
    Call PopulatePapers
End Sub

Private Sub cmbPrinter_Click()
    Call PopulatePapers
End Sub

