VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDayEndSummeryCounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Day End Counts"
   ClientHeight    =   7755
   ClientLeft      =   450
   ClientTop       =   -2445
   ClientWidth     =   12495
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
   ScaleHeight     =   7755
   ScaleWidth      =   12495
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11160
      TabIndex        =   9
      Top             =   7200
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
      Left            =   9840
      TabIndex        =   8
      Top             =   7200
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
      Height          =   6975
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   12255
      Begin VB.ComboBox cmbType 
         Height          =   360
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2520
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   78643203
         CurrentDate     =   39969
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   78643202
         CurrentDate     =   39969
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   78643202
         CurrentDate     =   39969
      End
      Begin VB.Label Label11 
         Caption         =   "From"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "To"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Bill Type"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblCashTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   2145
         TabIndex        =   46
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label lblCreditTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   3435
         TabIndex        =   45
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label lblCardTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   4740
         TabIndex        =   44
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label lblChequeTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   6045
         TabIndex        =   43
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label lblSlipsTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   7335
         TabIndex        =   42
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label lblGrandTotalN 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   8640
         TabIndex        =   41
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Net"
         Height          =   255
         Left            =   360
         TabIndex        =   40
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Cancellations"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label lblCashCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   38
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblCreditCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3435
         TabIndex        =   37
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblCardCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   4740
         TabIndex        =   36
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblChequeCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   6045
         TabIndex        =   35
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblSlipsCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   7335
         TabIndex        =   34
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label lblCancellationTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   8640
         TabIndex        =   33
         Top             =   5400
         Width           =   1695
      End
      Begin VB.Label lblSubtopic 
         Alignment       =   2  'Center
         Caption         =   "Topic"
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   600
         Width           =   10095
      End
      Begin VB.Label lblTopic 
         Alignment       =   2  'Center
         Caption         =   "Topic"
         Height          =   375
         Left            =   360
         TabIndex        =   31
         Top             =   360
         Width           =   10095
      End
      Begin VB.Label Label42 
         Caption         =   "Total"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label lblGrandTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   8640
         TabIndex        =   29
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label lblSlipsTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   7335
         TabIndex        =   28
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label lblChequeTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   6045
         TabIndex        =   27
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label lblCardTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   4740
         TabIndex        =   26
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label lblCreditTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3435
         TabIndex        =   25
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label lblCashTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2145
         TabIndex        =   24
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label lblReturnTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   8640
         TabIndex        =   23
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label lblSlipsReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   7335
         TabIndex        =   22
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label lblChequeReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   6045
         TabIndex        =   21
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label lblCardReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   4740
         TabIndex        =   20
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label lblCreditReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   3435
         TabIndex        =   19
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label lblCashReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
         Height          =   255
         Left            =   9240
         TabIndex        =   17
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Returns"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Slips"
         Height          =   255
         Left            =   7800
         TabIndex        =   15
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Cheque"
         Height          =   255
         Left            =   6600
         TabIndex        =   14
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Card"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cash"
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmDayEndSummeryCounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsTem As New ADODB.Recordset

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()


    Dim MyControl As Control
    
    For Each MyControl In Controls
        If TypeOf MyControl Is Label Then
            If MyControl.Alignment = 0 Then
                Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
            ElseIf MyControl.Alignment = 1 Then
                Printer.CurrentX = (MyControl.Left / Frame1.Width * Printer.Width) + (MyControl.Width / Frame1.Width * Printer.Width) - Printer.TextWidth(MyControl.Caption)
            ElseIf MyControl.Alignment = 2 Then
                Printer.CurrentX = (MyControl.Left / Frame1.Width * Printer.Width) + ((MyControl.Width / Frame1.Width * Printer.Width) / 2) - Printer.TextWidth(MyControl.Caption)
            End If
            Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
            Printer.Print MyControl.Caption
        ElseIf TypeOf MyControl Is DTPicker Then
            Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
            Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
            If MyControl.Format = dtpTime Then
                Printer.Print MyControl.Value
            Else
                Printer.Print Format(MyControl.Value, "dd MMMM yyyy")
            End If
        ElseIf TypeOf MyControl Is ComboBox Then
            Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
            Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
            Printer.Print MyControl.Text
        End If
    Next
    
    Printer.EndDoc
End Sub

Private Sub cmbType_Change()
    Call FillDetails
End Sub

Private Sub cmbType_Click()
    Call FillDetails
End Sub

Private Sub dtpDate_Change()
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
    
    lblCashTotal.Caption = Format(CashTotal, "0")
    lblCreditTotal.Caption = Format(CreditTotal, "0")
    lblChequeTotal.Caption = Format(ChequeTotal, "0")
    lblSlipsTotal.Caption = Format(SlipsTotal, "0")
    lblCardTotal.Caption = Format(CardTotal, "0")
    
    lblCashReturn.Caption = Format(CashTotalR, "0")
    lblCreditReturn.Caption = Format(CreditTotalR, "0")
    lblChequeReturn.Caption = Format(ChequeTotalR, "0")
    lblSlipsReturn.Caption = Format(SlipsTotalR, "0")
    lblCardReturn.Caption = Format(CardTotalR, "0")
    
    lblCashCancellation.Caption = Format(CashTotalC, "0")
    lblCreditCancellation.Caption = Format(CreditTotalC, "0")
    lblChequeCancellation.Caption = Format(ChequeTotalC, "0")
    lblSlipsCancellation.Caption = Format(SlipsTotalC, "0")
    lblCardCancellation.Caption = Format(CardTotalC, "0")
    
    lblCashTotalN.Caption = Format(CashTotal - (CashTotalC + CashTotalR), "0")
    lblCreditTotalN.Caption = Format(CreditTotal - (CreditTotalC + CreditTotalR), "0")
    lblChequeTotalN.Caption = Format(ChequeTotal - (ChequeTotalC + ChequeTotalR), "0")
    lblSlipsTotalN.Caption = Format(SlipsTotal - (SlipsTotalC + SlipsTotalR), "0")
    lblCardTotalN.Caption = Format(CardTotal - (CardTotalC + CardTotalR), "0")
    
    
    GrandTotal = CashTotal + CreditTotal + ChequeTotal + SlipsTotal + CardTotal
    GrandTotalR = CashTotalR + CreditTotalR + ChequeTotalR + SlipsTotalR + CardTotalR
    GrandTotalC = CashTotalC + CreditTotalC + ChequeTotalC + SlipsTotalC + CardTotalC
    
    lblGrandTotal.Caption = Format(GrandTotal, "0")
    lblReturnTotal.Caption = Format(GrandTotalR, "0")
    lblCancellationTotal.Caption = Format(GrandTotalC, "0")

    lblGrandTotalN.Caption = Format(GrandTotal - (GrandTotalR + GrandTotalC), "0")


End Sub



Private Function CatIncome(ByVal IncomeCategory As String, PaymentMethodID As Long) As Double
    CatIncome = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select Count(tblIncomeBill.NetTotal) AS SumOfNetTotal From tblIncomeBill "
        If IncomeCategory = "All" Then
            temSql = temSql & "WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.CompletedTime)>='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpFrom.Value & "') AND ((tblIncomeBill.CompletedTime)<='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTo.Value & "') AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.CompletedTime)>='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpFrom.Value & "') AND ((tblIncomeBill.CompletedTime)<='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTo.Value & "') AND ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1)AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
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
        temSql = "Select Count(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue " & _
                    "FROM tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID "
        If IncomeCategory <> "All" Then
            temSql = temSql & "WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.ReturnTime)>='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpFrom.Value & "') AND ((tblIncomeReturnBill.ReturnTime)<='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTo.Value & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ") AND ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        Else
            temSql = temSql & "WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.ReturnTime)>='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpFrom.Value & "') AND ((tblIncomeReturnBill.ReturnTime)<='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTo.Value & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ") AND ((tblIncomeReturnBill.Cancelled)=0))"
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
        temSql = "SELECT Count(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue " & _
                    "FROM tblIncomeReturnBill INNER JOIN tblBHT ON tblIncomeReturnBill.BHTID = tblBHT.BHTID "
        If IncomeCategory = "InwardPayment" Then
            temSql = temSql & "Where (((tblIncomeReturnBill.ReturnTime) between '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpFrom.Value & "' AND  '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTo.Value & "'  ) AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblBHT.IsBHT)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        ElseIf IncomeCategory = "GS" Then
            temSql = temSql & "Where (((tblIncomeReturnBill.ReturnTime) between '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpFrom.Value & "' AND  '" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTo.Value & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblBHT.IsGSB)=1) AND  ((tblIncomeReturnBill.Cancelled)=0))"
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
        temSql = "Select Count(tblIncomeBill.NetTotal) AS SumOfNetTotal " & _
                    "From tblIncomeBill "
        If IncomeCategory <> "All" Then
            temSql = temSql & "WHERE (((tblIncomeBill.CancelledDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.CancelledTime)>='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpFrom.Value & "') AND ((tblIncomeBill.CancelledTime)<='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTo.Value & "') AND ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeBill.Cancelled)=1) AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "WHERE (((tblIncomeBill.CancelledDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.CancelledTime)>='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpFrom.Value & "') AND ((tblIncomeBill.CancelledTime)<='" & Format(dtpDate.Value, "dd MMMM yyyy") & " " & dtpTo.Value & "') AND ((tblIncomeBill.Cancelled)=1) AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
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

Private Sub dtpFrom_Change()
    Call FillDetails

End Sub

Private Sub dtpTo_Change()
    Call FillDetails

End Sub

Private Sub Form_Load()
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
    lblSubtopic.Caption = "Day End Counts Summery - " & Format(dtpDate.Value, "dd MMMM yyyy")
    dtpFrom.Value = GetSetting(App.EXEName, Me.Name, dtpFrom.Name, "00:00:00")
    dtpTo.Value = GetSetting(App.EXEName, Me.Name, dtpTo.Name, "00:00:00")
    
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, dtpFrom.Name, dtpFrom.Value
    SaveSetting App.EXEName, Me.Name, dtpTo.Name, dtpTo.Value
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub
