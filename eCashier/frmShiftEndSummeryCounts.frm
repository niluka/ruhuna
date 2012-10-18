VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShiftEndSummeryCounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift End Summery"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10770
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
   ScaleWidth      =   10770
   Begin MSDataListLib.DataCombo cmbUser 
      Height          =   360
      Left            =   2160
      TabIndex        =   47
      Top             =   2280
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.ComboBox cmbType 
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   44
      Top             =   1800
      Width           =   3615
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9480
      TabIndex        =   27
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
      Left            =   8160
      TabIndex        =   26
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
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1080
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   75431939
         CurrentDate     =   39969
      End
      Begin VB.Label lblCashier 
         Height          =   255
         Left            =   2040
         TabIndex        =   46
         Top             =   2280
         Width           =   3615
      End
      Begin VB.Label Label9 
         Caption         =   "Cashier"
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblBillType 
         Caption         =   "Label2"
         Height          =   375
         Left            =   2040
         TabIndex        =   43
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label8 
         Caption         =   "Bill Type"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   1680
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
         Left            =   2140
         TabIndex        =   41
         Top             =   4440
         Width           =   1215
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
         Left            =   3440
         TabIndex        =   40
         Top             =   4440
         Width           =   1215
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
         Left            =   4740
         TabIndex        =   39
         Top             =   4440
         Width           =   1215
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
         Left            =   6040
         TabIndex        =   38
         Top             =   4440
         Width           =   1215
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
         Left            =   7340
         TabIndex        =   37
         Top             =   4440
         Width           =   1215
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
         Left            =   8640
         TabIndex        =   36
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Net Income"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Cancellations"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label lblCashCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   2160
         TabIndex        =   33
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblCreditCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3440
         TabIndex        =   32
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblCardCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4740
         TabIndex        =   31
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblChequeCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   6040
         TabIndex        =   30
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblSlipsCancellation 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   7340
         TabIndex        =   29
         Top             =   3960
         Width           =   1215
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
         Left            =   8640
         TabIndex        =   28
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label lblSubtopic 
         Alignment       =   2  'Center
         Caption         =   "Topic"
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   600
         Width           =   10095
      End
      Begin VB.Label lblTopic 
         Alignment       =   2  'Center
         Caption         =   "Topic"
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   10095
      End
      Begin VB.Label Label42 
         Caption         =   "Total"
         Height          =   255
         Left            =   360
         TabIndex        =   23
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
         Left            =   8640
         TabIndex        =   22
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblSlipsTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   7340
         TabIndex        =   21
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblChequeTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   6040
         TabIndex        =   20
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblCardTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4740
         TabIndex        =   19
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblCreditTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3440
         TabIndex        =   18
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblCashTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   2140
         TabIndex        =   17
         Top             =   3240
         Width           =   1215
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
         Left            =   8640
         TabIndex        =   16
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label lblSlipsReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   7340
         TabIndex        =   15
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblChequeReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   6040
         TabIndex        =   14
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblCardReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4740
         TabIndex        =   13
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblCreditReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   3440
         TabIndex        =   12
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblCashReturn 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   2140
         TabIndex        =   11
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
         Height          =   255
         Left            =   9240
         TabIndex        =   10
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Returns"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Slips"
         Height          =   255
         Left            =   7800
         TabIndex        =   8
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Cheque"
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Card"
         Height          =   255
         Left            =   4920
         TabIndex        =   6
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cash"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblDate 
         Caption         =   "Label2"
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmShiftEndSummeryCounts"
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
        If Left(MyControl.Name, 3) = "lbl" Or Left(MyControl.Name, 3) = "Lab" Then
            If MyControl.Alignment = 0 Then
                Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
            ElseIf MyControl.Alignment = 1 Then
                Printer.CurrentX = (MyControl.Left / Frame1.Width * Printer.Width) + (MyControl.Width / Frame1.Width * Printer.Width) - Printer.TextWidth(MyControl.Caption)
            ElseIf MyControl.Alignment = 2 Then
                Printer.CurrentX = (MyControl.Left / Frame1.Width * Printer.Width) + ((MyControl.Width / Frame1.Width * Printer.Width) / 2) - Printer.TextWidth(MyControl.Caption)
            End If
            Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
            Printer.Print MyControl.Caption
        ElseIf Left(MyControl.Name, 3) = "cmb" Then
            Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
            Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
            Printer.Print MyControl.Text
        End If
    Next
    
    Printer.EndDoc
End Sub

Private Sub cmbType_Change()
    lblBillType.Caption = cmbType.Text
    Call FillDetails
End Sub

Private Sub cmbType_Click()
    lblBillType.Caption = cmbType.Text
    Call FillDetails
End Sub

Private Sub cmbUser_Click(Area As Integer)
    lblCashier.Caption = cmbUser.Text
    Call FillDetails
End Sub

Private Sub dtpDate_Change()
    lblDate.Caption = Format(dtpDate.Value, "dd MMMM yyyy")
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
        Case "Screen Bills": TemString = "MedicalTest"
        Case "All Bills": TemString = "All"
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
        temSql = "SELECT Count(tblIncomeBill.NetTotal) AS SumOfNetTotal From tblIncomeBill "
        If IncomeCategory = "All" Then
            temSql = temSql & "HAVING (((tblIncomeBill.CompletedDate)=#" & lblDate.Caption & "#) AND ((tblIncomeBill.UserID)=" & Val(cmbUser.BoundText) & ") AND ((tblIncomeBill.Completed)=True) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "HAVING (((tblIncomeBill.CompletedDate)=#" & lblDate.Caption & "#)  AND ((tblIncomeBill.UserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=True)AND ((tblIncomeBill.Completed)=True) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumofNetTotal) = False Then
                CatIncome = !SumofNetTotal
            End If
        End If
        .Close
    End With
End Function

Private Function CatReturn(ByVal IncomeCategory As String, PaymentMethodID As Long) As Double
    CatReturn = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Count(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue " & _
                    "FROM tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID "
        If IncomeCategory <> "All" Then
            temSql = temSql & "HAVING (((tblIncomeReturnBill.ReturnDate)=#" & lblDate.Caption & "#) AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblIncomeReturnBill.ReturnUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=True) AND ((tblIncomeReturnBill.Cancelled)=False))"
        Else
            temSql = temSql & "HAVING (((tblIncomeReturnBill.ReturnDate)=#" & lblDate.Caption & "#) AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblIncomeReturnBill.ReturnUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeReturnBill.Cancelled)=False))"
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

Private Function CatCancellation(ByVal IncomeCategory As String, PaymentMethodID As Long) As Double
    CatCancellation = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Count(tblIncomeBill.NetTotal) AS SumOfNetTotal " & _
                    "From tblIncomeBill "
        If IncomeCategory <> "All" Then
            temSql = temSql & "HAVING (((tblIncomeBill.CancelledDate)=#" & lblDate.Caption & "#)  AND ((tblIncomeBill.CancelledUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=True) AND ((tblIncomeBill.Cancelled)=True) AND ((tblIncomeBill.Completed)=True) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "HAVING (((tblIncomeBill.CancelledDate)=#" & lblDate.Caption & "#) AND ((tblIncomeBill.Cancelled)=True)  AND ((tblIncomeBill.CancelledUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Completed)=True) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumofNetTotal) = False Then
                CatCancellation = !SumofNetTotal
            End If
        End If
        .Close
    End With
End Function

Private Sub Form_Load()
    lblDate.Caption = Format(dtpDate.Value, "dd MMMM yyyy")
    Call FillCombos
    Call GetSettings
    lblDate.Caption = Format(dtpDate.Value, "dd MMMM yyyy")
    cmbType.AddItem "OPD Bills"
    cmbType.AddItem "Lab Bills"
    cmbType.AddItem "Pharmacy Bills"
    cmbType.AddItem "Inward Bills"
    cmbType.AddItem "Medical Test Bills"
    cmbType.AddItem "All Bills"
    cmbType.Text = "All Bills"
    Call FillDetails
End Sub

Private Sub FillCombos()
    Dim Cashier As New clsFillCombos
    Cashier.FillSpecificField cmbUser, "Staff", "Name", False
End Sub

Private Sub GetSettings()
    dtpDate.Value = Date
    lblTopic.Caption = HospitalName
    lblSubtopic.Caption = "Shift End Summery - " & lblDate.Caption
    cmbUser.BoundText = UserID
End Sub
