VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDayEndPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Day End Payments"
   ClientHeight    =   7110
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
   ScaleHeight     =   7110
   ScaleWidth      =   10770
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9480
      TabIndex        =   6
      Top             =   6480
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
      TabIndex        =   5
      Top             =   6480
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
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.ComboBox cmbType 
         Height          =   360
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2160
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   1200
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   69795843
         CurrentDate     =   39969
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   2280
         TabIndex        =   23
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   69795843
         CurrentDate     =   39969
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Remaining Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cash"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Cheque"
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Slips"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label lblSlipsPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4200
         TabIndex        =   15
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label lblChequePaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblCashPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4185
         TabIndex        =   13
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3720
         Width           =   2295
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
         Left            =   4200
         TabIndex        =   11
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Bill Type"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblToPay 
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
         Left            =   4680
         TabIndex        =   9
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblRemaining 
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
         Left            =   4680
         TabIndex        =   8
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Total Due Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblSubtopic 
         Alignment       =   2  'Center
         Caption         =   "Topic"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   10095
      End
      Begin VB.Label lblTopic 
         Alignment       =   2  'Center
         Caption         =   "Topic"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   10095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
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
         Left            =   1800
         TabIndex        =   2
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmDayEndPayments"
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
        ElseIf TypeOf MyControl Is ComboBox Then
            Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
            Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
            Printer.Print MyControl.Text
        ElseIf TypeOf MyControl Is DTPicker Then
            Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
            Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
            Printer.Print Format(MyControl.Value, "dd MMMM yyyy")
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

Private Sub dtpDate_Change()
    Call FillDetails
End Sub

Private Sub FillDetails()
    Dim TemString As String
    
    Dim ToPay As Double
    Dim Remaining As Double

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
        Case "Screen Bills": TemString = "MedicalTest"
        Case "All Bills": TemString = "All"
        Case "Green Sheet Bills": TemString = "GS"
        Case Else:  Exit Sub
    End Select
    
    ToPay = CatToPay(TemString, 1)
    CashTotalP = CatPay(TemString, 1)
    CreditTotalP = CatPay(TemString, 4)
    ChequeTotalP = CatPay(TemString, 5)
    SlipsTotalP = CatPay(TemString, 7)
    CardTotalP = CatPay(TemString, 3)
    GrandTotalP = CashTotalP + CreditTotalP + ChequeTotalP + SlipsTotalP + CardTotalP
    
    Remaining = ToPay - GrandTotalP
    
    
    lblCashPaid.Caption = Format(CashTotalP, "#,##0.00")
    lblChequePaid.Caption = Format(ChequeTotalP, "#,##0.00")
    lblSlipsPaid.Caption = Format(SlipsTotalP, "#,##0.00")
    lblPaidTotal.Caption = Format(GrandTotalP, "#,##0.00")
    lblToPay.Caption = Format(ToPay, "#,##0.00")
    lblRemaining.Caption = Format(Remaining, "#,##0.00")


End Sub

Private Function CatToPay(ByVal IncomeCategory As String) As Double
    CatToPay = 0
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeBill.NetTotal) AS SumOfNetTotal From tblIncomeBill "
        If IncomeCategory = "All" Then
            temSql = temSql & "HAVING (((tblIncomeBill.CompletedDate)=#" & lblDate.Caption & "#) AND ((tblIncomeBill.Completed)=True) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "HAVING (((tblIncomeBill.CompletedDate)=#" & lblDate.Caption & "#) AND ((tblIncomeBill.Is" & IncomeCategory & "Bill)=True)AND ((tblIncomeBill.Completed)=True) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumofNetTotal) = False Then
                CatToPay = !SumofNetTotal
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
            temSql = temSql & "HAVING (((tblProfessionalPaymentBill.Date)=#" & lblDate.Caption & "#) AND ((tblProfessionalPaymentBill.Is" & IncomeCategory & "Bill)=True) AND ((tblProfessionalPaymentBill.PaymentMethodID)=" & PaymentMethodID & "))"
        Else
            temSql = temSql & "HAVING (((tblProfessionalPaymentBill.Date)=#" & lblDate.Caption & "#) AND ((tblProfessionalPaymentBill.PaymentMethodID)=" & PaymentMethodID & "))"
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
    Call GetSettings
    Call FillDetails
End Sub

Private Sub FillCombos()
    cmbType.AddItem "OPD Bills"
    cmbType.AddItem "Green Sheet Bills"
'    cmbType.AddItem "Lab Bills"
'    cmbType.AddItem "Pharmacy Bills"
    cmbType.AddItem "Inward Bills"
    cmbType.AddItem "Medical Test Bills"
    cmbType.AddItem "All Bills"
    cmbType.Text = "All Bills"
End Sub

Private Sub GetSettings()
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
    lblTopic.Caption = HospitalName
    lblSubtopic.Caption = "Payment Summery - " & lblDate.Caption
End Sub
