VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPeriodPaymentSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Period Payment Summery"
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
   Begin MSDataListLib.DataCombo cmbUser 
      Height          =   360
      Left            =   2400
      TabIndex        =   7
      Top             =   2760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9480
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   8
      Top             =   120
      Width           =   10575
      Begin VB.ComboBox cmbType 
         Height          =   360
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2160
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   1200
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   78381059
         CurrentDate     =   39969
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   78381059
         CurrentDate     =   39969
      End
      Begin VB.Label Label10 
         Caption         =   "User"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "To"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Remaining Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cash"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Cheque"
         Height          =   255
         Left            =   2160
         TabIndex        =   23
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Slips"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label lblSlipsPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label lblChequePaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4200
         TabIndex        =   20
         Top             =   4440
         Width           =   2535
      End
      Begin VB.Label lblCashPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   255
         Left            =   4185
         TabIndex        =   19
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Payments"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4080
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
         TabIndex        =   17
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Bill Type"
         Height          =   255
         Left            =   360
         TabIndex        =   4
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
         TabIndex        =   16
         Top             =   3480
         Width           =   2535
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
         Left            =   4560
         TabIndex        =   15
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Total Due Payments during Period"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3480
         Width           =   4335
      End
      Begin VB.Label lblSubtopic 
         Alignment       =   2  'Center
         Caption         =   "Topic"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   10095
      End
      Begin VB.Label lblTopic 
         Alignment       =   2  'Center
         Caption         =   "Topic"
         Height          =   375
         Left            =   360
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "From"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   1200
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmPeriodPaymentSummery"
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
        Case "Medical Test Bills": TemString = "MedicalTest"
        Case "Green Sheet Bills": TemString = "GS"
        Case "All Bills": TemString = "All"
        Case "Agent Bills": TemString = "Agent"
        Case "Expence Bills": TemString = "Expence"
        Case "Roentgents Bills": TemString = "R"
        Case "Health Scheme Supplier Payments": TemString = "HSSPayment"
        Case Else:  Exit Sub
    End Select
    
    
    ToPay = CatToPay(TemString)
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
        temSql = "SELECT Sum(tblProfessionalCharges.Fee) AS SumOfFee FROM tblProfessionalCharges "
        If IncomeCategory = "All" Then
            temSql = temSql & "WHERE tblProfessionalCharges.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' AND tblProfessionalCharges.Cancelled = 0  AND tblProfessionalCharges.PaidCancelled = 0 "
        Else
            temSql = temSql & "WHERE tblProfessionalCharges.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' AND tblProfessionalCharges.Cancelled = 0  AND tblProfessionalCharges.Is" & IncomeCategory & "Bill = 1  AND tblProfessionalCharges.PaidCancelled = 0 "
        End If
        If IsNumeric(cmbUser.BoundText) = True Then
            temSql = temSql & " And UserID  = " & cmbUser.BoundText & " "
        End If
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfFee) = False Then
                CatToPay = !SumOfFee
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
            temSql = temSql & "WHERE tblProfessionalPaymentBill.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' AND tblProfessionalPaymentBill.Is" & IncomeCategory & "Bill = 1 AND tblProfessionalPaymentBill.PaymentMethodID = " & PaymentMethodID & " "
        Else
            temSql = temSql & "WHERE tblProfessionalPaymentBill.Date Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' AND tblProfessionalPaymentBill.PaymentMethodID = " & PaymentMethodID & " "
        End If
        If IsNumeric(cmbUser.BoundText) = True Then
            temSql = temSql & " And UserID  = " & cmbUser.BoundText & " "
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

Private Sub cmbUser_Change()
    Call FillDetails

End Sub

Private Sub cmbUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbUser.Text = Empty
    End If
End Sub

Private Sub dtpFrom_Change()
    Call FillDetails
End Sub

Private Sub dtpTo_Change()
    Call FillDetails
End Sub


Private Sub Form_Load()
    Call FillCombos
    Call GetSettings
    Call FillDetails
End Sub

Private Sub FillCombos()
    cmbType.AddItem "Inward Bills"
    cmbType.AddItem "Green Sheet Bills"
    cmbType.AddItem "OPD Bills"
    cmbType.AddItem "Roentgents Bills"
    cmbType.AddItem "Lab Bills"
    cmbType.AddItem "Pharmacy Bills"
    cmbType.AddItem "Medical Test Bills"
'    cmbType.AddItem "Agent Bills"
'    cmbType.AddItem "Expence Bills"
'    cmbType.AddItem "Health Scheme Supplier Payments"
'    cmbType.AddItem "Health Screening Test Bills"
    cmbType.AddItem "All Bills"
'    cmbType.AddItem "Health Scheme Supplier Payments"
'    cmbType.AddItem "Health Screening Test Bills"
    
    cmbType.Text = "All Bills"
    
    Dim MyUser As New clsFillCombos
    MyUser.FillBoolCombo cmbUser, "Staff", "Name", "IsAUser", False
    
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpFrom.Value = Date
    dtpTo.Value = Date
    lblTopic.Caption = HospitalName
    lblSubtopic.Caption = "Payment Summery - From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    'cmbUser.BoundText = UserID
End Sub
