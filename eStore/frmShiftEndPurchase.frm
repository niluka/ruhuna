VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmShiftEndPurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shift End Purchase"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
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
   ScaleHeight     =   5565
   ScaleWidth      =   7680
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   4680
      TabIndex        =   46
      Top             =   4800
      Visible         =   0   'False
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   6000
      TabIndex        =   45
      Top             =   4800
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
   Begin btButtonEx.ButtonEx bttnCashSale 
      Height          =   255
      Left            =   3240
      TabIndex        =   35
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print"
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
   Begin MSDataListLib.DataCombo dtcStaff 
      Height          =   360
      Left            =   2040
      TabIndex        =   10
      Top             =   1440
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Frame Frame2 
      Caption         =   "To"
      Height          =   1215
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   3495
      Begin MSComCtl2.DTPicker dtpToTime 
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   73334786
         CurrentDate     =   39617
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   73334787
         CurrentDate     =   29224
      End
      Begin VB.Label Label4 
         Caption         =   "Time"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "From"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin MSComCtl2.DTPicker dtpFromTime 
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         Format          =   73334786
         CurrentDate     =   39617
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   73334787
         CurrentDate     =   29224
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Time"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
   End
   Begin btButtonEx.ButtonEx bttnCreditSale 
      Height          =   255
      Left            =   3240
      TabIndex        =   36
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print"
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
   Begin btButtonEx.ButtonEx bttnCardSale 
      Height          =   255
      Left            =   3240
      TabIndex        =   37
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print"
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
   Begin btButtonEx.ButtonEx bttnChequeSale 
      Height          =   255
      Left            =   3240
      TabIndex        =   38
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print"
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
   Begin btButtonEx.ButtonEx bttnSale 
      Height          =   255
      Left            =   3240
      TabIndex        =   39
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print"
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
   Begin btButtonEx.ButtonEx bttnCashRefund 
      Height          =   255
      Left            =   5160
      TabIndex        =   40
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print"
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
   Begin btButtonEx.ButtonEx bttnCreditRefund 
      Height          =   255
      Left            =   5160
      TabIndex        =   41
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print"
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
   Begin btButtonEx.ButtonEx bttnCardRefund 
      Height          =   255
      Left            =   5160
      TabIndex        =   42
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print"
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
   Begin btButtonEx.ButtonEx bttnChequeRefund 
      Height          =   255
      Left            =   5160
      TabIndex        =   43
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print"
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
   Begin btButtonEx.ButtonEx bttnRefund 
      Height          =   255
      Left            =   5160
      TabIndex        =   44
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Print"
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
   Begin VB.Label lblNetIncome 
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
      Left            =   6000
      TabIndex        =   34
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblRefund 
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
      Left            =   3840
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblSale 
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
      Left            =   1800
      TabIndex        =   32
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblCheque 
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
      Left            =   6000
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblChequeSale 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblChequeRefund 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   3840
      TabIndex        =   29
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCard 
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
      Left            =   6000
      TabIndex        =   28
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCardSale 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblCardRefund 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCredit 
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
      Left            =   6000
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCreditSale 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   1800
      TabIndex        =   24
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblCreditRefund 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCash 
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
      Left            =   6000
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblCashSale 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblCashRefund 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Total"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Net Income"
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
      Left            =   6000
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Purchase"
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Refund"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Cash"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Credit Card"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Cheque"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "Credit"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Staff User"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   3375
   End
End
Attribute VB_Name = "frmShiftEndPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim temTopic As String
    Dim temSubTopic As String
    Dim rsReport As New ADODB.Recordset
    Dim rsViewStaff As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
    Dim CSetPrinter As New cSetDfltPrinter

'
'Private Sub bttnCardRefund_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With rsReport
'                If .State = 1 Then .Close
'                temSQL = "SELECT tblReturnBill.ReturnBillID, tblReturnBill.Date AS ReturnDate, tblReturnBill.Time AS ReturnTime, tblRefundStaff.Name AS ReturnName, tblSaleBill.SaleBillID AS ReturnSaleBillID, tblSaleBill.Date AS SaleDate, tblSaleBill.Time AS SaleTime, tblSaleStaff.Name AS SaleName, tblReturnBill.NetPrice AS ReturnPrice " & _
'                            "FROM ((((tblReturnBill LEFT JOIN tblSaleBill ON tblReturnBill.SaleBillID = tblSaleBill.SaleBillID) LEFT JOIN tblPaymentMethod AS tblSalePaymentMethod ON tblReturnBill.PaymentMethodID = tblSalePaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblSaleStaff ON tblSaleBill.StaffID = tblSaleStaff.StaffID) LEFT JOIN tblStaff AS tblRefundStaff ON tblReturnBill.StaffID = tblRefundStaff.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundPaymentMethod ON tblSaleBill.PaymentMethodID = tblRefundPaymentMethod.PaymentMethodID " & _
'                            "WHERE tblRefundPaymentMethod.PaymentMethod ='Credit Card' "
'                If IsNumeric(dtcStaff.BoundText) = True Then temSQL = temSQL & " AND tblSaleBill.StaffID = " & dtcStaff.BoundText & " "
'                If dtpFromDate.Value = dtpToDate.Value Then
'                    temSQL = temSQL & " And  tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time Between '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'  "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
'                    temSQL = temSQL & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
'                    temSQL = temSQL & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblReturnBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                End If
'                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'            End With
'            With dtrCashReturnDayEnd
'                Set .DataSource = rsReport
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - All Credit Card Returns - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & format(dtptoDate.value, "dd MMMM yyyy") & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'End Sub
'
'Private Sub bttnCardSale_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With rsReport
'                If .State = 1 Then .Close
'                temSQL = "SELECT tblRefillBill.*, tblStaff.Name AS StaffUser, tblBHT.BHT, tblBHTPatient.FirstName AS BHTPatient, tblRefillBill.Date, tblStaffCustomer.Name AS StaffCustomer, tblOutPatient.FirstName AS OutPatient, tblRefillBill.Date AS BillDate " & _
'                            "FROM ((tblPatientMainDetails AS tblOutPatient RIGHT JOIN (((tblRefillBill LEFT JOIN tblStaff ON tblRefillBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN (tblBHT LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID) ON tblRefillBill.BilledBHTID = tblBHT.BHTID) ON tblOutPatient.PatientID = tblRefillBill.BilledOutPatientID) LEFT JOIN tblStaff AS tblStaffCustomer ON tblRefillBill.BilledStaffID = tblStaffCustomer.StaffID) LEFT JOIN tblSaleCategory ON tblRefillBill.SaleCategoryID = tblSaleCategory.SaleCategoryID " & _
'                            "WHERE tblPaymentMethod.PaymentMethod='Credit Card' "
'                If IsNumeric(dtcStaff.BoundText) = True Then temSQL = temSQL & " AND tblRefillBill.StaffID = " & dtcStaff.BoundText & " "
'                If dtpFromDate.Value = dtpToDate.Value Then
'                    temSQL = temSQL & " And  tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time Between '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'  "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
'                    temSQL = temSQL & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
'                    temSQL = temSQL & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblRefillBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                End If
'                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'            End With
'            With dtrCashSaleShiftEnd
'                Set .DataSource = rsReport
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - All Credit Card Sales - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'
'End Sub
'
'Private Sub bttnCashRefund_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With rsReport
'                If .State = 1 Then .Close
'                temSQL = "SELECT tblReturnBill.ReturnBillID, tblReturnBill.Date AS ReturnDate, tblReturnBill.Time AS ReturnTime, tblRefundStaff.Name AS ReturnName, tblSaleBill.SaleBillID AS ReturnSaleBillID, tblSaleBill.Date AS SaleDate, tblSaleBill.Time AS SaleTime, tblSaleStaff.Name AS SaleName, tblReturnBill.NetPrice AS ReturnPrice " & _
'                            "FROM ((((tblReturnBill LEFT JOIN tblSaleBill ON tblReturnBill.SaleBillID = tblSaleBill.SaleBillID) LEFT JOIN tblPaymentMethod AS tblSalePaymentMethod ON tblReturnBill.PaymentMethodID = tblSalePaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblSaleStaff ON tblSaleBill.StaffID = tblSaleStaff.StaffID) LEFT JOIN tblStaff AS tblRefundStaff ON tblReturnBill.StaffID = tblRefundStaff.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundPaymentMethod ON tblSaleBill.PaymentMethodID = tblRefundPaymentMethod.PaymentMethodID " & _
'                            "WHERE tblRefundPaymentMethod.PaymentMethod ='Cash' "
'                If IsNumeric(dtcStaff.BoundText) = True Then temSQL = temSQL & " AND tblSaleBill.StaffID = " & dtcStaff.BoundText & " "
'                If dtpFromDate.Value = dtpToDate.Value Then
'                    temSQL = temSQL & " And  tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time Between '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'  "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
'                    temSQL = temSQL & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
'                    temSQL = temSQL & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblReturnBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                End If
'                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'            End With
'            With dtrCashReturnDayEnd
'                Set .DataSource = rsReport
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - All Cash Returns - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'End Sub
'
'Private Sub bttnCashSale_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With rsReport
'                If .State = 1 Then .Close
'                temSQL = "SELECT tblRefillBill.*, tblStaff.Name AS StaffUser, tblBHT.BHT, tblBHTPatient.FirstName AS BHTPatient, tblRefillBill.Date, tblStaffCustomer.Name AS StaffCustomer, tblOutPatient.FirstName AS OutPatient, tblRefillBill.Date AS BillDate " & _
'                            "FROM ((tblPatientMainDetails AS tblOutPatient RIGHT JOIN (((tblRefillBill LEFT JOIN tblStaff ON tblRefillBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN (tblBHT LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID) ON tblRefillBill.BilledBHTID = tblBHT.BHTID) ON tblOutPatient.PatientID = tblRefillBill.BilledOutPatientID) LEFT JOIN tblStaff AS tblStaffCustomer ON tblRefillBill.BilledStaffID = tblStaffCustomer.StaffID) LEFT JOIN tblSaleCategory ON tblRefillBill.SaleCategoryID = tblSaleCategory.SaleCategoryID " & _
'                            "WHERE tblPaymentMethod.PaymentMethod='Cash' "
'                If IsNumeric(dtcStaff.BoundText) = True Then temSQL = temSQL & " AND tblRefillBill.StaffID = " & dtcStaff.BoundText & " "
'                If dtpFromDate.Value = dtpToDate.Value Then
'                    temSQL = temSQL & " And  tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time Between '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'  "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
'                    temSQL = temSQL & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
'                    temSQL = temSQL & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblRefillBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                End If
'                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'            End With
'            With dtrCashSaleShiftEnd
'                Set .DataSource = rsReport
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - All Cash Sales - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'End Sub
'
'Private Sub bttnChequeRefund_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With rsReport
'                If .State = 1 Then .Close
'                temSQL = "SELECT tblReturnBill.ReturnBillID, tblReturnBill.Date AS ReturnDate, tblReturnBill.Time AS ReturnTime, tblRefundStaff.Name AS ReturnName, tblSaleBill.SaleBillID AS ReturnSaleBillID, tblSaleBill.Date AS SaleDate, tblSaleBill.Time AS SaleTime, tblSaleStaff.Name AS SaleName, tblReturnBill.NetPrice AS ReturnPrice " & _
'                            "FROM ((((tblReturnBill LEFT JOIN tblSaleBill ON tblReturnBill.SaleBillID = tblSaleBill.SaleBillID) LEFT JOIN tblPaymentMethod AS tblSalePaymentMethod ON tblReturnBill.PaymentMethodID = tblSalePaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblSaleStaff ON tblSaleBill.StaffID = tblSaleStaff.StaffID) LEFT JOIN tblStaff AS tblRefundStaff ON tblReturnBill.StaffID = tblRefundStaff.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundPaymentMethod ON tblSaleBill.PaymentMethodID = tblRefundPaymentMethod.PaymentMethodID " & _
'                            "WHERE tblRefundPaymentMethod.PaymentMethod ='Cheque' "
'                If IsNumeric(dtcStaff.BoundText) = True Then temSQL = temSQL & " AND tblSaleBill.StaffID = " & dtcStaff.BoundText & " "
'                If dtpFromDate.Value = dtpToDate.Value Then
'                    temSQL = temSQL & " And  tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time Between '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'  "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
'                    temSQL = temSQL & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
'                    temSQL = temSQL & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblReturnBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                End If
'                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'            End With
'            With dtrCashReturnDayEnd
'                Set .DataSource = rsReport
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - All Cheque Returns - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'
'End Sub
'
'Private Sub bttnChequeSale_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With rsReport
'                If .State = 1 Then .Close
'                temSQL = "SELECT tblRefillBill.*, tblStaff.Name AS StaffUser, tblBHT.BHT, tblBHTPatient.FirstName AS BHTPatient, tblRefillBill.Date, tblStaffCustomer.Name AS StaffCustomer, tblOutPatient.FirstName AS OutPatient, tblRefillBill.Date AS BillDate " & _
'                            "FROM ((tblPatientMainDetails AS tblOutPatient RIGHT JOIN (((tblRefillBill LEFT JOIN tblStaff ON tblRefillBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN (tblBHT LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID) ON tblRefillBill.BilledBHTID = tblBHT.BHTID) ON tblOutPatient.PatientID = tblRefillBill.BilledOutPatientID) LEFT JOIN tblStaff AS tblStaffCustomer ON tblRefillBill.BilledStaffID = tblStaffCustomer.StaffID) LEFT JOIN tblSaleCategory ON tblRefillBill.SaleCategoryID = tblSaleCategory.SaleCategoryID " & _
'                            "WHERE tblPaymentMethod.PaymentMethod='Cheque' "
'                If IsNumeric(dtcStaff.BoundText) = True Then temSQL = temSQL & " AND tblRefillBill.StaffID = " & dtcStaff.BoundText & " "
'                If dtpFromDate.Value = dtpToDate.Value Then
'                    temSQL = temSQL & " And  tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time Between '" & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'  "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
'                    temSQL = temSQL & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & dtpFromTime.Value & "'     )    or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
'                    temSQL = temSQL & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblRefillBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                End If
'                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'            End With
'            With dtrCashSaleShiftEnd
'                Set .DataSource = rsReport
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - All Cheque Sales - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'
'End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub
'
'Private Sub bttnCreditRefund_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With rsReport
'                If .State = 1 Then .Close
'                temSQL = "SELECT tblReturnBill.ReturnBillID, tblReturnBill.Date AS ReturnDate, tblReturnBill.Time AS ReturnTime, tblRefundStaff.Name AS ReturnName, tblSaleBill.SaleBillID AS ReturnSaleBillID, tblSaleBill.Date AS SaleDate, tblSaleBill.Time AS SaleTime, tblSaleStaff.Name AS SaleName, tblReturnBill.NetPrice AS ReturnPrice " & _
'                            "FROM ((((tblReturnBill LEFT JOIN tblSaleBill ON tblReturnBill.SaleBillID = tblSaleBill.SaleBillID) LEFT JOIN tblPaymentMethod AS tblSalePaymentMethod ON tblReturnBill.PaymentMethodID = tblSalePaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblSaleStaff ON tblSaleBill.StaffID = tblSaleStaff.StaffID) LEFT JOIN tblStaff AS tblRefundStaff ON tblReturnBill.StaffID = tblRefundStaff.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundPaymentMethod ON tblSaleBill.PaymentMethodID = tblRefundPaymentMethod.PaymentMethodID " & _
'                            "WHERE tblRefundPaymentMethod.PaymentMethod ='Credit' "
'                If IsNumeric(dtcStaff.BoundText) = True Then temSQL = temSQL & " AND tblSaleBill.StaffID = " & dtcStaff.BoundText & " "
'                If dtpFromDate.Value = dtpToDate.Value Then
'                    temSQL = temSQL & " And  tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time Between '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'  "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
'                    temSQL = temSQL & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
'                    temSQL = temSQL & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblReturnBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                End If
'                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'            End With
'            With dtrCashReturnDayEnd
'                Set .DataSource = rsReport
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - All Credit Returns - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'End Sub
'
'Private Sub bttnCreditSale_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With rsReport
'                If .State = 1 Then .Close
'                temSQL = "SELECT tblRefillBill.*, tblStaff.Name AS StaffUser, tblBHT.BHT, tblBHTPatient.FirstName AS BHTPatient, tblRefillBill.Date, tblStaffCustomer.Name AS StaffCustomer, tblOutPatient.FirstName AS OutPatient, tblRefillBill.Date AS BillDate " & _
'                            "FROM ((tblPatientMainDetails AS tblOutPatient RIGHT JOIN (((tblRefillBill LEFT JOIN tblStaff ON tblRefillBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN (tblBHT LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID) ON tblRefillBill.BilledBHTID = tblBHT.BHTID) ON tblOutPatient.PatientID = tblRefillBill.BilledOutPatientID) LEFT JOIN tblStaff AS tblStaffCustomer ON tblRefillBill.BilledStaffID = tblStaffCustomer.StaffID) LEFT JOIN tblSaleCategory ON tblRefillBill.SaleCategoryID = tblSaleCategory.SaleCategoryID " & _
'                            "WHERE tblPaymentMethod.PaymentMethod='Credit' "
'                If IsNumeric(dtcStaff.BoundText) = True Then temSQL = temSQL & " AND tblRefillBill.StaffID = " & dtcStaff.BoundText & " "
'                If dtpFromDate.Value = dtpToDate.Value Then
'                    temSQL = temSQL & " And  tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time Between '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'  "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
'                    temSQL = temSQL & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
'                    temSQL = temSQL & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblRefillBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                End If
'                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'            End With
'            With dtrCashSaleShiftEnd
'                Set .DataSource = rsReport
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - All Credit Sales - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'End Sub
''
'Private Sub bttnPrint_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With dtrSummery1
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Sections("Section2").Controls.Item("lblCashSale").Caption = Me.lblCashSale.Caption
'                .Sections("Section2").Controls.Item("lblCashRefund").Caption = Me.lblCashRefund.Caption
'                .Sections("Section2").Controls.Item("lblCash").Caption = Me.lblCash.Caption
'                .Sections("Section2").Controls.Item("lblCreditSale").Caption = Me.lblCreditSale.Caption
'                .Sections("Section2").Controls.Item("lblCreditRefund").Caption = Me.lblCreditRefund.Caption
'                .Sections("Section2").Controls.Item("lblCredit").Caption = Me.lblCredit.Caption
'                .Sections("Section2").Controls.Item("lblCardSale").Caption = Me.lblCardSale.Caption
'                .Sections("Section2").Controls.Item("lblCardRefund").Caption = Me.lblCardRefund.Caption
'                .Sections("Section2").Controls.Item("lblCard").Caption = Me.lblCard.Caption
'                .Sections("Section2").Controls.Item("lblChequeSale").Caption = Me.lblChequeSale.Caption
'                .Sections("Section2").Controls.Item("lblChequeRefund").Caption = Me.lblChequeRefund.Caption
'                .Sections("Section2").Controls.Item("lblCheque").Caption = Me.lblCheque.Caption
'                .Sections("Section2").Controls.Item("lblSale").Caption = Me.lblSale.Caption
'                .Sections("Section2").Controls.Item("lblRefund").Caption = Me.lblRefund.Caption
'                .Sections("Section2").Controls.Item("lblNetIncome").Caption = Me.lblNetIncome.Caption
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'End Sub
'
'Private Sub bttnRefund_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With rsReport
'                If .State = 1 Then .Close
'                temSQL = "SELECT tblReturnBill.ReturnBillID, tblReturnBill.Date AS ReturnDate, tblReturnBill.Time AS ReturnTime, tblRefundStaff.Name AS ReturnName, tblSaleBill.SaleBillID AS ReturnSaleBillID, tblSaleBill.Date AS SaleDate, tblSaleBill.Time AS SaleTime, tblSaleStaff.Name AS SaleName, tblReturnBill.NetPrice AS ReturnPrice " & _
'                            "FROM ((((tblReturnBill LEFT JOIN tblSaleBill ON tblReturnBill.SaleBillID = tblSaleBill.SaleBillID) LEFT JOIN tblPaymentMethod AS tblSalePaymentMethod ON tblReturnBill.PaymentMethodID = tblSalePaymentMethod.PaymentMethodID) LEFT JOIN tblStaff AS tblSaleStaff ON tblSaleBill.StaffID = tblSaleStaff.StaffID) LEFT JOIN tblStaff AS tblRefundStaff ON tblReturnBill.StaffID = tblRefundStaff.StaffID) LEFT JOIN tblPaymentMethod AS tblRefundPaymentMethod ON tblSaleBill.PaymentMethodID = tblRefundPaymentMethod.PaymentMethodID " & _
'                            "WHERE tblRefundPaymentMethod.PaymentMethod is not null "
'                If IsNumeric(dtcStaff.BoundText) = True Then temSQL = temSQL & " AND tblSaleBill.StaffID = " & dtcStaff.BoundText & " "
'                If dtpFromDate.Value = dtpToDate.Value Then
'                    temSQL = temSQL & " And  tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time Between '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'  "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
'                    temSQL = temSQL & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
'                    temSQL = temSQL & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblReturnBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                End If
'                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'            End With
'            With dtrCashReturnDayEnd
'                Set .DataSource = rsReport
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - All Returns - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'
'End Sub
'
'Private Sub bttnSale_Click()
'    If CanDisplay = False Then Exit Sub
'    Dim TemResponce As Long
'    Dim RetVal As Integer
'    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'    RetVal = SelectForm(ReportPaperName, Me.hwnd)
'    Select Case RetVal
'        Case FORM_NOT_SELECTED   ' 0
'            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'        Case FORM_SELECTED   ' 1
'            With rsReport
'                If .State = 1 Then .Close
'                temSQL = "SELECT  tblRefillBill.*, tblStaff.Name AS StaffUser, tblBHT.BHT, tblBHTPatient.FirstName AS BHTPatient, tblRefillBill.Date, tblStaffCustomer.Name AS StaffCustomer, tblOutPatient.FirstName AS OutPatient, tblRefillBill.Date AS BillDate " & _
'                            "FROM ((tblPatientMainDetails AS tblOutPatient RIGHT JOIN (((tblRefillBill LEFT JOIN tblStaff ON tblRefillBill.StaffID = tblStaff.StaffID) LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID) LEFT JOIN (tblBHT LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID) ON tblRefillBill.BilledBHTID = tblBHT.BHTID) ON tblOutPatient.PatientID = tblRefillBill.BilledOutPatientID) LEFT JOIN tblStaff AS tblStaffCustomer ON tblRefillBill.BilledStaffID = tblStaffCustomer.StaffID) LEFT JOIN tblSaleCategory ON tblRefillBill.SaleCategoryID = tblSaleCategory.SaleCategoryID " & _
'                            "WHERE tblPaymentMethod.PaymentMethod  Is Not Null "

'                If IsNumeric(dtcStaff.BoundText) = True Then temSQL = temSQL & " AND tblRefillBill.StaffID = " & dtcStaff.BoundText & " "
'                If dtpFromDate.Value = dtpToDate.Value Then
'                    temSQL = temSQL & " And  tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time Between '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'  "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
'                    temSQL = temSQL & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
'                    temSQL = temSQL & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblRefillBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & "'    ) ) "
'                End If
'                .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'            End With
'            With dtrSaleShiftEnd
'                Set .DataSource = rsReport
'                .Sections("Section4").Controls.Item("lblNaME").Caption = HospitalName
'                .Sections("Section4").Controls.Item("lblContact").Caption = HospitalAddress
'                temTopic = "Shift End Summery - All Sales - " & dtcStaff.Text
'                If dtpFromDate.Value <> dtpToDate.Value Then
'                    temSubTopic = "From " & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & " onward on " & Format(dtpFromDate.Value, LongDateFormat) & " till " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                Else
'                    temSubTopic = "From " & format(dtpfromDate.value, "dd MMMM yyyy") & " " & dtpFromTime.Value & " to " & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtptoTime.Value & " on " & Format(dtpToDate.Value, LongDateFormat)
'                End If
'                .Sections("Section4").Controls.Item("lblTopic").Caption = temTopic
'                .Sections("Section4").Controls.Item("lblSubTopic").Caption = temSubTopic
'                .Caption = temTopic & " - " & temSubTopic
'                .Show
'            End With
'        Case FORM_ADDED   ' 2
'            TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'    End Select
'End Sub
'
'
Private Sub dtcStaff_Change()
    Call DisplayValues
End Sub

Private Sub dtpFromDate_Change()
    Call DisplayValues
End Sub


Private Sub dtpFromTime_Change()
    Call DisplayValues
End Sub

Private Sub dtpToDate_Change()
    Call DisplayValues
End Sub

Private Sub dtpToTime_Change()
    Call DisplayValues
End Sub

Private Sub Form_Load()
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
    dtpFromTime.Value = GetSetting(App.EXEName, "Options", "Time1", "00:00:00 AM")
    dtpToTime.Value = GetSetting(App.EXEName, "Options", "Time2", "00:00:00 AM")
    Call FillCombos
End Sub

Private Sub FillCombos()
    With rsViewStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblstaff order by listedname"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcStaff
        Set .RowSource = rsViewStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
End Sub

Private Sub ClearValues()
    lblCashSale.Caption = "0.00"
    lblCreditSale.Caption = "0.00"
    lblCardSale.Caption = "0.00"
    lblChequeSale.Caption = "0.00"
    lblCashRefund.Caption = "0.00"
    lblCreditRefund.Caption = "0.00"
    lblCardRefund.Caption = "0.00"
    lblChequeRefund.Caption = "0.00"
    lblCash.Caption = "0.00"
    lblCredit.Caption = "0.00"
    lblCard.Caption = "0.00"
    lblCheque.Caption = "0.00"
    lblSale.Caption = "0.00"
    lblRefund.Caption = "0.00"
    lblNetIncome.Caption = "0.00"
End Sub

Private Function CanDisplay() As Boolean
    CanDisplay = False
    Dim tr As Integer
    If dtpFromDate.Value > dtpToDate.Value Then
        tr = MsgBox("The time period you selected is not valied", vbCritical, "Adjust Dates")
        dtpFromDate.SetFocus
        Exit Function
    End If
'    If IsNumeric(dtcStaff.BoundText) = False Then
'        tr = MsgBox("You have not selected a staff member", vbCritical, "Staff Member")
'        dtpFromDate.SetFocus
'        Exit Function
'    End If
    CanDisplay = True
End Function

Private Sub DisplayValues()
    If CanDisplay = False Then Exit Sub
    Me.MousePointer = vbHourglass
    Call ClearValues
    Call CashSale
    Call CashRefund
    Call Cash
    Call CreditSale
    Call CreditRefund
    Call Credit
    Call CardSale
    Call CardRefund
    Call Card
    Call ChequeSale
    Call ChequeRefund
    Call Cheque
    Call Sale
    Call Refund
    Call Net
    Me.MousePointer = vbDefault
End Sub

Private Sub Sale()
    lblSale.Caption = Format(Val(lblCashSale.Caption) + Val(lblCreditSale.Caption) + Val(lblCardSale.Caption) + Val(lblChequeSale.Caption), "0.00")
End Sub

Private Sub Refund()
    lblRefund.Caption = Format(Val(lblCashRefund.Caption) + Val(lblCreditRefund.Caption) + Val(lblCardRefund.Caption) + Val(lblChequeRefund.Caption), "0.00")
End Sub

Private Sub Net()
    lblNetIncome.Caption = Format(Val(lblSale.Caption) - Val(lblRefund.Caption), "0.00")
End Sub

Private Sub CashSale()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum([tblRefillBill].[NetPrice]) AS TotalAmount " & _
                    "FROM tblRefillBill LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID " & _
                    "WHERE tblPaymentMethod.PaymentMethod='Cash' "
        If IsNumeric(dtcStaff.BoundText) = True Then temSql = temSql & " AND tblRefillBill.StaffID = " & dtcStaff.BoundText & " "
        If dtpFromDate.Value = dtpToDate.Value Then
            temSql = temSql & " And  tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtpToTime.Value & "'  "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
            temSql = temSql & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtpToTime.Value & "'    ) ) "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
            temSql = temSql & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblRefillBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtpToTime.Value & "'    ) ) "
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!TotalAmount) = False Then
                lblCashSale.Caption = !TotalAmount
            Else
                lblCashSale.Caption = "0.00"
            End If
        Else
            lblCashSale.Caption = "0.00"
        End If
    End With
End Sub

Private Sub CashRefund()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum([tblReturnBill].[NetPrice]) AS TotalAmount " & _
                    "FROM tblPaymentMethod RIGHT JOIN tblReturnBill ON tblPaymentMethod.PaymentMethodID = tblReturnBill.PaymentMethodID " & _
                    "WHERE tblPaymentMethod.PaymentMethod='Cash' "
        If IsNumeric(dtcStaff.BoundText) = True Then temSql = temSql & " AND tblReturnBill.StaffID = " & dtcStaff.BoundText & " "
        If dtpFromDate.Value = dtpToDate.Value Then
            temSql = temSql & " And  tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & " " & dtpToTime.Value & "'  "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
            temSql = temSql & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
            temSql = temSql & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblReturnBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!TotalAmount) = False Then
                lblCashRefund.Caption = !TotalAmount
            Else
                lblCashRefund.Caption = "0.00"
            End If
        Else
            lblCashRefund.Caption = "0.00"
        End If
    End With
End Sub

Private Sub CreditSale()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum([tblRefillBill].[NetPrice]) AS TotalAmount " & _
                    "FROM tblRefillBill LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID " & _
                    "WHERE tblPaymentMethod.PaymentMethod='Credit' "
        If IsNumeric(dtcStaff.BoundText) = True Then temSql = temSql & " AND tblRefillBill.StaffID = " & dtcStaff.BoundText & " "
        If dtpFromDate.Value = dtpToDate.Value Then
            temSql = temSql & " And  tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'  "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
            temSql = temSql & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
            temSql = temSql & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblRefillBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!TotalAmount) = False Then
                lblCreditSale.Caption = !TotalAmount
            Else
                lblCreditSale.Caption = "0.00"
            End If
        Else
            lblCreditSale.Caption = "0.00"
        End If
    End With
End Sub

Private Sub CreditRefund()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum([tblReturnBill].[NetPrice]) AS TotalAmount " & _
                    "FROM tblPaymentMethod RIGHT JOIN tblReturnBill ON tblPaymentMethod.PaymentMethodID = tblReturnBill.PaymentMethodID " & _
                    "WHERE tblPaymentMethod.PaymentMethod='Credit' "
        If IsNumeric(dtcStaff.BoundText) = True Then temSql = temSql & " AND tblReturnBill.StaffID = " & dtcStaff.BoundText & " "
        If dtpFromDate.Value = dtpToDate.Value Then
            temSql = temSql & " And  tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'  "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
            temSql = temSql & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
            temSql = temSql & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblReturnBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!TotalAmount) = False Then
                lblCreditRefund.Caption = !TotalAmount
            Else
                lblCreditRefund.Caption = "0.00"
            End If
        Else
            lblCreditRefund.Caption = "0.00"
        End If
    End With
End Sub


Private Sub CardSale()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum([tblRefillBill].[NetPrice]) AS TotalAmount " & _
                    "FROM tblRefillBill LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID " & _
                    "WHERE tblPaymentMethod.PaymentMethod='Credit Card' "
        If IsNumeric(dtcStaff.BoundText) = True Then temSql = temSql & " AND tblRefillBill.StaffID = " & dtcStaff.BoundText & " "
        If dtpFromDate.Value = dtpToDate.Value Then
            temSql = temSql & " And  tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'  "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
            temSql = temSql & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
            temSql = temSql & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblRefillBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!TotalAmount) = False Then
                lblCardSale.Caption = !TotalAmount
            Else
                lblCardSale.Caption = "0.00"
            End If
        Else
            lblCardSale.Caption = "0.00"
        End If
    End With
End Sub

Private Sub CardRefund()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum([tblReturnBill].[NetPrice]) AS TotalAmount " & _
                    "FROM tblPaymentMethod RIGHT JOIN tblReturnBill ON tblPaymentMethod.PaymentMethodID = tblReturnBill.PaymentMethodID " & _
                    "WHERE tblPaymentMethod.PaymentMethod='Credit Card' "
        If IsNumeric(dtcStaff.BoundText) = True Then temSql = temSql & " AND tblReturnBill.StaffID = " & dtcStaff.BoundText & " "
        If dtpFromDate.Value = dtpToDate.Value Then
            temSql = temSql & " And  tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'  "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
            temSql = temSql & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
            temSql = temSql & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblReturnBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!TotalAmount) = False Then
                lblCardRefund.Caption = !TotalAmount
            Else
                lblCardRefund.Caption = "0.00"
            End If
        Else
            lblCardRefund.Caption = "0.00"
        End If
    End With
End Sub


Private Sub ChequeSale()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum([tblRefillBill].[NetPrice]) AS TotalAmount " & _
                    "FROM tblRefillBill LEFT JOIN tblPaymentMethod ON tblRefillBill.PaymentMethodID = tblPaymentMethod.PaymentMethodID " & _
                    "WHERE tblPaymentMethod.PaymentMethod='Cheque' "
        If IsNumeric(dtcStaff.BoundText) = True Then temSql = temSql & " AND tblRefillBill.StaffID = " & dtcStaff.BoundText & " "
        If dtpFromDate.Value = dtpToDate.Value Then
            temSql = temSql & " And  tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'  "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
            temSql = temSql & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
            temSql = temSql & " And (  (     tblRefillBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblRefillBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblRefillBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblRefillBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!TotalAmount) = False Then
                lblChequeSale.Caption = !TotalAmount
            Else
                lblChequeSale.Caption = "0.00"
            End If
        Else
            lblChequeSale.Caption = "0.00"
        End If
    End With
End Sub


Private Sub ChequeRefund()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum([tblReturnBill].[NetPrice]) AS TotalAmount " & _
                    "FROM tblPaymentMethod RIGHT JOIN tblReturnBill ON tblPaymentMethod.PaymentMethodID = tblReturnBill.PaymentMethodID " & _
                    "WHERE tblPaymentMethod.PaymentMethod='Cheque' "
        If IsNumeric(dtcStaff.BoundText) = True Then temSql = temSql & " AND tblReturnBill.StaffID = " & dtcStaff.BoundText & " "
        If dtpFromDate.Value = dtpToDate.Value Then
            temSql = temSql & " And  tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time Between '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "' And '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'  "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) = 1 Then
            temSql = temSql & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        ElseIf DateDiff("d", dtpFromDate.Value, dtpToDate.Value) > 1 Then
            temSql = temSql & " And (  (     tblReturnBill.Date = '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time > '" & Format(dtpFromDate.Value, "dd MMMM yyyy") & " " & dtpFromTime.Value & "'     )    or   (    tblReturnBill.Date Between '" & Format(dtpFromDate.Value + 1, "dd MMMM yyyy") & "'  And  '" & Format(dtpToDate.Value - 1, "dd MMMM yyyy") & "'      )   or     (     tblReturnBill.Date = '" & Format(dtpToDate.Value, "dd MMMM yyyy") & "' And tblReturnBill.Time < '" & Format(dtpToDate.Value, "dd MMMM yyyy") & " " & dtpToTime.Value & "'    ) ) "
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!TotalAmount) = False Then
                lblChequeRefund.Caption = !TotalAmount
            Else
                lblChequeRefund.Caption = "0.00"
            End If
        Else
            lblChequeRefund.Caption = "0.00"
        End If
    End With
End Sub


Private Sub Cash()
    lblCash.Caption = Format(Val(lblCashSale.Caption) - Val(lblCashRefund.Caption), "0.00")
End Sub

Private Sub Credit()
    lblCredit.Caption = Format(Val(lblCreditSale.Caption) - Val(lblCreditRefund.Caption), "0.00")
End Sub

Private Sub Card()
    lblCard.Caption = Format(Val(lblCardSale.Caption) - Val(lblCardRefund.Caption), "0.00")
End Sub

Private Sub Cheque()
    lblCheque.Caption = Format(Val(lblChequeSale.Caption) - Val(lblChequeRefund.Caption), "0.00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "Options", "Time1", dtpFromTime.Value
    SaveSetting App.EXEName, "Options", "Time2", dtpToTime.Value
End Sub
