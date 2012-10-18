VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPurchaseBillSettlements 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Bill Selttlements"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10500
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
   ScaleHeight     =   7725
   ScaleWidth      =   10500
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9120
      TabIndex        =   6
      Top             =   7080
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
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   6600
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   66912259
      CurrentDate     =   39697
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   66912259
      CurrentDate     =   39697
   End
   Begin MSFlexGridLib.MSFlexGrid GridPurchase 
      Height          =   5415
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9551
      _Version        =   393216
      ScrollTrack     =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPurchaseBillSettlements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub FillGrid()
    Dim TotalValue As Double
    Dim rsPurchase As New ADODB.Recordset
    Dim i As Integer
    
    With GridPurchase
        .Clear
        .Rows = 1
        .Cols = 5
        .Visible = False
        .Row = 0
        .Col = 0
        .Text = "No"
        .CellAlignment = 4
        .Col = 1
        .Text = "Date"
        .CellAlignment = 4
        .Col = 2
        .Text = "Supplier"
        .CellAlignment = 4
        .Col = 3
        .Text = "Comments"
        .CellAlignment = 4
        .Col = 4
        .Text = "Value"
        .CellAlignment = 4
        .ColWidth(0) = 600
        .ColWidth(1) = 1400
        .ColWidth(2) = 3600
        .ColWidth(3) = 2600
        .ColWidth(4) = 1600
        
        With rsPurchase
            temSql = "SELECT tblDistrubutor.DistributorName, tblDistributorPayment.PaymentDate, tblDistributorPayment.PaymentComments, tblDistributorPayment.PaymentValue "
            temSql = temSql & "FROM tblDistributorPayment LEFT JOIN tblDistrubutor ON tblDistributorPayment.DistributorID = tblDistrubutor.DistributorID "
            temSql = temSql & "WHERE (((tblDistributorPayment.PaymentDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblDistributorPayment.Cancelled)=0)) "
            temSql = temSql & "ORDER BY tblDistributorPayment.PaymentDate "
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                .MoveLast
                .MoveFirst
                GridPurchase.Rows = .RecordCount + 1
                i = 1
                While .EOF = False
                    GridPurchase.Row = i
                    Dim n As Long
                    
                    For n = 0 To GridPurchase.Cols - 1
                        GridPurchase.Col = n
                        If i Mod 2 = 1 Then
                            GridPurchase.CellBackColor = RGB(255, 255, 200)
                        Else
                            GridPurchase.CellBackColor = RGB(255, 255, 30)
                        End If
                    Next
                    GridPurchase.TextMatrix(i, 0) = i
                    If Not IsNull(!DistributorName) Then GridPurchase.TextMatrix(i, 2) = !DistributorName
                    If Not IsNull(!PaymentDate) Then GridPurchase.TextMatrix(i, 1) = Format(!PaymentDate, "dd MMMM yyyy")
                    If Not IsNull(!PaymentComments) Then GridPurchase.TextMatrix(i, 3) = !PaymentComments
                    If Not IsNull(!PaymentValue) Then
                        GridPurchase.TextMatrix(i, 4) = Format(!PaymentValue, "0.00")
                        TotalValue = TotalValue + !PaymentValue
                    End If
                    i = i + 1
                    .MoveNext
                Wend
            End If
        End With
        .Visible = True
    End With
    txtTotal.Text = Format(TotalValue, "#,##0.00")
   
End Sub

Private Sub dtpFrom_Change()
    Call FillGrid
End Sub

Private Sub dtpTo_Change()
    Call FillGrid
End Sub

Private Sub Form_Load()
    dtpTo.Value = Date
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    Call FillGrid
End Sub
