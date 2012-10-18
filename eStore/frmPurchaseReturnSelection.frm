VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPurchaseReturnSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Bill Return Selection"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
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
   ScaleHeight     =   1065
   ScaleWidth      =   6390
   Begin btButtonEx.ButtonEx bttnRefillBillIDSearch 
      Default         =   -1  'True
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   255
      Caption         =   "Search"
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
   Begin VB.TextBox txtRefillBillID 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Refill Bill ID"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmPurchaseReturnSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsRefillBill As New ADODB.Recordset
    Dim rsDistributorOrderBill As New ADODB.Recordset
    
    Dim temSql As String

Private Sub bttnRefillBillIDSearch_Click()
    Dim i As Integer
    With rsRefillBill
        If .State = 1 Then .Close
        temSql = "SELECT * from tblRefillBill where RefillBillID = " & Val(txtRefillBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
'            If !Cancelled = True Then
'                MsgBox "YOU HAVE ALREADY CANCELLED THIS GRN"
'                txtRefillBillID.SetFocus
'                SendKeys "{HOME}+{END}"
'                Exit Sub
'            ElseIf !Returned = True Then
'                MsgBox "YOU HAVE ALREADY RETURNED THIS GRN"
'                txtRefillBillID.SetFocus
'                SendKeys "{HOME}+{END}"
'                Exit Sub
'            Else
                TxRefillBillID = Val(txtRefillBillID.Text)
                Unload frmPurchaseCancellation
                Unload frmPurchaseReturn
                frmPurchaseReturn.Show
                frmPurchaseReturn.ZOrder 0
                Unload Me
'            End If
        Else
            i = MsgBox("No Such Refill Bill ID")
            txtRefillBillID.SetFocus
            SendKeys "{Home}+{END}"
            Exit Sub
        End If
    End With
End Sub
