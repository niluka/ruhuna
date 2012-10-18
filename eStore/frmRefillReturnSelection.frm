VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmRefillReturnSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase / Good Receive Cancellation"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
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
   ScaleHeight     =   1710
   ScaleWidth      =   6165
   Begin VB.TextBox txtDistributorOrderID 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   600
      Width           =   2535
   End
   Begin btButtonEx.ButtonEx bttnRefillBillIDSearch 
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Search"
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
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin btButtonEx.ButtonEx bttnDistributorOrderIDSearch 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Search"
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
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Search"
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
   Begin VB.Label Label2 
      Caption         =   "Distributor Order ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Refill Bill ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmRefillReturnSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsRefillBill As New ADODB.Recordset
    Dim rsDistributorOrderID As New ADODB.Recordset

Private Sub bttnRefillBillIDSearch_Click()
    Dim tr As Integer
    With rsRefillBill
        If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblRefillBill where RefillBillID = " & Val(txtRefillBillID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 1 Then
            tr = MsgBox("No Such Refill Bill ID. Please Recheck", vbCritical, "Wrong Refill Bill ID")
            txtRefillBillID.SetFocus
            SendKeys "{Home}+{End}"
            Exit Sub
        End If
        If !Cancelled = True Then
            tr = MsgBox("This Refill Bill is Cancelled. Please Recheck", vbCritical, "Cancelled Refill Bill ID")
            txtRefillBillID.SetFocus
            SendKeys "{Home}+{End}"
            Exit Sub
        End If
        If !Returned = True Then
            tr = MsgBox("This Refill Bill is Returned. Please Recheck", vbCritical, "Returned Refill Bill ID")
            txtRefillBillID.SetFocus
            SendKeys "{Home}+{End}"
            Exit Sub
        End If
        TxRefillBillID = Val(txtRefillBillID.Text)
        frmPurchaseReturn.Show
        frmPurchaseReturn.ZOrder 0
    End With
End Sub

Private Sub Form_Load()

End Sub
