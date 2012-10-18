VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmRefillSearchForCancellations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for Purchase Cancellation"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
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
   ScaleWidth      =   4680
   Begin btButtonEx.ButtonEx bttnSearch 
      Default         =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
   Begin VB.TextBox txtSaleBillID 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Bill ID"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmRefillSearchForCancellations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsSaleBill As New ADODB.Recordset
    Dim temSQL As String
    
Private Sub bttnSearch_Click()
    With rsSaleBill
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblSaleBill where SaleBillID = " & Val(txtSaleBillID.Text)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            TxSaleBillID = Val(txtSaleBillID.Text)
            Unload frmSaleReturn
            On Error Resume Next
            frmSaleCancellations.Show
            .Close
            Unload Me
            Exit Sub
        Else
            Dim tr As Integer
            tr = MsgBox("You have not entered a valis Sale Bill ID. Please recheck", vbCritical, "Sale Bill ID")
            txtSaleBillID.SetFocus
            SendKeys "{Home}+{end}"
        End If
        .Close
    End With
End Sub

