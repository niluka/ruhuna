VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSearchMTBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Medical Test Bills"
   ClientHeight    =   1695
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
   ScaleHeight     =   1695
   ScaleWidth      =   4680
   Begin btButtonEx.ButtonEx btnSearch 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   240
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
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin VB.Label Label1 
      Caption         =   "Bill &ID"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSearchMTBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSearch_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill Where IsMedicalTestBill = 1 AND DisplayBillID = " & Val(txtID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            On Error Resume Next
            
            Unload frmMTBillReprint
            
            frmMTBillReprint.txtBillID.Text = !IncomeBillID
            frmMTBillReprint.Show
            frmMTBillReprint.ZOrder 0
            frmMTBillReprint.Top = 0
            frmMTBillReprint.Left = 0
            txtID.Text = Empty
        Else
            MsgBox "No Such Bill Number"
            txtID.SetFocus
            SendKeys "{Home}+{End}"
        End If
        .Close
    End With
End Sub

Private Sub Form_Load()
   GetCommonSettings Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub


Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnSearch_Click
    End If
End Sub
