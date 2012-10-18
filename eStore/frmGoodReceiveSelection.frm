VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmGoodReceiveSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Good Receive Selection"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3765
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
   ScaleHeight     =   1740
   ScaleWidth      =   3765
   Begin btButtonEx.ButtonEx bttnCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Cancel"
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
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin btButtonEx.ButtonEx bttnSearch 
      Default         =   -1  'True
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
   Begin VB.Label Label1 
      Caption         =   "Enter the Distributor Order No"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmGoodReceiveSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTemOrder As New ADODB.Recordset
    Dim Temsql As String
    Public TemBillOrderID As Long
    Public TemDistributorOrderID As Long
    Public TemDistributorID As Long

Private Sub bttnCancel_Click()
Unload Me
End Sub

Private Sub bttnSearch_Click()
Dim TR As Integer

If Not IsNumeric(txtID.Text) Then
    TR = MsgBox("The Distributor Order Number you entered is not valid", vbCritical, "Wrong ID")
    txtID.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If

With rsTemOrder
    If .State = 1 Then .Close
    Temsql = "SELECT tblDistributorOrder.* " & _
                " From tblDistributorOrder " & _
                " WHERE (((tblDistributorOrder.DistributorOrderID)=" & Val(txtID.Text) & "))"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount < 1 Then
        TR = MsgBox("There is no such Distributor Order Number. Please reckeck", vbCritical, "No Such ID")
        txtID.SetFocus
        SendKeys "{Home}+{End}"
        If .State = 1 Then .Close
        Exit Sub
    ElseIf !ReceivedComplete = True Then
        TR = MsgBox("This Order Number is already processed. Please reckeck", vbCritical, "ALready Received")
        txtID.SetFocus
        SendKeys "{Home}+{End}"
        If .State = 1 Then .Close
        Exit Sub
    
    Else
        TemBillOrderID = !OrderBillID
        TemDistributorID = !DistributorID
        TemDistributorOrderID = !DistributorOrderID
        If .State = 1 Then .Close
        frmGoodReceive.Show
        Unload Me
    End If
    
End With


End Sub
