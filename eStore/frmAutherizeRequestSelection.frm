VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAutherizeRequestSelection 
   Caption         =   "Approvel Requests - Selection"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   9855
   Begin MSFlexGridLib.MSFlexGrid GridOrder 
      Height          =   3975
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7011
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   4440
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
   Begin btButtonEx.ButtonEx bttnSelect 
      Height          =   495
      Left            =   7080
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Select"
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
      Caption         =   "Requests awaiting approval"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmAutherizeRequestSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsOrderBill As New ADODB.Recordset
    Dim TemSql As String

Private Sub bttnCancel_Click()
    Unload Me
End Sub

Private Sub bttnSelect_Click()
    If GridOrder.Rows <= 1 Then Exit Sub
    If GridOrder.Row < 1 Then Exit Sub
    GridOrder.Col = 0
    If Not IsNumeric(GridOrder.Text) Then Exit Sub
    OrderBillID = Val(GridOrder.Text)
    Unload Me
    frmAutherizeRequests.Show
End Sub

Private Sub Form_Load()
    bttnSelect.Enabled = False
    Call FillOrders
End Sub

Private Sub FillOrders()

    With GridOrder
        .Cols = 4
        .Rows = 1
        
        .ColWidth(0) = 800
        .ColWidth(1) = 2600
        .ColWidth(3) = 2400
        .ColWidth(2) = .Width - (.ColWidth(0) + .ColWidth(1) + .ColWidth(3) + 150)
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Order ID"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Requested Date"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Requested By"
        
        .Col = 3
        .CellAlignment = 4
        .Text = "Generetion Mode"
        
    
    End With
    
    With rsOrderBill
        If .State = 1 Then .Close
        TemSql = "SELECT tblOrderBill.OrderBillID, tblOrderBill.RequestDate, tblStaff.Name, tblOrderBill.AutoRequest FROM tblStaff RIGHT JOIN tblOrderBill ON tblStaff.StaffID = tblOrderBill.RequestStaffID WHERE (((tblOrderBill.RequestComplete)=True) AND ((tblOrderBill.ApprovedComplete)=False)) ORDER BY tblOrderBill.RequestDate"
        .Open TemSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            bttnSelect.Enabled = True
            While .EOF = False
                GridOrder.Rows = GridOrder.Rows + 1
                GridOrder.Row = GridOrder.Rows - 1
                GridOrder.Col = 0
                GridOrder.CellAlignment = 1
                GridOrder.Text = !OrderBillID
                GridOrder.Col = 1
                GridOrder.CellAlignment = 4
                GridOrder.Text = Format(!requestdate, LongDateFormat)
                GridOrder.Col = 2
                GridOrder.CellAlignment = 1
                GridOrder.Text = !Name
                GridOrder.Col = 3
                GridOrder.CellAlignment = 4
                If !Autorequest = True Then
                    GridOrder.Text = "Auto generated order"
                Else
                    GridOrder.Text = "Manually generated order"
                End If
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub
