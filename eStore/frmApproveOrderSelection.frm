VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmApproveOrderSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Approvel Requests - Selection"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12135
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
   ScaleHeight     =   7200
   ScaleWidth      =   12135
   Begin MSFlexGridLib.MSFlexGrid GridOrder 
      Height          =   6135
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10821
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   495
      Left            =   10800
      TabIndex        =   0
      Top             =   6600
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
      Left            =   9480
      TabIndex        =   1
      Top             =   6600
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
Attribute VB_Name = "frmApproveOrderSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsOrderBill As New ADODB.Recordset
    Dim temSql As String

Private Sub bttnCancel_Click()
    Unload Me
End Sub

Private Sub bttnSelect_Click()
    If gridOrder.Rows <= 1 Then Exit Sub
    If gridOrder.Row < 1 Then Exit Sub
    gridOrder.Col = 0
    If Not IsNumeric(gridOrder.Text) Then Exit Sub
    OrderBillID = Val(gridOrder.Text)
    Unload Me
    frmApproveOrdering.Show
End Sub

Private Sub Form_Load()
    GetCommonSettings Me
    bttnSelect.Enabled = False
    Call FillOrders
End Sub

Private Sub FillOrders()

    With gridOrder
        .Cols = 4
        .Rows = 1
        
        .ColWidth(0) = 800
        .ColWidth(1) = 2600
        .ColWidth(3) = 0
        .ColWidth(2) = .Width - (.ColWidth(0) + .ColWidth(1) + .ColWidth(3) + 150)
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Order ID"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Requested Date"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Distributors"
        
'        .Col = 3
'        .CellAlignment = 4
'        .Text = "Generetion Mode"
        
    
    End With
    
    With rsOrderBill
        If .State = 1 Then .Close
        temSql = "SELECT tblOrderBill.OrderBillID, tblOrderBill.RequestDate, tblStaff.Name, tblOrderBill.AutoRequest FROM tblStaff RIGHT JOIN tblOrderBill ON tblStaff.StaffID = tblOrderBill.RequestStaffID WHERE (((tblOrderBill.RequestComplete)=1) AND ((tblOrderBill.ApprovedComplete)=0)) ORDER BY tblOrderBill.RequestDate"
        
        temSql = "SELECT TOP 100 PERCENT dbo.tblOrderBill.OrderBillID, dbo.tblOrderBill.RequestDate, dbo.tblDistrubutor.DistributorName "
        temSql = temSql & "FROM         dbo.tblOrder LEFT OUTER JOIN                      dbo.tblDistrubutor ON dbo.tblOrder.RequestDistributorID = dbo.tblDistrubutor.DistributorID LEFT OUTER JOIN                      dbo.tblOrderBill ON dbo.tblOrder.OrderBillID = dbo.tblOrderBill.OrderBillID "
        temSql = temSql & "WHERE     (dbo.tblOrderBill.RequestComplete = 1) AND (dbo.tblOrderBill.ApprovedComplete = 0) "
        temSql = temSql & "GROUP BY dbo.tblOrderBill.OrderBillID, dbo.tblOrderBill.RequestDate, dbo.tblDistrubutor.DistributorName "
        temSql = temSql & " "
        
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        If .RecordCount > 0 Then
            bttnSelect.Enabled = True
            While .EOF = False
                gridOrder.Rows = gridOrder.Rows + 1
                gridOrder.Row = gridOrder.Rows - 1
                gridOrder.Col = 0
                gridOrder.CellAlignment = 1
                gridOrder.Text = !OrderBillID
                gridOrder.Col = 1
                gridOrder.CellAlignment = 4
                gridOrder.Text = Format(!requestdate, LongDateFormat)
                gridOrder.Col = 2
                gridOrder.CellAlignment = 1
                gridOrder.Text = Format(!DistributorName, "")
                gridOrder.Col = 3
                gridOrder.CellAlignment = 4
'                If !Autorequest = True Then
'                    gridOrder.Text = "Auto generated order"
'                Else
'                    gridOrder.Text = "Manually generated order"
'                End If
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub
