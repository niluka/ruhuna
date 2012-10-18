VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmViewOrders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Orders"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9015
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
   ScaleHeight     =   6660
   ScaleWidth      =   9015
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   6000
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
      Left            =   6240
      TabIndex        =   7
      Top             =   6000
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
   Begin MSFlexGridLib.MSFlexGrid gridOrder 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7646
      _Version        =   393216
      FocusRect       =   2
      HighLight       =   2
      SelectionMode   =   1
   End
   Begin MSDataListLib.DataCombo cmbDistributor 
      Height          =   360
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   78249987
      CurrentDate     =   39836
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   78249987
      CurrentDate     =   39836
   End
   Begin VB.Label Label3 
      Caption         =   "Distributor"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmViewOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTem As New ADODB.Recordset
    Dim rsViewDistributor As New ADODB.Recordset
    Dim temSql As String
    
Private Sub btnPrint_Click()
    Dim temRow As Long
    Dim DIsOrderID As Long
    Dim OrderID As Long
    With GridOrder
        temRow = .Row
        DIsOrderID = Val(.TextMatrix(temRow, 2))
        OrderID = Val(.TextMatrix(temRow, 3))
    End With
    
    Dim TemDistributorId As Long
    Dim TemDistributor As String
    Dim TemAddress As String
    Dim TemFax As String
    Dim TemTel As String
    Dim TemDistributorOrderID As Long
    Dim TemResponce As Long
    Dim RetVal As Integer
    Dim tr As Integer
    
    Dim rsTemDistributor As New ADODB.Recordset
    Dim rsTemOrder As New ADODB.Recordset
    Dim CsetPrinter As New cSetDfltPrinter
    
    With rsTemDistributor
        If .State = 1 Then .Close
        temSql = "SELECT tblDistrubutor.* " & _
                    "FROM tblDistributorOrder LEFT JOIN tblDistrubutor ON tblDistributorOrder.DistributorID = tblDistrubutor.DistributorID " & _
                    "WHERE (((tblDistributorOrder.DistributorOrderID)=" & DIsOrderID & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            TemDistributorId = !DistributorID
            TemDistributor = !DistributorName
            TemAddress = !DistributorAddress
            TemTel = !DistributorTelephone
            TemFax = !DistributorFax
        End If
    End With
    TemDistributorOrderID = DIsOrderID
    With Dataenvironment1.rscmmdDistributorOrder
        If .State = 1 Then .Close
        .Source = "SELECT tblItem.Display, tblItem.AMPP, [ApprovedAmount]/[tblItem].[IssueUnitsPerPack] AS AAinPUnit, tblPackUnit.PackUnit " & _
                    " FROM ((tblPackUnit RIGHT JOIN (tblOrder LEFT JOIN tblItem ON tblOrder.ItemID = tblItem.ItemID) ON tblPackUnit.PackUnitID = tblItem.PackUnitID) LEFT JOIN tblIssueUnit ON tblItem.IssueUnitID = tblIssueUnit.IssueUnitID) LEFT JOIN tblDistrubutor ON tblOrder.ApprovedDistributorID = tblDistrubutor.DistributorID " & _
                    " WHERE (((tblOrder.OrderBillID)= " & OrderID & ") AND ((tblOrder.ApprovedAmount) > 0) AND ((tblOrder.ApprovedDistributorID)=" & TemDistributorId & ")) " & _
                    " Order by tblItem.Display"
      
        .Open
         If .RecordCount > 0 Then
             CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
             Set dtrDistributorOrdering.DataSource = Dataenvironment1.rscmmdDistributorOrder
             dtrDistributorOrdering.Sections("Section4").Controls("lblName").Caption = HospitalName
             dtrDistributorOrdering.Sections("Section4").Controls("lblContact").Caption = HospitalAddress
             dtrDistributorOrdering.Sections("Section4").Controls("lblTopic").Caption = "Order Request Forms"
             dtrDistributorOrdering.Sections("Section4").Controls("lblSUbtopic").Caption = Empty
             dtrDistributorOrdering.Sections("Section4").Controls("lblTo").Caption = TemDistributor
             dtrDistributorOrdering.Sections("Section4").Controls("lblAddress").Caption = TemAddress
             dtrDistributorOrdering.Sections("Section4").Controls("lblTel").Caption = TemTel
             dtrDistributorOrdering.Sections("Section4").Controls("lblFax").Caption = TemFax
             dtrDistributorOrdering.Sections("Section4").Controls("lblDate").Caption = Format(Date, LongDateFormat)
             dtrDistributorOrdering.Sections("Section4").Controls("lblOrderID").Caption = TemDistributorOrderID
             dtrDistributorOrdering.Sections("Section3").Controls("lblAd").Caption = LongAd
'            dtrDistributorOrdering.Sections("Section5").Controls("lblmsg1").Caption = "Please be kind enough to supply the above stocks between " & TemSTime & " on " & Format(TemSDate, LongDateFormat)
'            dtrDistributorOrdering.Sections("Section5").Controls("lblmsg2").Caption = "and " & TemETime & " on " & Format(TemEDate, LongDateFormat) & "."
             RetVal = SelectForm(ReportPaperName, Me.hwnd)
             If RetVal = FORM_SELECTED Then
                 dtrDistributorOrdering.Show
             Else
                 TemResponce = MsgBox("An Error in the report printer", vbCritical, "Printer Error")
                 Exit Sub
             End If
            End If
        End With
End Sub

Private Sub cmbDistributor_Change()
    Call FormatGrid
    Call FillGrid
End Sub


Private Sub cmbDistributor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmbDistributor.Text = Empty
    End If
End Sub

Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
End Sub


Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Call FormatGrid
    Call FillCombos
End Sub

Private Sub FillCombos()
    With rsViewDistributor
        If .State = 1 Then .Close
        temSql = "SELECT tblDistrubutor.* From tblDistrubutor ORDER BY tblDistrubutor.DistributorName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbDistributor
        Set .RowSource = rsViewDistributor
        .ListField = "DistributorName"
        .BoundColumn = "DistributorID"
    End With
End Sub

Private Sub FillGrid()

    Dim RowCount As Long
    Dim i As Integer
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(cmbDistributor.BoundText) = False Then
            temSql = "SELECT tblDistrubutor.DistributorName, tblOrderBill.ApprovedDate, tblDistributorOrder.DistributorOrderID, tblOrderBill.OrderBillID " & _
                        "FROM (tblDistributorOrder LEFT JOIN tblDistrubutor ON tblDistributorOrder.DistributorID = tblDistrubutor.DistributorID) RIGHT JOIN tblOrderBill ON tblDistributorOrder.OrderBillID = tblOrderBill.OrderBillID " & _
                        "WHERE (((tblOrderBill.ApprovedDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND ((tblOrderBill.ApprovedComplete)=1))"
        Else
            temSql = "SELECT tblDistrubutor.DistributorName, tblOrderBill.ApprovedDate, tblDistributorOrder.DistributorOrderID, tblOrderBill.OrderBillID " & _
                        "FROM (tblDistributorOrder LEFT JOIN tblDistrubutor ON tblDistributorOrder.DistributorID = tblDistrubutor.DistributorID) RIGHT JOIN tblOrderBill ON tblDistributorOrder.OrderBillID = tblOrderBill.OrderBillID " & _
                        "WHERE (((tblOrderBill.ApprovedDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND  ((tblDistributorOrder.DistributorID)= " & Val(cmbDistributor.BoundText) & ")  AND ((tblOrderBill.ApprovedComplete)=1))"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            RowCount = .RecordCount
            .MoveFirst
            GridOrder.Rows = RowCount + 1
            For i = 1 To RowCount
                GridOrder.TextMatrix(i, 0) = Format(!ApprovedDate, "dd MMM yyyy")
                GridOrder.TextMatrix(i, 1) = Format(!DistributorName, "")
                GridOrder.TextMatrix(i, 2) = Format(!DistributorOrderID, "0")
                GridOrder.TextMatrix(i, 3) = Format(!OrderBillID, "0")
                .MoveNext
            Next
        End If
        .Close

    End With
    
    
End Sub

Private Sub FormatGrid()
    With GridOrder
        .Clear
        
        .Cols = 4
        .Rows = 1
    
        .ColWidth(0) = 2000
        .ColWidth(1) = 6000
        .ColWidth(2) = 0
        .ColWidth(3) = 0
    
        .TextMatrix(0, 0) = "Date"
        .TextMatrix(0, 1) = "Distributor"
        .TextMatrix(0, 0) = "Distributor ID"
        .TextMatrix(0, 0) = "Order ID"
    
    End With
    
    
End Sub
