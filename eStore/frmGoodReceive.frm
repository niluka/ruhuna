VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGoodReceive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Good Receive"
   ClientHeight    =   10920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
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
   ScaleHeight     =   10920
   ScaleWidth      =   15240
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   13680
      TabIndex        =   64
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtMargin 
      Height          =   375
      Left            =   13680
      TabIndex        =   63
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   13680
      TabIndex        =   62
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRow 
      Height          =   375
      Left            =   13680
      TabIndex        =   61
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   13320
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtPPrice 
      Height          =   375
      Left            =   13320
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtSPrice 
      Height          =   375
      Left            =   13320
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtBatch 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox txtIFree 
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtIQty 
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtPFree 
      Height          =   375
      Left            =   8880
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtPQty 
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo dtcItem 
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   120
      TabIndex        =   21
      Top             =   7200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6376
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Cost"
      TabPicture(0)   =   "frmGoodReceive.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNetTotal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblGrossTotal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDiscount"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Distributor"
      TabPicture(1)   =   "frmGoodReceive.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label33"
      Tab(1).Control(1)=   "lblCity"
      Tab(1).Control(2)=   "lblFax"
      Tab(1).Control(3)=   "Label31"
      Tab(1).Control(4)=   "lblAddress"
      Tab(1).Control(5)=   "lblTelNo"
      Tab(1).Control(6)=   "lblBalance"
      Tab(1).Control(7)=   "Label27"
      Tab(1).Control(8)=   "Label26"
      Tab(1).Control(9)=   "Label25"
      Tab(1).Control(10)=   "Label20"
      Tab(1).Control(11)=   "lblDistributor"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "frmGoodReceive.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label22"
      Tab(2).Control(1)=   "Label21"
      Tab(2).Control(2)=   "Label4"
      Tab(2).Control(3)=   "Label3"
      Tab(2).Control(4)=   "dtcAStaff"
      Tab(2).Control(5)=   "dtcStaff"
      Tab(2).Control(6)=   "dtcChecked"
      Tab(2).Control(7)=   "txtInvoice"
      Tab(2).ControlCount=   8
      Begin VB.TextBox txtInvoice 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   -72720
         TabIndex        =   40
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   3960
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   1080
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo dtcChecked 
         Height          =   360
         Left            =   -72720
         TabIndex        =   41
         Top             =   1080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcStaff 
         Height          =   360
         Left            =   -72720
         TabIndex        =   42
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dtcAStaff 
         Height          =   360
         Left            =   -72720
         TabIndex        =   46
         Top             =   1560
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Autherised by"
         Height          =   255
         Left            =   -74280
         TabIndex        =   47
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Received by"
         Height          =   255
         Left            =   -74280
         TabIndex        =   45
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Checked by"
         Height          =   255
         Left            =   -74280
         TabIndex        =   44
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Invoice No."
         Height          =   255
         Left            =   -74280
         TabIndex        =   43
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "City"
         Height          =   255
         Left            =   -74640
         TabIndex        =   39
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblCity 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -73200
         TabIndex        =   38
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label lblFax 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -73200
         TabIndex        =   37
         Top             =   2400
         Width           =   4815
      End
      Begin VB.Label Label31 
         Caption         =   "Fax No"
         Height          =   255
         Left            =   -74640
         TabIndex        =   36
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -73200
         TabIndex        =   35
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label lblTelNo 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -73200
         TabIndex        =   34
         Top             =   1920
         Width           =   4815
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   -73200
         TabIndex        =   33
         Top             =   2880
         Width           =   4815
      End
      Begin VB.Label Label27 
         Caption         =   "Balance"
         Height          =   255
         Left            =   -74640
         TabIndex        =   32
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "Tel No"
         Height          =   255
         Left            =   -74640
         TabIndex        =   31
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Address"
         Height          =   255
         Left            =   -74640
         TabIndex        =   30
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "Distributor"
         Height          =   255
         Left            =   -74640
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblDistributor 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -73200
         TabIndex        =   28
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label1 
         Caption         =   "Gross Total"
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblGrossTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   375
         Left            =   3840
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Net Total"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblNetTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   375
         Left            =   3840
         TabIndex        =   24
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Discount"
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.TextBox txtDataEntry 
      Height          =   375
      Left            =   13320
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin btButtonEx.ButtonEx bttnReceive 
      Height          =   375
      Left            =   12120
      TabIndex        =   13
      Top             =   7320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Receive"
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
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   4815
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8493
      _Version        =   393216
      WordWrap        =   -1  'True
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   375
      Left            =   13680
      TabIndex        =   14
      Top             =   7320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx bttnDuplicate 
      Height          =   375
      Left            =   13680
      TabIndex        =   16
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Duplicate"
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
   Begin MSComCtl2.DTPicker dtpDOM 
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM yyyy"
      Format          =   72548355
      CurrentDate     =   39545
   End
   Begin MSComCtl2.DTPicker dtpDOE 
      Height          =   375
      Left            =   8880
      TabIndex        =   9
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM yyyy"
      Format          =   72548355
      CurrentDate     =   39545
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   13680
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add"
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
   Begin VB.Label Label13 
      Caption         =   "Value"
      Height          =   255
      Left            =   11520
      TabIndex        =   60
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Purchase Price:"
      Height          =   255
      Left            =   11520
      TabIndex        =   59
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Sale Price"
      Height          =   255
      Left            =   11520
      TabIndex        =   58
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Batch"
      Height          =   375
      Left            =   120
      TabIndex        =   57
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "DOM"
      Height          =   375
      Left            =   5160
      TabIndex        =   56
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "DOE"
      Height          =   375
      Left            =   8040
      TabIndex        =   55
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblIUnit1 
      Height          =   255
      Left            =   10080
      TabIndex        =   54
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblPUnit1 
      Height          =   255
      Left            =   10080
      TabIndex        =   53
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblIUnit 
      Height          =   255
      Left            =   6960
      TabIndex        =   52
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblPunit 
      Height          =   255
      Left            =   6960
      TabIndex        =   51
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Free:"
      Height          =   255
      Left            =   8160
      TabIndex        =   50
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Qty :"
      Height          =   255
      Left            =   5160
      TabIndex        =   49
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Item :"
      Height          =   375
      Left            =   120
      TabIndex        =   48
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblOrderID 
      Height          =   375
      Left            =   8160
      TabIndex        =   20
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label28 
      Caption         =   "Order ID"
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblThisDistributor 
      Height          =   375
      Left            =   1560
      TabIndex        =   18
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label23 
      Caption         =   "Distributor"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmGoodReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
    Dim TemOrderBillID As Long
    Dim temRefillBillID As Long
    Dim TemDistributorId As Long
    Dim TemDistributorOrderID As Long
    Dim EditingData As Boolean
    Dim TemContent(22) As String
    Dim CurrentRow As Integer
    Dim TemCellContent As String
    
    Dim NewItem As New Item
    
    Dim rsTemOrder As New ADODB.Recordset
    Dim rsTemPrice As New ADODB.Recordset
    Dim rsTemDistributor As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    Dim rsStaff As New ADODB.Recordset
    Dim rsTemOrderBill As New ADODB.Recordset
    Dim rsTemDistributorOrder As New ADODB.Recordset
    Dim rsTemRefill As New ADODB.Recordset
    Dim rsTemRefillBill As New ADODB.Recordset
    Dim rsSPrice As New ADODB.Recordset
    Dim rsPPrice As New ADODB.Recordset
    Dim rsViewItems As New ADODB.Recordset
    
    Public DOM As String
    Public DOE As String
    Public Batch As String
    Public Item As String
    
    Dim CsetPrinter As New cSetDfltPrinter
    
Private Sub bttnAdd_Click()
    If CanAdd = False Then Exit Sub
    With GridItem
        .TextMatrix(Val(txtRow.Text), 15) = Val(txtIQty.Text)
        .TextMatrix(Val(txtRow.Text), 16) = Val(txtIFree.Text)
        .TextMatrix(Val(txtRow.Text), 5) = Val(txtPQty.Text)
        .TextMatrix(Val(txtRow.Text), 7) = Val(txtPFree.Text)
        .TextMatrix(Val(txtRow.Text), 9) = txtBatch.Text
        .TextMatrix(Val(txtRow.Text), 10) = Val(txtPPrice.Text) / NewItem.IssueUnitsPerPack
        .TextMatrix(Val(txtRow.Text), 11) = Val(txtSPrice.Text)
        .TextMatrix(Val(txtRow.Text), 18) = Format(Val(txtValue.Text), "#,#0.00")
        .TextMatrix(Val(txtRow.Text), 19) = Val(txtValue.Text)
        .TextMatrix(Val(txtRow.Text), 20) = dtpDOM.Value
        .TextMatrix(Val(txtRow.Text), 21) = dtpDOE.Value
    End With
    Call CalculateTotal
    Call ClearAddValues
End Sub

Private Sub ClearAddValues()
    txtBatch.Text = Empty
    txtIFree.Text = Empty
    txtIQty.Text = Empty
    txtPQty.Text = Empty
    txtPFree.Text = Empty
    dtpDOE.Value = Date
    dtpDOM.Value = Date
    txtRow.Text = Empty
    txtPPrice.Text = Empty
    txtSPrice.Text = Empty
    txtValue.Text = Empty
End Sub

Private Function CanAdd() As Boolean
    CanAdd = False
    Dim tr As Integer
    If IsNumeric(dtcItem.BoundText) = False Then
        MsgBox "No Item"
        dtcItem.SetFocus
        Exit Function
    End If
    If IsNumeric(txtPQty.Text) = False Then
        MsgBox "Wrong Pack Quentity"
        txtPQty.SetFocus
        Exit Function
    End If
    If IsNumeric(txtIQty.Text) = False Then
        MsgBox "Wrong Issue Quentity"
        txtIQty.SetFocus
        Exit Function
    End If
    If Trim(txtBatch.Text) = Empty Then
        MsgBox "No Batch number"
        txtBatch.SetFocus
        Exit Function
    End If
    If IsNumeric(txtPPrice.Text) = False Then
        MsgBox "Wrong Purchase Price"
        txtPPrice.SetFocus
        Exit Function
    End If
    If IsNumeric(txtSPrice.Text) = False Then
        MsgBox "Wrong Sale Price"
        txtSPrice.SetFocus
        Exit Function
    End If
    CanAdd = True
End Function


Private Sub bttnCancel_Click()
    Unload Me
End Sub

Private Sub bttnDuplicate_Click()
    With GridItem
        If .Rows < 2 Or .Row < 1 Then Exit Sub
        EditingData = False
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Col = i
            TemContent(i) = .Text
        Next
        .AddItem "", .Row + 1
        .Row = .Row + 1
        For i = 1 To .Cols - 1
            .Col = i
            .Text = TemContent(i)
        Next
    End With
    EditingData = True
End Sub


Private Sub bttnReceive_Click()
    If CanReceive = False Then Exit Sub
    Dim tr As Integer
    Dim i As Integer
    Dim DiscountPercent As Double
    With rsTemRefillBill
        If .State = 1 Then .Close
        temSql = "SELECT * from tblRefillBill"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !DistributorID = TemDistributorId
        !StoreID = UserStoreID
        !StaffID = UserID
        If IsNumeric(dtcChecked.BoundText) = True Then
            !CheckedStaffID = dtcChecked.BoundText
        End If
        !Price = Val(lblGrossTotal.Caption)
        !Discount = Val(txtDiscount.Text)
        DiscountPercent = (Val(txtDiscount.Text) / Val(lblGrossTotal.Caption)) * 100
        !DiscountPercent = DiscountPercent
        !NetPrice = Val(lblNetTotal.Caption)
        !Date = Date
        !Time = Now
        !PaymentMethodID = 4
        !FullyPaid = False
        !OrderBillID = TemOrderBillID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        temRefillBillID = !NewID
        .Close
    End With
    With GridItem
        For i = 1 To .Rows - 1
            If rsTemOrder.State = 1 Then rsTemOrder.Close
            temSql = "SELECT * from tblorder where orderid = " & Val(.TextMatrix(i, 13))
            rsTemOrder.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If rsTemRefill.State = 1 Then rsTemRefill.Close
            temSql = "SELECT * FROM tblRefill"
            rsTemRefill.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If rsTemOrder.RecordCount > 0 Then
                rsTemRefill.AddNew
                rsTemRefill!ItemID = Val(.TextMatrix(i, 2))
                rsTemRefill!StoreID = UserStoreID
                rsTemRefill!Date = Date
                rsTemRefill!Time = Now
                rsTemRefill!StaffID = UserID
                rsTemRefill!DistributorID = TemDistributorId
                rsTemRefill!Price = Val(.TextMatrix(i, 19))
                
                rsTemRefill!SPrice = Val(.TextMatrix(i, 11))
                rsTemRefill!PPrice = Val(.TextMatrix(i, 10))
                rsTemRefill!PackPPrice = Val(.TextMatrix(i, 10) * Val(.TextMatrix(i, 17)))
                
                rsTemRefill!DiscountPercent = DiscountPercent
                rsTemRefill!NetPrice = (Val(.TextMatrix(i, 19))) - (Val(.TextMatrix(i, 19)) * DiscountPercent / 100)
                rsTemRefill!RefillBillID = temRefillBillID
                rsTemRefill!OrderBillID = TemOrderBillID
                rsTemRefill!OrderID = Val(.TextMatrix(i, 13))
                rsTemRefill!Amount = Val(.TextMatrix(i, 15))
                rsTemRefill!FreeAmount = Val(.TextMatrix(i, 16))
                rsTemRefill!DOM = .TextMatrix(i, 20)
                rsTemRefill!DOE = .TextMatrix(i, 21)
                rsTemRefill!LastPPrice = Val(.TextMatrix(i, 22))
                rsTemRefill!LastSPrice = Val(.TextMatrix(i, 23))
                
                If .TextMatrix(i, 0) <> Empty Then
                    rsTemOrder!ReceivedAmount = (Val(.TextMatrix(i, 15))) ' * Val(.TextMatrix(i, 17)))
                    rsTemOrder!ReceivedFreeAmount = (Val(.TextMatrix(i, 16))) ' * Val(.TextMatrix(i, 17)))
                Else
                    rsTemOrder!ReceivedAmount = rsTemOrder!ReceivedAmount + (Val(.TextMatrix(i, 15))) ' * Val(.TextMatrix(i, 17)))
                    rsTemOrder!ReceivedFreeAmount = rsTemOrder!ReceivedFreeAmount + (Val(.TextMatrix(i, 16))) ' * Val(.TextMatrix(i, 17)))
                End If
                rsTemOrder!ReceivedDate = Date
                rsTemOrder!ReceivedTime = Now
                rsTemOrder!ReceivedSTaffID = UserID
                rsTemOrder!ReceivedStoreID = UserStoreID
                If IsNumeric(dtcChecked.BoundText) Then
                    rsTemOrder!ReceivedCheckedStaffID = dtcChecked.BoundText
                    rsTemRefill!CheckedStaffID = dtcChecked.BoundText
                End If
                rsTemOrder!ReceivedDistributorID = TemDistributorId
                rsTemOrder!ReceivedInvoice = txtInvoice.Text
                
                
                
                rsTemOrder!ReceivedComplete = True
                
                rsTemOrder.Update
                Dim ThisBatch As Long
                ThisBatch = BatchExist(.TextMatrix(i, 9), Val(.TextMatrix(i, 2)))
                If ThisBatch <> 0 Then
                    rsTemRefill!BatchID = ThisBatch
                    If AddToStock(ThisBatch, UserStoreID, Val(.TextMatrix(i, 15)) + Val(.TextMatrix(i, 15))) = False Then
                        MsgBox "Error"
                        Exit For
                    End If
                Else
                    ThisBatch = AddBatch(.TextMatrix(i, 9), Val(.TextMatrix(i, 2)), .TextMatrix(i, 20), .TextMatrix(i, 21))
                    rsTemRefill!BatchID = ThisBatch
                    
                    If AddToStock(ThisBatch, UserStoreID, Val(.TextMatrix(i, 15)) + Val(.TextMatrix(i, 16))) = False Then
                        MsgBox "Error"
                        Exit For
                    End If
                End If
                rsTemRefill.Update
                rsTemOrder.Close
                
                If rsSPrice.State = 1 Then rsSPrice.Close
                temSql = "SELECT tblSalePrice.ItemID, tblSalePrice.SPrice, tblSalePrice.SetDate, tblSalePrice.SetTime, tblSalePrice.StaffID FROM tblSalePrice "
                rsSPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                rsSPrice.AddNew
                rsSPrice!ItemID = Val(.TextMatrix(i, 2))
                rsSPrice!SPrice = Val(.TextMatrix(i, 11))
                rsSPrice!setdate = Date
                rsSPrice!SetTime = Now
                rsSPrice!StaffID = UserID
                rsSPrice.Update
                rsSPrice.Close
                
                If rsSPrice.State = 1 Then rsSPrice.Close
                temSql = "SELECT * FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
                rsSPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If rsSPrice.RecordCount < 1 Then
                    rsSPrice.AddNew
                    rsSPrice!ItemID = Val(.TextMatrix(i, 2))
                    rsSPrice!SPrice = Val(.TextMatrix(i, 11))
                    rsSPrice!setdate = Date
                    rsSPrice!SetTime = Now
                    rsSPrice!StaffID = UserID
                    rsSPrice.Update
                ElseIf rsSPrice.RecordCount = 1 Then
                    rsSPrice!SPrice = Val(.TextMatrix(i, 11))
                    rsSPrice!setdate = Date
                    rsSPrice!SetTime = Now
                    rsSPrice!StaffID = UserID
                    rsSPrice.Update
                Else
                    If rsSPrice.State = 1 Then rsSPrice.Close
                    temSql = "Delete FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
                    rsSPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If rsSPrice.State = 1 Then rsSPrice.Close
                    temSql = "SELECT * FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
                    rsSPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    rsSPrice.AddNew
                    rsSPrice!ItemID = Val(.TextMatrix(i, 2))
                    rsSPrice!SPrice = Val(.TextMatrix(i, 11))
                    rsSPrice!setdate = Date
                    rsSPrice!SetTime = Now
                    rsSPrice!StaffID = UserID
                    rsSPrice.Update
                End If
                rsSPrice.Close
                
                If rsPPrice.State = 1 Then rsPPrice.Close
                temSql = "SELECT tblPurchasePrice.ItemID, tblPurchasePrice.PPrice, tblPurchasePrice.SetDate, tblPurchasePrice.SetTime, tblPurchasePrice.StaffID FROM tblPurchasePrice"
                rsPPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                rsPPrice.AddNew
                rsPPrice!ItemID = Val(.TextMatrix(i, 2))
                rsPPrice!PPrice = Val(.TextMatrix(i, 10))
                rsPPrice!setdate = Date
                rsPPrice!SetTime = Now
                rsPPrice!StaffID = UserID
                rsPPrice.Update
                rsPPrice.Close
                
                If rsPPrice.State = 1 Then rsPPrice.Close
                temSql = "SELECT * FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2))
                rsPPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If rsPPrice.RecordCount < 1 Then
                    rsPPrice.AddNew
                    rsPPrice!ItemID = Val(.TextMatrix(i, 2))
                    rsPPrice!PPrice = Val(.TextMatrix(i, 10)) / NewItem.IssueUnitsPerPack
                    rsPPrice!setdate = Date
                    rsPPrice!SetTime = Now
                    rsPPrice!StaffID = UserID
                    rsPPrice.Update
                ElseIf rsPPrice.RecordCount = 1 Then
                    rsPPrice!PPrice = Val(.TextMatrix(i, 10)) / NewItem.IssueUnitsPerPack
                    rsPPrice!setdate = Date
                    rsPPrice!SetTime = Now
                    rsPPrice!StaffID = UserID
                    rsPPrice.Update
                Else
                    If rsPPrice.State = 1 Then rsPPrice.Close
                    temSql = "Delete FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2))
                    rsPPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    If rsPPrice.State = 1 Then rsPPrice.Close
                    temSql = "SELECT * FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2))
                    rsPPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    rsPPrice.AddNew
                    rsPPrice!ItemID = Val(.TextMatrix(i, 2))
                    rsPPrice!PPrice = Val(.TextMatrix(i, 10)) / NewItem.IssueUnitsPerPack
                    rsPPrice!setdate = Date
                    rsPPrice!SetTime = Now
                    rsPPrice!StaffID = UserID
                    rsPPrice.Update
                End If
                rsPPrice.Close
                
            End If
        Next
    End With
    With rsTemOrderBill
        If .State = 1 Then .Close
        temSql = "SELECT * from tblOrderBill where ORderBillID = " & TemOrderBillID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
                !ReceivedDate = Date
                !ReceivedTime = Now
                !ReceivedSTaffID = UserID
                If IsNumeric(dtcChecked.BoundText) Then
                    !ReceivedCheckedStaffID = dtcChecked.BoundText
                End If
                !ReceivedDistributorID = TemDistributorId
                !ReceivedInvoice = txtInvoice.Text
                !ReceivedComplete = True
                .Update
        End If
        .Close
    End With
    With rsTemDistributorOrder
        If .State = 1 Then .Close
        temSql = "SELECT * from tblDistributorOrder where distributorOrderID = " & TemDistributorOrderID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !ReceivedComplete = True
            .Update
        End If
        .Close
    End With
    Call PrintGoodReceive
    tr = MsgBox("The Goods Received and added to stocks successfully", vbInformation, "Success")
    GridItem.Clear
    GridItem.Rows = 1
    GridItem.Cols = 1
    Unload Me
End Sub

Private Sub PrintGoodReceive()
    Dim RetVal As Integer
    Dim TemResponce     As Integer
     With Dataenvironment1.rscmmdGoodReceive
         If .State = 1 Then .Close
         .Source = "SELECT tblOrder.OrderID, tblItem.ItemID, tblItem.Display, tblItem.AMPP, [tblOrder].[ReceivedAmount]/[tblItem].[IssueUnitsPerPack] AS AmountInPackUnit    , [tblOrder].[ReceivedFreeAmount]/[tblItem].[IssueUnitsPerPack] AS FreeAmountInPackUnit   , tblOrder.ReceivedAmount, tblItem.IssueUnitsPerPack, tblPackUnit.PackUnit, tblIssueUnit.IssueUnit, tblItemCategory.SalesMargin, tblOrder.ReceivedFreeAmount, tblRefill.* " & _
                     " FROM tblRefill RIGHT JOIN ((((tblItem LEFT JOIN tblIssueUnit ON tblItem.IssueUnitID = tblIssueUnit.IssueUnitID) LEFT JOIN tblPackUnit ON tblItem.PackUnitID = tblPackUnit.PackUnitID) RIGHT JOIN tblOrder ON tblItem.ItemID = tblOrder.ItemID) LEFT JOIN tblItemCategory ON tblItem.ItemCategoryID = tblItemCategory.ItemCategoryID) ON tblRefill.OrderID = tblOrder.OrderID " & _
                     " WHERE (((tblOrder.OrderBillID)= " & TemOrderBillID & ") AND ((tblOrder.ReceivedAmount) > 0) AND ((tblOrder.ApprovedDistributorID)=" & TemDistributorId & "))"
         .Open
         If .RecordCount > 0 Then
        CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
            With dtrDistributorGoodReceive
                Set .DataSource = Dataenvironment1.rscmmdGoodReceive
                .Sections("Section4").Controls("lblName").Caption = HospitalName
                .Sections("Section4").Controls("lblContact").Caption = HospitalAddress
                .Sections("Section4").Controls("lblTopic").Caption = "Good Receive Note"
                .Sections("Section4").Controls("lblSUbtopic").Caption = Empty
                .Sections("Section4").Controls("lblTo").Caption = lblDistributor.Caption
                .Sections("Section4").Controls("lblAddress").Caption = lblAddress.Caption
                .Sections("Section4").Controls("lblTel").Caption = lblTelNo.Caption
                .Sections("Section4").Controls("lblFax").Caption = lblFax.Caption
                .Sections("Section4").Controls("lblDate").Caption = Format(Date, LongDateFormat)
                .Sections("Section4").Controls("lblOrderID").Caption = TemOrderBillID
                .Sections("Section4").Controls("lblDistributorOrderID").Caption = TemDistributorOrderID
                .Sections("Section4").Controls("lblRefillID").Caption = temRefillBillID
                .Sections("Section5").Controls("lblPayee").Caption = lblDistributor.Caption
                .Sections("Section5").Controls("lblTotalAmount").Caption = lblGrossTotal.Caption
                .Sections("Section5").Controls("lblDiscount").Caption = txtDiscount.Text
                .Sections("Section5").Controls("lblNetTotal").Caption = lblNetTotal.Caption
                .Sections("Section5").Controls("lblCheckedBy").Caption = dtcChecked.Text
                .Sections("Section5").Controls("lblAutherisedBy").Caption = dtcAStaff.Text
                .Sections("Section5").Controls("lblreceivedBy").Caption = dtcStaff.Text
                RetVal = SelectForm(ReportPaperName, Me.hwnd)
                If RetVal = FORM_SELECTED Then
                    .Show
                Else
                    TemResponce = MsgBox("An Error in the report printer", vbCritical, "Printer Error")
                    Exit Sub
                End If
            End With
         End If
    End With
End Sub

Private Function CanReceive() As Boolean
    Dim i As Integer
    Dim tr As Integer
    CanReceive = False
    With GridItem
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 9)) = Empty Then
                tr = MsgBox("You have not entered the batch number", vbCritical, "Batch?")
                .Row = i
                GridItem_Click
                txtBatch.SetFocus
                Exit Function
            End If
            If IsNumeric(.TextMatrix(i, 5)) = False Or Val(.TextMatrix(i, 5)) <= 0 Then
                tr = MsgBox("You have not entered the purchase quentity", vbCritical, "Batch?")
                .Row = i
                GridItem_Click
                txtPQty.SetFocus
                Exit Function
            End If
            If IsNumeric(.TextMatrix(i, 10)) = False Or Val(.TextMatrix(i, 10)) <= 0 Then
                tr = MsgBox("You have not entered the purchase price", vbCritical, "Batch?")
                .Row = i
                GridItem_Click
                txtPPrice.SetFocus
                Exit Function
            End If
            If IsNumeric(.TextMatrix(i, 11)) = False Or Val(.TextMatrix(i, 11)) <= 0 Then
                tr = MsgBox("You have not entered the sale price", vbCritical, "Batch?")
                .Row = i
                GridItem_Click
                txtSPrice.SetFocus
                Exit Function
            End If
            If Val(.TextMatrix(i, 11)) <= Val(.TextMatrix(i, 10)) Then
                tr = MsgBox("The sale price is less than the purchase price", vbCritical, "Wrong Sales Price")
                .Row = i
                GridItem_Click
                txtSPrice.SetFocus
                Exit Function
            End If
        
    '   0   No
    '   1   Item
    '   2   ItemID
    '   3   RequestedQuentity
    '   4   PUnit
    '   5   PurchaseQuentity
    '   6   PUnit
    '   7   FreeQuentity
    '   8   PUnit
    '   9   Batch
    '   10  Purchase Price
    '   11  Sales Price
    '   12  Sales Margin
    '   13  OrderID
    '   14  IRequested
    '   15  IReceived
    '   16  IFreeReceived
    '   17  IUnitsPerPack
    '   18  Display Price
    '   19  Actual Price
    '   20  DOM
    '   21  DOE
    '   22  Last Purchase Price
    '   23  Lasr Sale Price
        
        Next i
    End With
    CanReceive = True
End Function



Private Sub Form_Load()
    TemOrderBillID = frmGoodReceiveSelection.TemBillOrderID
    TemDistributorId = frmGoodReceiveSelection.TemDistributorId
    TemDistributorOrderID = frmGoodReceiveSelection.TemDistributorOrderID
    Call FillCombos
    Call FillGrid
    Call FindDetails
End Sub

Private Sub FindDetails()
    DistributorDetails (TemDistributorId)
    lblThisDistributor.Caption = lblDistributor.Caption
    lblOrderID.Caption = TemDistributorOrderID
End Sub

Private Sub FillCombos()
    With rsStaff
        If .State = 1 Then .Close
        temSql = "SELECT * from tblstaff order by listedname"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcStaff
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With dtcChecked
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With dtcAStaff
        Set .RowSource = rsStaff
        .ListField = "ListedName"
        .BoundColumn = "StaffID"
    End With
    With rsViewItems
        If .State = 1 Then .Close
        temSql = "SELECT * from tblItem order by Display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcItem
        Set .RowSource = rsViewItems
        .ListField = "Display"
        .BoundColumn = "ItemID"
    End With
End Sub

Private Sub FillGrid()
    EditingData = False
    With GridItem
        .Cols = 24
        .Rows = 1
        .Row = 0
        .Col = 0
        .FixedCols = 0
        
        .RowHeight(0) = .RowHeight(0) * 3
        
        Dim i As Integer
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            Select Case i
                Case 0:     .Text = "No"
                            .ColWidth(i) = 400
                Case 1:     .Text = "Item"
                            .ColWidth(i) = 3600
                Case 3:     .Text = "Requested"
                            .ColWidth(i) = 900
                Case 4:     .Text = "Unit"
                            .ColWidth(i) = 900
                Case 5:     .Text = "Supplied"
                            .ColWidth(i) = 900
                Case 6:     .Text = "Pack Unit"
                            .ColWidth(i) = 900
                Case 7:     .Text = "Free"
                            .ColWidth(i) = 900
                Case 8:     .Text = "Pack Unit"
                            .ColWidth(i) = 900
                Case 9:     .Text = "Batch"
                            .ColWidth(i) = 900
                Case 10:     .Text = "Pruchase Price Per Pack"
                            .ColWidth(i) = 900
                Case 11:     .Text = "Slaes Price Per Unit Sale"
                            .ColWidth(i) = 900
                Case 18:    .ColWidth(i) = 1200
                            .Text = "Total Pruchase Value"
                Case 15: .ColWidth(i) = 600
                
                Case Else:  .ColWidth(i) = 1
            End Select
        Next i
    
    End With
    With rsTemOrder
        If .State = 1 Then .Close
        temSql = "SELECT tblOrder.OrderID, tblItem.ItemID, tblItem.Display, tblItem.AMPP, [tblOrder].[ApprovedAmount]/[tblItem].[IssueUnitsPerPack] AS AmountInPackUnit, tblOrder.ApprovedAmount, tblItem.IssueUnitsPerPack, tblPackUnit.PackUnit, tblIssueUnit.IssueUnit, tblItemCategory.SalesMargin, tblCurrentPurchasePrice.PPrice, tblCurrentSalePrice.SPrice " & _
                    " FROM (((((tblItem LEFT JOIN tblIssueUnit ON tblItem.IssueUnitID = tblIssueUnit.IssueUnitID) LEFT JOIN tblPackUnit ON tblItem.PackUnitID = tblPackUnit.PackUnitID) RIGHT JOIN tblOrder ON tblItem.ItemID = tblOrder.ItemID) LEFT JOIN tblItemCategory ON tblItem.ItemCategoryID = tblItemCategory.ItemCategoryID) LEFT JOIN tblCurrentPurchasePrice ON tblItem.ItemID = tblCurrentPurchasePrice.ItemID) LEFT JOIN tblCurrentSalePrice ON tblItem.ItemID = tblCurrentSalePrice.ItemID " & _
                    " WHERE (((tblOrder.OrderBillID)=" & TemOrderBillID & ") AND ((tblOrder.ApprovedDistributorID)=" & TemDistributorId & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                GridItem.Rows = GridItem.Rows + 1
                GridItem.Row = GridItem.Rows - 1
                
                GridItem.Col = 0
                GridItem.CellAlignment = 7
                GridItem.Text = GridItem.Row
                
                GridItem.Col = 1
                GridItem.CellAlignment = 1
                GridItem.Text = ![AMPP]
                
                GridItem.Col = 2
                GridItem.Text = ![ItemID]
                
                GridItem.Col = 3
                GridItem.CellAlignment = 7
                GridItem.Text = ![AmountInPackUnit]
                
                GridItem.Col = 4
                GridItem.CellAlignment = 1
                GridItem.Text = ![PackUnit]
                GridItem.Col = 6
                GridItem.CellAlignment = 1
                GridItem.Text = ![PackUnit]
                GridItem.Col = 8
                GridItem.CellAlignment = 1
                GridItem.Text = ![PackUnit]
                
                GridItem.Col = 5
                GridItem.CellAlignment = 7
                GridItem.Text = ![AmountInPackUnit]
                
                GridItem.Col = 7
                GridItem.CellAlignment = 7
                GridItem.Text = 0
                
                GridItem.Col = 9
                GridItem.CellAlignment = 4
                GridItem.Text = Empty
                
                GridItem.Col = 10
                GridItem.CellAlignment = 7
                GridItem.Text = "0.00"
                
                GridItem.Col = 11
                GridItem.CellAlignment = 7
                GridItem.Text = "0.00"
                
                GridItem.Col = 12
                GridItem.Text = ![SalesMargin]
                
                GridItem.Col = 13
                GridItem.Text = ![OrderID]
                
                GridItem.Col = 14
                GridItem.Text = ![ApprovedAmount]
                
                GridItem.Col = 15
                GridItem.Text = ![ApprovedAmount]
                    
                GridItem.Col = 16
                GridItem.Text = Empty
                
                GridItem.Col = 17
                GridItem.Text = ![IssueUnitsPerPack]
                
                GridItem.Col = 22
                If Not IsNull(!PPrice) Then GridItem.Text = ![PPrice]
    
                GridItem.Col = 23
                If Not IsNull(!SPrice) Then GridItem.Text = ![SPrice]
    
    
    '   0   No
    '   1   Item
    '   2   ItemID
    '   3   RequestedQuentity
    '   4   PUnit
    '   5   PurchaseQuentity
    '   6   PUnit
    '   7   FreeQuentity
    '   8   PUnit
    '   9   Batch
    '   10  Purchase Price
    '   11  Sales Price
    '   12  Sales Margin
    '   13  OrderID
    '   14  IRequested
    '   15  IReceived
    '   16  IFreeReceived
    '   17  IUnitsPerPack
    '   18  Display Price
    '   19  Actual Price
    '   20  DOM
    '   21  DOE
    '   22  Last Purchase Price
    '   23  Lasr Sale Price
                .MoveNext
            Wend
        End If
    End With
    GridItem.Col = 0
    GridItem.Row = 0
    EditingData = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim tr As Integer
If GridItem.Rows > 1 Then
    tr = MsgBox("There are items to be received. Are You sure you want to exit?", vbYesNo + vbQuestion, "Exit?")
    If tr = vbNo Then Cancel = True: Exit Sub
End If
End Sub

Private Sub GridItem_Click()
    With GridItem
        txtRow.Text = .Row
        dtcItem.BoundText = .TextMatrix(.Row, 2)
        NewItem.ID = dtcItem.BoundText
        txtIQty.Text = .TextMatrix(.Row, 15)
        txtIFree.Text = .TextMatrix(.Row, 16)
        txtBatch.Text = .TextMatrix(.Row, 9)
        txtMargin.Text = .TextMatrix(.Row, 12)
        If IsDate(.TextMatrix(.Row, 21)) Then dtpDOE.Value = .TextMatrix(.Row, 21)
        If IsDate(.TextMatrix(.Row, 20)) Then dtpDOM.Value = .TextMatrix(.Row, 20)
        txtPPrice.Text = Format(.TextMatrix(.Row, 10) * NewItem.IssueUnitsPerPack, "0.00")
        txtSPrice.Text = Format((.TextMatrix(.Row, 11)), "0.00")
        lblIUnit.Caption = NewItem.IUnit
        lblIUnit1.Caption = NewItem.IUnit
        lblPunit.Caption = NewItem.PUnit
        lblPUnit1.Caption = NewItem.PUnit
        txtValue.Text = Format(.TextMatrix(.Row, 19), "0.00")
    End With
End Sub

'Private Sub GridItem_EnterCell()
'If EditingData = False Then Exit Sub
'        Dim TemCol As Long
'        Dim TemId As Long
'
'With GridItem
'    If .Col = 9 Then
'        DOE = .TextMatrix(.Row, 21)
'        DOM = .TextMatrix(.Row, 20)
'        Batch = .TextMatrix(.Row, 9)
'        Item = .TextMatrix(.Row, 1)
'        frmBatch.Show 1
'        .TextMatrix(.Row, 21) = frmBatch.DOE
'        .TextMatrix(.Row, 20) = frmBatch.DOM
'        .TextMatrix(.Row, 9) = frmBatch.Batch
'        txtDataEntry.Text = .TextMatrix(.Row, 9)
'        Exit Sub
'    End If
'
'    If .Row = 0 Or .Row = 0 Then
'        txtDataEntry.Visible = False
'        If .Row <> CurrentRow Then
'            EditingData = False
'            TemCol = 2
'            TemId = Val(.TextMatrix(.Row, TemCol))
'            NewItem.ID = TemId
'            EditingData = True
'            CurrentRow = .Row
'        End If
'        Exit Sub
'    End If
'
'    Select Case .CellAlignment
'        Case 0, 1, 2: txtDataEntry.Alignment = 0
'        Case 3, 4, 5: txtDataEntry.Alignment = 2
'        Case 6, 7, 8: txtDataEntry.Alignment = 1
'    End Select
'
'    If .Col = 5 Or .Col = 7 Or .Col = 9 Or .Col = 10 Or .Col = 11 Then
'        txtDataEntry.Locked = False
'    Else
'        txtDataEntry.Locked = True
'    End If
'
'    txtDataEntry.Text = Empty
'    txtDataEntry.Visible = False
'    txtDataEntry.Top = .Top + .CellTop
'    txtDataEntry.Left = .Left + .CellLeft
'    txtDataEntry.Width = .CellWidth
'    txtDataEntry.Height = .CellHeight
'    TemCellContent = .Text
'    txtDataEntry.Text = .Text
'    txtDataEntry.Visible = True
'    txtDataEntry.SetFocus
'
'    If .Row <> CurrentRow Then
'        EditingData = False
'        TemCol = 2
'        TemId = Val(.TextMatrix(.Row, TemCol))
'        NewItem.ID = TemId
'        EditingData = True
'        CurrentRow = .Row
'    End If
'
'
'End With
'End Sub
'
'Private Sub GridItem_LeaveCell()
'Dim TemRow As Integer
'If EditingData = False Then Exit Sub
'With GridItem
'        GridItem.Text = txtDataEntry.Text
'        Call CalculateQuentities
'        If TemCellContent <> .Text Then
'            TemRow = .Row
'            Select Case .Col
'                Case 10:
'
'                            Dim TemPPrice As Double
'                            Dim TemSMargin As Double
'                            Dim TemSPrice As Double
'                            Dim TemUnitPrice As Double
'                            Dim TemUnitsPerPasck As Double
'
'                            TemPPrice = Val(.TextMatrix(TemRow, 10))
'                            TemSMargin = 1 + ((Val(.TextMatrix(TemRow, 12))) / 100)
'                            TemUnitsPerPasck = Val(.TextMatrix(TemRow, 17))
'                            TemUnitPrice = Format((TemPPrice * TemSMargin / TemUnitsPerPasck), "0.00")
'
'                            .TextMatrix(TemRow, 11) = TemUnitPrice
'
'                            .TextMatrix(TemRow, 18) = Format((Val(.TextMatrix(TemRow, 10)) * Val(.TextMatrix(TemRow, 5))), "#,###.00")
'                            .TextMatrix(TemRow, 19) = Format((Val(.TextMatrix(TemRow, 10)) * Val(.TextMatrix(TemRow, 5))), "####.00")
'
'                            CalculateTotal
'            End Select
'
'            TemCellContent = .Text
'        End If
'
'    '   0   No
'    '   1   Item
'    '   2   ItemID
'    '   3   RequestedQuentity
'    '   4   PUnit
'    '   5   PurchaseQuentity
'    '   6   PUnit
'    '   7   FreeQuentity
'    '   8   PUnit
'    '   9   Batch
'    '   10  Purchase Price
'    '   11  Sales Price
'    '   12  Sales Margin
'    '   13  OrderID
'    '   14  IRequested
'    '   15  IReceived
'    '   16  IFreeReceived
'    '   17  IUnitsPerPack
'    '   18  Display Price
'    '   19  Actual Price
'    '   20  DOM
'    '   21  DOE
'    '   22  Last Purchase Price
'    '   23  Lasr Sale Price
'
'End With
'
'
'End Sub

Private Sub CalculateTotal()
    Dim i As Integer
    Dim GrossTotal As Double
    Dim NetTotal As Double
    With GridItem
        For i = 1 To GridItem.Rows - 1
            GrossTotal = GrossTotal + Val(.TextMatrix(i, 19))
        Next
        lblGrossTotal.Caption = Format(GrossTotal, "#,###.00")
        NetTotal = GrossTotal - Val(txtDiscount.Text)
        lblNetTotal.Caption = Format(NetTotal, "#,###.00")
    End With
End Sub

Private Sub CalculateQuentities()
    Dim temRow As Integer
    With GridItem
    temRow = .Row
        If temRow = 0 Then Exit Sub
        .TextMatrix(temRow, 15) = Val(.TextMatrix(temRow, 5)) * Val(.TextMatrix(temRow, 17))
        .TextMatrix(temRow, 16) = Val(.TextMatrix(temRow, 7)) * Val(.TextMatrix(temRow, 17))
        .TextMatrix(temRow, 18) = Format((Val(.TextMatrix(temRow, 10)) * Val(.TextMatrix(temRow, 5))), "#,###.00")
        .TextMatrix(temRow, 19) = Format((Val(.TextMatrix(temRow, 10)) * Val(.TextMatrix(temRow, 5))), "####.00")
        CalculateTotal
    End With
End Sub

Private Sub GetOrderDetails(OrderID As Long)
    With rsTemOrder
        If .State = 1 Then .Close
        temSql = "SELECT * from tblOrder where orderID = " & OrderID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 1 Then Exit Sub
        NewItem.ID = !ItemID
    End With
End Sub


Private Sub DistributorDetails(ByVal DistributorID As Long)
    With rsTemDistributor
       If .State = 1 Then .Close
       temSql = "SELECT tblDistrubutor.*, tblCity.City FROM tblCity RIGHT JOIN tblDistrubutor ON tblCity.CityId = tblDistrubutor.DistributorCityID Where DistributorId = " & DistributorID & ""
       .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
       If .RecordCount = 0 Then Exit Sub
       If Not IsNull(!DistributorName) Then lblDistributor.Caption = !DistributorName
       If Not IsNull(!Balance) Then lblBalance.Caption = Format(!Balance, "0.00")
       If Not IsNull(!DistributorTelephone) Then lblTelNo.Caption = !DistributorTelephone
       If Not IsNull(!DistributorFax) Then lblFax.Caption = !DistributorFax
       If Not IsNull(!DistributorAddress) Then lblAddress.Caption = !DistributorAddress
       If Not IsNull(!City) Then lblCity.Caption = !City
       If .State = 1 Then .Close
    End With
End Sub

Private Sub lblGrossTotal_Change()
    lblNetTotal.Caption = Format((Val(lblGrossTotal.Caption) - Val(txtDiscount.Text)), "0.00")
End Sub


Private Sub txtDiscount_Change()
    lblNetTotal.Caption = Format((Val(lblGrossTotal.Caption) - Val(txtDiscount.Text)), "0.00")
End Sub

Private Sub txtDiscount_LostFocus()
    txtDiscount.Text = Format(txtDiscount.Text, "0.00")
End Sub

Private Sub txtIQty_Change()
    txtPQty.Text = Val(txtIQty.Text) / NewItem.IssueUnitsPerPack
End Sub


Private Sub txtPPrice_Change()
    txtValue.Text = Format(Val(txtPQty.Text) * Val(txtPPrice.Text), "0.00")
    txtSPrice.Text = Format((Val(txtMargin.Text) + 100) * Val(txtPPrice.Text) / (100 * NewItem.IssueUnitsPerPack), "0.00")
End Sub

Private Sub txtPQty_LostFocus()
    txtIQty.Text = txtPQty.Text * NewItem.IssueUnitsPerPack
    txtValue.Text = Format(Val(txtPQty.Text) * Val(txtPPrice.Text), "0.00")
End Sub

Private Sub txtIFree_Change()
    txtPFree.Text = Val(txtIFree.Text) / NewItem.IssueUnitsPerPack
End Sub


Private Sub txtPFree_LostFocus()
    txtIFree.Text = Val(txtPFree.Text) * NewItem.IssueUnitsPerPack
End Sub

