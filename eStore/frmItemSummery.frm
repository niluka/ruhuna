VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmItemSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Summery"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   14205
   Begin btButtonEx.ButtonEx btnFill 
      Height          =   375
      Left            =   10800
      TabIndex        =   49
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Fill"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "frmItemSummery.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtAMPP"
      Tab(0).Control(1)=   "txtVMPP"
      Tab(0).Control(2)=   "txtAMP"
      Tab(0).Control(3)=   "txtVMP"
      Tab(0).Control(4)=   "txtVTM"
      Tab(0).Control(5)=   "txtDisplay"
      Tab(0).Control(6)=   "txtIStore"
      Tab(0).Control(7)=   "Label13"
      Tab(0).Control(8)=   "Label12"
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(10)=   "Label10"
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(12)=   "Label6"
      Tab(0).Control(13)=   "Label14"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Stocks"
      TabPicture(1)   =   "frmItemSummery.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridTotal"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Usage"
      TabPicture(2)   =   "frmItemSummery.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dtpUFrom"
      Tab(2).Control(1)=   "GridUsage"
      Tab(2).Control(2)=   "dtpUTo"
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(4)=   "Label3"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Ordering"
      TabPicture(3)   =   "frmItemSummery.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dtpOFrom"
      Tab(3).Control(1)=   "GridOrdering"
      Tab(3).Control(2)=   "dtpOTo"
      Tab(3).Control(3)=   "Label15"
      Tab(3).Control(4)=   "Label8"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Prices"
      TabPicture(4)   =   "frmItemSummery.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label16"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label17"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label18"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label19"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "dtpPTo"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "dtpPFrom"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "GridSPrice"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "GridPPrice"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Good Receive"
      TabPicture(5)   =   "frmItemSummery.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "dtpPurchaseFrom"
      Tab(5).Control(1)=   "GridPurchase"
      Tab(5).Control(2)=   "dtpPurchaseTo"
      Tab(5).Control(3)=   "Label2"
      Tab(5).Control(4)=   "Label1"
      Tab(5).ControlCount=   5
      Begin VB.TextBox txtAMPP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2880
         Width           =   6615
      End
      Begin VB.TextBox txtVMPP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2400
         Width           =   6615
      End
      Begin VB.TextBox txtAMP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1920
         Width           =   6615
      End
      Begin VB.TextBox txtVMP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   6615
      End
      Begin VB.TextBox txtVTM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   6615
      End
      Begin VB.TextBox txtDisplay 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   6615
      End
      Begin VB.TextBox txtIStore 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   3360
         Width           =   6615
      End
      Begin MSFlexGridLib.MSFlexGrid GridPPrice 
         Height          =   4815
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   8493
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpUFrom 
         Height          =   375
         Left            =   -74040
         TabIndex        =   9
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   189792259
         CurrentDate     =   39540
      End
      Begin MSFlexGridLib.MSFlexGrid GridUsage 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   10
         Top             =   840
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   9128
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridTotal 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   10186
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpUTo 
         Height          =   375
         Left            =   -71040
         TabIndex        =   12
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   189792259
         CurrentDate     =   39540
      End
      Begin MSComCtl2.DTPicker dtpOFrom 
         Height          =   375
         Left            =   -74040
         TabIndex        =   13
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   189792259
         CurrentDate     =   39540
      End
      Begin MSFlexGridLib.MSFlexGrid GridOrdering 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   14
         Top             =   840
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   9128
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpOTo 
         Height          =   375
         Left            =   -71040
         TabIndex        =   15
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   189792259
         CurrentDate     =   39540
      End
      Begin MSFlexGridLib.MSFlexGrid GridSPrice 
         Height          =   4815
         Left            =   8040
         TabIndex        =   16
         Top             =   1080
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   8493
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpPFrom 
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   189792259
         CurrentDate     =   39540
      End
      Begin MSComCtl2.DTPicker dtpPTo 
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   189792259
         CurrentDate     =   39540
      End
      Begin MSComCtl2.DTPicker dtpPurchaseFrom 
         Height          =   375
         Left            =   -74040
         TabIndex        =   34
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   189792259
         CurrentDate     =   39540
      End
      Begin MSFlexGridLib.MSFlexGrid GridPurchase 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   35
         Top             =   960
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   9128
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpPurchaseTo 
         Height          =   375
         Left            =   -71040
         TabIndex        =   36
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   189792259
         CurrentDate     =   39540
      End
      Begin VB.Label Label2 
         Caption         =   "From :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "To :"
         Height          =   255
         Left            =   -71520
         TabIndex        =   37
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label19 
         Caption         =   "From :"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "To :"
         Height          =   255
         Left            =   3360
         TabIndex        =   32
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Purchase Prices"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label16 
         Caption         =   "Sales Prices"
         Height          =   255
         Left            =   8160
         TabIndex        =   30
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label15 
         Caption         =   "From :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "To :"
         Height          =   255
         Left            =   -71520
         TabIndex        =   28
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "To :"
         Height          =   255
         Left            =   -71520
         TabIndex        =   27
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "From :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label13 
         Caption         =   "Actual Medicinal Product Pack :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   25
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label12 
         Caption         =   "Virtual Medicinal Product Pack :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label Label11 
         Caption         =   "Actual Medicinal Product:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Virtual Medicinal Product:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Virtual Therapeutic Moiety:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Display Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label14 
         Caption         =   "Store :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   3360
         Width           =   3255
      End
   End
   Begin MSDataListLib.DataCombo dtcCatogery 
      Height          =   315
      Left            =   1320
      TabIndex        =   39
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItem 
      Height          =   465
      Left            =   1320
      TabIndex        =   40
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCode 
      Height          =   465
      Left            =   1320
      TabIndex        =   41
      Top             =   840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   8520
      TabIndex        =   45
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   190119939
      CurrentDate     =   39540
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   8520
      TabIndex        =   46
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   188940291
      CurrentDate     =   39540
   End
   Begin VB.Label lblCategory 
      Height          =   255
      Left            =   3480
      TabIndex        =   50
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label20 
      Caption         =   "To :"
      Height          =   255
      Left            =   7800
      TabIndex        =   48
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "From :"
      Height          =   255
      Left            =   7800
      TabIndex        =   47
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label45 
      Caption         =   "C&ode"
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label29 
      Caption         =   "&Item"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "&Category"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmItemSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCategory As New ADODB.Recordset
    Dim rsCode As New ADODB.Recordset


    Dim temSql As String
    
    Dim CsetPrinter As New cSetDfltPrinter
    
    Dim TemOrderBillID As Long
    Dim TemDistributorId As Long
    Dim TemDistributorOrderID As Long
    Dim EditingData As Boolean
    Dim TemContent(22) As String
    Dim CurrentRow As Integer
    Dim TemCellContent As String
    Dim temRefillBillID As Long
    
    Dim rsTemPurchase As New ADODB.Recordset
    
    Dim NewItem As New Item
    
    Dim rsStaff As New ADODB.Recordset
    Dim rsSPrice As New ADODB.Recordset
    Dim rsPPrice As New ADODB.Recordset
    Dim rsCC As New ADODB.Recordset
    Dim rsBanks As New ADODB.Recordset
    Dim rsCreditCards As New ADODB.Recordset
    Dim rsCities As New ADODB.Recordset
    Dim rsPayment As New ADODB.Recordset
    Dim rsDistributor As New ADODB.Recordset
    
    Dim rsTemOrder As New ADODB.Recordset
    Dim rsTemPrice As New ADODB.Recordset
    Dim rsTemDistributor As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    Dim rsTemOrderBill As New ADODB.Recordset
    Dim rsTemDistributorOrder As New ADODB.Recordset
    Dim rsTemRefill As New ADODB.Recordset
    Dim rsTemRefillBill As New ADODB.Recordset
    Dim rsTemCash As New ADODB.Recordset
    Dim rsTemCredit As New ADODB.Recordset
    Dim rsTemCheque As New ADODB.Recordset


Private Sub btnFill_Click()
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    dtcCode.BoundText = dtcItem.BoundText
    Dim temID As Long
    temID = Val(dtcItem.BoundText)
    NewItem.ID = temID
    Call FillGridPurchase
    Call FillOrdering(temID)
    Call FillPrice(temID)
    Call FillStocks(temID)
    Call FillUsage(temID)
End Sub

Private Sub dtcCatogery_LostFocus()
    If IsNumeric(dtcCatogery.BoundText) Then
        ListSelectedItems
    Else
        ListAllItems
    End If

End Sub

Private Sub dtcItem_Change()
'    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    dtcCode.BoundText = dtcItem.BoundText
'    Dim temID As Long
'    temID = Val(dtcItem.BoundText)
'    Call FillGridPurchase
'    Call FillOrdering(temID)
'    Call FillPrice(temID)
'    Call FillStocks(temID)
'    Call FillUsage(temID)
End Sub

Private Sub dtcCode_Change()
    dtcItem.BoundText = dtcCode.BoundText
End Sub

Private Sub dtpFrom_Change()
    dtpOFrom.Value = dtpFrom.Value
    dtpPFrom.Value = dtpFrom.Value
    dtpUFrom.Value = dtpFrom.Value
    dtpPurchaseFrom.Value = dtpFrom.Value
End Sub





Private Sub dtpTo_Change()
    dtpOTo.Value = dtpTo.Value
    dtpUTo.Value = dtpTo.Value
    dtpPTo.Value = dtpTo.Value
    dtpPurchaseTo.Value = dtpTo.Value
End Sub


Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    dtpOFrom.Value = Date
    dtpOTo.Value = Date
    dtpPFrom.Value = Date
    dtpPTo.Value = Date
    dtpPurchaseFrom.Value = Date
    dtpPurchaseTo.Value = Date
    Call FillCombos
    GetCommonSettings Me
End Sub

Private Sub dtcCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtcCode.Text = Empty
        KeyCode = Empty
    End If
End Sub


Private Sub dtcCatogery_Change()
'    If IsNumeric(dtcCatogery.BoundText) Then
'        ListSelectedItems
'    Else
'        ListAllItems
'    End If
    dtcItem.Text = Empty
    dtcCode.Text = Empty
    Dim rsIC As New ADODB.Recordset
    With rsIC
        If .State = 1 Then .Close
        temSql = "Select * from tblItemCategory where ItemCategoryID = " & Val(dtcCatogery.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            lblCategory.Caption = !ItemCategory
        End If
        .Close
    End With

End Sub


Private Sub ListSelectedItems()
With rsItem
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by display"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "Display"
    .BoundColumn = "ItemID"
End With
With rsCode
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by code"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcCode
    Set .RowSource = rsCode
    .ListField = "Code"
    .BoundColumn = "ItemID"
End With

End Sub

Private Sub ListAllItems()
With rsItem
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem order by display"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcItem
    Set .RowSource = rsItem
    .ListField = "display"
    .BoundColumn = "ItemID"
End With
With rsCode
    If .State = 1 Then .Close
    temSql = "SELECT * from tblitem order by code"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
End With
With dtcCode
    Set .RowSource = rsCode
    .ListField = "Code"
    .BoundColumn = "ItemID"
End With
End Sub

Private Sub dtcCatogery_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtcCatogery.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcItem.SetFocus
    End If
End Sub

Private Sub FillCombos()
    With rsItem
        If .State = 1 Then .Close
        temSql = "SELECT * from tblitem order by display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcItem
        Set .RowSource = rsItem
        .ListField = "display"
        .BoundColumn = "ItemID"
    End With
    With rsItemCategory
        If .State = 1 Then .Close
        temSql = "SELECT * from tblItemCategory order by categoryCode"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCatogery
        Set .RowSource = rsItemCategory
        .ListField = "CategoryCode"
        .BoundColumn = "ItemCategoryID"
    End With
    With rsCode
        If .State = 1 Then .Close
        temSql = "SELECT * from tblitem order by code"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCode
        Set .RowSource = rsCode
        .ListField = "code"
        .BoundColumn = "ItemID"
    End With

End Sub

Private Sub FillUsage(ByVal ItemID As Long)
    '0 Store
    '1 Sale
    '2 Consum
    '3 Discard
    '4 Adjustments
    '5 Total
    Dim StoreConsumption As Double
    Dim StoreSale As Double
    Dim StoreAdjustment As Double
    Dim StoreDiscard As Double
    Dim StoreUsage As Double
    Dim TotalConsumption As Double
    Dim TotalSale As Double
    Dim TotalAdjustment As Double
    Dim TotalDiscard As Double
    Dim TotalUsage As Double
    Dim TemStore As String
    
    With GridUsage
        .Cols = 6
        .Rows = 1
        .FixedCols = 0
        
        .ColWidth(0) = 3000
        
        .ColWidth(1) = (.Width - (.ColWidth(0) + 100)) / 5
        .ColWidth(2) = .ColWidth(1)
        .ColWidth(3) = .ColWidth(1)
        .ColWidth(4) = .ColWidth(1)
        .ColWidth(5) = .ColWidth(1)
        
        Dim i As Long
        
        For i = 0 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            Select Case i
                Case 0: .Text = "Store"
                Case 1: .Text = "Sale"
                Case 2: .Text = "Consumption"
                Case 3: .Text = "Discard"
                Case 4: .Text = "Adjustment"
                Case 5: .Text = "Total"
            End Select
        Next i
        
    End With

    With rsTemStore
        If .State = 1 Then .Close
        temSql = "SELECT * from tblStore order by store"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
        
            While .EOF = False
                
                TemStore = !Store
                
                StoreUsage = 0
                
                StoreConsumption = CalculateConsumption(ItemID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreConsumption
                TotalConsumption = TotalConsumption + StoreConsumption
                
                StoreSale = CalculateSale(ItemID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreSale
                TotalSale = TotalSale + StoreSale
                
                StoreDiscard = CalculateDiscard(ItemID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreDiscard
                TotalDiscard = TotalDiscard + StoreDiscard
                
                StoreAdjustment = CalculateAdjustment(ItemID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreAdjustment
                TotalAdjustment = TotalAdjustment + StoreAdjustment
                
                With GridUsage
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = 0
                    .CellAlignment = 1
                    .Text = TemStore
                    .Col = 1
                    .Text = StoreSale & " " & NewItem.IUnit
                    .Col = 2
                    .Text = StoreConsumption & " " & NewItem.IUnit
                    .Col = 3
                    .Text = StoreDiscard & " " & NewItem.IUnit
                    .Col = 4
                    .Text = StoreAdjustment & " " & NewItem.IUnit
                    .Col = 5
                    .Text = StoreUsage & " " & NewItem.IUnit
                End With
                .MoveNext
            Wend
            With GridUsage
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0
            .CellAlignment = 1
            .Text = "Total"
            .Col = 1
            .Text = TotalSale & " " & NewItem.IUnit
            .Col = 2
            .Text = TotalConsumption & " " & NewItem.IUnit
            .Col = 3
            .Text = TotalDiscard & " " & NewItem.IUnit
            .Col = 4
            .Text = TotalAdjustment & " " & NewItem.IUnit
            TotalUsage = TotalConsumption + TotalSale + TotalDiscard + TotalAdjustment
            .Col = 5
            .Text = TotalUsage & " " & NewItem.IUnit
            End With
        End If
        .Close
    End With
    
End Sub


Private Sub FillOrdering(ByVal ItemID As Long)
    With GridOrdering
        .Rows = 1
        .Cols = 8
        .FixedCols = 0
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Requested On"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Approved On"

        .Col = 2
        .CellAlignment = 4
        .Text = "Received On"
        
        .Col = 3
        .CellAlignment = 4
        .Text = "Requested Amount"
        
        .Col = 4
        .CellAlignment = 4
        .Text = "Approved Amount"
        
        .Col = 5
        .CellAlignment = 4
        .Text = "Received Amount"
        
        .Col = 6
        .CellAlignment = 4
        .Text = "Requested Distributor"
        
        .Col = 7
        .CellAlignment = 4
        .Text = "Approved Distributor"
        
        Dim i As Integer
'
'        For i = 0 To .Cols - 1
'            .ColWidth(i) = (.Width - 100) / 8
'        Next i
'
        .ColWidth(5) = 0
        
    End With
    With rsTemOrder
        If .State = 1 Then .Close
        temSql = "SELECT tblOrder.RequestDate, tblOrder.ApprovedDate, tblOrder.ReceivedDate, tblOrder.RequestAmount, tblOrder.ApprovedAmount, tblOrder.ReceivedAmount, tblRDistrubutor.DistributorName as RDistributorName , tblADistrubutor.DistributorName as ADistributorName FROM (tblDistrubutor AS tblRDistrubutor RIGHT JOIN tblOrder ON tblRDistrubutor.DistributorID = tblOrder.ApprovedDistributorID) LEFT JOIN tblDistrubutor AS tblADistrubutor ON tblOrder.RequestDistributorID = tblADistrubutor.DistributorID WHERE (((tblOrder.ItemID)=" & ItemID & ") AND ((tblOrder.RequestDate) Between '" & Format(dtpOFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpOTo.Value, "dd MMMM yyyy") & "')) ORDER BY tblOrder.RequestDate"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount >= 1 Then
            While .EOF = False
                GridOrdering.Rows = GridOrdering.Rows + 1
                GridOrdering.Row = GridOrdering.Rows - 1
                GridOrdering.Col = 0
                GridOrdering.CellAlignment = 1
                GridOrdering.Text = Format(!requestdate, ShortDateFormat)
                GridOrdering.Col = 1
                GridOrdering.CellAlignment = 1
                If Not IsNull(!ApprovedDate) Then
                    GridOrdering.Text = Format(!ApprovedDate, ShortDateFormat)
                Else
                    GridOrdering.Text = "Not Approved"
                End If
                GridOrdering.Col = 2
                GridOrdering.CellAlignment = 1
                If Not IsNull(!ReceivedDate) Then
                    GridOrdering.Text = Format(!ReceivedDate, ShortDateFormat)
                Else
                    GridOrdering.Text = "Not Received"
                End If
                GridOrdering.Col = 3
                GridOrdering.CellAlignment = 7
                If Not IsNull(!RequestAmount) Then
                    GridOrdering.Text = !RequestAmount & " " & NewItem.IUnit
                Else
                    GridOrdering.Text = "Not Requested"
                End If
                GridOrdering.Col = 4
                GridOrdering.CellAlignment = 7
                If Not IsNull(!ApprovedAmount) Then
                    GridOrdering.Text = !ApprovedAmount & " " & NewItem.IUnit
                Else
                    GridOrdering.Text = "Not Approved"
                End If
                GridOrdering.Col = 5
                GridOrdering.CellAlignment = 7
                If Not IsNull(!ReceivedAmount) Then
                    GridOrdering.Text = !ReceivedAmount & " " & NewItem.IUnit
                Else
                    GridOrdering.Text = "Not Received"
                End If
                GridOrdering.Col = 6
                GridOrdering.CellAlignment = 7
                If Not IsNull(.Fields("RDistributorName").Value) Then
                    GridOrdering.Text = .Fields("RDistributorName").Value
                Else
                    GridOrdering.Text = "Not Requested"
                End If
                GridOrdering.Col = 7
                GridOrdering.CellAlignment = 7
                If Not IsNull(.Fields("ADistributorName").Value) Then
                    GridOrdering.Text = .Fields("ADistributorName").Value
                Else
                    GridOrdering.Text = "Not Approved"
                End If
                .MoveNext
            Wend
        End If
    End With
End Sub

Private Sub FillPrice(ByVal ItemID As Long)
With GridPPrice
    .Cols = 2
    .Rows = 1
    .FixedCols = 0
    
    .Row = 0
    
    .Col = 0
    .CellAlignment = 4
    .Text = "Starting Date"
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Purchase Price per " & NewItem.PUnit
    
    .ColWidth(0) = (.Width - 100) / 2
    .ColWidth(1) = (.Width - 100) / 2
    
End With

With rsTemPrice
    If .State = 1 Then .Close
    temSql = "SELECT tblPurchasePrice.SetDate, tblPurchasePrice.PPrice FROM tblPurchasePrice WHERE (((tblPurchasePrice.ItemID)=" & ItemID & ") AND ((tblPurchasePrice.SetDate) Between '" & Format(dtpPFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpPTo.Value, "dd MMMM yyyy") & "')) ORDER BY tblPurchasePrice.SetDate DESC"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        While .EOF = False
            With GridPPrice
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .CellAlignment = 1
                .Text = Format(rsTemPrice!setdate, LongDateFormat)
                .Col = 1
                .CellAlignment = 7
                .Text = Format(rsTemPrice!PPrice * NewItem.IssueUnitsPerPack, "#,#00.00")
            End With
            .MoveNext
        Wend
    End If
End With


With GridSPrice
    .Cols = 2
    .Rows = 1
    .FixedCols = 0
    
    .Row = 0
    
    .Col = 0
    .CellAlignment = 4
    .Text = "Starting Date"
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Sales Price per " & NewItem.IUnit
    
    .ColWidth(0) = (.Width - 100) / 2
    .ColWidth(1) = (.Width - 100) / 2
    
End With

With rsTemPrice
    If .State = 1 Then .Close
    temSql = "SELECT tblSalePrice.SetDate, tblSalePrice.SPrice FROM tblSalePrice WHERE (((tblSalePrice.ItemID)=" & ItemID & ") AND ((tblSalePrice.SetDate) Between '" & Format(dtpPFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpPTo.Value, "dd MMMM yyyy") & "')) ORDER BY tblSalePrice.SetDate DESC"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        While .EOF = False
            With GridSPrice
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .CellAlignment = 1
                .Text = Format(rsTemPrice!setdate, LongDateFormat)
                .Col = 1
                .CellAlignment = 7
                .Text = Format(rsTemPrice!SPrice, "#,#00.00")
            End With
            .MoveNext
        Wend
    End If
End With


End Sub


Private Sub FillStocks(ByVal ItemID As Long)
    With GridTotal
        .Visible = False
    
        .Clear
        .Cols = 4
        .Rows = 1
        .Row = 0
        .FixedCols = 0
        
        .Col = 0
        .CellAlignment = 4
        .Text = "Batch"
        
        .Col = 1
        .CellAlignment = 4
        .Text = "Stock (" & NewItem.IUnit & ")"
        
        .Col = 2
        .CellAlignment = 4
        .Text = "Expiary"
        
        .Col = 3
        .CellAlignment = 4
        .Text = "Department"
        
        .ColWidth(1) = 1600
        .ColWidth(2) = 1600
        .ColWidth(3) = 1600
        .ColWidth(0) = .Width - (.ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 100)
        
    End With
    With rsTemStore
        If .State = 1 Then .Close
        temSql = "SELECT tblBatch.Batch, tblBatch.DOE, tblBatchStock.Stock, tblStore.Store, tblBatch.ItemID " & _
                    " FROM tblStore RIGHT JOIN (tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) ON tblStore.StoreID = tblBatchStock.StoreID " & _
                    " WHERE tblBatch.ItemID=" & ItemID & " AND tblBatchStock.Stock > 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                GridTotal.Rows = GridTotal.Rows + 1
                GridTotal.Row = GridTotal.Rows - 1
                GridTotal.Col = 0
                GridTotal.CellAlignment = 1
                GridTotal.Text = !Batch
                GridTotal.Col = 1
                GridTotal.CellAlignment = 7
                If Not IsNull(!Stock) Then
                    GridTotal.Text = !Stock
                Else
                    GridTotal.Text = 0
                End If
                GridTotal.Col = 2
                GridTotal.CellAlignment = 1
                GridTotal.Text = Format(!DOE, ShortDateFormat)
                GridTotal.Col = 3
                GridTotal.CellAlignment = 1
                If Not IsNull(!Store) Then
                    GridTotal.Text = !Store
                Else
                    GridTotal.Text = Empty
                End If
                .MoveNext
            Wend
        End If
        GridTotal.Visible = True
        .Close
    End With
    
End Sub


Private Sub GetItemDetails(ItemID As Long)
    NewItem.ID = ItemID
    txtAMP.Text = NewItem.AMP
    txtAMPP.Text = NewItem.AMPP
    txtVMP.Text = NewItem.VMP
    txtVMPP.Text = NewItem.VMPP
    txtVTM.Text = NewItem.Generic
    txtDisplay.Text = NewItem.Display
End Sub


Private Sub FillGridPurchase()
    With GridPurchase
        .Clear
        .Rows = 2
        .Cols = 8
        
        .Row = 0
        
        .Col = 0
        .Text = "GRN Date"
        
        .Col = 1
        .Text = "Supplier"
        
        .Col = 2
        .Text = "GRN No"
        
        .Col = 3
        .Text = "Invoice No"
        
        .Col = 4
        .Text = "Quentity"
        
        .Col = 5
        .Text = "Free Quentity"
        
        .Col = 6
        .Text = "Price"
        
        .Col = 7
        .Text = "Expiary"
        
        .ColWidth(1) = 3000
    
    End With
    
    temSql = "SELECT tblRefillBill.Date, tblDistrubutor.DistributorName, tblRefillBill.RefillBillID, tblRefillBill.InvoiceNo, tblRefill.Amount, tblRefill.FreeAmount, tblRefill.Price, tblRefill.DOE " & _
                "FROM (tblRefill LEFT JOIN tblRefillBill ON tblRefill.RefillBillID = tblRefillBill.RefillBillID) LEFT JOIN tblDistrubutor ON tblRefillBill.DistributorID = tblDistrubutor.DistributorID " & _
                "WHERE (((tblRefillBill.Date) Between '" & Format(dtpPurchaseFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpPurchaseTo.Value, "dd MMMM yyyy") & "') AND ((tblRefill.ItemID)=" & Val(dtcItem.BoundText) & ")) " & _
                "Order by tblRefillBill.Date"
    With rsTemPurchase
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            GridPurchase.Rows = .RecordCount + 1
            Dim Pi As Integer
            For Pi = 1 To .RecordCount
                If Not IsNull(!Date) Then GridPurchase.TextMatrix(Pi, 0) = !Date
                If Not IsNull(!DistributorName) Then GridPurchase.TextMatrix(Pi, 1) = !DistributorName
                If Not IsNull(!RefillBillID) Then GridPurchase.TextMatrix(Pi, 2) = !RefillBillID
                If Not IsNull(!InvoiceNo) Then GridPurchase.TextMatrix(Pi, 3) = !InvoiceNo
                If Not IsNull(!Amount) Then GridPurchase.TextMatrix(Pi, 4) = !Amount
                If Not IsNull(!FreeAmount) Then GridPurchase.TextMatrix(Pi, 5) = !FreeAmount
                If Not IsNull(!Price) Then GridPurchase.TextMatrix(Pi, 6) = Format(!Price, "0.00")
                If Not IsNull(!DOE) Then GridPurchase.TextMatrix(Pi, 7) = !DOE
                .MoveNext
            Next
        End If
        .Close
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub
