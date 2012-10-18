VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmItemsDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Details"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
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
   ScaleHeight     =   8025
   ScaleWidth      =   15270
   Begin btButtonEx.ButtonEx btnFill 
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Fill"
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
      Left            =   14160
      TabIndex        =   12
      Top             =   7560
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSDataListLib.DataCombo dtcCategory 
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "frmItemsDetails.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Stocks"
      TabPicture(1)   =   "frmItemsDetails.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblStoreStock"
      Tab(1).Control(1)=   "GridTotal"
      Tab(1).Control(2)=   "txtStoreStock"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Usage"
      TabPicture(2)   =   "frmItemsDetails.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtStoreUsage"
      Tab(2).Control(1)=   "GridUsage"
      Tab(2).Control(2)=   "lblStoreUsage"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Ordering"
      TabPicture(3)   =   "frmItemsDetails.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GridOrdering"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Purchase"
      TabPicture(4)   =   "frmItemsDetails.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GridPurchase"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Prices"
      TabPicture(5)   =   "frmItemsDetails.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "GridPPrice"
      Tab(5).Control(1)=   "GridSPrice"
      Tab(5).Control(2)=   "Label16"
      Tab(5).Control(3)=   "Label17"
      Tab(5).ControlCount=   4
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   120
         TabIndex        =   53
         Top             =   4080
         Width           =   6855
         Begin VB.TextBox txtROQ 
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox txtROL 
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label8 
            Caption         =   "ROQ:"
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
            Left            =   120
            TabIndex        =   57
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label5 
            Caption         =   "ROL:"
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
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3735
         Left            =   7080
         TabIndex        =   35
         Top             =   360
         Width           =   7695
         Begin VB.Label Label33 
            Caption         =   "City"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblCity 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   46
            Top             =   1800
            Width           =   3375
         End
         Begin VB.Label lblFax 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   45
            Top             =   2760
            Width           =   3375
         End
         Begin VB.Label Label31 
            Caption         =   "Fax No"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label lblAddress 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   975
            Left            =   1560
            TabIndex        =   43
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label lblTelNo 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   42
            Top             =   2280
            Width           =   3375
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   375
            Left            =   1560
            TabIndex        =   41
            Top             =   3240
            Width           =   3375
         End
         Begin VB.Label Label27 
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "Tel No"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Address"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Distributor"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblDistributor 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   36
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3735
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   6855
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   3120
            Width           =   4095
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   4095
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   720
            Width           =   4095
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1200
            Width           =   4095
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1680
            Width           =   4095
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   2160
            Width           =   4095
         End
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   2640
            Width           =   4095
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
            Left            =   120
            TabIndex        =   34
            Top             =   3120
            Width           =   3255
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
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   2535
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
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   2535
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
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   2295
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
            Left            =   120
            TabIndex        =   30
            Top             =   1680
            Width           =   2175
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
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   3855
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
            Left            =   120
            TabIndex        =   28
            Top             =   2640
            Width           =   3255
         End
      End
      Begin VB.TextBox txtStoreUsage 
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
         Left            =   -67440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtStoreStock 
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
         Left            =   -69960
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   400
         Width           =   2295
      End
      Begin MSFlexGridLib.MSFlexGrid GridUsage 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   13
         Top             =   840
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7858
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridTotal 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   14
         Top             =   840
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7858
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridOrdering 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   15
         Top             =   840
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7858
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPPrice 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   48
         Top             =   840
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7223
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridSPrice 
         Height          =   4095
         Left            =   -66840
         TabIndex        =   49
         Top             =   840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7223
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPurchase 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   52
         Top             =   840
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7858
         _Version        =   393216
      End
      Begin VB.Label Label16 
         Caption         =   "Sales Prices"
         Height          =   255
         Left            =   -66840
         TabIndex        =   51
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label17 
         Caption         =   "Purchase Prices"
         Height          =   255
         Left            =   -74880
         TabIndex        =   50
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label lblStoreUsage 
         Alignment       =   2  'Center
         Caption         =   "Store Stock"
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
         Left            =   -69960
         TabIndex        =   19
         Top             =   435
         Width           =   2415
      End
      Begin VB.Label lblStoreStock 
         Alignment       =   2  'Center
         Caption         =   "Store Stock"
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
         Left            =   -72480
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
   End
   Begin MSDataListLib.DataCombo dtcItem 
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCode 
      Height          =   360
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFDate 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   22020099
      CurrentDate     =   39540
   End
   Begin MSComCtl2.DTPicker dtpTDate 
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   22020099
      CurrentDate     =   39540
   End
   Begin VB.Label Label3 
      Caption         =   "From :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "To :"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "&Code"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "&Item"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Ca&tegory"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmItemsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
    Dim CsetPrinter As New cSetDfltPrinter
    
    Dim NewItem As New Item
    
    Dim rsSPrice As New ADODB.Recordset
    Dim rsPPrice As New ADODB.Recordset
    Dim rsCC As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim rsItemCategory As New ADODB.Recordset
    Dim rsCode As New ADODB.Recordset
    Dim rsDistributor As New ADODB.Recordset
    
    Dim rsTemOrder As New ADODB.Recordset
    Dim rsTemPrice As New ADODB.Recordset
    Dim rsTemDistributor As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    Dim rsTemOrderBill As New ADODB.Recordset
    Dim rsTemDistributorOrder As New ADODB.Recordset
    Dim rsTemRefill As New ADODB.Recordset
    Dim rsTemRefillBill As New ADODB.Recordset

Private Sub ListSelectedItems()
    With rsItem
        If .State = 1 Then .Close
        temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcCategory.BoundText & " order by display"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcItem
        Set .RowSource = rsItem
        .ListField = "Display"
        .BoundColumn = "ItemID"
    End With
    With rsCode
        If .State = 1 Then .Close
        temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcCategory.BoundText & " order by code"
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

Private Sub btnFill_Click()
    dtcCode.BoundText = dtcItem.BoundText
    Call FillDetails
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub dtcCategory_Change()
    If IsNumeric(dtcCategory.BoundText) Then
        ListSelectedItems
    Else
        ListAllItems
    End If
    dtcItem.Text = Empty
    dtcCode.Text = Empty
End Sub

Private Sub dtcCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcItem.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtcCategory.Text = Empty
    End If
End Sub

Private Sub dtcCode_Change()
    If IsNumeric(dtcCode.BoundText) = True Then dtcItem.BoundText = dtcCode.BoundText
End Sub

Private Sub dtcCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        btnFill.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtcCode.Text = Empty
    End If
End Sub

Private Sub dtcItem_Change()
    If IsNumeric(dtcItem.BoundText) = True Then dtcCode.BoundText = dtcItem.BoundText
End Sub

Private Sub dtcItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpFDate.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtcItem.Text = Empty
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
        temSql = "SELECT * from tblItemCategory order by ItemCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCategory
        Set .RowSource = rsItemCategory
        .ListField = "ItemCategory"
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
    

Private Sub ClearItemDetails()

End Sub

    
Private Sub FillDetails()
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    NewItem.ID = dtcItem.BoundText
    Call FormatGrids
    Call FillLabels
    Call GetItemDetails(NewItem.ID)
    Call FillStocks(dtcItem.BoundText)
    Call FillPurchase(dtcItem.BoundText)
    Call FillPrice(dtcItem.BoundText)
    Call GetItemDetails(dtcItem.BoundText)
    Call FillOrdering(dtcItem.BoundText)
    Call FillUsage(dtcItem.BoundText)
    Dim rsDI As New ADODB.Recordset
    With rsDI
        If .State = 1 Then .Close
        temSql = "SELECT tblItemDistributor.DistributorID FROM tblItemDistributor WHERE (((tblItemDistributor.ItemID)=" & Val(dtcItem.BoundText) & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            DistributorDetails (!DistributorID)
        End If
        .Close
    End With
End Sub

Private Sub FillLabels()
    txtAMP.Text = NewItem.AMP
    txtAMPP.Text = NewItem.AMPP
    txtDisplay.Text = NewItem.Display
    txtROL.Text = NewItem.ROL
    txtROQ.Text = NewItem.ROQ
    txtVMP.Text = NewItem.VMP
    txtVMPP.Text = NewItem.VMPP
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
    

    With rsTemStore
        If .State = 1 Then .Close
        temSql = "SELECT * from tblStore order by store"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
        
            While .EOF = False
                
                TemStore = !Store
                
                StoreUsage = 0
                
                StoreConsumption = CalculateConsumption(NewItem.ID, dtpFDate.Value, dtpTDate.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreConsumption
                TotalConsumption = TotalConsumption + StoreConsumption
                
                StoreSale = CalculateSale(NewItem.ID, dtpFDate.Value, dtpTDate.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreSale
                TotalSale = TotalSale + StoreSale
                
                StoreDiscard = CalculateDiscard(NewItem.ID, dtpFDate.Value, dtpTDate.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreDiscard
                TotalDiscard = TotalDiscard + StoreDiscard
                
                StoreAdjustment = CalculateAdjustment(NewItem.ID, dtpFDate.Value, dtpTDate.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreAdjustment
                TotalAdjustment = TotalAdjustment + StoreAdjustment
                
                If !StoreID = UserStoreID Then
                    lblStoreUsage.Caption = !Store
                    txtStoreUsage.Text = StoreUsage
                End If
                
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
    With rsTemOrder
        If .State = 1 Then .Close
        temSql = "SELECT tblOrder.RequestDate, tblOrder.ApprovedDate, tblOrder.ReceivedDate, tblOrder.RequestAmount, tblOrder.ApprovedAmount, tblOrder.ReceivedAmount, tblRDistrubutor.DistributorName as RDistributorName , tblADistrubutor.DistributorName  as  ADistributorName FROM (tblDistrubutor AS tblRDistrubutor RIGHT JOIN tblOrder ON tblRDistrubutor.DistributorID = tblOrder.ApprovedDistributorID) LEFT JOIN tblDistrubutor AS tblADistrubutor ON tblOrder.RequestDistributorID = tblADistrubutor.DistributorID WHERE tblOrder.ItemID  = " & ItemID & "AND tblOrder.RequestDate between '" & Format(dtpFDate.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTDate.Value, "dd MMMM yyyy") & "' Order by tblOrder.ApprovedDate"
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
                If Not IsNull(!RequestAmount) Then
                    GridOrdering.Text = !RequestAmount & " " & NewItem.IUnit
                Else
                    GridOrdering.Text = "Not Requested"
                End If
                GridOrdering.Col = 3
                GridOrdering.CellAlignment = 7
                If Not IsNull(!ApprovedAmount) Then
                    GridOrdering.Text = !ApprovedAmount & " " & NewItem.IUnit
                Else
                    GridOrdering.Text = "Not Approved"
                End If
                GridOrdering.Col = 4
                GridOrdering.CellAlignment = 7
                If Not IsNull(.Fields("RDistributorName").Value) Then
                    GridOrdering.Text = .Fields("RDistributorName").Value
                Else
                    GridOrdering.Text = "Not Requested"
                End If
                GridOrdering.Col = 5
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

Private Sub FillPurchase(ItemID As Long)
    With rsTemOrder
        If .State = 1 Then .Close
        temSql = "SELECT tblRefill.Date, tblBatch.Batch, tblRefill.Amount, tblRefill.FreeAmount, tblRefill.DOE " & _
                    "FROM tblRefill LEFT JOIN tblBatch ON tblRefill.BatchID = tblBatch.BatchID " & _
                    "WHERE (((tblRefill.Date) Between '" & Format(dtpFDate.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTDate.Value, "dd MMMM yyyy") & "') AND ((tblRefill.ItemID)=" & ItemID & ")) " & _
                    "Order by tblRefill.Date"

        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                GridPurchase.Rows = GridPurchase.Rows + 1
                GridPurchase.Row = GridPurchase.Rows - 1
                GridPurchase.Col = 0
                GridPurchase.CellAlignment = 4
                GridPurchase.Text = Format(!Date, ShortDateFormat)
                GridPurchase.Col = 1
                GridPurchase.CellAlignment = 4
                GridPurchase.Text = Format(!Batch, "")
                GridPurchase.Col = 2
                GridPurchase.CellAlignment = 7
                GridPurchase.Text = !Amount
                GridPurchase.Col = 3
                GridPurchase.CellAlignment = 7
                GridPurchase.Text = !FreeAmount
                GridPurchase.Col = 4
                GridPurchase.CellAlignment = 4
                GridPurchase.Text = Format(!DOE, ShortDateFormat)
                .MoveNext
            Wend
        End If
    End With

End Sub

Private Sub FillPrice(ByVal ItemID As Long)
    With rsTemPrice
        If .State = 1 Then .Close
        temSql = "SELECT tblPurchasePrice.SetDate, tblPurchasePrice.PPrice FROM tblPurchasePrice WHERE tblPurchasePrice.ItemID = " & ItemID & " AND tblPurchasePrice.SetDate between '" & Format(dtpFDate.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTDate.Value, "dd MMMM yyyy") & "'"
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
    With rsTemPrice
        If .State = 1 Then .Close
        temSql = "SELECT tblSalePrice.SetDate, tblSalePrice.SPrice FROM tblSalePrice WHERE tblSalePrice.ItemID = " & ItemID & "   AND tblSalePrice.SetDate between '" & Format(dtpFDate.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTDate.Value, "dd MMMM yyyy") & "'"
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

Private Sub FillStocks(ByVal ItemID As Long)
    Dim StoreStock As Double
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
                    If !Store = UserStore Then
                        If Not IsNull(!Stock) Then
                            StoreStock = StoreStock + !Stock
                        End If
                    End If
                Else
                    GridTotal.Text = Empty
                End If
                .MoveNext
            Wend
        End If
        GridTotal.Visible = True
        .Close
    End With
    lblStoreStock.Caption = UserStore
    txtStoreStock.Text = StoreStock
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

Private Sub Form_Load()
    dtpFDate.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTDate.Value = Date
    FillCombos
    FormatGrids
    GetCommonSettings Me
End Sub

Private Sub FormatGrids()
    txtVMP.Text = Empty
    txtVMPP.Text = Empty
    txtVTM.Text = Empty
    txtAMP.Text = Empty
    txtAMPP.Text = Empty
    txtDisplay.Text = Empty
    
    
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
    
    With GridTotal
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
        .Text = "Requested Amount"
        .Col = 3
        .CellAlignment = 4
        .Text = "Approved Amount"
        .Col = 4
        .CellAlignment = 4
        .Text = "Requested Distributor"
        .Col = 5
        .CellAlignment = 4
        .Text = "Approved Distributor"
        For i = 0 To .Cols - 1
            .ColWidth(i) = (.Width - 100) / 6
        Next i
    End With

    With GridPurchase
        .Rows = 1
        .Cols = 5
        .FixedCols = 0
        .Col = 0
        .CellAlignment = 4
        .Text = "Date"
        .Col = 1
        .CellAlignment = 4
        .Text = "Batch"
        .Col = 2
        .CellAlignment = 4
        .Text = "Quentity"
        .Col = 3
        .CellAlignment = 4
        .Text = "Free"
        .Col = 4
        .CellAlignment = 4
        .Text = "Expiary"
        For i = 0 To .Cols - 1
            .ColWidth(i) = (.Width - 100) / 5
        Next i
    End With


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub
