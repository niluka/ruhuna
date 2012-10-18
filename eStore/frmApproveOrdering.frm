VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmApproveOrdering 
   Caption         =   "Autherize Requests"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
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
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   6720
      TabIndex        =   77
      Top             =   6000
      Width           =   8535
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   720
         TabIndex        =   79
         Top             =   240
         Width           =   4935
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   6240
         TabIndex        =   78
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Paper"
         Height          =   255
         Left            =   5760
         TabIndex        =   80
         Top             =   240
         Width           =   1815
      End
   End
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   375
      Left            =   11880
      TabIndex        =   5
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16576
      Caption         =   "Edit"
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
   Begin VB.TextBox txtAppCost 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   11760
      TabIndex        =   33
      Top             =   4080
      Width           =   3135
   End
   Begin VB.TextBox txtTemRow 
      Height          =   360
      Left            =   4800
      TabIndex        =   31
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtTemTopRow 
      Height          =   360
      Left            =   6120
      TabIndex        =   30
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame FrameAction 
      Height          =   735
      Left            =   11760
      TabIndex        =   28
      Top             =   4560
      Width           =   3135
      Begin btButtonEx.ButtonEx bttnConfirm 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16576
         Caption         =   "Approve"
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
      Begin btButtonEx.ButtonEx bttnExit 
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Appearance      =   3
         BorderColor     =   16576
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
   End
   Begin VB.TextBox txtIQty 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtOrderID 
      Height          =   375
      Left            =   7800
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame FrameExpected 
      Caption         =   "Excepted Duration"
      Height          =   1215
      Left            =   11760
      TabIndex        =   24
      Top             =   2400
      Width           =   3135
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yy"
         Format          =   68812803
         CurrentDate     =   39539
      End
      Begin MSComCtl2.DTPicker dtpFromTime 
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yy"
         Format          =   68812802
         CurrentDate     =   39539
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yy"
         Format          =   68812803
         CurrentDate     =   39539
      End
      Begin MSComCtl2.DTPicker dtpToTime 
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yy"
         Format          =   68812802
         CurrentDate     =   39539
      End
   End
   Begin VB.Frame FrameOrder 
      Caption         =   "Order By"
      Height          =   1575
      Left            =   11760
      TabIndex        =   22
      Top             =   720
      Width           =   3135
      Begin VB.OptionButton OptRequestedDate 
         Caption         =   "Requested Date"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton OptDistributorOrder 
         Caption         =   "Distributor"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton OptItemOrder 
         Caption         =   "Item"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSDataListLib.DataCombo dtcDistributor 
      Height          =   360
      Left            =   9120
      TabIndex        =   4
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
   End
   Begin VB.TextBox txtPQty 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtItem 
      Enabled         =   0   'False
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9128
      _Version        =   393216
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   375
      Left            =   11880
      TabIndex        =   7
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16576
      Caption         =   "Save"
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
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   375
      Left            =   13440
      TabIndex        =   6
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16576
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
   Begin VB.Frame FrameShow 
      Caption         =   "Items Not Requested"
      Height          =   1095
      Left            =   11760
      TabIndex        =   23
      Top             =   720
      Width           =   3135
      Begin VB.OptionButton OptRequestedItems 
         Caption         =   "Do Not Show"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optAllItems 
         Caption         =   "Show"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   34
      Top             =   6840
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   16576
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "frmApproveOrdering.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Stocks"
      TabPicture(1)   =   "frmApproveOrdering.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridTotal"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Usage"
      TabPicture(2)   =   "frmApproveOrdering.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblStoreUsage"
      Tab(2).Control(1)=   "GridUsage"
      Tab(2).Control(2)=   "txtStoreUsage"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Ordering"
      TabPicture(3)   =   "frmApproveOrdering.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "GridOrdering"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Purchase"
      TabPicture(4)   =   "frmApproveOrdering.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GridPurchase"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Prices"
      TabPicture(5)   =   "frmApproveOrdering.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "GridPPrice"
      Tab(5).Control(1)=   "GridSPrice"
      Tab(5).Control(2)=   "Label17"
      Tab(5).Control(3)=   "Label16"
      Tab(5).ControlCount=   4
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
         TabIndex        =   63
         Top             =   600
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   48
         Top             =   360
         Width           =   6855
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
            TabIndex        =   55
            Top             =   2640
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
            TabIndex        =   54
            Top             =   2160
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
            TabIndex        =   53
            Top             =   1680
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
            TabIndex        =   52
            Top             =   1200
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
            TabIndex        =   51
            Top             =   720
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
            TabIndex        =   50
            Top             =   240
            Width           =   4095
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   3120
            Width           =   4095
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
            TabIndex        =   62
            Top             =   2640
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
            Left            =   120
            TabIndex        =   61
            Top             =   2160
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
            Left            =   120
            TabIndex        =   60
            Top             =   1680
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
            Left            =   120
            TabIndex        =   59
            Top             =   1200
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
            Left            =   120
            TabIndex        =   58
            Top             =   720
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
            Left            =   120
            TabIndex        =   57
            Top             =   240
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
            Left            =   120
            TabIndex        =   56
            Top             =   3120
            Width           =   3255
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3735
         Left            =   -67800
         TabIndex        =   35
         Top             =   360
         Width           =   7695
         Begin VB.Label lblDistributor 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   47
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label20 
            Caption         =   "Distributor"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Address"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "Tel No"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            Height          =   375
            Left            =   1560
            TabIndex        =   42
            Top             =   3240
            Width           =   3375
         End
         Begin VB.Label lblTelNo 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   41
            Top             =   2280
            Width           =   3375
         End
         Begin VB.Label lblAddress 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   975
            Left            =   1560
            TabIndex        =   40
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label Label31 
            Caption         =   "Fax No"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label lblFax 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   38
            Top             =   2760
            Width           =   3375
         End
         Begin VB.Label lblCity 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   37
            Top             =   1800
            Width           =   3375
         End
         Begin VB.Label Label33 
            Caption         =   "City"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1800
            Width           =   1335
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridUsage 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   64
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridTotal 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   65
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridOrdering 
         Height          =   3015
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPPrice 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   67
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridSPrice 
         Height          =   3015
         Left            =   -66960
         TabIndex        =   68
         Top             =   1080
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPurchase 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   69
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5318
         _Version        =   393216
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
         TabIndex        =   72
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Purchase Prices"
         Height          =   255
         Left            =   -74880
         TabIndex        =   71
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label16 
         Caption         =   "Sales Prices"
         Height          =   255
         Left            =   -66840
         TabIndex        =   70
         Top             =   600
         Width           =   4815
      End
   End
   Begin MSComCtl2.DTPicker dtpFDate 
      Height          =   375
      Left            =   1560
      TabIndex        =   73
      Top             =   6360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16576
      CalendarTitleForeColor=   16576
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   68812803
      CurrentDate     =   39540
   End
   Begin MSComCtl2.DTPicker dtpTDate 
      Height          =   375
      Left            =   4560
      TabIndex        =   74
      Top             =   6360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16576
      CalendarTitleForeColor=   16576
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   68812803
      CurrentDate     =   39540
   End
   Begin VB.Label Label7 
      Caption         =   "To :"
      Height          =   255
      Left            =   4080
      TabIndex        =   76
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "From :"
      Height          =   255
      Left            =   120
      TabIndex        =   75
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label21 
      Caption         =   "App. Cost of Purchase"
      Height          =   375
      Left            =   11760
      TabIndex        =   32
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Issue Quentity"
      Height          =   255
      Left            =   5880
      TabIndex        =   27
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblIUnit 
      Height          =   375
      Left            =   7560
      TabIndex        =   26
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Distributor"
      Height          =   255
      Left            =   9120
      TabIndex        =   21
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPunit 
      Height          =   375
      Left            =   4440
      TabIndex        =   20
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Purchase Quentity"
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmApproveOrdering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTemOrder As New ADODB.Recordset
    Dim rsDistributors As New ADODB.Recordset
    Dim rsTemDistributor As New ADODB.Recordset
    Dim rsTemItem As New ADODB.Recordset
    Dim rsTemStocks As New ADODB.Recordset
    Dim temSql As String
    Dim NewItem As New Item
    Dim TemString As String
    Dim TemQty As Double
    Dim rsDistributor As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    Dim rsTemPrice As New ADODB.Recordset
    Dim rsTemDistributorOrde As New ADODB.Recordset
    Dim CsetPrinter As New cSetDfltPrinter
    Dim TemSDate As Date
    Dim TemEDate As Date
    Dim TemSTime As Date
    Dim TemETime As Date
    
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    
    

Private Sub dtpOTom_Change()
    Call FillOrdering(NewItem.ID)
End Sub

Private Sub bttnCancel_Click()
    Call BeforeEdit
End Sub

Private Sub bttnConfirm_Click()
    Dim tr As Integer
'On Error GoTo eh

With rsTemOrder
    If .State = 1 Then .Close
    temSql = "SELECT tblOrderBill.* FROM tblOrderBill where orderbillID = " & OrderBillID
    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
    If .RecordCount = 0 Then Exit Sub
    !ApprovedComplete = True
    !ApprovedDate = Date
    !ApprovedTime = Now
    !ApprovedStaffID = UserID
    !ApprovedStoreID = UserStoreID
    !expecteddate1 = dtpFromDate.Value
    !expecteddate2 = dtpToDate.Value
    !expectedtime1 = dtpFromTime.Value
    !expectedtime2 = dtpToTime.Value
    .Update
End With

    TemSDate = dtpFromDate.Value
    TemEDate = dtpToDate.Value
    TemSTime = dtpFromTime.Value
    TemETime = dtpToTime.Value

With GridItem
    .Visible = False
    Dim i As Integer
    Dim TemOrderID As Long
    For i = 1 To .Rows - 1
        .Row = i
        .Col = 1
        If IsNumeric(.Text) = True Then
            TemOrderID = Val(.Text)
            With rsTemOrder
                If .State = 1 Then .Close
                temSql = "SELECT * from tblOrder where OrderID = " & TemOrderID
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !ApprovedDate = Date
                    !ApprovedTime = Now
                    !ApprovedStaffID = UserID
                    !ApprovedStoreID = UserStoreID
                    GridItem.Col = 6
                    !ApprovedAmount = Val(GridItem.Text)
                    !ReceivedAmount = Val(GridItem.Text)
                    !ApprovedComplete = True
                    GridItem.Col = 7
                    !ApprovedDistributorID = Val(GridItem.Text)
                    !expecteddate1 = dtpFromDate.Value
                    !expecteddate2 = dtpToDate.Value
                    !expectedtime1 = dtpFromTime.Value
                    !expectedtime2 = dtpToTime.Value
                    .Update
                End If
            End With
            
        End If
    Next i
    .Visible = False
End With
    GridItem.Clear
    GridItem.Rows = 1
    Me.Hide
    tr = MsgBox("Approval of the request was successfully confirmed and dealerwise request forms will be generated", vbInformation, "Successful")
    Call DistributorOrders
    Unload Me
Exit Sub

eh:
    tr = MsgBox("An Error occured during Approval", vbCritical, "Error")
    If rsTemOrder.State = 1 Then rsTemOrder.CancelUpdate
    If rsTemOrder.State = 1 Then rsTemOrder.Close
    Exit Sub
End Sub

Private Sub DistributorOrders()
    Dim TemDistributorId As Long
    Dim TemDistributor As String
    Dim TemAddress As String
    Dim TemFax As String
    Dim TemTel As String
    Dim TemDistributorOrderID As Long
    Dim TemResponce As Long
    Dim RetVal As Integer
    Dim tr As Integer
    
    
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    On Error Resume Next
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    
    If SelectForm(cmbPaper.Text, Me.hwnd) <> 1 Then
        MsgBox "Printer Error"
    End If
    
    With rsTemDistributor
        
        
        If .State = 1 Then .Close
        temSql = "SELECT (tblDistrubutor.DistributorID) AS FirstOfDistributorID,  (tblDistrubutor.DistributorName) AS FirstOfDistributorName,  (tblDistrubutor.DistributorTelephone) AS FirstOfDistributorTelephone1, MIN(tblDistrubutor.DistributorFax) AS FirstOfDistributorFax " & _
                    " FROM tblOrder LEFT JOIN tblDistrubutor ON tblOrder.ApprovedDistributorID = tblDistrubutor.DistributorID " & _
                    " Where tblOrder.OrderBillID = " & OrderBillID & " AND (tblOrder.ApprovedAmount > 0)  " & _
                    " GROUP BY tblDistrubutor.DistributorName, tblDistrubutor.DistributorID, tblDistrubutor.DistributorTelephone, tblDistrubutor.DistributorFax  " & _
                    " ORDER BY tblDistrubutor.DistributorName  "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        
        
        If .RecordCount > 0 Then
            
            
            While .EOF = False
            
            TemDistributorId = .Fields("FirstOfDistributorID").Value
            If Not IsNull(!FirstOfDistributorName) Then
                TemDistributor = !FirstOfDistributorName
            Else
                TemDistributor = Empty
            End If
            
'            If Not IsNull(!FirstOfDistributorAddress) Then
'                TemAddress = !FirstOfDistributorAddress
'            Else
'                TemAddress = Empty
'            End If
            
            If Not IsNull(!FirstOfDistributorTelephone1) Then
                TemTel = !FirstOfDistributorTelephone1
            Else
                TemTel = Empty
            End If
            
            If Not IsNull(!FirstOfDistributorFax) Then
                TemFax = !FirstOfDistributorFax
            Else
                TemFax = Empty
            End If
                With rsTemOrder
                    If .State = 1 Then .Close
                    temSql = "SELECT tblDistributorOrder.* FROM tblDistributorOrder"
                    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                    .AddNew
                    !OrderBillID = OrderBillID
                    !DistributorID = TemDistributorId
                    .Update
                    temSql = "SELECT @@IDENTITY AS NewID"
                    .Close
                    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    TemDistributorOrderID = !NewID
                    .Close

                    With Dataenvironment1.rscmmdDistributorOrder
                        If .State = 1 Then .Close
                        .Source = "SELECT tblItem.Display, tblItem.AMPP, [ApprovedAmount]/[tblItem].[IssueUnitsPerPack] AS AAinPUnit, tblPackUnit.PackUnit " & _
                                    " FROM ((tblPackUnit RIGHT JOIN (tblOrder LEFT JOIN tblItem ON tblOrder.ItemID = tblItem.ItemID) ON tblPackUnit.PackUnitID = tblItem.PackUnitID) LEFT JOIN tblIssueUnit ON tblItem.IssueUnitID = tblIssueUnit.IssueUnitID) LEFT JOIN tblDistrubutor ON tblOrder.ApprovedDistributorID = tblDistrubutor.DistributorID " & _
                                    " WHERE (((tblOrder.OrderBillID)= " & OrderBillID & ") AND ((tblOrder.ApprovedAmount) > 0) AND ((tblOrder.ApprovedDistributorID)=" & TemDistributorId & ")) " & _
                                    " Order by tblItem.Display"
                      
                        .Open

                    
                        If .RecordCount > 0 Then
'                                CsetPrinter.SetPrinterAsDefault (ReportPrinterName)

                                CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
                                'On Error Resume Next
                                PrinterName = Printer.DeviceName
                                If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
                                    ClosePrinter (PrinterHandle)
                                End If
                                CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
                                
                                
                                PrinterName = Printer.DeviceName
                                If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
                                    ClosePrinter (PrinterHandle)
                                End If
                                CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
                                
                                
                                CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
                                'On Error Resume Next
                                PrinterName = Printer.DeviceName
                                If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
                                    ClosePrinter (PrinterHandle)
                                End If
                                CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
                                
                                
                                If SelectForm(cmbPaper.Text, Me.hwnd) <> 1 Then
                                    MsgBox "Printer Error"
                                End If


                                        With dtrDistributorOrdering
                                            Set .DataSource = Dataenvironment1.rscmmdDistributorOrder
                                            .Sections("Section4").Controls("lblName").Caption = HospitalName
                                            .Sections("Section4").Controls("lblContact").Caption = HospitalAddress
                                            .Sections("Section4").Controls("lblTopic").Caption = "Order Request Forms"
                                            .Sections("Section4").Controls("lblSUbtopic").Caption = Empty
                                            .Sections("Section4").Controls("lblTo").Caption = TemDistributor
                                            .Sections("Section4").Controls("lblAddress").Caption = TemAddress
                                            .Sections("Section4").Controls("lblTel").Caption = TemTel
                                            .Sections("Section4").Controls("lblFax").Caption = TemFax
                                            .Sections("Section4").Controls("lblDate").Caption = Format(Date, LongDateFormat)
                                            .Sections("Section4").Controls("lblOrderID").Caption = TemDistributorOrderID
                                            .Sections("Section3").Controls("lblAd").Caption = LongAd
                                            
                                            .Sections("Section5").Controls("lblmsg1").Caption = "Please be kind enough to supply the above stocks between " & TemSTime & " on " & Format(TemSDate, LongDateFormat)
                                            .Sections("Section5").Controls("lblmsg2").Caption = "and " & TemETime & " on " & Format(TemEDate, LongDateFormat) & "."
                                            
                                CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
                                'On Error Resume Next
                                PrinterName = Printer.DeviceName
                                If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
                                    ClosePrinter (PrinterHandle)
                                End If
                                CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
                                
                                
                                PrinterName = Printer.DeviceName
                                If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
                                    ClosePrinter (PrinterHandle)
                                End If
                                CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
                                
                                
                                CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
                                'On Error Resume Next
                                PrinterName = Printer.DeviceName
                                If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
                                    ClosePrinter (PrinterHandle)
                                End If
                                CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
                                
                                
                                If SelectForm(cmbPaper.Text, Me.hwnd) <> 1 Then
                                    MsgBox "Printer Error"
                                End If
                                            
                                                
                                                .Show
                                                tr = MsgBox("Order Forms" & vbNewLine & "Distributor : " & vbTab & TemDistributor & vbNewLine & "Distributor Order ID : " & vbTab & vbNewLine & TemDistributorOrderID & vbNewLine & "Print Order Form ?", vbQuestion + vbYesNo, "Print?")
                                                If tr = vbYes Then .PrintReport False
                                                tr = MsgBox("Order Forms" & vbNewLine & "Distributor : " & vbTab & TemDistributor & vbNewLine & "Distributor Order ID : " & vbTab & vbNewLine & TemDistributorOrderID & vbNewLine & "Save Order Form ?", vbQuestion + vbYesNo, "Save?")
                                                If tr = vbYes Then .ExportReport , App.Path & "\" & TemDistributor & " " & Format(Date, "dd MMMM yyyy"), True, True

                                        End With


                        End If

                   End With


                End With


                .MoveNext
            
            
            Wend
    
    
        End If
    
    
    
    End With
End Sub

'
'Private Sub SetReportPaper()
'Dim TemResponce As Long
'Dim RetVal As Integer
'RetVal = SelectForm(BillPaperName, Me.hwnd)
'Select Case RetVal
'    Case FORM_NOT_SELECTED   ' 0
'        TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Bill Not Printed")
'    Case FORM_SELECTED   ' 1
'        Call SelectPrint
'    Case FORM_ADDED   ' 2
'        TemResponce = MsgBox("New paper size added.", vbExclamation, "New Paper size")
'End Select
'End Sub


Private Sub bttnEdit_Click()
    If GridItem.Rows < 2 Then Exit Sub
    If GridItem.Row < 1 Then Exit Sub
    If txtOrderID.Text = Empty Then Exit Sub
    If Not IsNumeric(txtOrderID.Text) Then Exit Sub
   
    GridItem.Col = 1
    If Not IsNumeric(GridItem.Text) Then Exit Sub
    Call AfterEdit
    txtPQty.SetFocus
    SendKeys "{home}+{end}"
End Sub

Private Sub bttnExit_Click()
    Unload Me
End Sub

Private Sub bttnSave_Click()
    With rsTemOrder
        If .State = 1 Then .Close
        temSql = "SELECT * from tblOrder where orderid = " & Val(txtOrderID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !ApprovedAmount = Val(txtIQty.Text)
            !ApprovedDistributorID = Val(dtcDistributor.BoundText)
            .Update
        End If
    End With
    Call BeforeEdit
    Call FillGrid
    On Error Resume Next
    GridItem.TopRow = Val(txtTemTopRow.Text)
    GridItem.Row = Val(txtTemRow.Text)
    GridItem.Col = GridItem.Cols - 1
    GridItem.ColSel = 0
    GridItem.SetFocus
End Sub

Private Sub cmbPrinter_Change()
    Call PopulatePapers
End Sub

Private Sub cmbPrinter_Click()
    Call PopulatePapers
End Sub

Private Sub dtcDistributor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        bttnSave_Click
    End If
End Sub

Private Sub dtpOFrom_Change()
    Call FillOrdering(NewItem.ID)
End Sub

Private Sub dtpUFrom_Change()
    Call FillUsage(NewItem.ID)
End Sub

Private Sub dtpUTo_Change()
    Call FillUsage(NewItem.ID)
End Sub


Private Sub dtpFDate_Change()
    GridItem_Click
End Sub


Private Sub dtpTDate_Change()
    Call GridItem_Click
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
'    Call DistributorDetails
'    Call FillDetails
    Call GetSettings
    Call BeforeEdit
    Call FillGrid
    
    dtpFromDate.Value = Date + Val(GetSetting(App.EXEName, "Options", "txtFromDate", 3))
    dtpToDate.Value = Date + Val(GetSetting(App.EXEName, "Options", "txtToDate", 5))
    dtpFromTime.Value = GetSetting(App.EXEName, "Options", "dtpSTime", "13:30")
    dtpToTime.Value = GetSetting(App.EXEName, "Options", "dtpETime", "13:30")
    dtpTDate.Value = Date - Val(GetSetting(App.EXEName, "Options", "UsageDays", 30))
    dtpTDate.Value = Date
End Sub

Private Sub FillCombos()
    With rsDistributor
        If .State = 1 Then .Close
        temSql = "SELECT tblDistrubutor.* From tblDistrubutor ORDER BY tblDistrubutor.DistributorName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcDistributor
        Set .RowSource = rsDistributor
        .ListField = "DistributorName"
        .BoundColumn = "DistributorID"
    End With
End Sub

Private Sub GetSettings()
    optAllItems.Value = GetSetting(App.EXEName, Me.Caption, "optAllItems", True)
    OptDistributorOrder.Value = GetSetting(App.EXEName, Me.Caption, "OptDistributorOrder", False)
    OptItemOrder.Value = GetSetting(App.EXEName, Me.Caption, "OptItemOrder", True)
    OptRequestedItems.Value = GetSetting(App.EXEName, Me.Caption, "OptRequestedItems", False)

    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "")


End Sub

Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub


Private Sub PopulatePapers()
    cmbPaper.Clear
    SetPrinter = False
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
'        With FormSize
'            .cx = BillPaperHeight
'            .cy = BillPaperWidth
'        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                cmbPaper.AddItem PtrCtoVbString(.pName)
'                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Caption, "optAllItems", optAllItems.Value
    SaveSetting App.EXEName, Me.Caption, "OptDistributorOrder", OptDistributorOrder
    SaveSetting App.EXEName, Me.Caption, "OptItemOrder", OptItemOrder.Value
    SaveSetting App.EXEName, Me.Caption, "OptRequestedItems", OptRequestedItems.Value
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
End Sub

Private Sub BeforeEdit()
    FrameAction.Enabled = True
    FrameExpected.Enabled = True
    FrameOrder.Enabled = True
    FrameShow.Enabled = True
    bttnSave.Visible = False
    bttnCancel.Visible = False
    bttnEdit.Visible = True
    GridItem.Enabled = True
    txtIQty.Enabled = False
    txtPQty.Enabled = False
    dtcDistributor.Enabled = False
End Sub

Private Sub AfterEdit()
    FrameAction.Enabled = False
    FrameExpected.Enabled = False
    FrameOrder.Enabled = False
    FrameShow.Enabled = False
    bttnSave.Visible = True
    bttnCancel.Visible = True
    bttnEdit.Visible = False
    GridItem.Enabled = False
    txtIQty.Enabled = True
    txtPQty.Enabled = True
    dtcDistributor.Enabled = True
End Sub

Private Sub FillGrid()
    Me.MousePointer = vbHourglass
    Dim TotalCost As Double
    With GridItem
        DoEvents
        .Clear
        .Cols = 11
        .Rows = 1
        .ColWidth(0) = 600
        .ColWidth(1) = 1
        .ColWidth(3) = 3200
        .ColWidth(4) = 2200
        .ColWidth(5) = 1
        .ColWidth(6) = 1
        .ColWidth(7) = 1
        .ColWidth(8) = 1
        .ColWidth(9) = 1600
        .ColWidth(10) = 1600
        .ColWidth(2) = .Width - (.ColWidth(0) + .ColWidth(1) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7) + .ColWidth(8) + .ColWidth(9) + .ColWidth(10) + 100)
        .Row = 0
        Dim i As Integer
        For i = 0 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            Select Case i
                Case 0: .Text = "No."
                Case 2: .Text = "Item"
                Case 3: .Text = "Quentity"
                Case 4: .Text = "Distributor"
                Case 9: .Text = "App. Cost"
                Case 10:    .Text = "Request Date"
            End Select
        Next i
    End With
    
    With rsTemOrder
'        If .State = 1 Then .Close
'        TemSql = "DELETE tblOrder.* from tblORder orderbillid = " & OrderBillID & " And tblOrder.RequestAmount > 0 where (((tblOrder.AutoRequestComplete) = True) And ((tblOrder.RequestComplete) = False) And ((tblOrder.AutoRequestAmount) <= 0)) "
'        .Open TemSql, cnnStores, adOpenStatic, adLockOptimistic
        
        If .State = 1 Then .Close
        
        temSql = "SELECT tblOrder.*, tblItem.Display, tblDistrubutor.DistributorName FROM tblDistrubutor RIGHT JOIN (tblOrder LEFT JOIN tblItem ON tblOrder.ItemID = tblItem.ItemID) ON tblDistrubutor.DistributorID = tblOrder.ApprovedDistributorID Where orderbillid = " & OrderBillID & " And tblOrder.RequestAmount > 0 "           ' (((tblOrder.AutoRequestComplete) = True) And ((tblOrder.RequestComplete) = False) And ((tblOrder.AutoRequestAmount) > 0)) "
        
        If OptItemOrder.Value = True Then
            temSql = temSql & " ORDER BY tblItem.Display"
        ElseIf OptDistributorOrder.Value = True Then
            temSql = temSql & " ORDER BY tblDistrubutor.DistributorName"
        Else
            temSql = temSql & " ORDER BY tblOrder.RequestDate"
        End If
        
    '   0   No.
    '   1   OrderID
    '   2   Item
    '   3   Quentity
    '   4   DIstributor
    '   5   ItemID
    '   6   Quqntity Value
    '   7   DistributorID
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            dtpFromDate.Value = !expecteddate1
            dtpToDate.Value = !expecteddate2
            dtpFromTime.Value = !expectedtime1
            dtpToTime.Value = !expectedtime2
            .MoveLast
            .MoveFirst
            GridItem.Rows = .RecordCount + 1
            For i = 1 To .RecordCount
                GridItem.TextMatrix(i, 0) = i
                GridItem.TextMatrix(i, 1) = !OrderID
                GridItem.TextMatrix(i, 2) = !Display
                NewItem.ID = !ItemID
                TemString = ""
                TemQty = !ApprovedAmount
                TemString = Format((TemQty / NewItem.IssueUnitsPerPack), "0") & " " & NewItem.PUnit & " (" & TemQty & " " & NewItem.IUnit & ")"
                GridItem.TextMatrix(i, 3) = TemString
                If IsNull(!DistributorName) Then
                    GridItem.TextMatrix(i, 4) = "Not Selected"
                Else
                    GridItem.TextMatrix(i, 4) = !DistributorName
                End If
                GridItem.TextMatrix(i, 5) = NewItem.ID
                GridItem.TextMatrix(i, 6) = TemQty
                GridItem.TextMatrix(i, 7) = !ApprovedDistributorID
                GridItem.TextMatrix(i, 9) = Format(TemQty * NewItem.PPrice, "0.00")
                TotalCost = TotalCost + TemQty * NewItem.PPrice
                GridItem.TextMatrix(i, 10) = !requestdate
                .MoveNext
            Next
        End If
    End With
    Me.MousePointer = vbDefault
    txtAppCost.Text = Format(TotalCost, "#,##0.00")
    DoEvents
End Sub

Private Sub ClearItemValues()
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim tr As Integer
    If GridItem.Rows > 1 Then
        tr = MsgBox("You have not Approved the ordering request. Are you sure you want to exit?", vbQuestion + vbYesNo, "Quit?")
        If tr = vbNo Then Cancel = True: Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSettings
End Sub

Private Sub GridItem_Click()
Dim i As Integer
With GridItem
    txtTemRow.Text = .Row
    txtTemTopRow.Text = .TopRow
    .Col = 0
    .ColSel = .Cols - 1
    i = .Row
    If IsNumeric(.TextMatrix(i, 1)) = False Then Exit Sub
    Call GetOrderDetails(Val(.TextMatrix(i, 1)))
    Call FormatGrids
    Call FillStocks(Val(.TextMatrix(i, 5)))
    Call FillUsage(Val(.TextMatrix(i, 5)))
    Call FillOrdering(Val(.TextMatrix(i, 5)))
    Call FillPrice(Val(.TextMatrix(i, 5)))
    Call FillPurchase(Val(.TextMatrix(i, 5)))
    Call DistributorDetails(Val(.TextMatrix(i, 7)))
End With
End Sub

Private Sub GridItem_DblClick()
    Call GridItem_Click
    Call bttnEdit_Click
End Sub

Private Sub GridItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GridItem_DblClick
    Else
        GridItem_Click
    End If
End Sub

Private Sub OptAllItems_Click()
    Call FillGrid
End Sub

Private Sub OptDistributorOrder_Click()
    Call FillGrid
End Sub

Private Sub OptItemOrder_Click()
    Call FillGrid
End Sub

Private Sub OptRequestedItems_Click()
    Call FillGrid
End Sub

Private Sub FillStocks(ByVal ItemID As Long)
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
'                        lblStoreStock.Caption = !Store
                        If Not IsNull(!Stock) Then
'                            txtStoreStock.Text = !Stock
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


Private Sub GetItemDetails(ItemID As Long)
    NewItem.ID = ItemID
    txtAMP.Text = NewItem.AMP
    txtAMPP.Text = NewItem.AMPP
    txtVMP.Text = NewItem.VMP
    txtVMPP.Text = NewItem.VMPP
    txtVTM.Text = NewItem.Generic
    txtDisplay.Text = NewItem.Display
End Sub

Private Sub GetOrderDetails(OrderID As Long)
With rsTemOrder
    If .State = 1 Then .Close
    temSql = "SELECT * from tblOrder where orderID = " & OrderID
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount < 1 Then Exit Sub
    NewItem.ID = !ItemID
        lblIUnit.Caption = NewItem.IUnit
    lblPunit.Caption = NewItem.PUnit
    txtAMP.Text = NewItem.AMP
    txtAMPP.Text = NewItem.AMPP
    txtVMP.Text = NewItem.VMP
    txtVMPP.Text = NewItem.VMPP
    txtVTM.Text = NewItem.Generic
    txtDisplay.Text = NewItem.Display
    lblPunit.Caption = NewItem.PUnit
    lblIUnit.Caption = NewItem.IUnit
    txtIQty.Text = !ApprovedAmount
    txtPQty.Text = Val(txtIQty.Text) / NewItem.IssueUnitsPerPack
    txtItem.Text = txtDisplay.Text
    txtOrderID.Text = OrderID
    dtcDistributor.BoundText = !ApprovedDistributorID
    
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
        temSql = "SELECT tblOrder.RequestDate, tblOrder.ApprovedDate, tblOrder.ReceivedDate, tblOrder.RequestAmount, tblOrder.ApprovedAmount, tblOrder.ReceivedAmount, tblRDistrubutor.DistributorName as RDistributorName, tblADistrubutor.DistributorName as ADistributorName FROM (tblDistrubutor AS tblRDistrubutor RIGHT JOIN tblOrder ON tblRDistrubutor.DistributorID = tblOrder.ApprovedDistributorID) LEFT JOIN tblDistrubutor AS tblADistrubutor ON tblOrder.RequestDistributorID = tblADistrubutor.DistributorID WHERE tblOrder.ItemID  = " & ItemID & "AND tblOrder.RequestDate between '" & Format(dtpFDate.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTDate.Value, "dd MMMM yyyy") & "'"
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
                    "WHERE (((tblRefill.Date) Between '" & Format(dtpFDate.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTDate.Value, "dd MMMM yyyy") & "') AND ((tblRefill.ItemID)=" & ItemID & "))"

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
                If IsNull(!Batch) = False Then
                    GridPurchase.Text = !Batch
                End If
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

'Private Sub FillDetails()
'    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
'    NewItem.ID = dtcItem.BoundText
'    Call FormatGrids
'    Call FillLabels
'    Call GetItemDetails(NewItem.ID)
'    Call FillStocks(dtcItem.BoundText)
'    Call FillPurchase(dtcItem.BoundText)
'    Call FillPrice(dtcItem.BoundText)
'    Call GetItemDetails(dtcItem.BoundText)
'    Call FillOrdering(dtcItem.BoundText)
'    Call FillUsage(dtcItem.BoundText)
'    Dim rsDI As New ADODB.Recordset
'    With rsDI
'        If .State = 1 Then .Close
'        temSQL = "SELECT tblItemDistributor.DistributorID FROM tblItemDistributor WHERE (((tblItemDistributor.ItemID)=" & Val(dtcItem.BoundText) & "))"
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            DistributorDetails (!DistributorID)
'        End If
'        .Close
'    End With
'End Sub

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

Private Sub txtIQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcDistributor.SetFocus
    End If
End Sub

Private Sub txtIQty_LostFocus()
    txtPQty.Text = Val(txtIQty.Text) / NewItem.IssueUnitsPerPack
End Sub

Private Sub txtPQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtIQty.SetFocus
        SendKeys "{Home}+{end}"
    End If
End Sub

Private Sub txtPQty_LostFocus()
    txtIQty.Text = Val(txtPQty.Text) * NewItem.IssueUnitsPerPack
End Sub
