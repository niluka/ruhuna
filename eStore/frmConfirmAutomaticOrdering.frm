VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmConfirmAutomaticOrdering 
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
   Begin VB.Frame FrameAction 
      Height          =   735
      Left            =   11760
      TabIndex        =   74
      Top             =   5040
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
         Caption         =   "Confirm"
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
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtOrderID 
      Height          =   375
      Left            =   7800
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   18
      Top             =   6000
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "frmConfirmAutomaticOrdering.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(7)=   "txtIStore"
      Tab(0).Control(8)=   "txtDisplay"
      Tab(0).Control(9)=   "txtVTM"
      Tab(0).Control(10)=   "txtVMP"
      Tab(0).Control(11)=   "txtAMP"
      Tab(0).Control(12)=   "txtVMPP"
      Tab(0).Control(13)=   "txtAMPP"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Stocks"
      TabPicture(1)   =   "frmConfirmAutomaticOrdering.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridTotal"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Usage"
      TabPicture(2)   =   "frmConfirmAutomaticOrdering.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(1)=   "Label7"
      Tab(2).Control(2)=   "dtpUTo"
      Tab(2).Control(3)=   "GridUsage"
      Tab(2).Control(4)=   "dtpUFrom"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Ordering"
      TabPicture(3)   =   "frmConfirmAutomaticOrdering.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label8"
      Tab(3).Control(1)=   "Label15"
      Tab(3).Control(2)=   "dtpOTo"
      Tab(3).Control(3)=   "GridOrdering"
      Tab(3).Control(4)=   "dtpOFrom"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Prices"
      TabPicture(4)   =   "frmConfirmAutomaticOrdering.frx":0070
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
      TabCaption(5)   =   "Distributor"
      TabPicture(5)   =   "frmConfirmAutomaticOrdering.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label33"
      Tab(5).Control(1)=   "lblCity"
      Tab(5).Control(2)=   "lblFax"
      Tab(5).Control(3)=   "Label31"
      Tab(5).Control(4)=   "lblAddress"
      Tab(5).Control(5)=   "lblTelNo"
      Tab(5).Control(6)=   "lblBalance"
      Tab(5).Control(7)=   "Label27"
      Tab(5).Control(8)=   "Label26"
      Tab(5).Control(9)=   "Label25"
      Tab(5).Control(10)=   "Label20"
      Tab(5).Control(11)=   "lblDistributor"
      Tab(5).ControlCount=   12
      Begin MSFlexGridLib.MSFlexGrid GridPPrice 
         Height          =   2775
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4895
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpUFrom 
         Height          =   375
         Left            =   -74040
         TabIndex        =   27
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   75038723
         CurrentDate     =   39540
      End
      Begin MSFlexGridLib.MSFlexGrid GridUsage 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   29
         Top             =   840
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridTotal 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   26
         Top             =   360
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   6165
         _Version        =   393216
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
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   25
         Top             =   3360
         Width           =   6615
      End
      Begin MSComCtl2.DTPicker dtpUTo 
         Height          =   375
         Left            =   -71040
         TabIndex        =   28
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   75038723
         CurrentDate     =   39540
      End
      Begin MSComCtl2.DTPicker dtpOFrom 
         Height          =   375
         Left            =   -74040
         TabIndex        =   30
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   75038723
         CurrentDate     =   39540
      End
      Begin MSFlexGridLib.MSFlexGrid GridOrdering 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   32
         Top             =   840
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpOTo 
         Height          =   375
         Left            =   -71040
         TabIndex        =   31
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   75038723
         CurrentDate     =   39540
      End
      Begin MSFlexGridLib.MSFlexGrid GridSPrice 
         Height          =   2775
         Left            =   5040
         TabIndex        =   37
         Top             =   1080
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4895
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpPFrom 
         Height          =   375
         Left            =   840
         TabIndex        =   33
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   75038723
         CurrentDate     =   39540
      End
      Begin MSComCtl2.DTPicker dtpPTo 
         Height          =   375
         Left            =   3840
         TabIndex        =   35
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   75038723
         CurrentDate     =   39540
      End
      Begin VB.Label lblDistributor 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -71520
         TabIndex        =   38
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label20 
         Caption         =   "Distributor"
         Height          =   255
         Left            =   -72960
         TabIndex        =   73
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Address"
         Height          =   255
         Left            =   -72960
         TabIndex        =   72
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label26 
         Caption         =   "Tel No"
         Height          =   255
         Left            =   -72960
         TabIndex        =   71
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Balance"
         Height          =   255
         Left            =   -72960
         TabIndex        =   70
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   375
         Left            =   -71520
         TabIndex        =   43
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Label lblTelNo 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -71520
         TabIndex        =   41
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   -71520
         TabIndex        =   39
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label31 
         Caption         =   "Fax No"
         Height          =   255
         Left            =   -72960
         TabIndex        =   69
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblFax 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -71520
         TabIndex        =   42
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label lblCity 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -71520
         TabIndex        =   40
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label33 
         Caption         =   "City"
         Height          =   255
         Left            =   -72960
         TabIndex        =   68
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "From :"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "To :"
         Height          =   255
         Left            =   3360
         TabIndex        =   66
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Purchase Prices"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label16 
         Caption         =   "Sales Prices"
         Height          =   255
         Left            =   5040
         TabIndex        =   64
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label15 
         Caption         =   "From :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   63
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "To :"
         Height          =   255
         Left            =   -71520
         TabIndex        =   62
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "To :"
         Height          =   255
         Left            =   -71520
         TabIndex        =   61
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "From :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
         Top             =   3360
         Width           =   3255
      End
   End
   Begin VB.Frame FrameExpected 
      Caption         =   "Excepted Duration"
      Height          =   1215
      Left            =   11760
      TabIndex        =   49
      Top             =   3480
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
         Format          =   75038723
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
         Format          =   75038722
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
         Format          =   75038723
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
         Format          =   75038722
         CurrentDate     =   39539
      End
   End
   Begin VB.Frame FrameShow 
      Caption         =   "Items Not Requested"
      Height          =   1095
      Left            =   11760
      TabIndex        =   48
      Top             =   960
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
   Begin VB.Frame FrameOrder 
      Caption         =   "Order By"
      Height          =   1215
      Left            =   11760
      TabIndex        =   47
      Top             =   2160
      Width           =   3135
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
      Text            =   ""
   End
   Begin VB.TextBox txtPQty 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtItem 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8705
      _Version        =   393216
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
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   375
      Left            =   11880
      TabIndex        =   7
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
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
   Begin VB.Label Label5 
      Caption         =   "Issue Quentity"
      Height          =   255
      Left            =   5880
      TabIndex        =   52
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblIUnit 
      Height          =   375
      Left            =   7560
      TabIndex        =   51
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Distributor"
      Height          =   255
      Left            =   9120
      TabIndex        =   46
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPunit 
      Height          =   375
      Left            =   4440
      TabIndex        =   45
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Purchase Quentity"
      Height          =   255
      Left            =   2760
      TabIndex        =   44
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmConfirmAutomaticOrdering"
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
    Dim Temsql As String
    Dim NewItem As New Item
    Dim TemString As String
    Dim TemQty As Double
    Dim rsDistributor As New ADODB.Recordset
    Dim rsTemStore As New ADODB.Recordset
    Dim rsTemPrice As New ADODB.Recordset

Private Sub dtpOTom_Change()
    Call FillOrdering(NewItem.ID)
End Sub

Private Sub bttnCancel_Click()
    Call BeforeEdit
End Sub

Private Sub bttnConfirm_Click()
    Dim TR As Integer

'On Error GoTo EH
If GridItem.Rows <= 1 Then
    TR = MsgBox("There are no items to be request", vbCritical, "No items")
    Unload Me
    Exit Sub
End If

With rsTemOrder
    If .State = 1 Then .Close
    Temsql = "SELECT tblOrderBill.* FROM tblOrderBill where orderbillID = " & OrderBillID
    .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic
    If .RecordCount = 0 Then Exit Sub
    !RequestComplete = True
    !requestdate = Date
    !RequestTime = Time
    !RequestStaffID = UserID
    !RequestStoreID = UserStoreID
    !expecteddate1 = dtpFromDate.Value
    !expecteddate2 = dtpToDate.Value
    !expectedtime1 = dtpFromTime.Value
    !expectedtime2 = dtpToTime.Value
    .Update
End With


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
                Temsql = "SELECT * from tblOrder where OrderID = " & TemOrderID
                .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !requestdate = Date
                    !RequestTime = Time
                    !RequestStaffID = UserID
                    !RequestStoreID = UserStoreID
                    GridItem.Col = 6
                    !RequestAmount = Val(GridItem.Text)
                    !RequestComplete = True
                    !ApprovedAmount = Val(GridItem.Text)
                    GridItem.Col = 7
                    !RequestDistributorID = Val(GridItem.Text)
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
    TR = MsgBox("Automatic order was successfully confirmed and will be awaiting for approval", vbInformation, "Successful")
    Unload Me
Exit Sub

eh:
    TR = MsgBox("An Error occured during confirmation", vbCritical, "Error")
    If rsTemOrder.State = 1 Then rsTemOrder.CancelUpdate
    If rsTemOrder.State = 1 Then rsTemOrder.Close
    Exit Sub
End Sub

Private Sub bttnEdit_Click()
    If GridItem.Rows < 2 Then Exit Sub
    If GridItem.Row < 1 Then Exit Sub
    If txtOrderID.Text = Empty Then Exit Sub
    If Not IsNumeric(txtOrderID.Text) Then Exit Sub
   
    GridItem.Col = 1
    If Not IsNumeric(GridItem.Text) Then Exit Sub
    Call AfterEdit
End Sub

Private Sub bttnExit_Click()
    Unload Me
End Sub

Private Sub bttnSave_Click()
    With rsTemOrder
        If .State = 1 Then .Close
        Temsql = "SELECT * from tblOrder where orderid = " & Val(txtOrderID.Text)
        .Open Temsql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !RequestAmount = Val(txtIQty.Text)
            !RequestDistributorID = Val(dtcDistributor.BoundText)
            .Update
        End If
    End With
    Call BeforeEdit
    Call FillGrid
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

Private Sub Form_Load()
    Call FillCombos
    Call GetSettings
    Call BeforeEdit
    Call FillGrid
    dtpFromDate.Value = Date + Val(GetSetting(App.EXEName, "Options", "txtFromDate", 3))
    dtpToDate.Value = Date + Val(GetSetting(App.EXEName, "Options", "txtToDate", 5))
    dtpFromTime.Value = GetSetting(App.EXEName, "Options", "dtpSTime", "16:00")
    dtpToTime.Value = GetSetting(App.EXEName, "Options", "dtpETime", "16:00")
    dtpUFrom.Value = Date - Val(GetSetting(App.EXEName, "Options", "UsageDays", 30))
    dtpUTo.Value = Date
    dtpOTo.Value = Date
    dtpOFrom.Value = Date - Val(GetSetting(App.EXEName, "Options", "OrderingDays", 30))
    dtpPFrom.Value = Date - Val(GetSetting(App.EXEName, "Options", "PriceDays", 30))
    dtpPTo.Value = Date
End Sub

Private Sub FillCombos()
    With rsDistributor
        If .State = 1 Then .Close
        Temsql = "SELECT tblDistrubutor.* From tblDistrubutor ORDER BY tblDistrubutor.DistributorName"
        .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcDistributor
        Set .RowSource = rsDistributor
        .ListField = "DistributorName"
        .BoundColumn = "DistributorID"
    End With
End Sub

Private Sub GetSettings()
    OptAllItems.Value = GetSetting(App.EXEName, Me.Caption, "optAllItems", True)
    OptDistributorOrder.Value = GetSetting(App.EXEName, Me.Caption, "OptDistributorOrder", False)
    OptItemOrder.Value = GetSetting(App.EXEName, Me.Caption, "OptItemOrder", True)
    OptRequestedItems.Value = GetSetting(App.EXEName, Me.Caption, "OptRequestedItems", False)
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Caption, "optAllItems", OptAllItems.Value
    SaveSetting App.EXEName, Me.Caption, "OptDistributorOrder", OptDistributorOrder
    SaveSetting App.EXEName, Me.Caption, "OptItemOrder", OptItemOrder.Value
    SaveSetting App.EXEName, Me.Caption, "OptRequestedItems", OptRequestedItems.Value
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
    '   0   No.
    '   1   OrderID
    '   2   Item
    '   3   Quentity
    '   4   DIstributor
    '   5   ItemID
    '   6   Quqntity Value
    '   7   DistributorID
    Me.MousePointer = vbHourglass
    
    With GridItem
    
        .Visible = False
        DoEvents
        .Clear
        .Cols = 8
        .Rows = 1
        .ColWidth(0) = 600
        .ColWidth(1) = 1
        .ColWidth(3) = 3500
        .ColWidth(4) = 2600
        .ColWidth(5) = 1
        .ColWidth(6) = 1
        .ColWidth(7) = 1
        .ColWidth(2) = .Width - (.ColWidth(0) + .ColWidth(1) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6) + .ColWidth(7) + 100)
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
            End Select
        Next i
    End With
'    SELECT tblOrder.*, tblItem.Display, tblDistrubutor.DistributorName FROM tblDistrubutor RIGHT JOIN (tblOrder LEFT JOIN tblItem ON tblOrder.ItemID = tblItem.ItemID) ON tblDistrubutor.DistributorID = tblOrder.RequestDistributorID Where (((tblOrder.AutoRequestComplete) = True) And ((tblOrder.RequestComplete) = False) And ((tblOrder.AutoRequestAmount) > 0)) ORDER BY tblDistrubutor.DistributorName, tblItem.Display

    With rsTemOrder
        If .State = 1 Then .Close
        If OptAllItems.Value = True Then
            Temsql = "SELECT tblOrder.*, tblItem.Display, tblDistrubutor.DistributorName FROM tblDistrubutor RIGHT JOIN (tblOrder LEFT JOIN tblItem ON tblOrder.ItemID = tblItem.ItemID) ON tblDistrubutor.DistributorID = tblOrder.RequestDistributorID Where  orderbillID = " & OrderBillID & " " ' (((tblOrder.AutoRequestComplete) = True) And ((tblOrder.RequestComplete) = False)) "
        Else
            Temsql = "SELECT tblOrder.*, tblItem.Display, tblDistrubutor.DistributorName FROM tblDistrubutor RIGHT JOIN (tblOrder LEFT JOIN tblItem ON tblOrder.ItemID = tblItem.ItemID) ON tblDistrubutor.DistributorID = tblOrder.RequestDistributorID Where tblOrder.orderbillid = " & OrderBillID & " AND tblOrder.RequestAmount > 0 "
        End If
        
        If OptItemOrder.Value = True Then
            Temsql = Temsql & " ORDER BY tblItem.Display"
        Else
            Temsql = Temsql & " ORDER BY tblDistrubutor.DistributorName"
        End If
        
    '   0   No.
    '   1   OrderID
    '   2   Item
    '   3   Quentity
    '   4   DIstributor
    '   5   ItemID
    '   6   Quqntity Value
    '   7   DistributorID
        
        .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                GridItem.Rows = GridItem.Rows + 1
                GridItem.Row = GridItem.Rows - 1
                GridItem.Col = 0
                GridItem.CellAlignment = 1
                GridItem.Text = GridItem.Row
                GridItem.Col = 1
                GridItem.Text = !OrderID
                GridItem.Col = 2
                GridItem.CellAlignment = 1
                NewItem.ID = !ItemID
                GridItem.Text = NewItem.Display
                GridItem.Col = 3
                GridItem.CellAlignment = 1
                TemString = ""
                TemQty = !RequestAmount
                TemString = Format((TemQty / NewItem.IssueUnitsPerPack), "0") & " " & NewItem.PUnit & " (" & TemQty & " " & NewItem.IUnit & ")"
                GridItem.Text = TemString
                GridItem.Col = 4
                GridItem.CellAlignment = 1
                If IsNull(!DistributorName) Then
                    GridItem.Text = "Not Selected"
                Else
                    GridItem.Text = !DistributorName
                End If
                GridItem.Col = 5
                GridItem.Text = NewItem.ID
                GridItem.Col = 6
                GridItem.Text = TemQty
                GridItem.Col = 7
                GridItem.Text = !RequestDistributorID
                .MoveNext
            Wend
        Else
            
        End If
    End With
    
    GridItem.Visible = True
    Me.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub ClearItemValues()
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim TR As Integer
    If GridItem.Rows > 1 Then
        TR = MsgBox("You have not confirmed the auto ordering request. Are you sure you want to exit?", vbQuestion + vbYesNo, "Quit?")
        If TR = vbNo Then Cancel = True: Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSettings
End Sub

Private Sub GridItem_Click()
With GridItem
    .Col = 1
    If Not IsNumeric(.Text) Then Exit Sub
    Call GetOrderDetails(Val(.Text))
    .Col = 5
    Call FillStocks(Val(.Text))
    Call FillUsage(Val(.Text))
    Call FillOrdering(Val(.Text))
    Call FillPrice(Val(.Text))
    .Col = 7
    Call DistributorDetails(Val(.Text))
    .Col = 0
    .ColSel = .Cols - 1
End With
End Sub

Private Sub GetOrderDetails(OrderID As Long)
With rsTemOrder
    If .State = 1 Then .Close
    Temsql = "SELECT * from tblOrder where orderID = " & OrderID
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount < 1 Then Exit Sub
    NewItem.ID = !ItemID
        lblIUnit.Caption = NewItem.IUnit
    lblPUnit.Caption = NewItem.PUnit
    txtAMP.Text = NewItem.AMP
    txtAMPP.Text = NewItem.AMPP
    txtVMP.Text = NewItem.VMP
    txtVMPP.Text = NewItem.VMPP
    txtVTM.Text = NewItem.Generic
    txtDisplay.Text = NewItem.Display
    lblPUnit.Caption = NewItem.PUnit
    lblIUnit.Caption = NewItem.IUnit
    txtIQty.Text = !RequestAmount
    txtPQty.Text = Val(txtIQty.Text) / NewItem.IssueUnitsPerPack
    txtItem.Text = txtDisplay.Text
    txtOrderID.Text = OrderID
    dtcDistributor.BoundText = !RequestDistributorID
    
End With
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
    Temsql = "SELECT tblBatch.Batch, tblBatch.DOE, tblBatchStock.Stock, tblStore.Store, tblBatch.ItemID " & _
                " FROM tblStore RIGHT JOIN (tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) ON tblStore.StoreID = tblBatchStock.StoreID " & _
                " WHERE tblBatch.ItemID=" & ItemID & " AND tblBatchStock.Stock > 0 "
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
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
        Temsql = "SELECT * from tblStore order by store"
        .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
        
            While .EOF = False
                
                TemStore = !Store
                
                StoreUsage = 0
                
                StoreConsumption = CalculateConsumption(NewItem.ID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreConsumption
                TotalConsumption = TotalConsumption + StoreConsumption
                
                StoreSale = CalculateSale(NewItem.ID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreSale
                TotalSale = TotalSale + StoreSale
                
                StoreDiscard = CalculateDiscard(NewItem.ID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
                StoreUsage = StoreUsage + StoreDiscard
                TotalDiscard = TotalDiscard + StoreDiscard
                
                StoreAdjustment = CalculateAdjustment(NewItem.ID, dtpUFrom.Value, dtpUTo.Value, , !StoreID)
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
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = (.Width - 100) / 8
        Next i
        
    End With
    With rsTemOrder
        If .State = 1 Then .Close
        Temsql = "SELECT tblOrder.RequestDate, tblOrder.ApprovedDate, tblOrder.ReceivedDate, tblOrder.RequestAmount, tblOrder.ApprovedAmount, tblOrder.ReceivedAmount, tblRDistrubutor.DistributorName, tblADistrubutor.DistributorName FROM (tblDistrubutor AS tblRDistrubutor RIGHT JOIN tblOrder ON tblRDistrubutor.DistributorID = tblOrder.ApprovedDistributorID) LEFT JOIN tblDistrubutor AS tblADistrubutor ON tblOrder.RequestDistributorID = tblADistrubutor.DistributorID WHERE (((tblOrder.ItemID)=" & ItemID & ") AND ((tblOrder.RequestDate) Between #" & Format(dtpOFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpOTo.Value, "dd MMMM yyyy") & "#)) ORDER BY tblOrder.RequestDate"
        .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 1 Then
            While .EOF = False
                GridOrdering.Rows = GridOrdering.Rows + 1
                GridOrdering.Row = GridOrdering.Rows - 1
                GridOrdering.Col = 0
                GridOrdering.CellAlignment = 1
                GridOrdering.Text = Format(!requestdate, ShortDateFormat)
                GridOrdering.Col = 1
                GridOrdering.CellAlignment = 1
                If Not IsNull(!Approveddate) Then
                    GridOrdering.Text = Format(!Approveddate, ShortDateFormat)
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
                If Not IsNull(.Fields("tblRDistrubutor.DistributorName").Value) Then
                    GridOrdering.Text = .Fields("tblRDistrubutor.DistributorName").Value
                Else
                    GridOrdering.Text = "Not Requested"
                End If
                GridOrdering.Col = 7
                GridOrdering.CellAlignment = 7
                If Not IsNull(.Fields("tblADistrubutor.DistributorName").Value) Then
                    GridOrdering.Text = .Fields("tblADistrubutor.DistributorName").Value
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
    Temsql = "SELECT tblCurrentPurchasePrice.SetDate, tblCurrentPurchasePrice.PPrice FROM tblCurrentPurchasePrice WHERE (((tblCurrentPurchasePrice.ItemID)=" & ItemID & ") AND ((tblCurrentPurchasePrice.SetDate) Between #" & Format(dtpPFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpPTo.Value, "dd MMMM yyyy") & "#)) ORDER BY tblCurrentPurchasePrice.SetDate DESC"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        While .EOF = False
            With GridPPrice
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .CellAlignment = 1
                .Text = Format(rsTemPrice!SetDate, LongDateFormat)
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
    Temsql = "SELECT tblCurrentSalePrice.SetDate, tblCurrentSalePrice.SPrice FROM tblCurrentSalePrice WHERE (((tblCurrentSalePrice.ItemID)=" & ItemID & ") AND ((tblCurrentSalePrice.SetDate) Between #" & Format(dtpPFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpPTo.Value, "dd MMMM yyyy") & "#)) ORDER BY tblCurrentSalePrice.SetDate DESC"
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        While .EOF = False
            With GridSPrice
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .CellAlignment = 1
                .Text = Format(rsTemPrice!SetDate, LongDateFormat)
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
    Temsql = "SELECT tblDistrubutor.*, tblCity.City FROM tblCity RIGHT JOIN tblDistrubutor ON tblCity.CityId = tblDistrubutor.DistributorCityID Where DistributorId = " & DistributorID & ""
    .Open Temsql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    lblDistributor.Caption = !DistributorName
    lblBalance.Caption = Format(!balance, "#0.00")
    lblTelNo.Caption = !distributorTelephone
    lblFax.Caption = !distributorFax
    lblAddress.Caption = !distributorAddress
    lblCity.Caption = !City
    If .State = 1 Then .Close
 End With
End Sub

Private Sub txtIQty_LostFocus()
    txtPQty.Text = Val(txtIQty.Text) / NewItem.IssueUnitsPerPack
End Sub

Private Sub txtPQty_LostFocus()
    txtIQty.Text = Val(txtPQty.Text) * NewItem.IssueUnitsPerPack
End Sub
