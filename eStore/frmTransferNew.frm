VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTransferNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer"
   ClientHeight    =   12360
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
   MDIChild        =   -1  'True
   ScaleHeight     =   12360
   ScaleWidth      =   15240
   Begin VB.TextBox txtInvoice 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   11880
      TabIndex        =   71
      Top             =   3540
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   11880
      TabIndex        =   26
      Top             =   6240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   21102593
      CurrentDate     =   39691
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   1095
      Left            =   5040
      TabIndex        =   25
      Top             =   120
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "By Issue Units"
      TabPicture(0)   =   "frmTransferNew.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtQty"
      Tab(0).Control(1)=   "Label32"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "By Pack Units"
      TabPicture(1)   =   "frmTransferNew.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label53"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtPQty"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtPQty 
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtQty 
         Height          =   375
         Left            =   -73200
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label53 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label32 
         Caption         =   "Quantity"
         Height          =   375
         Left            =   -74760
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   255
      Left            =   11880
      TabIndex        =   24
      Top             =   6720
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.TextBox txtBatch 
      Height          =   375
      Left            =   12600
      TabIndex        =   11
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtDataEntry 
      Height          =   375
      Left            =   1560
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin btButtonEx.ButtonEx bttnReceive 
      Height          =   375
      Left            =   11880
      TabIndex        =   19
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Purchase"
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
      Height          =   5895
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   10398
      _Version        =   393216
      WordWrap        =   -1  'True
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   375
      Left            =   13560
      TabIndex        =   20
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
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
   Begin MSDataListLib.DataCombo dtcCatogery 
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcItem 
      Height          =   360
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcCode 
      Height          =   360
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDOM 
      Height          =   375
      Left            =   12600
      TabIndex        =   13
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "yyyy MMMM"
      Format          =   21102595
      CurrentDate     =   39545
   End
   Begin MSComCtl2.DTPicker dtpDOE 
      Height          =   375
      Left            =   12600
      TabIndex        =   15
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "yyyy MMMM"
      Format          =   21102595
      CurrentDate     =   39545
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   10440
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   375
      Left            =   10440
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Delete"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   27
      Top             =   8040
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Item"
      TabPicture(0)   =   "frmTransferNew.frx":0038
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Stocks"
      TabPicture(1)   =   "frmTransferNew.frx":0054
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridTotal"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Usage"
      TabPicture(2)   =   "frmTransferNew.frx":0070
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblStoreUsage"
      Tab(2).Control(1)=   "GridUsage"
      Tab(2).Control(2)=   "txtStoreUsage"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Ordering"
      TabPicture(3)   =   "frmTransferNew.frx":008C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GridOrdering"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Purchase"
      TabPicture(4)   =   "frmTransferNew.frx":00A8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GridPurchase"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Prices"
      TabPicture(5)   =   "frmTransferNew.frx":00C4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label17"
      Tab(5).Control(1)=   "Label16"
      Tab(5).Control(2)=   "GridSPrice"
      Tab(5).Control(3)=   "GridPPrice"
      Tab(5).ControlCount=   4
      Begin VB.Frame Frame2 
         Height          =   3735
         Left            =   7200
         TabIndex        =   44
         Top             =   360
         Width           =   7695
         Begin VB.Label Label33 
            Caption         =   "City"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lblCity 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   55
            Top             =   1800
            Width           =   3375
         End
         Begin VB.Label lblFax 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   54
            Top             =   2760
            Width           =   3375
         End
         Begin VB.Label Label31 
            Caption         =   "Fax No"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label lblAddress 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   975
            Left            =   1560
            TabIndex        =   52
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label lblTelNo 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   51
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
            TabIndex        =   50
            Top             =   3240
            Width           =   3375
         End
         Begin VB.Label Label27 
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "Tel No"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label25 
            Caption         =   "Address"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Distributor"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblDistributor 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1560
            TabIndex        =   45
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3735
         Left            =   240
         TabIndex        =   29
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
         TabIndex        =   28
         Top             =   600
         Width           =   2295
      End
      Begin MSFlexGridLib.MSFlexGrid GridUsage 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   57
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridTotal 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   58
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridOrdering 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   59
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPPrice 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   60
         Top             =   1080
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridSPrice 
         Height          =   3015
         Left            =   -66960
         TabIndex        =   61
         Top             =   1080
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPurchase 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   62
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin VB.Label Label16 
         Caption         =   "Sales Prices"
         Height          =   255
         Left            =   -66840
         TabIndex        =   65
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label17 
         Caption         =   "Purchase Prices"
         Height          =   255
         Left            =   -74880
         TabIndex        =   64
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
         TabIndex        =   63
         Top             =   720
         Width           =   2415
      End
   End
   Begin MSComCtl2.DTPicker dtpFDate 
      Height          =   375
      Left            =   1560
      TabIndex        =   66
      Top             =   7560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   21102595
      CurrentDate     =   39540
   End
   Begin MSComCtl2.DTPicker dtpTDate 
      Height          =   375
      Left            =   4560
      TabIndex        =   67
      Top             =   7560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   21102595
      CurrentDate     =   39540
   End
   Begin MSDataListLib.DataCombo dtcChecked 
      Height          =   360
      Left            =   11880
      TabIndex        =   72
      Top             =   5340
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dtcStaff 
      Height          =   360
      Left            =   11880
      TabIndex        =   73
      Top             =   4380
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbFrom 
      Height          =   360
      Left            =   11880
      TabIndex        =   74
      Top             =   2100
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbTo 
      Height          =   360
      Left            =   11880
      TabIndex        =   79
      Top             =   2820
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "To"
      Height          =   255
      Left            =   11880
      TabIndex        =   80
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label22 
      Caption         =   "Invoice No."
      Height          =   255
      Left            =   11880
      TabIndex        =   78
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "Checked by"
      Height          =   255
      Left            =   11880
      TabIndex        =   77
      Top             =   4980
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Received by"
      Height          =   255
      Left            =   11880
      TabIndex        =   76
      Top             =   4020
      Width           =   1455
   End
   Begin VB.Label Label24 
      Caption         =   "From"
      Height          =   255
      Left            =   11880
      TabIndex        =   75
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblCategory 
      Height          =   375
      Left            =   2160
      TabIndex        =   70
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "From :"
      Height          =   255
      Left            =   240
      TabIndex        =   69
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "To :"
      Height          =   255
      Left            =   4080
      TabIndex        =   68
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label lblFQtyUnit 
      Height          =   375
      Left            =   9000
      TabIndex        =   23
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblQtyUnit 
      Height          =   375
      Left            =   9000
      TabIndex        =   22
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label38 
      Caption         =   "Batch"
      Height          =   375
      Left            =   10800
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label37 
      Caption         =   "Date of Manufacture"
      Height          =   375
      Left            =   10800
      TabIndex        =   12
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label36 
      Caption         =   "Date of Expiary"
      Height          =   375
      Left            =   10800
      TabIndex        =   14
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label35 
      Caption         =   "Code"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "&Item"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label28 
      Caption         =   "Catogery"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmTransferNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'    Dim temSql As String
'
'    Dim CsetPrinter As New cSetDfltPrinter
'
'    Dim TemOrderBillID As Long
'    Dim TemDistributorId As Long
'    Dim TemDistributorOrderID As Long
'    Dim EditingData As Boolean
'    Dim TemContent(22) As String
'    Dim CurrentRow As Integer
'    Dim TemCellContent As String
'    Dim temRefillBillID As Long
'
'    Dim NewItem As New Item
'
'    Dim rsStaff As New ADODB.Recordset
'    Dim rsSPrice As New ADODB.Recordset
'    Dim rsPPrice As New ADODB.Recordset
'    Dim rsCC As New ADODB.Recordset
'    Dim rsItem As New ADODB.Recordset
'    Dim rsItemCategory As New ADODB.Recordset
'    Dim rsCode As New ADODB.Recordset
'    Dim rsBanks As New ADODB.Recordset
'    Dim rsCreditCards As New ADODB.Recordset
'    Dim rsCities As New ADODB.Recordset
'    Dim rsPayment As New ADODB.Recordset
'    Dim rsFrom As New ADODB.Recordset
'    Dim rsTo As New ADODB.Recordset
'
'    Dim rsTemOrder As New ADODB.Recordset
'    Dim rsTemPrice As New ADODB.Recordset
'    Dim rsTemDistributor As New ADODB.Recordset
'    Dim rsTemStore As New ADODB.Recordset
'    Dim rsTemOrderBill As New ADODB.Recordset
'    Dim rsTemDistributorOrder As New ADODB.Recordset
'    Dim rsTemRefill As New ADODB.Recordset
'    Dim rsTemRefillBill As New ADODB.Recordset
'    Dim rsTemCash As New ADODB.Recordset
'    Dim rsTemCredit As New ADODB.Recordset
'    Dim rsTemCheque As New ADODB.Recordset
'    Dim rsDI As New ADODB.Recordset
'
' Private Sub FillDetails()
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
'
'    With rsDI
'        If .State = 1 Then .Close
'        temSql = "SELECT tblItemDistributor.DistributorID FROM tblItemDistributor WHERE (((tblItemDistributor.ItemID)=" & Val(dtcItem.BoundText) & "))"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            DistributorDetails (!DistributorID)
'        End If
'        .Close
'    End With
'End Sub
'
'Private Sub bttnDelete_Click()
'    If GridItem.Rows <= 1 Then Exit Sub
'    If GridItem.Rows = 2 Then
'        FormatGrid1
'        FormatGrids
'    Else
'        GridItem.RemoveItem (GridItem.Row)
'    End If
'End Sub
'
'Private Sub dtcCatogery_Change()
''    If IsNumeric(dtcCatogery.BoundText) Then
''        ListSelectedItems
''    Else
''        ListAllItems
''    End If
''    Dim rsIC As New ADODB.Recordset
''    With rsIC
''        If .State = 1 Then .Close
''        temSql = "Select * from tblItemCategory where ItemCategoryID = " & Val(dtcCatogery.BoundText)
''        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
''        If .RecordCount > 0 Then
''            lblCategory.Caption = !ItemCategory
''        End If
''        .Close
''    End With
'
'    dtcItem.Text = Empty
'    dtcCode.Text = Empty
'End Sub
'
'
'Private Sub ListSelectedItems()
'With rsItem
'    If .State = 1 Then .Close
'    temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by display"
'    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'End With
'With dtcItem
'    Set .RowSource = rsItem
'    .ListField = "Display"
'    .BoundColumn = "ItemID"
'End With
'With rsCode
'    If .State = 1 Then .Close
'    temSql = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by code"
'    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'End With
'With dtcCode
'    Set .RowSource = rsCode
'    .ListField = "Code"
'    .BoundColumn = "ItemID"
'End With
'
'End Sub
'
'Private Sub ListAllItems()
'With rsItem
'    If .State = 1 Then .Close
'    temSql = "SELECT * from tblitem order by display"
'    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'End With
'With dtcItem
'    Set .RowSource = rsItem
'    .ListField = "display"
'    .BoundColumn = "ItemID"
'End With
'With rsCode
'    If .State = 1 Then .Close
'    temSql = "SELECT * from tblitem order by code"
'    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'End With
'With dtcCode
'    Set .RowSource = rsCode
'    .ListField = "Code"
'    .BoundColumn = "ItemID"
'End With
'End Sub
'
'Private Sub dtcCatogery_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then
'        dtcCatogery.Text = Empty
'    ElseIf KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        dtcItem.SetFocus
'    End If
'End Sub
'
'
'Private Sub dtcCatogery_LostFocus()
'    If IsNumeric(dtcCatogery.BoundText) Then
'        ListSelectedItems
'    Else
'        ListAllItems
'    End If
'End Sub
'
'Private Sub dtcCode_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If SSTab3.Tab = 0 Then
'            txtQty.SetFocus
'        ElseIf SSTab3.Tab = 1 Then
'            txtPQty.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        dtcCode.Text = Empty
'    End If
'End Sub
'
'Private Sub dtcItem_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        dtcCode.SetFocus
'    ElseIf KeyCode = vbKeyEscape Then
'        dtcItem.Text = Empty
'    End If
'End Sub
'
'Private Sub dtcItem_LostFocus()
'    If IsNumeric(dtcItem.BoundText) = False Then Exit Sub
'    dtcCode.BoundText = dtcItem.BoundText
'    NewItem.ID = Val(dtcItem.BoundText)
'    Call FillLabels
'    Call FormatGrids
'
'    Call GetItemDetails(NewItem.ID)
'    Call FillStocks(dtcItem.BoundText)
'    Call FillPurchase(dtcItem.BoundText)
'    Call FillUsage(dtcItem.BoundText)
'    Call FillPrice(dtcItem.BoundText)
'    Call GetItemDetails(dtcItem.BoundText)
'    Call FillOrdering(dtcItem.BoundText)
'    Dim rsDI As New ADODB.Recordset
'    With rsDI
'        If .State = 1 Then .Close
'        temSql = "SELECT tblItemDistributor.DistributorID FROM tblItemDistributor WHERE (((tblItemDistributor.ItemID)=" & Val(dtcItem.BoundText) & "))"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            DistributorDetails (!DistributorID)
'        End If
'        .Close
'    End With
'
'End Sub
'
'Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        bttnReceive_Click
'    End If
'End Sub
'
'Private Sub dtpDOE_GotFocus()
'    'On Error Resume Next
'    SendKeys "{RIGHT}"
'End Sub
'
'Private Sub dtpDOE_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        bttnAdd_Click
'    End If
'End Sub
'
'Private Sub Form_Load()
'    Call FillCombos
'    Call FillDetails
'    Call FormatGrid1
'    Call FormatGrids
'    Call SetValues
'    GridItem.RowHeight(0) = GridItem.RowHeight(0) * 3
'    SSTab1.Tab = 0
'    SSTab3.Tab = 0
'    dtpDate.Value = Date
'End Sub
'Private Sub FormatGrids()
'    txtVMP.Text = Empty
'    txtVMPP.Text = Empty
'    txtVTM.Text = Empty
'    txtAMP.Text = Empty
'    txtAMPP.Text = Empty
'    txtDisplay.Text = Empty
'
'
'    With GridUsage
'        .Cols = 6
'        .Rows = 1
'        .FixedCols = 0
'        .ColWidth(0) = 3000
'        .ColWidth(1) = (.Width - (.ColWidth(0) + 100)) / 5
'        .ColWidth(2) = .ColWidth(1)
'        .ColWidth(3) = .ColWidth(1)
'        .ColWidth(4) = .ColWidth(1)
'        .ColWidth(5) = .ColWidth(1)
'        Dim i As Long
'        For i = 0 To .Cols - 1
'            .Col = i
'            .CellAlignment = 4
'            Select Case i
'                Case 0: .Text = "Store"
'                Case 1: .Text = "Sale"
'                Case 2: .Text = "Consumption"
'                Case 3: .Text = "Discard"
'                Case 4: .Text = "Adjustment"
'                Case 5: .Text = "Total"
'            End Select
'        Next i
'    End With
'
'    With GridTotal
'        .Clear
'        .Cols = 4
'        .Rows = 1
'        .Row = 0
'        .FixedCols = 0
'        .Col = 0
'        .CellAlignment = 4
'        .Text = "Batch"
'        .Col = 1
'        .CellAlignment = 4
'        .Text = "Stock (" & NewItem.IUnit & ")"
'        .Col = 2
'        .CellAlignment = 4
'        .Text = "Expiary"
'        .Col = 3
'        .CellAlignment = 4
'        .Text = "Department"
'        .ColWidth(1) = 1600
'        .ColWidth(2) = 1600
'        .ColWidth(3) = 1600
'        .ColWidth(0) = .Width - (.ColWidth(1) + .ColWidth(2) + .ColWidth(3) + 100)
'    End With
'
'    With GridSPrice
'        .Cols = 2
'        .Rows = 1
'        .FixedCols = 0
'        .Row = 0
'        .Col = 0
'        .CellAlignment = 4
'        .Text = "Starting Date"
'        .Col = 1
'        .CellAlignment = 4
'        .Text = "Sales Price per " & NewItem.IUnit
'        .ColWidth(0) = (.Width - 100) / 2
'        .ColWidth(1) = (.Width - 100) / 2
'    End With
'
'    With GridPPrice
'        .Cols = 2
'        .Rows = 1
'        .FixedCols = 0
'        .Row = 0
'        .Col = 0
'        .CellAlignment = 4
'        .Text = "Starting Date"
'        .Col = 1
'        .CellAlignment = 4
'        .Text = "Purchase Price per " & NewItem.PUnit
'        .ColWidth(0) = (.Width - 100) / 2
'        .ColWidth(1) = (.Width - 100) / 2
'    End With
'
'    With GridOrdering
'        .Rows = 1
'        .Cols = 8
'        .FixedCols = 0
'        .Col = 0
'        .CellAlignment = 4
'        .Text = "Requested On"
'        .Col = 1
'        .CellAlignment = 4
'        .Text = "Approved On"
'        .Col = 2
'        .CellAlignment = 4
'        .Text = "Requested Amount"
'        .Col = 3
'        .CellAlignment = 4
'        .Text = "Approved Amount"
'        .Col = 4
'        .CellAlignment = 4
'        .Text = "Requested Distributor"
'        .Col = 5
'        .CellAlignment = 4
'        .Text = "Approved Distributor"
'        For i = 0 To .Cols - 1
'            .ColWidth(i) = (.Width - 100) / 6
'        Next i
'    End With
'
'    With GridPurchase
'        .Rows = 1
'        .Cols = 5
'        .FixedCols = 0
'        .Col = 0
'        .CellAlignment = 4
'        .Text = "Date"
'        .Col = 1
'        .CellAlignment = 4
'        .Text = "Batch"
'        .Col = 2
'        .CellAlignment = 4
'        .Text = "Quentity"
'        .Col = 3
'        .CellAlignment = 4
'        .Text = "Free"
'        .Col = 4
'        .CellAlignment = 4
'        .Text = "Expiary"
'        For i = 0 To .Cols - 1
'            .ColWidth(i) = (.Width - 100) / 5
'        Next i
'    End With
'
'
'End Sub
'
'Private Sub SetValues()
'    dtpDOE.Value = Date
'    dtpDOM.Value = Date
'    dtpDOE.MinDate = LastDateOfMonth(Date)
'    dtcStaff.BoundText = UserID
'    dtcChecked.BoundText = UserID
'    dtcStaff.Locked = True
'    dtpTDate.Value = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
'    dtpTDate.Value = Date
'    dtpFDate.Value = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
'    dtpFDate.Value = Date
'End Sub
'
'Private Sub FillCombos()
'    With rsStaff
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblstaff order by listedname"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With dtcStaff
'        Set .RowSource = rsStaff
'        .ListField = "ListedName"
'        .BoundColumn = "StaffID"
'    End With
'    With dtcChecked
'        Set .RowSource = rsStaff
'        .ListField = "ListedName"
'        .BoundColumn = "StaffID"
'    End With
'
'    With rsFrom
'        If .State = 1 Then .Close
'        temSql = "SELECT tblStore.* From tblStore ORDER BY Store"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With cmbFrom
'        Set .RowSource = rsFrom
'        .ListField = "Store"
'        .BoundColumn = "StoreID"
'    End With
'
'    With rsTo
'        If .State = 1 Then .Close
'        temSql = "SELECT tblStore.* To tblStore ORDER BY Store"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With cmbTo
'        Set .RowSource = rsTo
'        .ListField = "Store"
'        .BoundColumn = "StoreID"
'    End With
'
'    With rsCC
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblpaymentMethod " & _
'                    "ORDER BY PaymentMethod"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'
'    With dtcItem
'        Set .RowSource = rsItem
'        .ListField = "display"
'        .BoundColumn = "ItemID"
'    End With
'    With rsItemCategory
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblItemCategory order by categoryCode"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With dtcCatogery
'        Set .RowSource = rsItemCategory
'        .ListField = "CategoryCode"
'        .BoundColumn = "ItemCategoryID"
'    End With
'    With rsCode
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblitem order by code"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With dtcCode
'        Set .RowSource = rsCode
'        .ListField = "code"
'        .BoundColumn = "ItemID"
'    End With
'    With rsBanks
'        If .State = 1 Then .Close
'        temSql = "SELECT tblBank.* FROM tblBank ORDER BY tblBank.Bank"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'End Sub
'
'Private Sub FormatGrid1()
'    EditingData = False
'    With GridItem
'        .Cols = 24
'        .Rows = 1
'        .Row = 0
'        .Col = 0
'        .FixedCols = 0
'
''        .RowHeight(0) = .RowHeight(0) * 3
'
'        Dim i As Integer
'
'        For i = 0 To .Cols - 1
'            .Col = i
'            .CellAlignment = 4
'            Select Case i
'                Case 0:     .Text = "No"
'                            .ColWidth(i) = 400
'                Case 1:     .Text = "Item"
'                            .ColWidth(i) = 3600
'                Case 5:     .Text = "Purchased"
'                            .ColWidth(i) = 900
'                Case 6:     .Text = "Unit"
'                            .ColWidth(i) = 900
'                Case 7:     .Text = "Free"
'                            .ColWidth(i) = 900
'                Case 8:     .Text = "Unit"
'                            .ColWidth(i) = 900
'                Case 9:     .Text = "Batch"
'                            .ColWidth(i) = 900
'                Case 10:     .Text = "Pruchase Price Per Unit"
'                            .ColWidth(i) = 900
'                Case 11:     .Text = "Slaes Price Per Unit"
'                            .ColWidth(i) = 900
'                Case 13:     .Text = "Purchase Price Per Pack"
'                            .ColWidth(i) = 900
'                Case 18:    .ColWidth(i) = 1200
'                            .Text = "Total Pruchase Value"
'                Case 21:    .ColWidth(i) = 1200
'                            .Text = "DOE"
'                Case Else:  .ColWidth(i) = 1
'            End Select
'        Next i
'
'    End With
'    '   0   No
'    '   1   Item
'    '   2   ItemID
'    '   3   PackUnitID
'    '   4   IssueUnitID
'    '   5   PurchaseQuentity
'    '   6   IUnit
'    '   7   FreeQuentity
'    '   8   IUnit
'    '   9   Batch
'    '   10  Purchase Price Per Unit
'    '   11  Sales Price Per Unit
'    '   12  Sales Margin
'    '   13  Purchaes Price Per Pack
'    '   14
'    '   15  IPurchased
'    '   16  IFreePurchased
'    '   17  IUnitsPerPack
'    '   18  Display Price
'    '   19  Actual Price
'    '   20  DOM
'    '   21  DOE
'    '   22  Last Sale Price
'    '   23  Last Purchase Price
'
'    EditingData = True
'End Sub
'
' Private Sub bttnCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub bttnAdd_Click()
'    If CanAdd = False Then Exit Sub
'    EditingData = False
'    With GridItem
'        .Rows = .Rows + 1
'        .Row = .Rows - 1
'
'        .Col = 0
'        .CellAlignment = 7
'        .Text = .Row
'
'        .Col = 1
'        .CellAlignment = 1
'        .Text = NewItem.Display
'
'        .Col = 2
'        .Text = NewItem.ID
'
'        .Col = 3
'        .Text = NewItem.PUnitID
'
'        .Col = 4
'        .Text = NewItem.IUnitID
'
'        .Col = 5
'        .CellAlignment = 7
'        .Text = txtQty.Text
'
'        .Col = 6
'        .CellAlignment = 1
'        .Text = NewItem.IUnit
'
'
'        .Col = 8
'        .CellAlignment = 1
'        .Text = NewItem.IUnit
'
'        .Col = 9
'        .CellAlignment = 7
'        .Text = txtBatch.Text
'
'
'        .Col = 15
'        .Text = Val(txtQty.Text)
'
'
'        .Col = 17
'        .Text = NewItem.IssueUnitsPerPack
'
'
'
'        .Col = 20
'        .CellAlignment = 4
'        .Text = LastDateOfMonth(dtpDOM.Value)
'
'        .Col = 21
'        .CellAlignment = 7
'        .Text = Format(LastDateOfMonth(dtpDOE.Value), "dd MMM yyyy")
'
'
'
'    End With
'    Call ClearAddValues
'    Call ClearItemDetails
'    Call ClearGrids
'    Call CalculateTotal
'    dtcCatogery.SetFocus
'    EditingData = True
'End Sub
'
'
'Private Sub ClearAddValues()
'    txtQty.Text = Empty
'    dtcItem.Text = Empty
'    dtcCatogery.Text = Empty
'    dtcCode.Text = Empty
'    txtBatch.Text = Empty
'
'
'    txtPQty.Text = Empty
'
'
'
'    lblQtyUnit.Caption = Empty
'
'End Sub
'
'Private Sub ClearItemDetails()
'    txtVMP.Text = Empty
'    txtVMPP.Text = Empty
'    txtVTM.Text = Empty
'    txtAMP.Text = Empty
'    txtAMPP.Text = Empty
'    txtDisplay.Text = Empty
'End Sub
'
'Private Sub ClearGrids()
'    With GridOrdering
'        .Clear
'        .Rows = 1
'        .Cols = 1
'        .ColWidth(0) = .Width
'    End With
'    With GridPPrice
'        .Clear
'        .Rows = 1
'        .Cols = 1
'        .ColWidth(0) = .Width
'    End With
'    With GridSPrice
'        .Clear
'        .Rows = 1
'        .Cols = 1
'        .ColWidth(0) = .Width
'    End With
'    With GridTotal
'        .Clear
'        .Rows = 1
'        .Cols = 1
'        .ColWidth(0) = .Width
'    End With
'    With GridUsage
'        .Clear
'        .Rows = 1
'        .Cols = 1
'        .ColWidth(0) = .Width
'    End With
'End Sub
'
'
'Private Function CanAdd() As Boolean
'    CanAdd = False
'    Dim tr As Integer
'        If IsNumeric(dtcItem.BoundText) = False Then
'            tr = MsgBox("You have not entered the item to add", vbCritical, "Item?")
'            dtcItem.SetFocus
'            Exit Function
'        End If
'        If IsNumeric(txtQty.Text) = False Or Val(txtQty.Text) = 0 Then
'            tr = MsgBox("You have not entered the quentity", vbCritical, "Quentity?")
'            txtQty.SetFocus
'            Exit Function
'        End If
'        If dtpDOE.Value = Date Then
'            tr = MsgBox("You have not entered a Date of Expiary", vbCritical, "Expiary Date")
'            dtpDOE.SetFocus
'            Exit Function
'        End If
'        If Trim(txtBatch.Text) = Empty Then
'            tr = MsgBox("You have not entered a Batch number", vbCritical, "Expiary Date")
'            txtBatch.SetFocus
'            Exit Function
'        End If
'    CanAdd = True
'End Function
'
'Private Sub dtcItem_Change()
'    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
''    dtcCode.BoundText = dtcItem.BoundText
''    NewItem.ID = dtcItem.BoundText
''    Call FillLabels
''    Call FormatGrids
''    Call GetLastPrices(dtcItem.BoundText)
''    Call GetItemDetails(NewItem.ID)
''    Call FillStocks(dtcItem.BoundText)
''    Call FillPurchase(dtcItem.BoundText)
''    Call FillUsage(dtcItem.BoundText)
''    Call FillPrice(dtcItem.BoundText)
''    Call GetItemDetails(dtcItem.BoundText)
''    Call FillOrdering(dtcItem.BoundText)
''    Dim rsDI As New ADODB.Recordset
''    With rsDI
''        If .State = 1 Then .Close
''        temSql = "SELECT tblItemDistributor.DistributorID FROM tblItemDistributor WHERE (((tblItemDistributor.ItemID)=" & Val(dtcItem.BoundText) & "))"
''        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
''        If .RecordCount > 0 Then
''            DistributorDetails (!DistributorID)
''        End If
''        .Close
''    End With
'End Sub
'
'Private Sub FillLabels()
'    lblQtyUnit.Caption = NewItem.IUnit
'End Sub
'
'Private Sub GridItem_DblClick()
'    With GridItem
'        If IsNumeric(.TextMatrix(.Row, 2)) = False Then Exit Sub
'        dtcItem.BoundText = .TextMatrix(.Row, 2)
'        txtQty.Text = .TextMatrix(.Row, 15)
'        txtBatch.Text = .TextMatrix(.Row, 9)
'        dtpDOM.Value = .TextMatrix(.Row, 20)
'        dtpDOE.Value = .TextMatrix(.Row, 21)
'    End With
'    bttnDelete_Click
'End Sub
'
'Private Sub txtBatch_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        dtpDOE.SetFocus
'    End If
'End Sub
'
'Private Sub txtBatch_LostFocus()
'    txtBatch.Text = UCase(txtBatch.Text)
'End Sub
'
'Private Sub txtInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then
'        txtInvoice.Text = Empty
'    ElseIf KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        dtpDate.SetFocus
'    End If
'End Sub
'
'Private Sub txtPQty_Change()
'    If SSTab3.Tab = 1 Then
'        txtQty.Text = Val(txtPQty.Text) * NewItem.IssueUnitsPerPack
'    End If
'End Sub
'
'Private Sub txtPQty_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        'txtFPQty.SetFocus
'    End If
'End Sub
'
'Private Sub txtQty_Change()
'    If SSTab3.Tab = 0 Then
'        If NewItem.IssueUnitsPerPack <> 0 Then txtPQty.Text = Val(txtQty.Text) / NewItem.IssueUnitsPerPack
'    End If
'
'End Sub
'
'
'
'
'Private Function CanReceive() As Boolean
'    Dim i As Integer
'    Dim tr As Integer
'    CanReceive = False
'
'    If GridItem.Rows <= 1 Then
'        tr = MsgBox("There are no items to sell", vbCritical, "No Items")
'        dtcItem.SetFocus
'        Exit Function
'    End If
'
'
'    If IsNumeric(cmbFrom.BoundText) = False Then
'        tr = MsgBox("You have not selected the returning department", vbCritical, "From?")
'
'        cmbFrom.SetFocus
'        Exit Function
'    End If
'
'    If IsNumeric(cmbTo.BoundText) = False Then
'        tr = MsgBox("You have not selected the receiving department", vbCritical, "To?")
'
'        cmbTo.SetFocus
'        Exit Function
'    End If
'
'
'
'    If IsNumeric(dtcStaff.BoundText) = False Then
'        tr = MsgBox("You have not selected the user", vbCritical, "Issued by?")
'
'        dtcStaff.SetFocus
'        Exit Function
'    End If
'
'    If IsNumeric(dtcChecked.BoundText) = False Then
'        tr = MsgBox("You have not selected the name of the checked staff member", vbCritical, "Checked by?")
'
'        dtcChecked.SetFocus
'        Exit Function
'    End If
'
'    CanReceive = True
'End Function
'
'Private Function NoSameInnvoice(InvoiceNo As String, DistributorID As Long) As Boolean
'    Dim rsNTem As New ADODB.Recordset
'    With rsNTem
'        If .State = 1 Then .Close
'        temSql = "SELECT     DistributorID, InvoiceNo FROM         dbo.tblRefillBill WHERE     (DistributorID = " & DistributorID & ") AND (InvoiceNo = N'" & InvoiceNo & "')"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            NoSameInnvoice = False
'        Else
'            NoSameInnvoice = True
'        End If
'        .Close
'    End With
'End Function
'
'
'Private Sub bttnReceive_Click()
'
'    Dim i As Integer
'
'    If CanReceive = False Then Exit Sub
'
'
'    Dim tr As Integer
'    Dim DiscountPercent As Double
'
'    With rsTemRefillBill
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblRefillBill"
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        .AddNew
'        !DistributorID = dtcSupplier.BoundText
'        !StoreID = UserStoreID
'        !StaffID = UserID
'        If IsNumeric(dtcChecked.BoundText) = True Then
'            !CheckedStaffID = dtcChecked.BoundText
'        End If
'        !Price = Val(lblGrossTotal.Caption)
'        !Discount = Val(txtDiscount.Text)
'        DiscountPercent = (Val(txtDiscount.Text) / Val(lblGrossTotal.Caption)) * 100
'        !DiscountPercent = DiscountPercent
'        !NetPrice = Val(lblNetTotal.Caption)
'        !Date = Date
'        !Time = Now
'        !PaymentMethodID = dtcPayment.BoundText
'        !PaymentMethod = dtcPayment.Text
'        !InvoiceNo = txtInvoice.Text
'        !InvoiceDate = dtpDate.Value
'        If dtcPayment.Text = "Credit" Then
'            !FullyPaid = False
'        Else
'            !FullyPaid = True
'        End If
'        !purchase = True
'        !Autorequest = False
'        !ManualRequest = False
'        .Update
'        temSql = "SELECT @@IDENTITY AS NewID"
'        .Close
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        temRefillBillID = !NewID
'        .Close
'        Dim CashID As Long
'        Dim CreditID As Long
'        Dim ChequeID As Long
'
'        If dtcPayment.Text = "Cash" Then
'            CashID = IssueCash(temRefillBillID)
'        ElseIf dtcPayment.Text = "Credit" Then
'            CreditID = IssueCredit(temRefillBillID)
'        ElseIf dtcPayment.Text = "Cheque" Then
'            ChequeID = IssueCheque(temRefillBillID)
'        End If
'        If .State = 1 Then .Close
'        temSql = "Select * from tblRefillBill Where RefillBillID = " & temRefillBillID
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        If .RecordCount > 0 Then
'            If dtcPayment.Text = "Cash" Then
'                !IssuedCashID = CashID
'            ElseIf dtcPayment.Text = "Credit" Then
'                !IssuedCreditID = CreditID
'            ElseIf dtcPayment.Text = "Cheque" Then
'                !IssuedChequeID = ChequeID
'            End If
'            .Update
'        End If
'        .Close
'    End With
'
'    With GridItem
'        For i = 1 To .Rows - 1
'            If rsTemRefill.State = 1 Then rsTemRefill.Close
'            temSql = "SELECT * FROM tblRefill"
'            rsTemRefill.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'            rsTemRefill.AddNew
'            rsTemRefill!ItemID = Val(.TextMatrix(i, 2))
'            rsTemRefill!StoreID = UserStoreID
'            rsTemRefill!Date = Date
'            rsTemRefill!Time = Now
'            rsTemRefill!StaffID = UserID
'            rsTemRefill!DistributorID = dtcSupplier.BoundText
'            rsTemRefill!Price = Val(.TextMatrix(i, 19))
'            rsTemRefill!DiscountPercent = DiscountPercent
'            rsTemRefill!NetPrice = (Val(.TextMatrix(i, 19))) - (Val(.TextMatrix(i, 19)) * DiscountPercent / 100)
'            rsTemRefill!RefillBillID = temRefillBillID
'            rsTemRefill!purchase = True
'            rsTemRefill!Autorequest = False
'            rsTemRefill!ManualRequest = False
'            rsTemRefill!Amount = Val(.TextMatrix(i, 15))
'            rsTemRefill!FreeAmount = Val(.TextMatrix(i, 16))
'            rsTemRefill!CheckedStaffID = dtcChecked.BoundText
'
'            Dim ThisBatch As Long
'            ThisBatch = BatchExist(.TextMatrix(i, 9), Val(.TextMatrix(i, 2)))
'            If ThisBatch <> 0 Then
'                rsTemRefill!BatchID = ThisBatch
'                If AddToStock(ThisBatch, UserStoreID, Val(.TextMatrix(i, 15)) + Val(.TextMatrix(i, 16))) = False Then
'                    MsgBox "Error"
'                    Exit For
'                End If
'            Else
'                ThisBatch = AddBatch(.TextMatrix(i, 9), Val(.TextMatrix(i, 2)), .TextMatrix(i, 20), .TextMatrix(i, 21))
'                rsTemRefill!BatchID = ThisBatch
'                If AddToStock(ThisBatch, UserStoreID, Val(.TextMatrix(i, 15)) + Val(.TextMatrix(i, 16))) = False Then
'                    MsgBox "Error"
'                    Exit For
'                End If
'            End If
'            rsTemRefill!DOM = CDate(.TextMatrix(i, 20))
'            rsTemRefill!DOE = CDate(.TextMatrix(i, 21))
'            rsTemRefill!PackPPrice = Val(.TextMatrix(i, 13))
'            rsTemRefill!SPrice = Val(.TextMatrix(i, 11))
'            rsTemRefill!PPrice = Val(.TextMatrix(i, 10))
'            rsTemRefill!LastSPrice = Val(.TextMatrix(i, 22))
'            rsTemRefill!LastPPrice = Val(.TextMatrix(i, 23))
'
'            rsTemRefill.Update
'            rsTemRefill.Close
'
'            If rsSPrice.State = 1 Then rsSPrice.Close
'            temSql = "SELECT tblSalePrice.ItemID, tblSalePrice.SPrice, tblSalePrice.SetDate, tblSalePrice.SetTime, tblSalePrice.StaffID FROM tblSalePrice "
'            rsSPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'            rsSPrice.AddNew
'            rsSPrice!ItemID = Val(.TextMatrix(i, 2))
'            rsSPrice!SPrice = Val(.TextMatrix(i, 11))
'            rsSPrice!setdate = Date
'            rsSPrice!SetTime = Now
'            rsSPrice!StaffID = UserID
'            rsSPrice.Update
'            rsSPrice.Close
'
'            If rsSPrice.State = 1 Then rsSPrice.Close
'            temSql = "SELECT * FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
'            rsSPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'            If rsSPrice.RecordCount < 1 Then
'                rsSPrice.AddNew
'                rsSPrice!ItemID = Val(.TextMatrix(i, 2))
'                rsSPrice!SPrice = Val(.TextMatrix(i, 11))
'                rsSPrice!setdate = Date
'                rsSPrice!SetTime = Now
'                rsSPrice!StaffID = UserID
'                rsSPrice.Update
'            ElseIf rsSPrice.RecordCount = 1 Then
'                rsSPrice!SPrice = Val(.TextMatrix(i, 11))
'                rsSPrice!setdate = Date
'                rsSPrice!SetTime = Now
'                rsSPrice!StaffID = UserID
'                rsSPrice.Update
'            Else
'                If rsSPrice.State = 1 Then rsSPrice.Close
'                temSql = "Delete FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
'                rsSPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'                If rsSPrice.State = 1 Then rsSPrice.Close
'                temSql = "SELECT * FROM tblCurrentSalePrice Where ItemID = " & Val(.TextMatrix(i, 2))
'                rsSPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'                rsSPrice.AddNew
'                rsSPrice!ItemID = Val(.TextMatrix(i, 2))
'                rsSPrice!SPrice = Val(.TextMatrix(i, 11))
'                rsSPrice!setdate = Date
'                rsSPrice!SetTime = Now
'                rsSPrice!StaffID = UserID
'                rsSPrice.Update
'            End If
'            rsSPrice.Close
'
'            If rsPPrice.State = 1 Then rsPPrice.Close
'            temSql = "SELECT tblPurchasePrice.ItemID, tblPurchasePrice.PPrice, tblPurchasePrice.SetDate, tblPurchasePrice.SetTime, tblPurchasePrice.StaffID FROM tblPurchasePrice"
'            rsPPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'            rsPPrice.AddNew
'            rsPPrice!ItemID = Val(.TextMatrix(i, 2))
'            rsPPrice!PPrice = Val(.TextMatrix(i, 10))
'            rsPPrice!setdate = Date
'            rsPPrice!SetTime = Now
'            rsPPrice!StaffID = UserID
'            rsPPrice.Update
'            rsPPrice.Close
'
'            If rsPPrice.State = 1 Then rsPPrice.Close
'            temSql = "SELECT * FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2)) & " Order by SetDate Desc, SetTime Desc"
'            rsPPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'            If rsPPrice.RecordCount < 1 Then
'                rsPPrice.AddNew
'                rsPPrice!ItemID = Val(.TextMatrix(i, 2))
'                rsPPrice!PPrice = Val(.TextMatrix(i, 10))
'                rsPPrice!setdate = Date
'                rsPPrice!SetTime = Now
'                rsPPrice!StaffID = UserID
'                rsPPrice.Update
'            ElseIf rsPPrice.RecordCount = 1 Then
'                rsPPrice!PPrice = Val(.TextMatrix(i, 10))
'                rsPPrice!setdate = Date
'                rsPPrice!SetTime = Now
'                rsPPrice!StaffID = UserID
'                rsPPrice.Update
'            Else
'                If rsPPrice.State = 1 Then rsPPrice.Close
'                temSql = "Delete FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2)) & " Order by SetDate Desc, SetTime Desc"
'                rsPPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'                If rsPPrice.State = 1 Then rsPPrice.Close
'                temSql = "SELECT * FROM tblCurrentPurchasePrice WHERE ItemID =" & Val(.TextMatrix(i, 2)) & " Order by SetDate Desc, SetTime Desc"
'                rsPPrice.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'                rsPPrice.AddNew
'                rsPPrice!ItemID = Val(.TextMatrix(i, 2))
'                rsPPrice!PPrice = Val(.TextMatrix(i, 10))
'                rsPPrice!setdate = Date
'                rsPPrice!SetTime = Now
'                rsPPrice!StaffID = UserID
'                rsPPrice.Update
'            End If
'            rsPPrice.Close
'
'        Next
'    End With
'    If chkPrint.Value = 1 Then PrintPurchase
'
'    tr = MsgBox("The Goods Received and added to stocks successfully" & vbNewLine & "Bill ID - " & temRefillBillID, vbInformation, "Success")
'    Call FormatGrid1
'    Call FormatGrids
'    Call ClearSettleValues
'    dtcCatogery.SetFocus
'    SSTab2.Tab = 0
'End Sub
'
'
'
'Private Sub PrintPurchase()
'    Dim RetVal As Integer
'    Dim TemResponce     As Integer
'     With Dataenvironment1.rscmmdGoodReceive
'         If .State = 1 Then .Close
'         .Source = "SELECT tblItem.Display, tblRefill.DOE, tblRefill.Amount, tblRefill.FreeAmount, tblRefill.PPrice, tblRefill.Price, tblRefill.SPrice, tblRefill.LastPPrice " & _
'                     " FROM tblRefill LEFT JOIN tblItem ON tblRefill.ItemID = tblItem.ItemID " & _
'                     " WHERE (((tblRefill.RefillBillID)= " & temRefillBillID & ") AND ((tblRefill.Amount) > 0))"
'         .Open
'         If .RecordCount > 0 Then
'
'
'        CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
'
'
'        Dim MyPrinter As Printer
'
'        For Each MyPrinter In VB.Printers
'            If MyPrinter.DeviceName = ReportPaperName Then
'                Set Printer = MyPrinter
'            End If
'        Next
'
'            With dtrPurchase
'                Set .DataSource = Dataenvironment1.rscmmdGoodReceive
'                .Sections("Section4").Controls("lblName").Caption = HospitalName
'                .Sections("Section4").Controls("lblContact").Caption = HospitalAddress
'                .Sections("Section4").Controls("lblTopic").Caption = "Good Receive Note"
'                .Sections("Section4").Controls("lblSUbtopic").Caption = Empty
'                .Sections("Section4").Controls("lblTo").Caption = lblDistributor.Caption
'                .Sections("Section4").Controls("lblAddress").Caption = lblAddress.Caption
'                .Sections("Section4").Controls("lblTel").Caption = lblTelNo.Caption
'                .Sections("Section4").Controls("lblFax").Caption = lblFax.Caption
'                .Sections("Section4").Controls("lblDate").Caption = Format(Date, LongDateFormat)
'                .Sections("Section4").Controls("lblRefillID").Caption = temRefillBillID
'                .Sections("Section4").Controls("lblInvoiceDate").Caption = Format(dtpDate.Value, LongDateFormat)
'                .Sections("Section4").Controls("lblInvoiceNo").Caption = txtInvoice.Text
'                .Sections("Section5").Controls("lblPayee").Caption = lblDistributor.Caption
'                .Sections("Section5").Controls("lblTotalAmount").Caption = lblGrossTotal.Caption
'                .Sections("Section5").Controls("lblDiscount").Caption = txtDiscount.Text
'                .Sections("Section5").Controls("lblNetTotal").Caption = lblNetTotal.Caption
'                .Sections("Section5").Controls("lblOperatedBy").Caption = UserName
'                '
'                RetVal = SelectForm(ReportPaperName, Me.hwnd)
'
'                If RetVal = FORM_SELECTED Then
'                    .Show
'                Else
'                    TemResponce = MsgBox("An Error in the report printer", vbCritical, "Printer Error")
'                    Exit Sub
'                End If
'
'            End With
'         End If
'    End With
'End Sub
'
'
'Private Sub ClearSettleValues()
'    txtAMP.Text = Empty
'    txtAMPP.Text = Empty
'    txtBalance.Text = Empty
'    txtBatch.Text = Empty
'    txtCashPaid.Text = Empty
'    txtChequeNo.Text = Empty
'    txtCreditCardNo.Text = Empty
'    txtCreditCode.Text = Empty
'    txtCreditDue.Text = Empty
'    txtDataEntry.Text = Empty
'    txtDiscount.Text = Empty
'    txtDisplay.Text = Empty
'    txtDue.Text = Empty
'    txtFQty.Text = Empty
'    txtInvoice.Text = Empty
'    txtIStore.Text = Empty
'    txtPPrice.Text = Empty
'    txtPurchaseValue.Text = Empty
'    txtQty.Text = Empty
'    txtVMP.Text = Empty
'    txtVMPP.Text = Empty
'    txtVTM.Text = Empty
'
'
'
'    dtcSupplier.Text = Empty
'    dtcBank.Text = Empty
'    dtcBranch.Text = Empty
'    dtcCardBank.Text = Empty
'    dtcCatogery.Text = Empty
'    dtcCode.Text = Empty
'    dtcCreditCard.Text = Empty
'    dtcItem.Text = Empty
'    dtcPayment.Text = Empty
'
'
'    dtpChequeDate.Value = Date
'    dtpTDate.Value = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
'    dtpTDate.Value = Date
'    dtpFDate.Value = DateSerial(Year(Date) - 1, Month(Date), Day(Date))
'    dtpFDate.Value = Date
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Dim tr As Integer
'    If GridItem.Rows > 1 Then
'        tr = MsgBox("There are items to be received. Are You sure you want to exit?", vbYesNo + vbQuestion, "Exit?")
'        If tr = vbNo Then Cancel = True: Exit Sub
'    End If
'End Sub
'
'Private Function IssueCredit(RefillBillID As Long) As Long
'    With rsTemCredit
'        If .State = 1 Then .Close
'        temSql = "SELECT * FROM tblIssuedCredit"
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        .AddNew
'        !IssuedSTaffID = dtcStaff.BoundText
'        !IssuedDate = Date
'        !IssuedTime = Now
'        !Price = Val(lblNetTotal.Caption)
'        !StoreID = UserStoreID
'        !RefillBillID = RefillBillID
'        .Update
'        .Close
'        temSql = "SELECT @@IDENTITY AS NewID"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        IssueCredit = !NewID
'        .Close
'        Set rsTemCredit = Nothing
'    End With
'End Function
'
'
'Private Function IssueCheque(RefillBillID As Long) As Long
'    With rsTemCheque
'        If .State = 1 Then .Close
'        temSql = "SELECT * FROM tblIssuedCheque"
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        .AddNew
'        !IssuedSTaffID = dtcStaff.BoundText
'        !IssuedDate = Date
'        !IssuedTime = Now
'        !bankID = Val(dtcBank.BoundText)
'        If IsNumeric(dtcBranch.BoundText) = True Then
'            !BranchID = dtcBranch.BoundText
'        End If
'        !ChequeDate = dtpChequeDate.Value
'        !ChequeNo = txtChequeNo.Text
'        !Price = Val(lblNetTotal.Caption)
'        !StoreID = UserStoreID
'        !RefillBillID = RefillBillID
'        .Update
'        .Close
'        temSql = "SELECT @@IDENTITY AS NewID"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        IssueCheque = !NewID
'        .Close
'        Set rsTemCredit = Nothing
'    End With
'End Function
'
'
'Private Function IssueCash(RefillBillID As Long) As Long
'    With rsTemCash
'        If .State = 1 Then .Close
'        temSql = "SELECT tblIssuedCash.* FROM tblIssuedCash"
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        .AddNew
'        !IssuedSTaffID = dtcStaff.BoundText
'        !IssuedDate = Date
'        !IssuedTime = Now
'        !Price = Val(lblNetTotal.Caption)
'        !RefillBillID = RefillBillID
'        !StoreID = UserStoreID
'        .Update
'        .Close
'        temSql = "SELECT @@IDENTITY AS NewID"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        IssueCash = !NewID
'        .Close
'        Set rsTemCredit = Nothing
'    End With
'End Function
'
'Private Sub CalculateTotal()
'    Dim i As Integer
'    Dim GrossTotal As Double
'    Dim NetTotal As Double
'    With GridItem
'        For i = 1 To GridItem.Rows - 1
'            GrossTotal = GrossTotal + Val(.TextMatrix(i, 19))
'        Next
'        lblGrossTotal.Caption = Format(GrossTotal, "####.00")
'        NetTotal = GrossTotal - Val(txtDiscount.Text)
'        lblNetTotal.Caption = Format(NetTotal, "####.00")
'    End With
'End Sub
'
'
'Private Sub DistributorDetails(ByVal DistributorID As Long)
'    With rsTemDistributor
'        If .State = 1 Then .Close
'        temSql = "SELECT tblDistrubutor.*, tblCity.City FROM tblCity RIGHT JOIN tblDistrubutor ON tblCity.CityId = tblDistrubutor.DistributorCityID Where DistributorId = " & DistributorID & ""
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount = 0 Then Exit Sub
'        If Not IsNull(!DistributorName) Then lblDistributor.Caption = !DistributorName
'        If Not IsNull(!Balance) Then lblBalance.Caption = Format(!Balance, "0.00")
'        If Not IsNull(!DistributorTelephone) Then lblTelNo.Caption = !DistributorTelephone
'        If Not IsNull(!DistributorFax) Then lblFax.Caption = !DistributorFax
'        If Not IsNull(!DistributorAddress) Then lblAddress.Caption = !DistributorAddress
'        If Not IsNull(!City) Then lblCity.Caption = !City
'        If .State = 1 Then .Close
'    End With
'End Sub
'
'Private Sub FillUsage(ByVal ItemID As Long)
'    '0 Store
'    '1 Sale
'    '2 Consum
'    '3 Discard
'    '4 Adjustments
'    '5 Total
'    Dim StoreConsumption As Double
'    Dim StoreSale As Double
'    Dim StoreAdjustment As Double
'    Dim StoreDiscard As Double
'    Dim StoreUsage As Double
'    Dim TotalConsumption As Double
'    Dim TotalSale As Double
'    Dim TotalAdjustment As Double
'    Dim TotalDiscard As Double
'    Dim TotalUsage As Double
'    Dim TemStore As String
'
'
'    With rsTemStore
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblStore order by store"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'
'            While .EOF = False
'
'                TemStore = !Store
'
'                StoreUsage = 0
'
'                StoreConsumption = CalculateConsumption(NewItem.ID, dtpFDate.Value, dtpTDate.Value, , !StoreID)
'                StoreUsage = StoreUsage + StoreConsumption
'                TotalConsumption = TotalConsumption + StoreConsumption
'
'                StoreSale = CalculateSale(NewItem.ID, dtpFDate.Value, dtpTDate.Value, , !StoreID)
'                StoreUsage = StoreUsage + StoreSale
'                TotalSale = TotalSale + StoreSale
'
'                StoreDiscard = CalculateDiscard(NewItem.ID, dtpFDate.Value, dtpTDate.Value, , !StoreID)
'                StoreUsage = StoreUsage + StoreDiscard
'                TotalDiscard = TotalDiscard + StoreDiscard
'
'                StoreAdjustment = CalculateAdjustment(NewItem.ID, dtpFDate.Value, dtpTDate.Value, , !StoreID)
'                StoreUsage = StoreUsage + StoreAdjustment
'                TotalAdjustment = TotalAdjustment + StoreAdjustment
'
'                If !StoreID = UserStoreID Then
'                    lblStoreUsage.Caption = !Store
'                    txtStoreUsage.Text = StoreUsage
'                End If
'
'                With GridUsage
'                    .Rows = .Rows + 1
'                    .Row = .Rows - 1
'                    .Col = 0
'                    .CellAlignment = 1
'                    .Text = TemStore
'                    .Col = 1
'                    .Text = StoreSale & " " & NewItem.IUnit
'                    .Col = 2
'                    .Text = StoreConsumption & " " & NewItem.IUnit
'                    .Col = 3
'                    .Text = StoreDiscard & " " & NewItem.IUnit
'                    .Col = 4
'                    .Text = StoreAdjustment & " " & NewItem.IUnit
'                    .Col = 5
'                    .Text = StoreUsage & " " & NewItem.IUnit
'                End With
'                .MoveNext
'            Wend
'            With GridUsage
'            .Rows = .Rows + 1
'            .Row = .Rows - 1
'            .Col = 0
'            .CellAlignment = 1
'            .Text = "Total"
'            .Col = 1
'            .Text = TotalSale & " " & NewItem.IUnit
'            .Col = 2
'            .Text = TotalConsumption & " " & NewItem.IUnit
'            .Col = 3
'            .Text = TotalDiscard & " " & NewItem.IUnit
'            .Col = 4
'            .Text = TotalAdjustment & " " & NewItem.IUnit
'            TotalUsage = TotalConsumption + TotalSale + TotalDiscard + TotalAdjustment
'            .Col = 5
'            .Text = TotalUsage & " " & NewItem.IUnit
'            End With
'        End If
'        .Close
'    End With
'
'End Sub
'
'
'Private Sub FillOrdering(ByVal ItemID As Long)
'    With rsTemOrder
'        If .State = 1 Then .Close
'        temSql = "SELECT tblOrder.RequestDate, tblOrder.ApprovedDate, tblOrder.ReceivedDate, tblOrder.RequestAmount, tblOrder.ApprovedAmount, tblOrder.ReceivedAmount, tblRDistrubutor.DistributorName as RDistributorName , tblADistrubutor.DistributorName as ADistributorName FROM (tblDistrubutor AS tblRDistrubutor RIGHT JOIN tblOrder ON tblRDistrubutor.DistributorID = tblOrder.ApprovedDistributorID) LEFT JOIN tblDistrubutor AS tblADistrubutor ON tblOrder.RequestDistributorID = tblADistrubutor.DistributorID WHERE tblOrder.ItemID  = " & ItemID & "AND tblOrder.RequestDate between '" & Format(dtpFDate.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTDate.Value, "dd MMMM yyyy") & "'"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount >= 1 Then
'            While .EOF = False
'                GridOrdering.Rows = GridOrdering.Rows + 1
'                GridOrdering.Row = GridOrdering.Rows - 1
'                GridOrdering.Col = 0
'                GridOrdering.CellAlignment = 1
'                GridOrdering.Text = Format(!requestdate, ShortDateFormat)
'                GridOrdering.Col = 1
'                GridOrdering.CellAlignment = 1
'                If Not IsNull(!ApprovedDate) Then
'                    GridOrdering.Text = Format(!ApprovedDate, ShortDateFormat)
'                Else
'                    GridOrdering.Text = "Not Approved"
'                End If
'                GridOrdering.Col = 2
'                GridOrdering.CellAlignment = 1
'                If Not IsNull(!RequestAmount) Then
'                    GridOrdering.Text = !RequestAmount & " " & NewItem.IUnit
'                Else
'                    GridOrdering.Text = "Not Requested"
'                End If
'                GridOrdering.Col = 3
'                GridOrdering.CellAlignment = 7
'                If Not IsNull(!ApprovedAmount) Then
'                    GridOrdering.Text = !ApprovedAmount & " " & NewItem.IUnit
'                Else
'                    GridOrdering.Text = "Not Approved"
'                End If
'                GridOrdering.Col = 4
'                GridOrdering.CellAlignment = 7
'                If Not IsNull(.Fields("RDistributorName").Value) Then
'                    GridOrdering.Text = .Fields("RDistributorName").Value
'                Else
'                    GridOrdering.Text = "Not Requested"
'                End If
'                GridOrdering.Col = 5
'                GridOrdering.CellAlignment = 7
'                If Not IsNull(.Fields("ADistributorName").Value) Then
'                    GridOrdering.Text = .Fields("ADistributorName").Value
'                Else
'                    GridOrdering.Text = "Not Approved"
'                End If
'                .MoveNext
'            Wend
'        End If
'    End With
'End Sub
'Private Sub FillStocks(ByVal ItemID As Long)
'    With rsTemStore
'        If .State = 1 Then .Close
'        temSql = "SELECT tblBatch.Batch, tblBatch.DOE, tblBatchStock.Stock, tblStore.Store, tblBatch.ItemID " & _
'                    " FROM tblStore RIGHT JOIN (tblBatchStock RIGHT JOIN tblBatch ON tblBatchStock.BatchID = tblBatch.BatchID) ON tblStore.StoreID = tblBatchStock.StoreID " & _
'                    " WHERE tblBatch.ItemID=" & ItemID & " AND tblBatchStock.Stock > 0 "
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            While .EOF = False
'                GridTotal.Rows = GridTotal.Rows + 1
'                GridTotal.Row = GridTotal.Rows - 1
'                GridTotal.Col = 0
'                GridTotal.CellAlignment = 1
'                GridTotal.Text = !Batch
'                GridTotal.Col = 1
'                GridTotal.CellAlignment = 7
'                If Not IsNull(!Stock) Then
'                    GridTotal.Text = !Stock
'                Else
'                    GridTotal.Text = 0
'                End If
'                GridTotal.Col = 2
'                GridTotal.CellAlignment = 1
'                GridTotal.Text = Format(!DOE, ShortDateFormat)
'                GridTotal.Col = 3
'                GridTotal.CellAlignment = 1
'                If Not IsNull(!Store) Then
'                    GridTotal.Text = !Store
'                    If !Store = UserStore Then
''                        lblStoreStock.Caption = !Store
'                        If Not IsNull(!Stock) Then
''                            txtStoreStock.Text = !Stock
'                        End If
'                    End If
'                Else
'                    GridTotal.Text = Empty
'                End If
'                .MoveNext
'            Wend
'        End If
'        GridTotal.Visible = True
'        .Close
'    End With
'
'End Sub
'
'Private Sub FillPurchase(ItemID As Long)
'    With rsTemOrder
'        If .State = 1 Then .Close
'        temSql = "SELECT tblRefill.Date, tblBatch.Batch, tblRefill.Amount, tblRefill.FreeAmount, tblRefill.DOE " & _
'                    "FROM tblRefill LEFT JOIN tblBatch ON tblRefill.BatchID = tblBatch.BatchID " & _
'                    "WHERE (((tblRefill.Date) Between '" & Format(dtpFDate.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTDate.Value, "dd MMMM yyyy") & "') AND ((tblRefill.ItemID)=" & ItemID & "))"
'
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            While .EOF = False
'                GridPurchase.Rows = GridPurchase.Rows + 1
'                GridPurchase.Row = GridPurchase.Rows - 1
'                GridPurchase.Col = 0
'                GridPurchase.CellAlignment = 4
'                GridPurchase.Text = Format(!Date, ShortDateFormat)
'                GridPurchase.Col = 1
'                GridPurchase.CellAlignment = 4
'                If IsNull(!Batch) = False Then
'                    GridPurchase.Text = !Batch
'                End If
'                GridPurchase.Col = 2
'                GridPurchase.CellAlignment = 7
'                GridPurchase.Text = !Amount
'                GridPurchase.Col = 3
'                GridPurchase.CellAlignment = 7
'                GridPurchase.Text = !FreeAmount
'                GridPurchase.Col = 4
'                GridPurchase.CellAlignment = 4
'                GridPurchase.Text = Format(!DOE, ShortDateFormat)
'                .MoveNext
'            Wend
'        End If
'    End With
'
'End Sub
'
'Private Sub FillPrice(ByVal ItemID As Long)
'    With rsTemPrice
'        If .State = 1 Then .Close
'        temSql = "SELECT tblPurchasePrice.SetDate, tblPurchasePrice.PPrice FROM tblPurchasePrice WHERE tblPurchasePrice.ItemID = " & ItemID & " AND tblPurchasePrice.SetDate between '" & Format(dtpFDate.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTDate.Value, "dd MMMM yyyy") & "'"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            While .EOF = False
'                With GridPPrice
'                    .Rows = .Rows + 1
'                    .Row = .Rows - 1
'                    .Col = 0
'                    .CellAlignment = 1
'                    .Text = Format(rsTemPrice!setdate, LongDateFormat)
'                    .Col = 1
'                    .CellAlignment = 7
'                    .Text = Format(rsTemPrice!PPrice * NewItem.IssueUnitsPerPack, "#,#00.00")
'                End With
'                .MoveNext
'            Wend
'        End If
'    End With
'    With rsTemPrice
'        If .State = 1 Then .Close
'        temSql = "SELECT tblSalePrice.SetDate, tblSalePrice.SPrice FROM tblSalePrice WHERE tblSalePrice.ItemID = " & ItemID & "   AND tblSalePrice.SetDate between '" & Format(dtpFDate.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTDate.Value, "dd MMMM yyyy") & "'"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            While .EOF = False
'                With GridSPrice
'                    .Rows = .Rows + 1
'                    .Row = .Rows - 1
'                    .Col = 0
'                    .CellAlignment = 1
'                    .Text = Format(rsTemPrice!setdate, LongDateFormat)
'                    .Col = 1
'                    .CellAlignment = 7
'                    .Text = Format(rsTemPrice!SPrice, "#,#00.00")
'                End With
'                .MoveNext
'            Wend
'        End If
'    End With
'End Sub
'
'Private Sub GetItemDetails(ItemID As Long)
'    NewItem.ID = ItemID
'    txtAMP.Text = NewItem.AMP
'    txtAMPP.Text = NewItem.AMPP
'    txtVMP.Text = NewItem.VMP
'    txtVMPP.Text = NewItem.VMPP
'    txtVTM.Text = NewItem.Generic
'    txtDisplay.Text = NewItem.Display
'End Sub
'
'Private Sub lblGrossTotal_Change()
'    lblNetTotal.Caption = Format((Val(lblGrossTotal.Caption) - Val(txtDiscount.Text)), "0.00")
'End Sub
'
'
'Private Sub txtDiscount_Change()
'    lblNetTotal.Caption = Format((Val(lblGrossTotal.Caption) - Val(txtDiscount.Text)), "0.00")
'End Sub
'
'Private Sub txtDiscount_LostFocus()
'    txtDiscount.Text = Format(txtDiscount.Text, "0.00")
'End Sub
'
'Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        txtFQty.SetFocus
'    ElseIf KeyCode = vbKeyEscape Then
'        txtQty.Text = Empty
'    End If
'End Sub
'
'
'Private Sub txtSPrice_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        txtBatch.SetFocus
'    End If
'End Sub
'
'
