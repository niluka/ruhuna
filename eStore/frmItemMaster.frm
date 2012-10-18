VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmItemMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Master"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11505
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
   ScaleHeight     =   9555
   ScaleWidth      =   11505
   Begin VB.Frame FrameData 
      Height          =   8895
      Left            =   4320
      TabIndex        =   34
      Top             =   120
      Width           =   7095
      Begin TabDlg.SSTab SSTab1 
         Height          =   7935
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   13996
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "&Product"
         TabPicture(0)   =   "frmItemMaster.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "bttnCreate"
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(2)=   "txtIssueUnitsPerPack"
         Tab(0).Control(3)=   "txtItemCode"
         Tab(0).Control(4)=   "txtDisplayName"
         Tab(0).Control(5)=   "txtStrengthPerIssueUnit"
         Tab(0).Control(6)=   "dtcTradeName"
         Tab(0).Control(7)=   "dtcSUnit"
         Tab(0).Control(8)=   "dtcPUnit"
         Tab(0).Control(9)=   "dtcIUnit"
         Tab(0).Control(10)=   "dtcCatogery"
         Tab(0).Control(11)=   "dtcGeneric"
         Tab(0).Control(12)=   "dtcManufacturer"
         Tab(0).Control(13)=   "Label17"
         Tab(0).Control(14)=   "Label7"
         Tab(0).Control(15)=   "Label4"
         Tab(0).Control(16)=   "lblIssueUnitsToPack"
         Tab(0).Control(17)=   "Label8"
         Tab(0).Control(18)=   "Label19"
         Tab(0).Control(19)=   "lblStrengthToIssueUnit"
         Tab(0).Control(20)=   "Label6"
         Tab(0).Control(21)=   "Label5"
         Tab(0).Control(22)=   "Label3"
         Tab(0).Control(23)=   "Label2"
         Tab(0).Control(24)=   "Label1"
         Tab(0).ControlCount=   25
         TabCaption(1)   =   "&Ordering"
         TabPicture(1)   =   "frmItemMaster.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label14"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label15"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label16"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label18"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label20"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "dtcMinQtyi"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "dtcMinQtyp"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "dtcImporter"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "dtcDistributor"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "dtcROQi"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "dtcROLi"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "dtcROQp"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "dtcROLp"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "dtlDistributors"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txtROQp"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "txtROLp"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "txtROLi"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "txtROQi"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "dtlDistributorIDs"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "txtMinQtyi"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "txtMinQtyp"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).ControlCount=   21
         Begin VB.TextBox txtMinQtyp 
            Height          =   375
            Left            =   1920
            TabIndex        =   12
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox txtMinQtyi 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1335
         End
         Begin btButtonEx.ButtonEx bttnCreate 
            Height          =   255
            Left            =   -71880
            TabIndex        =   9
            Top             =   4320
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            Appearance      =   3
            Caption         =   "&Create Names"
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
         Begin MSDataListLib.DataList dtlDistributorIDs 
            Height          =   2460
            Left            =   6000
            TabIndex        =   55
            Top             =   4800
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   4339
            _Version        =   393216
         End
         Begin VB.TextBox txtROQi 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtROLi 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtROLp 
            Height          =   375
            Left            =   1920
            TabIndex        =   10
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtROQp 
            Height          =   375
            Left            =   1920
            TabIndex        =   11
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Frame Frame1 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   37
            Top             =   5040
            Width           =   6615
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
               Left            =   2400
               TabIndex        =   46
               Top             =   2160
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
               Left            =   2400
               TabIndex        =   44
               Top             =   1680
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
               Left            =   2400
               TabIndex        =   42
               Top             =   1200
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
               Left            =   2400
               TabIndex        =   40
               Top             =   720
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
               Left            =   2400
               TabIndex        =   38
               Top             =   240
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
               TabIndex        =   47
               Top             =   2160
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
               TabIndex        =   45
               Top             =   1680
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
               TabIndex        =   43
               Top             =   1200
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
               TabIndex        =   41
               Top             =   720
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
               TabIndex        =   39
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.TextBox txtIssueUnitsPerPack 
            Height          =   375
            Left            =   -72600
            TabIndex        =   6
            Top             =   2880
            Width           =   1095
         End
         Begin VB.TextBox txtItemCode 
            Height          =   375
            Left            =   -73440
            TabIndex        =   7
            Top             =   3360
            Width           =   5175
         End
         Begin VB.TextBox txtDisplayName 
            Height          =   375
            Left            =   -73440
            TabIndex        =   36
            Top             =   4680
            Width           =   5175
         End
         Begin VB.TextBox txtStrengthPerIssueUnit 
            Height          =   375
            Left            =   -72600
            TabIndex        =   5
            Top             =   2400
            Width           =   1095
         End
         Begin MSDataListLib.DataCombo dtcTradeName 
            Height          =   360
            Left            =   -73440
            TabIndex        =   0
            Top             =   600
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcSUnit 
            Height          =   360
            Left            =   -74880
            TabIndex        =   2
            Top             =   1920
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcPUnit 
            Height          =   360
            Left            =   -70320
            TabIndex        =   4
            Top             =   1920
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcIUnit 
            Height          =   360
            Left            =   -72600
            TabIndex        =   3
            Top             =   1920
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataList dtlDistributors 
            Height          =   2460
            Left            =   1920
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   4800
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   4339
            _Version        =   393216
         End
         Begin MSDataListLib.DataCombo dtcROLp 
            Height          =   360
            Left            =   3360
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcROQp 
            Height          =   360
            Left            =   3360
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1680
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcROLi 
            Height          =   360
            Left            =   3360
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1080
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcROQi 
            Height          =   360
            Left            =   3360
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   2160
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcDistributor 
            Height          =   360
            Left            =   1920
            TabIndex        =   53
            Top             =   4800
            Visible         =   0   'False
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcImporter 
            Height          =   360
            Left            =   1920
            TabIndex        =   13
            Top             =   4080
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcCatogery 
            Height          =   360
            Left            =   -73440
            TabIndex        =   1
            Top             =   1080
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcMinQtyp 
            Height          =   360
            Left            =   3360
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   2760
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcMinQtyi 
            Height          =   360
            Left            =   3360
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   3240
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcGeneric 
            Height          =   360
            Left            =   -73440
            TabIndex        =   61
            Top             =   1080
            Visible         =   0   'False
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dtcManufacturer 
            Height          =   360
            Left            =   -73440
            TabIndex        =   8
            Top             =   3840
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   635
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label Label17 
            Caption         =   "Manufacturer"
            Height          =   255
            Left            =   -74880
            TabIndex        =   26
            Top             =   3840
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Issue Unit:"
            Height          =   255
            Left            =   -72600
            TabIndex        =   21
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label20 
            Caption         =   "&Minimum Order"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Catogery:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   19
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblIssueUnitsToPack 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   -71400
            TabIndex        =   59
            Top             =   2880
            Width           =   3135
         End
         Begin VB.Label Label8 
            Caption         =   "Strength of an Issue Unit"
            Height          =   255
            Left            =   -74880
            TabIndex        =   23
            Top             =   2400
            Width           =   2415
         End
         Begin VB.Label Label19 
            Caption         =   "Pack Unit:"
            Height          =   255
            Left            =   -70200
            TabIndex        =   22
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label18 
            Caption         =   "&Importer"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   4080
            Width           =   2055
         End
         Begin VB.Label Label16 
            Caption         =   "&Dealers"
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   4800
            Width           =   1935
         End
         Begin VB.Label Label15 
            Caption         =   "&Reorder Leval"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "Reorder &Quentity"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblStrengthToIssueUnit 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   -71400
            TabIndex        =   35
            Top             =   2400
            Width           =   3135
         End
         Begin VB.Label Label6 
            Caption         =   "Strength Unit :"
            Height          =   255
            Left            =   -74880
            TabIndex        =   20
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Display Name :"
            Height          =   255
            Left            =   -74760
            TabIndex        =   27
            Top             =   4680
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Item Code:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   25
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Issue units per Pack"
            Height          =   255
            Left            =   -74880
            TabIndex        =   24
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Trade Name :"
            Height          =   255
            Left            =   -74880
            TabIndex        =   18
            Top             =   600
            Width           =   1695
         End
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   495
         Left            =   3960
         TabIndex        =   28
         Top             =   8280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "&Save"
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   495
         Left            =   3960
         TabIndex        =   30
         Top             =   8280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "&Save"
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
         Height          =   495
         Left            =   5520
         TabIndex        =   29
         Top             =   8280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "&Cancel"
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
   Begin VB.Frame FrameSearch 
      Height          =   8535
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   4095
      Begin btButtonEx.ButtonEx bttnEdit 
         Height          =   495
         Left            =   2400
         TabIndex        =   17
         Top             =   7920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "&Edit"
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
      Begin MSDataListLib.DataCombo dtcItem 
         Height          =   6900
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   12171
         _Version        =   393216
         Style           =   1
         Text            =   ""
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
      Begin btButtonEx.ButtonEx bttnAdd 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   7920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
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
      Begin MSDataListLib.DataCombo dtcItemCategory 
         Height          =   360
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   9840
      TabIndex        =   31
      Top             =   9120
      Width           =   1455
      _ExtentX        =   2566
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
End
Attribute VB_Name = "frmItemMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim NewItem As New Item
    Dim temSql As String
    Dim rsItem As New ADODB.Recordset
    Dim rsTemItem As New ADODB.Recordset
    Dim rsManufacturer As New ADODB.Recordset
    Dim rsImporter As New ADODB.Recordset
    Dim rsDistributor As New ADODB.Recordset
    Dim rsViewItemDistributor As New ADODB.Recordset
    Dim rsIUnit As New ADODB.Recordset
    Dim rsPUnit As New ADODB.Recordset
    Dim rsSUnit As New ADODB.Recordset
    Dim rsGeneric As New ADODB.Recordset
    Dim rsTrade As New ADODB.Recordset
    Dim rsTemTrade As New ADODB.Recordset
    Dim rsCatogery As New ADODB.Recordset
    
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
    With rsCatogery
        If .State = 1 Then .Close
        temSql = "SELECT tblItemCategory.ItemCategoryID, tblItemCategory.ItemCategory FROM tblItemCategory ORDER BY tblItemCategory.ItemCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcCatogery
        Set .RowSource = rsCatogery
        .ListField = "ItemCategory"
        .BoundColumn = "ItemCategoryID"
    End With
    
    With rsTrade
        If .State = 1 Then .Close
        temSql = "SELECT tblTradeName.* From tblTradeName ORDER BY tblTradeName.TradeName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcTradeName
        Set .RowSource = rsTrade
        .ListField = "TradeName"
        .BoundColumn = "TradeNameID"
    End With
    With rsGeneric
        If .State = 1 Then .Close
        temSql = "SELECT tblgenericName.* From tblgenericName ORDER BY genericName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcGeneric
        Set .RowSource = rsGeneric
        .ListField = "genericName"
        .BoundColumn = "genericNameID"
    End With
    With rsManufacturer
        If .State = 1 Then .Close
        temSql = "SELECT tblManufacturer.* FROM tblManufacturer ORDER BY tblManufacturer.ManufacturerName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcManufacturer
        Set .RowSource = rsManufacturer
        .ListField = "ManufacturerName"
        .BoundColumn = "ManufacturerID"
    End With
    With rsImporter
        If .State = 1 Then .Close
        temSql = "SELECT tblImporter.* From tblImporter ORDER BY tblImporter.ImporterName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcImporter
        Set .RowSource = rsImporter
        .ListField = "ImporterName"
        .BoundColumn = "ImporterID"
    End With
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
    
    With rsSUnit
        If .State = 1 Then .Close
        temSql = "SELECT tblstrengthUnit.* From tblstrengthUnit ORDER BY strengthUnit"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcSUnit
        Set .RowSource = rsSUnit
        .ListField = "StrengthUnit"
        .BoundColumn = "StrengthUnitID"
    End With
    
    
    With rsIUnit
        If .State = 1 Then .Close
        temSql = "SELECT tblissueUnit.* From tblissueUnit ORDER BY issueUnit"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcIUnit
        Set .RowSource = rsIUnit
        .ListField = "IssueUnit"
        .BoundColumn = "IssueUnitID"
    End With
    With dtcROLi
        Set .RowSource = rsIUnit
        .ListField = "IssueUnit"
        .BoundColumn = "IssueUnitID"
    End With
    With dtcROQi
        Set .RowSource = rsIUnit
        .ListField = "IssueUnit"
        .BoundColumn = "IssueUnitID"
    End With
    With dtcMinQtyi
        Set .RowSource = rsIUnit
        .ListField = "IssueUnit"
        .BoundColumn = "IssueUnitID"
    End With
    
    With rsPUnit
        If .State = 1 Then .Close
        temSql = "SELECT tblpackUnit.* From tblpackUnit ORDER BY packUnit"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcPUnit
        Set .RowSource = rsPUnit
        .ListField = "PackUnit"
        .BoundColumn = "packUnitID"
    End With
    With dtcROLp
        Set .RowSource = rsPUnit
        .ListField = "PackUnit"
        .BoundColumn = "packUnitID"
    End With
    With dtcROQp
        Set .RowSource = rsPUnit
        .ListField = "PackUnit"
        .BoundColumn = "packUnitID"
    End With
    With dtcMinQtyp
        Set .RowSource = rsPUnit
        .ListField = "PackUnit"
        .BoundColumn = "packUnitID"
    End With


End Sub

Private Sub BeforeAddEdit()
    FrameData.Enabled = False
    FrameSearch.Enabled = True
    bttnCancel.Visible = False
    bttnSave.Visible = False
    bttnChange.Visible = False
    SSTab1.Tab = 0
End Sub

Private Sub AfterAdd()
    FrameData.Enabled = True
    FrameSearch.Enabled = False
    bttnCancel.Visible = True
    bttnSave.Visible = True
    bttnChange.Visible = False
End Sub

Private Sub AfterEdit()
    FrameData.Enabled = True
    FrameSearch.Enabled = False
    bttnCancel.Visible = True
    bttnSave.Visible = False
    bttnChange.Visible = True
End Sub

Private Sub ClearData()
    NewItem.ID = Empty
    Me.txtAMP = Empty
    Me.txtAMPP.Text = Empty
    Me.txtDisplayName.Text = Empty
    Me.txtItemCode.Text = Empty
    Me.txtIssueUnitsPerPack = Empty
    Me.txtROLi.Text = Empty
    Me.txtROLp.Text = Empty
    Me.txtROQi.Text = Empty
    Me.txtROQp.Text = Empty
    Me.txtStrengthPerIssueUnit = Empty
    Me.txtVMP.Text = Empty
    Me.txtVMPP.Text = Empty
    Me.txtVTM.Text = Empty
    Me.txtMinQtyi.Text = Empty
    Me.txtMinQtyp.Text = Empty
    Me.dtcImporter.Text = Empty
    Me.dtcIUnit.Text = Empty
    Me.dtcManufacturer.Text = Empty
    Me.dtcPUnit.Text = Empty
    Me.dtcROLi.Text = Empty
    Me.dtcROLp.Text = Empty
    Me.dtcROQi.Text = Empty
    Me.dtcROQp.Text = Empty
    Me.dtcMinQtyi.Text = Empty
    Me.dtcMinQtyp.Text = Empty
    Me.dtcSUnit.Text = Empty
    Me.dtcTradeName.Text = Empty
    Me.dtcCatogery.Text = Empty
    Me.lblIssueUnitsToPack.Caption = Empty
    Me.lblStrengthToIssueUnit.Caption = Empty
    Set dtlDistributorIDs.RowSource = Nothing
    Set dtlDistributors.RowSource = Nothing
End Sub

Private Sub DisplayData()
    On Error GoTo eh
    If Not IsNumeric(dtcItem.BoundText) Then Exit Sub
    NewItem.ID = Val(dtcItem.BoundText)
    Me.txtAMP = NewItem.AMP
    Me.txtAMPP.Text = NewItem.AMPP
    Me.txtDisplayName.Text = NewItem.Display
    Me.txtItemCode.Text = NewItem.Code
    Me.txtIssueUnitsPerPack = NewItem.IssueUnitsPerPack
    Me.txtROLi.Text = NewItem.ROL
    Me.txtROLp.Text = NewItem.ROL / NewItem.IssueUnitsPerPack
    Me.txtROQi.Text = NewItem.ROQ
    Me.txtROQp.Text = NewItem.ROQ / NewItem.IssueUnitsPerPack
    Me.txtStrengthPerIssueUnit = NewItem.StrengthOfIssueUnit
    Me.txtVMP.Text = NewItem.VMP
    Me.txtVMPP.Text = NewItem.VMPP
    Me.txtVTM.Text = NewItem.Generic
    Me.dtcImporter.Text = NewItem.ImporterName
    Me.dtcIUnit.Text = NewItem.IUnit
    Me.dtcManufacturer.Text = NewItem.ManufacturerName
    Me.dtcPUnit.Text = NewItem.PUnit
    Me.dtcGeneric.Text = NewItem.Generic
    Me.dtcROLi.Text = NewItem.IUnit
    Me.dtcROLp.Text = NewItem.PUnit
    Me.dtcROQi.Text = NewItem.IUnit
    Me.dtcROQp.Text = NewItem.PUnit
    Me.dtcMinQtyi.Text = NewItem.IUnit
    Me.dtcMinQtyp.Text = NewItem.PUnit
    Me.txtMinQtyi.Text = NewItem.MinQty
    Me.txtMinQtyp.Text = NewItem.MinQty / NewItem.IssueUnitsPerPack
    Me.dtcSUnit.Text = NewItem.StrengthUnit
    Me.dtcTradeName.Text = NewItem.Trade
    Me.dtcCatogery.Text = NewItem.Category
    Me.lblIssueUnitsToPack.Caption = NewItem.IUnit & " / " & NewItem.PUnit
    Me.lblStrengthToIssueUnit.Caption = NewItem.StrengthUnit & " / " & NewItem.IUnit
    Exit Sub
eh:
    Exit Sub
End Sub

Private Sub SaveData()
With rsTemItem
    If .State = 1 Then .Close
    temSql = "SELECT tblItem.* FROM tblItem"
    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
    .AddNew
    !VTM = txtVTM.Text
    !VMP = txtVMP.Text
    !AMP = txtAMP.Text
    !AMPP = txtAMPP.Text
    !VMPP = txtVMPP.Text
    !Display = txtDisplayName.Text
    !Code = txtItemCode.Text
    !TradeNameID = dtcTradeName.BoundText
    !GenericNameID = dtcGeneric.BoundText
    !ItemCategoryID = dtcCatogery.BoundText
    !StrengthUnitID = dtcSUnit.BoundText
    !StrengthOfIssueUnit = Val(txtStrengthPerIssueUnit.Text)
    !IssueUnitsPerPack = Val(txtIssueUnitsPerPack.Text)
    !IssueUnitID = dtcIUnit.BoundText
    !PackUnitID = dtcPUnit.BoundText
    !ROL = Val(txtROLi.Text)
    !ROQ = Val(txtROQi.Text)
    !MinQty = Val(txtMinQtyi.Text)
    !ManufacturerID = dtcManufacturer.BoundText
    !ImporterID = dtcImporter.BoundText
    .Update
    If .State = 1 Then .Close
    Exit Sub
eh:
    Dim tr As Integer
    tr = MsgBox("Could Not Update the database" & vbNewLine & Err.Description, vbCritical, "Error")
    .CancelUpdate
    If .State = 1 Then .Close
    Exit Sub
End With
End Sub

Private Sub ChangeData()
With rsTemItem
    If .State = 1 Then .Close
    temSql = "SELECT tblItem.* FROM tblItem where itemid = " & dtcItem.BoundText
    .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
    !VTM = txtVTM.Text
    !VMP = txtVMP.Text
    !AMP = txtAMP.Text
    !AMPP = txtAMPP.Text
    !VMPP = txtVMPP.Text
    !Display = txtDisplayName.Text
    !Code = txtItemCode.Text
    !TradeNameID = dtcTradeName.BoundText
    !GenericNameID = dtcGeneric.BoundText
    !ItemCategoryID = dtcCatogery.BoundText
    !StrengthUnitID = dtcSUnit.BoundText
    !StrengthOfIssueUnit = Val(txtStrengthPerIssueUnit.Text)
    !IssueUnitsPerPack = Val(txtIssueUnitsPerPack.Text)
    !IssueUnitID = dtcIUnit.BoundText
    !PackUnitID = dtcPUnit.BoundText
    !ROL = Val(txtROLi.Text)
    !ROQ = Val(txtROQi.Text)
    !MinQty = Val(txtMinQtyi.Text)
    !ManufacturerID = dtcManufacturer.BoundText
    !ImporterID = dtcImporter.BoundText
    .Update
    If .State = 1 Then .Close
    Exit Sub
eh:
    Dim tr As Integer
    tr = MsgBox("Could Not Update the database" & vbNewLine & Err.Description, vbCritical, "Error")
    .CancelUpdate
    If .State = 1 Then .Close
    Exit Sub
End With
End Sub

Private Function CanAdd() As Boolean
    CanAdd = False
    Dim tr As Integer
    If Trim(txtVTM.Text) = "" Then
        SSTab1.Tab = 0
        txtVTM.SetFocus
        Exit Function
    End If
    If Trim(txtVMP.Text) = "" Then
        SSTab1.Tab = 0
        txtVMP.SetFocus
        Exit Function
    End If
    If Trim(txtAMP.Text) = "" Then
        SSTab1.Tab = 0
        txtAMP.SetFocus
        Exit Function
    End If
    If Trim(txtAMPP.Text) = "" Then
        SSTab1.Tab = 0
        txtAMPP.SetFocus
        Exit Function
    End If
    If Trim(txtVMPP.Text) = "" Then
        SSTab1.Tab = 0
        txtVMPP.SetFocus
        Exit Function
    End If
    If Trim(txtDisplayName.Text) = "" Then
        SSTab1.Tab = 0
        txtDisplayName.SetFocus
        Exit Function
    End If
    If Trim(txtItemCode.Text) = "" Then
        SSTab1.Tab = 0
        txtItemCode.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcTradeName.BoundText) Then
        SSTab1.Tab = 0
        dtcTradeName.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcSUnit.BoundText) Then
        SSTab1.Tab = 0
        dtcSUnit.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcCatogery.BoundText) Then
        SSTab1.Tab = 0
        dtcCatogery.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcGeneric.BoundText) Then
        Exit Function
    End If
    If Not IsNumeric(dtcPUnit.BoundText) Then
        SSTab1.Tab = 0
        dtcPUnit.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcIUnit.BoundText) Then
        SSTab1.Tab = 0
        dtcIUnit.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcImporter.BoundText) Then
        SSTab1.Tab = 1
        dtcImporter.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcManufacturer.BoundText) Then
        SSTab1.Tab = 0
        dtcManufacturer.SetFocus
        Exit Function
    End If
        
    If Not IsNumeric(txtStrengthPerIssueUnit.Text) Then
        SSTab1.Tab = 0
        txtStrengthPerIssueUnit.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtIssueUnitsPerPack.Text) Then
        SSTab1.Tab = 0
        txtIssueUnitsPerPack.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtROLi.Text) Then
        SSTab1.Tab = 1
        txtROLi.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtROQi.Text) Then
        SSTab1.Tab = 1
        txtROQi.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtMinQtyi.Text) Then
        SSTab1.Tab = 1
        txtMinQtyi.SetFocus
        Exit Function
    End If
    CanAdd = True
End Function


Private Sub FillItemDistributors()
If IsNumeric(dtcItem.BoundText) = False Then Exit Sub
With rsViewItemDistributor
    If .State = 1 Then .Close
    .Open "SELECT tblItemDistributor.*, tblDistrubutor.DistributorName FROM tblDistrubutor RIGHT JOIN tblItemDistributor ON tblDistrubutor.DistributorID = tblItemDistributor.DistributorID Where ItemID = " & dtcItem.BoundText & "", cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount = 0 Then Exit Sub
    Set dtlDistributors.RowSource = rsViewItemDistributor
    dtlDistributors.BoundColumn = "ItemDistributorID"
    dtlDistributors.ListField = "DistributorName"
End With
End Sub

Private Sub bttnCancel_Click()
    Call ClearData
    Call BeforeAddEdit
End Sub

Private Sub bttnChange_Click()
    If CanAdd = True Then
        Call ChangeData
        Call ClearData
        Call BeforeAddEdit
        Call FillCombos
    Else
        Dim tr As Integer
        tr = MsgBox("You have to give valid values for all the fields to save the record", vbCritical, "Not Valid")
        Exit Sub
    End If
End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnCreate_Click()
    If CanCreate = True Then
        Call CreateNames
    Else
        Dim tr As Integer
        tr = MsgBox("You have to give valid values for all this field to create names", vbCritical, "Not Enough Data")
        Exit Sub
    End If
End Sub

Private Sub CreateNames()
    txtVTM.Text = dtcGeneric.Text
    txtVMP.Text = dtcGeneric.Text & " " & txtStrengthPerIssueUnit.Text & dtcSUnit.Text & " " & dtcIUnit.Text
    txtAMP.Text = dtcTradeName.Text & " " & txtStrengthPerIssueUnit.Text & dtcSUnit.Text & " " & dtcIUnit.Text
    If txtAMP.Text = txtVMP.Text Then txtAMP.Text = dtcTradeName.Text & "(" & dtcManufacturer.Text & ") " & txtStrengthPerIssueUnit.Text & dtcSUnit.Text & " " & dtcIUnit.Text
    txtVMPP.Text = dtcGeneric.Text & " " & txtStrengthPerIssueUnit.Text & dtcSUnit.Text & " " & dtcIUnit.Text & " - " & txtIssueUnitsPerPack.Text & " " & dtcIUnit.Text & " " & dtcPUnit.Text
    txtAMPP.Text = dtcTradeName.Text & " " & txtStrengthPerIssueUnit.Text & dtcSUnit.Text & " " & dtcIUnit.Text & " - " & txtIssueUnitsPerPack.Text & " " & dtcIUnit.Text & " " & dtcPUnit.Text
    If txtAMPP.Text = txtVMPP.Text Then txtAMPP.Text = dtcTradeName.Text & "(" & dtcManufacturer.Text & ") " & txtStrengthPerIssueUnit.Text & dtcSUnit.Text & " " & dtcIUnit.Text & " - " & txtIssueUnitsPerPack.Text & " " & dtcIUnit.Text & " " & dtcPUnit.Text
    txtDisplayName.Text = txtAMP.Text
End Sub

Private Function CanCreate() As Boolean
    CanCreate = False
    Dim tr As Integer
    If Not IsNumeric(dtcTradeName.BoundText) Then
        SSTab1.Tab = 0
        dtcTradeName.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcSUnit.BoundText) Then
        SSTab1.Tab = 0
        dtcSUnit.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcCatogery.BoundText) Then
        SSTab1.Tab = 0
        dtcCatogery.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcGeneric.BoundText) Then
        SSTab1.Tab = 0
'        dtcGeneric.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcPUnit.BoundText) Then
        SSTab1.Tab = 0
        dtcPUnit.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcIUnit.BoundText) Then
        SSTab1.Tab = 0
        dtcIUnit.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtStrengthPerIssueUnit.Text) Then
        SSTab1.Tab = 0
        txtStrengthPerIssueUnit.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtIssueUnitsPerPack.Text) Then
        SSTab1.Tab = 0
        txtIssueUnitsPerPack.SetFocus
        Exit Function
    End If
    If Not IsNumeric(dtcManufacturer.BoundText) Then
        SSTab1.Tab = 0
        dtcManufacturer.SetFocus
        Exit Function
    End If
    CanCreate = True
End Function

Private Sub bttnCreate_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtROLi.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub bttnPrint_Click()
    dtrCategoryWise.Show
End Sub

Private Sub dtcCatogery_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtcSUnit.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub dtcImporter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        bttnSave.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub dtcItemCategory_Change()
    If IsNumeric(dtcItemCategory.BoundText) = False Then
        Call ListAllItems
    Else
        Call ListSelectedItems
    End If
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
'    With rsCode
'        If .State = 1 Then .Close
'        temSQL = "SELECT * from tblitem where ItemCategoryID = " & dtcCatogery.BoundText & " order by code"
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With dtcCode
'        Set .RowSource = rsCode
'        .ListField = "Code"
'        .BoundColumn = "ItemID"
'    End With
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
'With rsCode
'    If .State = 1 Then .Close
'    temSQL = "SELECT * from tblitem order by code"
'    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'End With
'With dtcCode
'    Set .RowSource = rsCode
'    .ListField = "Code"
'    .BoundColumn = "ItemID"
'End With
End Sub


Private Sub dtcItemCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtcItemCategory.Text = Empty
        KeyCode = Empty
    End If
End Sub

Private Sub dtcIUnit_Change()
    Call SetUnits
End Sub

Private Sub dtcIUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtcPUnit.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub dtcManufacturer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDisplayName.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub dtcPUnit_Change()
    Call SetUnits
End Sub

Private Sub dtcPUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtStrengthPerIssueUnit.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub dtcSUnit_Change()
    Call SetUnits
End Sub

Private Sub dtcSUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtcIUnit.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub dtcTradeName_Click(Area As Integer)
    If Not IsNumeric(dtcTradeName.BoundText) Then Exit Sub
    With rsTemTrade
        If .State = 1 Then .Close
        temSql = "SELECT tblTradeName.GenericNameID FROM tblTradeName WHERE (((tblTradeName.TradeNameID)=" & dtcTradeName.BoundText & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 1 Then Exit Sub
        dtcGeneric.BoundText = !GenericNameID
        If .State = 1 Then .Close
    End With
End Sub

Private Sub dtcTradeName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtcCatogery.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub Form_Load()
    Call BeforeAddEdit
    Call FillCombos
End Sub

Private Sub dtcItem_Click(Area As Integer)
    Call DisplayData
    Call FillItemDistributors
End Sub

Private Sub bttnAdd_Click()
    Call ClearData
    Call AfterAdd
    SSTab1.Tab = 0
    dtcTradeName.SetFocus
    txtItemCode.Text = LastCode + 1
End Sub

Private Function LastCode() As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem order by ItemID desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            LastCode = Val(!Code)
        Else
            LastCode = 0
        End If
    End With
End Function

Private Sub bttnEdit_Click()
    Call AfterEdit
    If SSTab1.Tab = 0 Then dtcTradeName.SetFocus
End Sub

Private Sub bttnSave_Click()
    If CanAdd = True Then
        Call SaveData
        Call ClearData
        Call BeforeAddEdit
        Call FillCombos
    Else
        Dim tr As Integer
        tr = MsgBox("You have to give valid values for all the fields to save the record", vbCritical, "Not Valid")
        Exit Sub
    End If
End Sub


Private Sub SetUnits()
    lblStrengthToIssueUnit.Caption = dtcSUnit.Text & " / " & dtcIUnit.Text
    lblIssueUnitsToPack.Caption = dtcIUnit.Text & " / " & dtcPUnit.Text
    dtcROLi.Text = dtcIUnit.Text
    dtcROLp.Text = dtcPUnit.Text
    dtcROQi.Text = dtcIUnit.Text
    dtcROQp.Text = dtcPUnit.Text
    dtcMinQtyi.Text = dtcIUnit.Text
    dtcMinQtyp.Text = dtcPUnit.Text
End Sub

Private Sub SetValues()
    txtROLp.Text = Val(txtROLi.Text) / Val(txtIssueUnitsPerPack.Text)
    txtROLi.Text = Val(txtROLp.Text) * Val(txtIssueUnitsPerPack.Text)
    txtROQp.Text = Val(txtROQi.Text) / Val(txtIssueUnitsPerPack.Text)
    txtROQi.Text = Val(txtROQp.Text) * Val(txtIssueUnitsPerPack.Text)
End Sub

Private Sub txtIssueUnitsPerPack_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtItemCode.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtcManufacturer.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub txtROLi_LostFocus()
    txtROLp.Text = Val(txtROLi.Text) / Val(txtIssueUnitsPerPack.Text)
End Sub

Private Sub txtROLp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtROQp.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub txtROLp_LostFocus()
    txtROLi.Text = Val(txtROLp.Text) * Val(txtIssueUnitsPerPack.Text)
End Sub

Private Sub txtROQi_LostFocus()
    txtROQp.Text = Val(txtROQi.Text) / Val(txtIssueUnitsPerPack.Text)
End Sub

Private Sub txtROQp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtMinQtyi.SetFocus
        KeyCode = Empty
    End If
End Sub

Private Sub txtROQp_LostFocus()
    txtROQi.Text = Val(txtROQp.Text) * Val(txtIssueUnitsPerPack.Text)
End Sub

Private Sub txtminqtyi_LostFocus()
    txtMinQtyp.Text = Val(txtMinQtyi.Text) / Val(txtIssueUnitsPerPack.Text)
End Sub

Private Sub txtMinQtyp_LostFocus()
    txtMinQtyi.Text = Val(txtMinQtyp.Text) * Val(txtIssueUnitsPerPack.Text)
End Sub

Private Sub txtStrengthPerIssueUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtIssueUnitsPerPack.SetFocus
        KeyCode = Empty
    End If
End Sub
