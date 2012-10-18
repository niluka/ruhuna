VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmIxFormat 
   Caption         =   "Field"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
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
   WindowState     =   2  'Maximized
   Begin VB.Frame frameSaveCancel 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4200
      TabIndex        =   52
      Top             =   8520
      Width           =   3255
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   120
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
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   120
         TabIndex        =   54
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
         Left            =   1680
         TabIndex        =   55
         Top             =   120
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
   End
   Begin MSDataListLib.DataCombo dtcIx 
      Bindings        =   "LabIxFormat.frx":0000
      Height          =   360
      Left            =   120
      TabIndex        =   59
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   "DataCombo1"
   End
   Begin VB.Frame FramePaper 
      Caption         =   "Paper"
      Height          =   10455
      Left            =   4680
      TabIndex        =   56
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtField 
         Height          =   360
         Index           =   0
         Left            =   1920
         TabIndex        =   57
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Index           =   0
         Left            =   480
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label lblLabels 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   58
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   480
         X2              =   3120
         Y1              =   2280
         Y2              =   2280
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab sstab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5953
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Labels"
      TabPicture(0)   =   "LabIxFormat.frx":001F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstLabels"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Fields"
      TabPicture(1)   =   "LabIxFormat.frx":003B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstFields"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Lines"
      TabPicture(2)   =   "LabIxFormat.frx":0057
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstLines"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.ListBox lstLines 
         Height          =   2700
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   3615
      End
      Begin VB.ListBox lstLabels 
         Height          =   2700
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   3615
      End
      Begin VB.ListBox lstFields 
         Height          =   2700
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   3615
      End
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Add"
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
      Left            =   1440
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Delete"
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
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Frame frameLabels 
      Caption         =   "Labels"
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   3735
      Begin VB.TextBox txtLabelName 
         Height          =   960
         Left            =   720
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   2895
      End
      Begin btButtonEx.ButtonEx bttnLabelPosRight 
         Height          =   375
         Left            =   2760
         TabIndex        =   26
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Æ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLabelPosUp 
         Height          =   375
         Left            =   2400
         TabIndex        =   27
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Ç"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLabelPosLeft 
         Height          =   375
         Left            =   2040
         TabIndex        =   28
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Å"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLabelPosDown 
         Height          =   375
         Left            =   2400
         TabIndex        =   29
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "È"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLabelSizeRight 
         Height          =   375
         Left            =   1560
         TabIndex        =   30
         Top             =   2520
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Æ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLabelSizeUp 
         Height          =   375
         Left            =   2400
         TabIndex        =   31
         Top             =   2520
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Ç"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLabelSizeLeft 
         Height          =   375
         Left            =   1200
         TabIndex        =   32
         Top             =   2520
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Å"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLabelSizeDown 
         Height          =   375
         Left            =   2760
         TabIndex        =   33
         Top             =   2520
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "È"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLabelFont 
         Height          =   375
         Left            =   840
         TabIndex        =   34
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Font"
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
      Begin btButtonEx.ButtonEx bttnLabelColour 
         Height          =   375
         Left            =   840
         TabIndex        =   35
         Top             =   3480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Colour"
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
      Begin VB.Label Label7 
         Caption         =   "Position"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Size"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frameFields 
      Caption         =   "Fields"
      Height          =   5775
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   3735
      Begin VB.TextBox txtFieldName 
         Height          =   360
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   2415
      End
      Begin btButtonEx.ButtonEx bttnFieldPosRight 
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Æ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnFieldPosUp 
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Ç"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnFieldPosLeft 
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Å"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnFieldPosDown 
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "È"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnFieldSizeRight 
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   2400
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Æ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnFieldSizeUp 
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Ç"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnFieldSizeLeft 
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   2400
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Å"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnFieldSizeDown 
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   2760
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "È"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnFieldFont 
         Height          =   375
         Left            =   840
         TabIndex        =   22
         Top             =   3600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Font"
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
      Begin btButtonEx.ButtonEx BttnFieldColour 
         Height          =   375
         Left            =   840
         TabIndex        =   23
         Top             =   4080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Colour"
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
      Begin btButtonEx.ButtonEx bttnDefaultValues 
         Height          =   375
         Left            =   840
         TabIndex        =   24
         Top             =   4560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Default Values"
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
      Begin VB.Label Label4 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Size"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Position"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame frameLines 
      Caption         =   "Lines"
      Height          =   5775
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   3735
      Begin VB.OptionButton optBox 
         Caption         =   "Box"
         Height          =   375
         Left            =   1200
         TabIndex        =   51
         Top             =   3360
         Width           =   2295
      End
      Begin VB.OptionButton optLine 
         Caption         =   "Line"
         Height          =   375
         Left            =   1200
         TabIndex        =   50
         Top             =   3000
         Width           =   1935
      End
      Begin btButtonEx.ButtonEx bttnLinePosRight 
         Height          =   375
         Left            =   3000
         TabIndex        =   39
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Æ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLinePosUp 
         Height          =   375
         Left            =   2640
         TabIndex        =   40
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Ç"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLinePosLeft 
         Height          =   375
         Left            =   2280
         TabIndex        =   41
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Å"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLinePosDown 
         Height          =   375
         Left            =   2640
         TabIndex        =   42
         Top             =   1080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "È"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLineSizeRight 
         Height          =   375
         Left            =   3000
         TabIndex        =   43
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Æ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLineSizeUp 
         Height          =   375
         Left            =   2640
         TabIndex        =   44
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Ç"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLineSizeLeft 
         Height          =   375
         Left            =   2280
         TabIndex        =   45
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Å"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnLineSizeDown 
         Height          =   375
         Left            =   2640
         TabIndex        =   46
         Top             =   2400
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "È"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx ButtonEx31 
         Height          =   375
         Left            =   1200
         TabIndex        =   49
         Top             =   4320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Default Values"
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
      Begin VB.Label Label9 
         Caption         =   "Position"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Size"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   2040
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmIxFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim LabelCount As Long
    Dim FieldCount As Long
    Dim LineCount As Long
    Dim MoveValue As Long
    Dim rsIxList As New ADODB.Recordset

Private Sub BeforeAddEdit()
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    bttnDelete.Enabled = True
    frameLines.Enabled = False
    sstab1.Enabled = True
    frameFields.Enabled = False
    frameLabels.Enabled = False
    frameSaveCancel.Enabled = False
    FramePaper.Enabled = False
End Sub

Private Sub AfterAdd()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    sstab1.Enabled = False
    frameLines.Enabled = True
    frameFields.Enabled = True
    frameLabels.Enabled = True
    frameSaveCancel.Enabled = True
    FramePaper.Enabled = False
    bttnSave.Visible = True
    bttnChange.Visible = False
End Sub

Private Sub AfterEdit()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    sstab1.Enabled = False
    frameLines.Enabled = True
    frameFields.Enabled = True
    frameLabels.Enabled = True
    frameSaveCancel.Enabled = True
    FramePaper.Enabled = False
    bttnSave.Visible = False
    bttnChange.Visible = True
End Sub

Private Sub bttnAdd_Click()
    Call AfterAdd
    Select Case sstab1
        Case 0: Call AddNewLabel
        Case 1: Call AddNewField
        Case 2: Call AddNewLine
    End Select
End Sub


Private Sub bttnLabelPosRight_Click()
    lblLabels(LabelCount).Left = lblLabels(LabelCount).Left + MoveValue
End Sub
Private Sub bttnLabelPosDown_Click()
    lblLabels(LabelCount).Top = lblLabels(LabelCount).Top + MoveValue
End Sub
Private Sub bttnLabelPosLeft_Click()
    lblLabels(LabelCount).Left = lblLabels(LabelCount).Left - MoveValue
End Sub
Private Sub bttnLabelPosUp_Click()
    lblLabels(LabelCount).Top = lblLabels(LabelCount).Top - MoveValue
End Sub

Private Sub bttnLabelSizeDown_Click()
    lblLabels(LabelCount).Top = lblLabels(LabelCount).Top + MoveValue / 2
    lblLabels(LabelCount).Height = lblLabels(LabelCount).Height - MoveValue
End Sub

Private Sub bttnLabelSizeLeft_Click()
    lblLabels(LabelCount).Left = lblLabels(LabelCount).Left - MoveValue / 2
    lblLabels(LabelCount).Width = lblLabels(LabelCount).Width + MoveValue
End Sub

Private Sub bttnLabelSizeRight_Click()
    lblLabels(LabelCount).Left = lblLabels(LabelCount).Left + MoveValue / 2
    lblLabels(LabelCount).Width = lblLabels(LabelCount).Width - MoveValue
End Sub

Private Sub bttnLabelSizeUp_Click()
    lblLabels(LabelCount).Top = lblLabels(LabelCount).Top - MoveValue / 2
    lblLabels(LabelCount).Height = lblLabels(LabelCount).Height + MoveValue
End Sub

Private Sub Form_Load()
    MoveValue = 20
    Call FillIxCombo
End Sub

Private Sub FillIxCombo()
    With rsIxList
        dtcIx.ListField = Empty
        dtcIx.BoundText = Empty
        If .State = 1 Then .Close
        .Open "SELECT * from tblIx order by Ix", dbHospital, adOpenStatic, adLockOptimistic
        Set dtcIx.RowSource = rsIxList
        dtcIx.ListField = "Ix"
        dtcIx.BoundColumn = "Ix_Id"
    End With
End Sub

Private Sub AddNewLabel()
    LabelCount = lstLabels.ListCount + 1
    txtLabelName.Text = "Label " & LabelCount
    txtLabelName.SetFocus
    SendKeys "{home}+{end}"
    Load lblLabels(LabelCount)
    lblLabels(LabelCount).Visible = True
End Sub

Private Sub AddNewField()

End Sub

Private Sub AddNewLine()

End Sub

Private Sub sstab1_DblClick()
    Select Case sstab1.Tab
        Case 0:
                frameLabels.Visible = True
                frameFields.Visible = False
                frameLines.Visible = False
        Case 1:
                frameLabels.Visible = False
                frameFields.Visible = True
                frameLines.Visible = False
        Case 2:
                frameLabels.Visible = False
                frameFields.Visible = False
                frameLines.Visible = True
    End Select

End Sub
