VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPatientFacilityBillPreferances 
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnSaveExit 
      Height          =   375
      Left            =   5880
      TabIndex        =   61
      Top             =   8760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save and Exit"
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
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Paper Area"
      TabPicture(0)   =   "frmPatientFacilityBillPreferances.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePaperSetting"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FramePaperArea"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Labels"
      TabPicture(1)   =   "frmPatientFacilityBillPreferances.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkLblPatientSex"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkLblPatientAge"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkLblPatientID"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkLblPatientName"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chkLblComments"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "checkHLine4"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "checkHLine3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "checkHLine2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "checkHLine1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "bttnHLineDown4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "bttnHLineUP4"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "bttnHLineDown3"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "bttnHLineUP3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "bttnHLineDown2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "bttnHLineUP2"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "bttnHLineDown1"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "bttnHLineUP1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkLblResults"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "bttnTextFont"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "bttnLabelFont"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "chkInstitutionAddress"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "CheckInstitutionName"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "bttnMessageFont"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "bttnAdvertiestmentFont"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "bttnFontInstitutionAddress"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "bttnFontInstitutionName"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "chkMessage"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtInstitutionName"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtInstitutionAddress"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "txtMessage1"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "chkPatientName"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "chkPatientID"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "chkPatientAge"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "chkPatientSex"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtLblPatientName"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "txtLblPatientID"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txtLblPatientAge"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "txtLblPatientSex"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "chkResults"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "txtLblComments"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "chkComments"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).ControlCount=   41
      Begin VB.CheckBox chkComments 
         Caption         =   "Comments"
         Height          =   255
         Left            =   -74880
         TabIndex        =   58
         Top             =   6240
         Width           =   3495
      End
      Begin VB.TextBox txtLblComments 
         Height          =   285
         Left            =   -73200
         TabIndex        =   57
         Top             =   5760
         Width           =   3975
      End
      Begin VB.CheckBox chkResults 
         Caption         =   "List of Results"
         Height          =   255
         Left            =   -72240
         TabIndex        =   44
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox txtLblPatientSex 
         Height          =   285
         Left            =   -72840
         TabIndex        =   36
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtLblPatientAge 
         Height          =   285
         Left            =   -72840
         TabIndex        =   35
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtLblPatientID 
         Height          =   285
         Left            =   -72840
         TabIndex        =   34
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtLblPatientName 
         Height          =   285
         Left            =   -72840
         TabIndex        =   33
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox chkPatientSex 
         Caption         =   "Sex"
         Height          =   255
         Left            =   -70320
         TabIndex        =   32
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CheckBox chkPatientAge 
         Caption         =   "Age"
         Height          =   255
         Left            =   -70320
         TabIndex        =   31
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox chkPatientID 
         Caption         =   "Patient ID"
         Height          =   255
         Left            =   -70320
         TabIndex        =   30
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox chkPatientName 
         Caption         =   "Patient Name "
         Height          =   255
         Left            =   -70320
         TabIndex        =   29
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtMessage1 
         Height          =   495
         Left            =   -73200
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   6600
         Width           =   3975
      End
      Begin VB.TextBox txtInstitutionAddress 
         Height          =   645
         Left            =   -72840
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox txtInstitutionName 
         Height          =   285
         Left            =   -72840
         TabIndex        =   20
         Top             =   480
         Width           =   3855
      End
      Begin VB.CheckBox chkMessage 
         Caption         =   "Message"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   6600
         Width           =   2655
      End
      Begin VB.Frame FramePaperArea 
         Caption         =   "Paper Area"
         Height          =   6855
         Left            =   960
         TabIndex        =   4
         Top             =   1200
         Width           =   5415
         Begin VB.Frame FramePrintableArea 
            Caption         =   "Printable Area"
            Height          =   6375
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   4695
            Begin VB.TextBox txtResultsDisplay 
               Height          =   975
               Left            =   120
               TabIndex        =   60
               Text            =   "values"
               Top             =   3000
               Width           =   5775
            End
            Begin VB.TextBox txtCommentsDisplay 
               Height          =   885
               Left            =   120
               TabIndex        =   18
               Text            =   "comments"
               Top             =   4440
               Width           =   5775
            End
            Begin VB.TextBox txtLblCommentsDisplay 
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   120
               TabIndex        =   17
               Text            =   "COMMENTS"
               Top             =   4080
               Width           =   1575
            End
            Begin VB.TextBox txtPatientIDDisplay 
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   2280
               TabIndex        =   16
               Text            =   "id"
               Top             =   2280
               Width           =   2775
            End
            Begin VB.TextBox txtPatientAgeDisplay 
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   2280
               TabIndex        =   15
               Text            =   "age"
               Top             =   1800
               Width           =   2775
            End
            Begin VB.TextBox txtPatientSexDisplay 
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   2280
               TabIndex        =   14
               Text            =   "sex"
               Top             =   2040
               Width           =   2775
            End
            Begin VB.TextBox txtLblPatientIDDisplay 
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Text            =   "ID"
               Top             =   2280
               Width           =   1455
            End
            Begin VB.TextBox txtAddressDisplay 
               Height          =   735
               Left            =   1320
               TabIndex        =   12
               Text            =   "Institution Address"
               Top             =   600
               Width           =   3255
            End
            Begin VB.TextBox txtInstitutionNameDisplay 
               Height          =   285
               Left            =   1320
               TabIndex        =   11
               Text            =   "Institution Name"
               Top             =   240
               Width           =   3255
            End
            Begin VB.TextBox txtMessageDisplay 
               Height          =   645
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   10
               Text            =   "frmPatientFacilityBillPreferances.frx":0038
               Top             =   5520
               Width           =   5775
            End
            Begin VB.TextBox txtLblPatientAgeDisplay 
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   120
               TabIndex        =   9
               Text            =   "AGE"
               Top             =   1800
               Width           =   1455
            End
            Begin VB.TextBox txtLblPatientSexDisplay 
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   120
               TabIndex        =   8
               Text            =   "SEX"
               Top             =   2040
               Width           =   1455
            End
            Begin VB.TextBox txtPatientNameDisplay 
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   2280
               TabIndex        =   7
               Text            =   "name"
               Top             =   1560
               Width           =   2775
            End
            Begin VB.TextBox txtLblPatientNameDisplay 
               BorderStyle     =   0  'None
               Height          =   195
               Left            =   120
               TabIndex        =   6
               Text            =   "NAME"
               Top             =   1560
               Width           =   1455
            End
            Begin VB.Line HLine2 
               X1              =   120
               X2              =   5880
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Line HLine1 
               X1              =   120
               X2              =   5880
               Y1              =   1440
               Y2              =   1440
            End
            Begin VB.Line HLine4 
               X1              =   120
               X2              =   5880
               Y1              =   5400
               Y2              =   5400
            End
            Begin VB.Line HLine3 
               X1              =   120
               X2              =   5880
               Y1              =   2640
               Y2              =   2640
            End
         End
      End
      Begin VB.Frame FramePaperSetting 
         Caption         =   "Paper Setting"
         Height          =   7995
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   6315
         Begin MSComctlLib.Slider SliderRight 
            Height          =   675
            Left            =   3240
            TabIndex        =   2
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1191
            _Version        =   393216
            Min             =   11
            Max             =   20
            SelStart        =   19
            Value           =   19
         End
         Begin MSComctlLib.Slider SliderBottom 
            Height          =   3495
            Left            =   120
            TabIndex        =   3
            Top             =   4200
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   6165
            _Version        =   393216
            Orientation     =   1
            Min             =   11
            Max             =   20
            SelStart        =   19
            Value           =   19
         End
         Begin MSComctlLib.Slider SliderLeft 
            Height          =   675
            Left            =   480
            TabIndex        =   62
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   1191
            _Version        =   393216
            Max             =   9
            SelStart        =   1
            Value           =   1
         End
         Begin MSComctlLib.Slider SliderTop 
            Height          =   3375
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   5953
            _Version        =   393216
            Orientation     =   1
            Max             =   9
            SelStart        =   1
            Value           =   1
         End
      End
      Begin btButtonEx.ButtonEx bttnFontInstitutionName 
         Height          =   255
         Left            =   -68880
         TabIndex        =   23
         Top             =   480
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "Font"
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
      Begin btButtonEx.ButtonEx bttnFontInstitutionAddress 
         Height          =   255
         Left            =   -68880
         TabIndex        =   24
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "Font"
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
      Begin btButtonEx.ButtonEx bttnAdvertiestmentFont 
         Height          =   255
         Left            =   -69000
         TabIndex        =   25
         Top             =   5760
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "Font"
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
      Begin btButtonEx.ButtonEx bttnMessageFont 
         Height          =   255
         Left            =   -67440
         TabIndex        =   26
         Top             =   9600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "Font"
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
      Begin VB.CheckBox CheckInstitutionName 
         Caption         =   "Institution Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chkInstitutionAddress 
         Caption         =   "Address"
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   840
         Width           =   2655
      End
      Begin btButtonEx.ButtonEx bttnLabelFont 
         Height          =   255
         Left            =   -74760
         TabIndex        =   37
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "Label Font"
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
      Begin btButtonEx.ButtonEx bttnTextFont 
         Height          =   255
         Left            =   -70320
         TabIndex        =   38
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "Detail Font"
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
      Begin VB.CheckBox chkLblResults 
         Caption         =   "Results Label"
         Height          =   255
         Left            =   -74760
         TabIndex        =   43
         Top             =   3600
         Width           =   2415
      End
      Begin btButtonEx.ButtonEx bttnHLineUP1 
         Height          =   375
         Left            =   -73080
         TabIndex        =   49
         Top             =   4080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "#"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnHLineDown1 
         Height          =   375
         Left            =   -73080
         TabIndex        =   50
         Top             =   4440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "$"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnHLineUP2 
         Height          =   375
         Left            =   -69600
         TabIndex        =   51
         Top             =   4080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "#"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnHLineDown2 
         Height          =   375
         Left            =   -69600
         TabIndex        =   52
         Top             =   4440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "$"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnHLineUP3 
         Height          =   375
         Left            =   -73080
         TabIndex        =   53
         Top             =   4920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "#"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnHLineDown3 
         Height          =   375
         Left            =   -73080
         TabIndex        =   54
         Top             =   5280
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "$"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnHLineUP4 
         Height          =   375
         Left            =   -69600
         TabIndex        =   55
         Top             =   4920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "#"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin btButtonEx.ButtonEx bttnHLineDown4 
         Height          =   375
         Left            =   -69600
         TabIndex        =   56
         Top             =   5280
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "$"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox checkHLine1 
         Caption         =   "1st Horizontal Line"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74760
         TabIndex        =   45
         Top             =   4200
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox checkHLine2 
         Caption         =   "2nd Horizontal Line"
         Height          =   495
         Left            =   -71280
         TabIndex        =   46
         Top             =   4080
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox checkHLine3 
         Caption         =   "3rd Horizontal Line"
         Height          =   495
         Left            =   -74760
         TabIndex        =   47
         Top             =   4920
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox checkHLine4 
         Caption         =   "4th Horizontal Line"
         Height          =   495
         Left            =   -71400
         TabIndex        =   48
         Top             =   5040
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkLblComments 
         Caption         =   "Comments Label"
         Height          =   255
         Left            =   -74880
         TabIndex        =   59
         Top             =   5760
         Width           =   3495
      End
      Begin VB.CheckBox chkLblPatientName 
         Caption         =   "Label Patient Name "
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CheckBox chkLblPatientID 
         Caption         =   "Label Patient ID"
         Height          =   255
         Left            =   -74760
         TabIndex        =   40
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox chkLblPatientAge 
         Caption         =   "Label Age"
         Height          =   255
         Left            =   -74760
         TabIndex        =   41
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox chkLblPatientSex 
         Caption         =   "Label Sex"
         Height          =   255
         Left            =   -74760
         TabIndex        =   42
         Top             =   2640
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmPatientFacilityBillPreferances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim InsideX As Long
Dim InsideY As Long
Dim temsource As TextBox
Dim VerticalExpansion As Double
Dim HorizontalExpansion As Double
Dim TemResponce As Byte




Private Sub bttnAdvertiestmentFont_Click()
CommonDialog1.FontName = PreferanceAdvertiestmentFontName
CommonDialog1.FontSize = PreferanceAdvertiestmentFontSize
CommonDialog1.FontBold = PreferanceAdvertiestmentFontBold
CommonDialog1.FontItalic = PreferanceAdvertiestmentFontItalic
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
PreferanceAdvertiestmentFontName = CommonDialog1.FontName
PreferanceAdvertiestmentFontSize = CommonDialog1.FontSize
PreferanceAdvertiestmentFontBold = CommonDialog1.FontBold
PreferanceAdvertiestmentFontItalic = CommonDialog1.FontItalic

End Sub



Private Sub bttnConfidentialFont_Click()
CommonDialog1.FontName = LblConfidentialFontName
CommonDialog1.FontSize = LblConfidentialFontSize
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
LblConfidentialFontName = CommonDialog1.FontName
LblConfidentialFontSize = CommonDialog1.FontSize
End Sub

Private Sub bttnDoctorMLTFont_Click()
CommonDialog1.FontName = PreferanceDoctorMLTFontName
CommonDialog1.FontSize = PreferanceDoctorMLTFontSize
CommonDialog1.FontBold = PreferanceDoctorMLTFontBold
CommonDialog1.FontItalic = PreferanceDoctorMLTFontItalic
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
PreferanceDoctorMLTFontName = CommonDialog1.FontName
PreferanceDoctorMLTFontSize = CommonDialog1.FontSize
PreferanceDoctorMLTFontBold = CommonDialog1.FontBold
PreferanceDoctorMLTFontItalic = CommonDialog1.FontItalic
End Sub


Private Sub bttnFontInstitutionAddress_Click()
CommonDialog1.FontName = PreferanceInstitutionAddressFontName
CommonDialog1.FontSize = PreferanceInstitutionAddressFontSize
CommonDialog1.FontBold = PreferanceInstitutionAddressFontBold
CommonDialog1.FontItalic = PreferanceInstitutionAddressFontItalic
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
PreferanceInstitutionAddressFontName = CommonDialog1.FontName
PreferanceInstitutionAddressFontSize = CommonDialog1.FontSize
PreferanceInstitutionAddressFontBold = CommonDialog1.FontBold
PreferanceInstitutionAddressFontItalic = CommonDialog1.FontItalic
End Sub

Private Sub bttnFontInstitutionContact_Click()
CommonDialog1.FontName = PreferanceInstitutionContactFontName
CommonDialog1.FontSize = PreferanceInstitutionContactFontSize
CommonDialog1.FontBold = PreferanceInstitutionContactFontBold
CommonDialog1.FontItalic = PreferanceInstitutionContactFontItalic
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
PreferanceInstitutionContactFontName = CommonDialog1.FontName
PreferanceInstitutionContactFontSize = CommonDialog1.FontSize
PreferanceInstitutionContactFontBold = CommonDialog1.FontBold
PreferanceInstitutionContactFontItalic = CommonDialog1.FontItalic
End Sub



Private Sub bttnLabelFont_Click()
CommonDialog1.FontName = PreferanceLabelFontName
CommonDialog1.FontSize = PreferanceLabelFontSize
CommonDialog1.FontBold = PreferanceLabelFontBold
CommonDialog1.FontItalic = PreferanceLabelFontItalic
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
PreferanceLabelFontName = CommonDialog1.FontName
PreferanceLabelFontSize = CommonDialog1.FontSize
PreferanceLabelFontBold = CommonDialog1.FontBold
PreferanceLabelFontItalic = CommonDialog1.FontItalic
End Sub


Private Sub bttnLabReportFont_Click()
CommonDialog1.FontName = LblReportFontName
CommonDialog1.FontSize = LblReportFontSize
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
LblReportFontName = CommonDialog1.FontName
LblReportFontSize = CommonDialog1.FontSize
End Sub

Private Sub bttnMessageFont_Click()
CommonDialog1.FontName = PreferanceMessageFontName
CommonDialog1.FontSize = PreferanceMessageFontSize
CommonDialog1.FontBold = PreferanceMessageFontBold
CommonDialog1.FontItalic = PreferanceMessageFontItalic
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
PreferanceMessageFontName = CommonDialog1.FontName
PreferanceMessageFontSize = CommonDialog1.FontSize
PreferanceMessageFontBold = CommonDialog1.FontBold
PreferanceMessageFontItalic = CommonDialog1.FontItalic

End Sub


Private Sub bttnFontInstitutionName_Click()
CommonDialog1.FontName = PreferanceInstitutionNameFontName
CommonDialog1.FontSize = PreferanceInstitutionNameFontSize
CommonDialog1.FontBold = PreferanceInstitutionNameFontBold
CommonDialog1.FontItalic = PreferanceInstitutionNameFontItalic
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
PreferanceInstitutionNameFontName = CommonDialog1.FontName
PreferanceInstitutionNameFontSize = CommonDialog1.FontSize
PreferanceInstitutionNameFontBold = CommonDialog1.FontBold
PreferanceInstitutionNameFontItalic = CommonDialog1.FontItalic
End Sub


Private Sub bttnTextFont_Click()
CommonDialog1.FontName = PreferanceTextFontName
CommonDialog1.FontSize = PreferanceTextFontSize
CommonDialog1.FontBold = PreferanceTextFontBold
CommonDialog1.FontItalic = PreferanceTextFontItalic
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
PreferanceTextFontName = CommonDialog1.FontName
PreferanceTextFontSize = CommonDialog1.FontSize
PreferanceTextFontBold = CommonDialog1.FontBold
PreferanceTextFontItalic = CommonDialog1.FontItalic
End Sub


Private Sub bttnTopicFont_Click()
CommonDialog1.FontName = TopicFontName
CommonDialog1.FontSize = TopicFontSize
CommonDialog1.FontBold = TopicFontBold
CommonDialog1.FontItalic = TopicFontItalic
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
TopicFontName = CommonDialog1.FontName
TopicFontSize = CommonDialog1.FontSize
TopicFontBold = CommonDialog1.FontBold
TopicFontItalic = CommonDialog1.FontItalic
End Sub


Private Sub bttnValueFont_Click()
CommonDialog1.FontName = ValueFontName
CommonDialog1.FontSize = ValueFontSize
CommonDialog1.FontBold = ValueFontBold
CommonDialog1.FontItalic = ValueFontItalic
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
ValueFontName = CommonDialog1.FontName
ValueFontSize = CommonDialog1.FontSize
ValueFontBold = CommonDialog1.FontBold
ValueFontItalic = CommonDialog1.FontItalic
End Sub

Private Sub bttnLogoRemove_Click()
ImageInstitutionLogo.Picture = LoadPicture()
PreferanceInstitutionLogoFileName = Empty
End Sub


Private Sub bttnLogoLoad_Click()
ImageInstitutionLogo.Stretch = True
CommonDialog1.Filter = "BMP|*.BMP|JPG|*.JPG;JPE;JPEG|GIF|*.GIF|All Images|*.BMP;*.JPG;*.JPE;*.JPGE;*.GIF|All Files|*.*"
CommonDialog1.ShowOpen
On Error GoTo PhotoError:
ImageInstitutionLogo.Picture = LoadPicture(CommonDialog1.FileName)
PreferanceInstitutionLogoFileName = CommonDialog1.FileName
Exit Sub
PhotoError:
If Err.Number = 481 Then
    TemResponce = MsgBox("The Photo you choose is not suitable, try using a medium size BMP, JPG or GIF file", vbOKOnly, "Photo Error")
ElseIf Err.Number = 53 Then
    TemResponce = MsgBox("No photo exist to selected, try to select again correctly.", vbOKOnly, "Photo Error")
Else
    TemResponce = MsgBox("An unknown error has occured, try again," & Chr(13) & Err.Description, vbOKOnly, "Photo Error")
End If

End Sub



Private Sub bttnSaveExit_Click()
Call SavePreferances
Call LoadPreferances
Unload Me
End Sub

Private Sub CheckInstitutionName_Click()
If CheckInstitutionName.Value = 1 Then
    txtInstitutionNameDisplay.Visible = True
Else
    txtInstitutionNameDisplay.Visible = False
End If
End Sub

Private Sub Form_Load()
Call LoadPreferances
Call SetPreferances

End Sub




Private Sub bttnHLineUP1_Click()
HLine1.Y1 = HLine1.Y1 - 100
HLine1.Y2 = HLine1.Y2 - 100
End Sub

Private Sub bttnHLineUP2_Click()
HLine2.Y1 = HLine2.Y1 - 100
HLine2.Y2 = HLine2.Y2 - 100
End Sub

Private Sub bttnHLineUP3_Click()
HLine3.Y1 = HLine3.Y1 - 100
HLine3.Y2 = HLine3.Y2 - 100
End Sub

Private Sub bttnHLineUP4_Click()
HLine4.Y1 = HLine4.Y1 - 100
HLine4.Y2 = HLine4.Y2 - 100
End Sub


Private Sub bttnHLineDown1_Click()
HLine1.Y1 = HLine1.Y1 + 100
HLine1.Y2 = HLine1.Y2 + 100
End Sub

Private Sub bttnHLineDown2_Click()
HLine2.Y1 = HLine2.Y1 + 100
HLine2.Y2 = HLine2.Y2 + 100
End Sub

Private Sub bttnHLineDown3_Click()
HLine3.Y1 = HLine3.Y1 + 100
HLine3.Y2 = HLine3.Y2 + 100
End Sub

Private Sub bttnHLineDown4_Click()
HLine4.Y1 = HLine4.Y1 + 100
HLine4.Y2 = HLine4.Y2 + 100
End Sub

Private Sub txtCommentsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtCommentsDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub


Private Sub txtLblCommentsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblCommentsDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub


Private Sub txtLblConfidentialDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblConfidentialDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblLabReportDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblLabReportDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblParametersDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblParametersDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtlblpatientnamedisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblPatientNameDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtlblpatientagedisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblPatientAgeDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtlblpatientsexdisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblPatientSexDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtlblpatientiddisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblPatientIDDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblDateDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblDateDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblReferancesDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblReferancesDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblSpecimanDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblSpecimanDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblSpecimanNoDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblSpecimanNoDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblTestDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblTestDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLbltimeDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblTimeDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblRDoctorDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblRDoctorDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblRInstitutionDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblRInstitutionDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblresultsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblResultsDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtLblUnitsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLblUnitsDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtParamatersDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtParamatersDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtpatientnamedisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtPatientNameDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtpatientagedisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtPatientAgeDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtpatientsexdisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtPatientSexDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtpatientiddisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtPatientIDDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtDateDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtDateDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtReferancesDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtReferancesDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtSpecimanDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtSpecimanDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtSpecimanNoDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtSpecimanNoDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtTestDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtTestDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txttimeDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtTimeDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtRDoctorDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtRDoctorDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub

Private Sub txtRInstitutionDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtRInstitutionDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
KeyCode = 0
End Sub


Private Sub txtAddressDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtAddressDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtLogo_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtLogo
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtAdvertiestmentDisplay1_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtAdvertiestmentDisplay1
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtAdvertiestmentDisplay2_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtAdvertiestmentDisplay2
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtContactDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtContactDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtDoctorMLTDetailsDisplay1_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtDoctorMLTDetailsDisplay1
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtDoctorMLTDetailsDisplay2_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtDoctorMLTDetailsDisplay2
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtDoctorMLTDetailsDisplay3_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtDoctorMLTDetailsDisplay3
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtDoctorMLTDetailsDisplay4_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtDoctorMLTDetailsDisplay4
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtInstitutionNameDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtInstitutionNameDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtMessage1_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtMessage1
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub



Private Sub SliderLeft_Change()
Call SetPrintableArea
Call SetLabels
End Sub
Private Sub Sliderright_Change()
Call SetPrintableArea
Call SetLabels
End Sub
Private Sub Slidertop_Change()
Call SetPrintableArea
Call SetLabels
End Sub
Private Sub Sliderbottom_Change()
Call SetPrintableArea
Call SetLabels
End Sub


Private Sub SetPrintableArea()
Dim PreviousWidth As Long
Dim PreviousHeight As Long
Dim PreviousLeft As Long
Dim PreviousTop As Long
Dim CurrentWidth As Long
Dim CurrentHeight As Long
Dim CurrentLeft As Long
Dim CurrentTop As Long

PreviousWidth = FramePrintableArea.Width
PreviousHeight = FramePrintableArea.Height

FramePrintableArea.Top = FramePaperArea.Height * SliderTop.Value / 20
FramePrintableArea.Left = FramePaperArea.Width * SliderLeft.Value / 20
FramePrintableArea.Height = (FramePaperArea.Height) * (SliderBottom.Value - SliderTop.Value) / 20
FramePrintableArea.Width = FramePaperArea.Width * (SliderRight.Value - SliderLeft.Value) / 20

CurrentWidth = FramePrintableArea.Width
CurrentHeight = FramePrintableArea.Height

VerticalExpansion = PreviousHeight / CurrentHeight
HorizontalExpansion = PreviousWidth / CurrentWidth

End Sub

Private Sub SetLabels()

Set temsource = txtInstitutionNameDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

Set temsource = txtAddressDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion
Set temsource = txtMessageDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

Set temsource = txtPatientNameDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

Set temsource = txtPatientAgeDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

Set temsource = txtPatientIDDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

Set temsource = txtPatientSexDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

Set temsource = txtLblPatientNameDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

Set temsource = txtLblPatientAgeDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

Set temsource = txtLblPatientIDDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

Set temsource = txtLblPatientSexDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

Set temsource = txtCommentsDisplay
temsource.Height = temsource.Height / VerticalExpansion
temsource.Top = temsource.Top / VerticalExpansion
temsource.Width = temsource.Width / HorizontalExpansion
temsource.Left = temsource.Left / HorizontalExpansion

HLine1.Y1 = HLine1.Y1 / VerticalExpansion
HLine1.Y2 = HLine1.Y2 / VerticalExpansion
HLine2.Y1 = HLine2.Y1 / VerticalExpansion
HLine2.Y2 = HLine2.Y2 / VerticalExpansion
HLine3.Y1 = HLine3.Y1 / VerticalExpansion
HLine3.Y2 = HLine3.Y2 / VerticalExpansion
HLine4.Y1 = HLine4.Y1 / VerticalExpansion
HLine4.Y2 = HLine4.Y2 / VerticalExpansion

HLine1.X1 = 0
HLine1.X2 = FramePrintableArea.Width
HLine2.X1 = 0
HLine2.X2 = FramePrintableArea.Width
HLine3.X1 = 0
HLine3.X2 = FramePrintableArea.Width
HLine4.X1 = 0
HLine4.X2 = FramePrintableArea.Width


End Sub

Private Sub LoadPreferances()

With DataEnvironment1.rscmmdPatientFacilityBillPreferances
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    .MoveFirst
    
        PreferanceLX = !LX
        PreferanceRX = !RX
        PreferanceTY = !TY
        PreferanceBY = !By
        
        If Not IsNull(!chkInstitutionName) Then PreferancechkInstitutionName = !chkInstitutionName
        If Not IsNull(!txtInstitutionName) Then PreferancetxtInstitutionName = DecreptedWord(!txtInstitutionName)
        If Not IsNull(!InstitutionNameFontName) Then PreferanceInstitutionNameFontName = !InstitutionNameFontName
        If Not IsNull(!InstitutionNameFontSize) Then PreferanceInstitutionNameFontSize = !InstitutionNameFontSize
        If Not IsNull(!InstitutionNameFontBold) Then PreferanceInstitutionNameFontBold = !InstitutionNameFontBold
        If Not IsNull(!InstitutionNameFontItalic) Then PreferanceInstitutionNameFontItalic = !InstitutionNameFontItalic
        If Not IsNull(!chkInstitutionAddress) Then PreferancechkInstitutionAddress = !chkInstitutionAddress
        If Not IsNull(!txtInstitutionAddress) Then PreferancetxtInstitutionAddress = !txtInstitutionAddress
        If Not IsNull(!InstitutionaddressFontName) Then PreferanceInstitutionAddressFontName = !InstitutionaddressFontName
        If Not IsNull(!InstitutionAddressFontSize) Then PreferanceInstitutionAddressFontSize = !InstitutionAddressFontSize
        If Not IsNull(!InstitutionAddressFontBole) Then PreferanceInstitutionAddressFontBold = !InstitutionAddressFontBole
        If Not IsNull(!InstitutionAddressFontItalic) Then PreferanceInstitutionAddressFontItalic = !InstitutionAddressFontItalic
        If Not IsNull(!Messagechk) Then PreferanceChkMessage = !Messagechk
        If Not IsNull(!Messagetxt) Then PreferancetxtMessage = !Messagetxt
        If Not IsNull(!MessageFontName) Then PreferanceMessageFontName = !MessageFontName
        If Not IsNull(!MessageFontsize) Then PreferanceMessageFontSize = !MessageFontsize
        If Not IsNull(!MessageFontbold) Then PreferanceMessageFontBold = !MessageFontbold
        If Not IsNull(!MessageFontitalic) Then PreferanceMessageFontItalic = !MessageFontitalic
        
        If Not IsNull(!InstitutionNameLX) Then PreferanceInstitutionNameLX = !InstitutionNameLX
        If Not IsNull(!InstitutionNameRX) Then PreferanceInstitutionNameRX = !InstitutionNameRX
        If Not IsNull(!InstitutionNameTY) Then PreferanceInstitutionNameTY = !InstitutionNameTY
        If Not IsNull(!InstitutionNameBY) Then PreferanceInstitutionNameBY = !InstitutionNameBY
        
        If Not IsNull(!InstitutionAddressLX) Then PreferanceInstitutionAddressLX = !InstitutionAddressLX
        If Not IsNull(!InstitutionAddressRX) Then PreferanceInstitutionAddressRX = !InstitutionAddressRX
        If Not IsNull(!InstitutionAddressTY) Then PreferanceInstitutionAddressTY = !InstitutionAddressTY
        If Not IsNull(!InstitutionAddressBY) Then PreferanceInstitutionAddressBY = !InstitutionAddressBY
        If Not IsNull(!InstitutionContactLX) Then PreferanceInstitutionContactLX = !InstitutionContactLX
        If Not IsNull(!InstitutionContactRX) Then PreferanceInstitutionContactRX = !InstitutionContactRX
        If Not IsNull(!InstitutionContactTY) Then PreferanceInstitutionContactTY = !InstitutionContactTY
        If Not IsNull(!InstitutionContactBY) Then PreferanceInstitutionContactBY = !InstitutionContactBY
        
        If Not IsNull(!messageLX) Then PreferanceMessageLX = !messageLX
        If Not IsNull(!messageRX) Then PreferanceMessageRX = !messageRX
        If Not IsNull(!messageTY) Then PreferanceMessageTY = !messageTY
        If Not IsNull(!messageBY) Then PreferanceMessageBY = !messageBY
        
        If Not IsNull(!sliderleftvalue) Then SliderLeft.Value = !sliderleftvalue
        If Not IsNull(!sliderrightvalue) Then SliderRight.Value = !sliderrightvalue
        If Not IsNull(!slidertopvalue) Then SliderTop.Value = !slidertopvalue
        If Not IsNull(!sliderbottomvalue) Then SliderBottom.Value = !sliderbottomvalue

        If Not IsNull(!PreferancechkPatientName) Then PreferancechkPatientName = !PreferancechkPatientName
        If Not IsNull(!PreferancechkPatientAge) Then PreferancechkPatientAge = !PreferancechkPatientAge
        If Not IsNull(!PreferancechkPatientID) Then PreferancechkPatientID = !PreferancechkPatientID
        If Not IsNull(!PreferancechkPatientSex) Then PreferancechkPatientSex = !PreferancechkPatientSex
        
        If Not IsNull(!PreferancechkLblPatientName) Then PreferancechkLblPatientName = !PreferancechkLblPatientName
        If Not IsNull(!PreferancechkLblPatientAge) Then PreferancechkLblPatientAge = !PreferancechkLblPatientAge
        If Not IsNull(!PreferancechkLblPatientID) Then PreferancechkLblPatientID = !PreferancechkLblPatientID
        If Not IsNull(!PreferancechkLblPatientSex) Then PreferancechkLblPatientSex = !PreferancechkLblPatientSex
        
        If Not IsNull(!PreferancetxtLblPatientName) Then PreferancetxtLblPatientName = !PreferancetxtLblPatientName
        If Not IsNull(!PreferancetxtLblPatientAge) Then PreferancetxtLblPatientAge = !PreferancetxtLblPatientAge
        If Not IsNull(!PreferancetxtLblPatientID) Then PreferancetxtLblPatientID = !PreferancetxtLblPatientID
        If Not IsNull(!PreferancetxtLblPatientSex) Then PreferancetxtLblPatientSex = !PreferancetxtLblPatientSex
        
        If Not IsNull(!PreferanceLblPatientNameLX) Then PreferanceLblPatientNameLX = !PreferanceLblPatientNameLX
        If Not IsNull(!PreferanceLblPatientAgeLX) Then PreferanceLblPatientAgeLX = !PreferanceLblPatientAgeLX
        If Not IsNull(!PreferanceLblPatientIDLX) Then PreferanceLblPatientIDLX = !PreferanceLblPatientIDLX
        If Not IsNull(!PreferanceLblPatientSexLX) Then PreferanceLblPatientSexLX = !PreferanceLblPatientSexLX
        
        If Not IsNull(!PreferancePatientNameLX) Then PreferancePatientNameLX = !PreferancePatientNameLX
        If Not IsNull(!PreferancePatientAgeLX) Then PreferancePatientAgeLX = !PreferancePatientAgeLX
        If Not IsNull(!PreferancePatientIDLX) Then PreferancePatientIDLX = !PreferancePatientIDLX
        If Not IsNull(!PreferancePatientSexLX) Then PreferancePatientSexLX = !PreferancePatientSexLX
        
        If Not IsNull(!PreferanceLblPatientNameTY) Then PreferanceLblPatientNameTY = !PreferanceLblPatientNameTY
        If Not IsNull(!PreferanceLblPatientAgeTY) Then PreferanceLblPatientAgeTY = !PreferanceLblPatientAgeTY
        If Not IsNull(!PreferanceLblPatientIDTY) Then PreferanceLblPatientIDTY = !PreferanceLblPatientIDTY
        If Not IsNull(!PreferanceLblPatientSexTY) Then PreferanceLblPatientSexTY = !PreferanceLblPatientSexTY
        
        If Not IsNull(!PreferancePatientNameTY) Then PreferancePatientNameTY = !PreferancePatientNameTY
        If Not IsNull(!PreferancePatientAgeTY) Then PreferancePatientAgeTY = !PreferancePatientAgeTY
        If Not IsNull(!PreferancePatientIDTY) Then PreferancePatientIDTY = !PreferancePatientIDTY
        If Not IsNull(!PreferancePatientSexTY) Then PreferancePatientSexTY = !PreferancePatientSexTY
        
        If Not IsNull(!PreferanceLabelFontName) Then PreferanceLabelFontName = !PreferanceLabelFontName
        If Not IsNull(!PreferanceLabelFontSize) Then PreferanceLabelFontSize = !PreferanceLabelFontSize
        If Not IsNull(!PreferanceLabelFontBold) Then PreferanceLabelFontBold = !PreferanceLabelFontBold
        If Not IsNull(!PreferanceLabelFontItalic) Then PreferanceLabelFontItalic = !PreferanceLabelFontItalic
        If Not IsNull(!PreferanceTextFontName) Then PreferanceTextFontName = !PreferanceTextFontName
        If Not IsNull(!PreferanceTextFontSize) Then PreferanceTextFontSize = !PreferanceTextFontSize
        If Not IsNull(!PreferanceTextFontBold) Then PreferanceTextFontBold = !PreferanceTextFontBold
        If Not IsNull(!PreferanceTextFontItalic) Then PreferanceTextFontItalic = !PreferanceTextFontItalic

        If Not IsNull(!PreferanceChkLblResults) Then PreferanceChkLblResults = !PreferanceChkLblResults
        If Not IsNull(!PreferanceChkLblComments) Then PreferanceChkLblComments = !PreferanceChkLblComments
        If Not IsNull(!PreferanceTxtLblResults) Then PreferanceTxtLblResults = !PreferanceTxtLblResults
        If Not IsNull(!PreferanceTxtLblComments) Then PreferanceTxtLblComments = !PreferanceTxtLblComments
        If Not IsNull(!PreferanceChkResults) Then PreferanceChkResults = !PreferanceChkResults
        If Not IsNull(!PreferanceChkComments) Then PreferanceChkComments = !PreferanceChkComments
        If Not IsNull(!LblResultsLX) Then LblResultsLX = !LblResultsLX
        If Not IsNull(!LblResultsTY) Then LblResultsTY = !LblResultsTY
        If Not IsNull(!LblCommentsLX) Then LblCommentsLX = !LblCommentsLX
        If Not IsNull(!LblCommentsTY) Then LblCommentsTY = !LblCommentsTY
        If Not IsNull(!ResultsLX) Then ResultsLX = !ResultsLX
        If Not IsNull(!ResultsRX) Then ResultsRX = !ResultsRX
        If Not IsNull(!ResultsTY) Then ResultsTY = !ResultsTY
        If Not IsNull(!ResultsBY) Then ResultsBY = !ResultsBY
        If Not IsNull(!CommentsLX) Then CommentsLX = !CommentsLX
        If Not IsNull(!CommentsRX) Then CommentsRX = !CommentsRX
        If Not IsNull(!CommentsTY) Then CommentsTY = !CommentsTY
        If Not IsNull(!CommentsBY) Then CommentsBY = !CommentsBY
        If Not IsNull(!TopicFontName) Then TopicFontName = !TopicFontName
        If Not IsNull(!TopicFontSize) Then TopicFontSize = !TopicFontSize
        If Not IsNull(!TopicFontBold) Then TopicFontBold = !TopicFontBold
        If Not IsNull(!TopicFontItalic) Then TopicFontItalic = !TopicFontItalic
        If Not IsNull(!ValueFontName) Then ValueFontName = !ValueFontName
        If Not IsNull(!ValueFontSize) Then ValueFontSize = !ValueFontSize
        If Not IsNull(!ValueFontBold) Then ValueFontBold = !ValueFontBold
        If Not IsNull(!ValueFontItalic) Then ValueFontItalic = !ValueFontItalic

        If Not IsNull(!InbetweenY) Then InbetweenY = !InbetweenY
        If Not IsNull(!HLineY1) Then HLineY1 = !HLineY1
        If Not IsNull(!HLineY2) Then HLineY2 = !HLineY2
        If Not IsNull(!HLineY3) Then HLineY3 = !HLineY3
        If Not IsNull(!HLineY4) Then HLineY4 = !HLineY4
        If Not IsNull(!chkHLine3) Then chkHLine3 = !chkHLine3
        If Not IsNull(!chkHLine2) Then chkHLine2 = !chkHLine2
        If Not IsNull(!chkHLine4) Then chkHLine4 = !chkHLine4
        .Close

End With
End Sub

Private Sub SetPreferances()

'Exit Sub

    FramePrintableArea.Left = PreferanceLX * FramePaperArea.Width
    FramePrintableArea.Top = PreferanceTY * FramePaperArea.Height
    FramePrintableArea.Width = (PreferanceRX - PreferanceLX) * FramePaperArea.Width
    FramePrintableArea.Height = (PreferanceBY - PreferanceTY) * FramePaperArea.Height
    
    Set temsource = txtInstitutionNameDisplay
    temsource.Left = FramePrintableArea.Width * PreferanceInstitutionNameLX
    temsource.Top = FramePrintableArea.Height * PreferanceInstitutionNameTY
    temsource.Width = FramePrintableArea.Width * (PreferanceInstitutionNameRX - PreferanceInstitutionNameLX)
    temsource.Height = FramePrintableArea.Height * (PreferanceInstitutionNameBY - PreferanceInstitutionNameTY)


    Set temsource = txtAddressDisplay
    temsource.Left = FramePrintableArea.Width * PreferanceInstitutionAddressLX
    temsource.Top = FramePrintableArea.Height * PreferanceInstitutionAddressTY
    temsource.Width = FramePrintableArea.Width * (PreferanceInstitutionAddressRX - PreferanceInstitutionAddressLX)
    temsource.Height = FramePrintableArea.Height * (PreferanceInstitutionAddressBY - PreferanceInstitutionAddressTY)

    Set temsource = txtLblPatientNameDisplay
    temsource.Left = FramePrintableArea.Width * PreferanceLblPatientNameLX
    temsource.Top = FramePrintableArea.Height * PreferanceLblPatientNameTY
    
    Set temsource = txtLblPatientAgeDisplay
    temsource.Left = FramePrintableArea.Width * PreferanceLblPatientAgeLX
    temsource.Top = FramePrintableArea.Height * PreferanceLblPatientAgeTY
    
    Set temsource = txtLblPatientSexDisplay
    temsource.Left = FramePrintableArea.Width * PreferanceLblPatientSexLX
    temsource.Top = FramePrintableArea.Height * PreferanceLblPatientSexTY
    
    Set temsource = txtLblPatientIDDisplay
    temsource.Left = FramePrintableArea.Width * PreferanceLblPatientIDLX
    temsource.Top = FramePrintableArea.Height * PreferanceLblPatientIDTY
    
    Set temsource = txtPatientNameDisplay
    temsource.Left = FramePrintableArea.Width * PreferancePatientNameLX
    temsource.Top = FramePrintableArea.Height * PreferancePatientNameTY
    
    Set temsource = txtPatientAgeDisplay
    temsource.Left = FramePrintableArea.Width * PreferancePatientAgeLX
    temsource.Top = FramePrintableArea.Height * PreferancePatientAgeTY
    
    Set temsource = txtPatientSexDisplay
    temsource.Left = FramePrintableArea.Width * PreferancePatientSexLX
    temsource.Top = FramePrintableArea.Height * PreferancePatientSexTY
    
    Set temsource = txtPatientIDDisplay
    temsource.Left = FramePrintableArea.Width * PreferancePatientIDLX
    temsource.Top = FramePrintableArea.Height * PreferancePatientIDTY
    
    
'    Set temsource = txtLblResultsDisplay
'    temsource.Left = FramePrintableArea.Width * LblResultsLX
'    temsource.Top = FramePrintableArea.Height * LblResultsTY
    
    Set temsource = txtLblCommentsDisplay
    temsource.Left = FramePrintableArea.Width * LblCommentsLX
    temsource.Top = FramePrintableArea.Height * LblCommentsTY

    Set temsource = txtResultsDisplay
    temsource.Left = FramePrintableArea.Width * ResultsLX
    temsource.Top = FramePrintableArea.Height * ResultsTY
    temsource.Width = FramePrintableArea.Width * (ResultsRX - ResultsLX)
    temsource.Height = FramePrintableArea.Height * (ResultsBY - ResultsTY)
    
    Set temsource = txtCommentsDisplay
    temsource.Left = FramePrintableArea.Width * CommentsLX
    temsource.Top = FramePrintableArea.Height * CommentsTY
    temsource.Width = FramePrintableArea.Width * (CommentsRX - CommentsLX)
    temsource.Height = FramePrintableArea.Height * (CommentsBY - CommentsTY)
    
     Set temsource = txtMessageDisplay
    temsource.Left = FramePrintableArea.Width * PreferanceMessageLX
    temsource.Top = FramePrintableArea.Height * PreferanceMessageTY
    temsource.Width = FramePrintableArea.Width * (PreferanceMessageRX - PreferanceMessageLX)
    temsource.Height = FramePrintableArea.Height * (PreferanceMessageBY - PreferanceMessageTY)
   
    
    HLine1.Y1 = FramePrintableArea.Height * HLineY1
    HLine1.Y2 = FramePrintableArea.Height * HLineY1
    HLine1.X1 = 0
    HLine1.X2 = FramePrintableArea.Width
    
    HLine2.Y1 = FramePrintableArea.Height * HLineY2
    HLine2.Y2 = FramePrintableArea.Height * HLineY2
    HLine2.X1 = 0
    HLine2.X2 = FramePrintableArea.Width
    
    HLine3.Y1 = FramePrintableArea.Height * HLineY3
    HLine3.Y2 = FramePrintableArea.Height * HLineY3
    HLine3.X1 = 0
    HLine3.X2 = FramePrintableArea.Width
    
    HLine4.Y1 = FramePrintableArea.Height * HLineY4
    HLine4.Y2 = FramePrintableArea.Height * HLineY4
    HLine4.X1 = 0
    HLine4.X2 = FramePrintableArea.Width
    
    
    If PreferancechkInstitutionName = True Then
        CheckInstitutionName.Value = 1
    Else
        CheckInstitutionName.Value = 0
    End If
    If PreferancechkInstitutionAddress = True Then
        chkInstitutionAddress.Value = 1
    Else
        chkInstitutionAddress.Value = 0
    End If
    If PreferancechkLblPatientName = True Then
        chkLblPatientName.Value = 1
    Else
        chkLblPatientName.Value = 0
    End If
    If PreferancechkLblPatientAge = True Then
        chkLblPatientAge.Value = 1
    Else
        chkLblPatientAge.Value = 0
    End If
    If PreferancechkLblPatientSex = True Then
        chkLblPatientSex.Value = 1
    Else
        chkLblPatientSex.Value = 0
    End If
    If PreferancechkLblPatientID = True Then
        chkLblPatientID.Value = 1
    Else
        chkLblPatientID.Value = 0
    End If
    
    If PreferancechkPatientName = True Then
        chkPatientName.Value = 1
    Else
        chkPatientName.Value = 0
    End If
    If PreferancechkPatientAge = True Then
        chkPatientAge.Value = 1
    Else
        chkPatientAge.Value = 0
    End If
    If PreferancechkPatientSex = True Then
        chkPatientSex.Value = 1
    Else
        chkPatientSex.Value = 0
    End If
    If PreferancechkPatientID = True Then
        chkPatientID.Value = 1
    Else
        chkPatientID.Value = 0
    End If
    If PreferanceChkResults = True Then
        chkResults.Value = 1
    Else
        chkResults.Value = 0
    End If
    If PreferanceChkLblComments = True Then
        chkLblComments.Value = 1
    Else
        chkLblComments.Value = 0
    End If
    
    If PreferanceChkComments = True Then
        chkComments.Value = 1
    Else
        chkComments.Value = 0
    End If
    
    
    If chkHLine2 = True Then
        checkHLine2.Value = 1
    Else
        checkHLine2.Value = 0
    End If
    
    If chkHLine3 = True Then
        checkHLine3.Value = 1
    Else
        checkHLine3.Value = 0
    End If
    
    If chkHLine4 = True Then
        checkHLine4.Value = 1
    Else
        checkHLine4.Value = 0
    End If
    
    txtInstitutionName.Text = PreferancetxtInstitutionName
    txtInstitutionAddress.Text = PreferancetxtInstitutionAddress
    txtMessage1.Text = PreferancetxtMessage
    txtLblPatientName.Text = PreferancetxtLblPatientName
    txtLblPatientAge.Text = PreferancetxtLblPatientAge
    txtLblPatientSex.Text = PreferancetxtLblPatientSex
    txtLblPatientID.Text = PreferancetxtLblPatientID
                
'    txtLblResults.Text = PreferanceTxtLblResults
    txtLblComments.Text = PreferanceTxtLblComments
    
    
End Sub

Private Sub SavePreferances()



With DataEnvironment1.rscmmdPatientFacilityBillPreferances
    If .State = 0 Then .Open
    If .RecordCount = 0 Then
        .AddNew
    Else
        .MoveFirst
    End If
    
    !LX = FramePrintableArea.Left / FramePaperArea.Width
    !TY = FramePrintableArea.Top / FramePaperArea.Height
    !RX = (FramePrintableArea.Left + FramePrintableArea.Width) / FramePaperArea.Width
    !By = (FramePrintableArea.Top + FramePrintableArea.Height) / FramePaperArea.Height
    
   
    Set temsource = txtInstitutionNameDisplay
    !InstitutionNameLX = temsource.Left / FramePrintableArea.Width
    !InstitutionNameTY = temsource.Top / FramePrintableArea.Height
    !InstitutionNameRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
    !InstitutionNameBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
    
    
    Set temsource = txtAddressDisplay
    !InstitutionAddressLX = temsource.Left / FramePrintableArea.Width
    !InstitutionAddressTY = temsource.Top / FramePrintableArea.Height
    !InstitutionAddressRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
    !InstitutionAddressBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
    Set temsource = txtMessageDisplay
    !messageLX = temsource.Left / FramePrintableArea.Width
    !messageTY = temsource.Top / FramePrintableArea.Height
    !messageRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
    !messageBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
    
'    Set temsource = txtLblResultsDisplay
'    !LblResultsLX = temsource.Left / FramePrintableArea.Width
'    !LblResultsTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblResultsDisplay
'    !LblResultsLX = temsource.Left / FramePrintableArea.Width
'    !LblResultsTY = temsource.Top / FramePrintableArea.Height
    
    Set temsource = txtLblCommentsDisplay
    !LblCommentsLX = temsource.Left / FramePrintableArea.Width
    !LblCommentsTY = temsource.Top / FramePrintableArea.Height
      
    Set temsource = txtResultsDisplay
    !ResultsLX = temsource.Left / FramePrintableArea.Width
    !ResultsTY = temsource.Top / FramePrintableArea.Height
    !ResultsRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
    !ResultsBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
    
    .Update
    
    Set temsource = txtCommentsDisplay
    !CommentsLX = temsource.Left / FramePrintableArea.Width
    !CommentsTY = temsource.Top / FramePrintableArea.Height
    !CommentsRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
    !CommentsBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
    
    
    Set temsource = txtPatientNameDisplay
    !PreferancePatientNameLX = txtPatientNameDisplay.Left / FramePrintableArea.Width
    !PreferancePatientNameTY = txtPatientNameDisplay.Top / FramePrintableArea.Height
    
    Set temsource = txtPatientAgeDisplay
    !PreferancePatientAgeLX = txtPatientAgeDisplay.Left / FramePrintableArea.Width
    !PreferancePatientAgeTY = txtPatientAgeDisplay.Top / FramePrintableArea.Height
    
    Set temsource = txtPatientSexDisplay
    !PreferancePatientSexLX = txtPatientSexDisplay.Left / FramePrintableArea.Width
    !PreferancePatientSexTY = txtPatientSexDisplay.Top / FramePrintableArea.Height
    
    Set temsource = txtPatientIDDisplay
    !PreferancePatientIDLX = txtPatientIDDisplay.Left / FramePrintableArea.Width
    !PreferancePatientIDTY = txtPatientIDDisplay.Top / FramePrintableArea.Height
    
   
    .Update
    
    Set temsource = txtLblPatientNameDisplay
    !PreferanceLblPatientNameLX = temsource.Left / FramePrintableArea.Width
    !PreferanceLblPatientNameTY = temsource.Top / FramePrintableArea.Height
    
    Set temsource = txtLblPatientAgeDisplay
    !PreferanceLblPatientAgeLX = temsource.Left / FramePrintableArea.Width
    !PreferanceLblPatientAgeTY = temsource.Top / FramePrintableArea.Height
    
    Set temsource = txtLblPatientSexDisplay
    !PreferanceLblPatientSexLX = temsource.Left / FramePrintableArea.Width
    !PreferanceLblPatientSexTY = temsource.Top / FramePrintableArea.Height
    
    Set temsource = txtLblPatientIDDisplay
    !PreferanceLblPatientIDLX = temsource.Left / FramePrintableArea.Width
    !PreferanceLblPatientIDTY = temsource.Top / FramePrintableArea.Height
    
    !HLineY1 = HLine1.Y1 / FramePrintableArea.Height
    !HLineY2 = HLine2.Y1 / FramePrintableArea.Height
    !HLineY3 = HLine3.Y1 / FramePrintableArea.Height
    !HLineY4 = HLine4.Y1 / FramePrintableArea.Height
       
    If checkHLine2.Value = 1 Then
        !chkHLine2 = True
    Else
        !chkHLine2 = False
    End If
       
    If checkHLine3.Value = 1 Then
        !chkHLine3 = True
    Else
        !chkHLine3 = False
    End If
    
    If checkHLine4.Value = 1 Then
        !chkHLine4 = True
    Else
        !chkHLine4 = False
    End If
    
    
    .Update
    
    !sliderleftvalue = SliderLeft.Value
    !sliderrightvalue = SliderRight.Value
    !slidertopvalue = SliderTop.Value
    !sliderbottomvalue = SliderBottom.Value
    
    
    ' ***************************
    .Update
    If chkMessage.Value = 1 Then
        !Messagechk = True
    Else
        !Messagechk = False
    End If
    
    
    !Messagetxt = txtMessage1.Text
    !MessageFontName = PreferanceMessageFontName
    !MessageFontsize = PreferanceMessageFontSize
    !MessageFontbold = PreferanceMessageFontBold
    !MessageFontitalic = PreferanceMessageFontItalic
    
    .Update


    If chkLblPatientName.Value = 1 Then
        !PreferancechkLblPatientName = True
    Else
        !PreferancechkLblPatientName = False
    End If
    
    If chkLblPatientAge.Value = 1 Then
        !PreferancechkLblPatientAge = True
    Else
        !PreferancechkLblPatientAge = False
    End If
    
    If chkLblPatientSex.Value = 1 Then
        !PreferancechkLblPatientSex = True
    Else
        !PreferancechkLblPatientSex = False
    End If
    
    If chkLblPatientID.Value = 1 Then
        !PreferancechkLblPatientID = True
    Else
        !PreferancechkLblPatientID = False
    End If
    
    If chkPatientName.Value = 1 Then
        !PreferancechkPatientName = True
    Else
        !PreferancechkPatientName = False
    End If
    
    If chkPatientAge.Value = 1 Then
        !PreferancechkPatientAge = True
    Else
        !PreferancechkPatientAge = False
    End If
    
    If chkPatientSex.Value = 1 Then
        !PreferancechkPatientSex = True
    Else
        !PreferancechkPatientSex = False
    End If
    
    If chkPatientID.Value = 1 Then
        !PreferancechkPatientID = True
    Else
        !PreferancechkPatientID = False
    End If
    
    !PreferancetxtLblPatientName = txtLblPatientName.Text
    !PreferancetxtLblPatientAge = txtLblPatientAge.Text
    !PreferancetxtLblPatientID = txtLblPatientID.Text
    !PreferancetxtLblPatientSex = txtLblPatientSex.Text
    
    !PreferanceTextFontName = PreferanceTextFontName
    !PreferanceTextFontSize = PreferanceTextFontSize
    !PreferanceTextFontBold = PreferanceTextFontBold
    !PreferanceTextFontItalic = PreferanceTextFontItalic
    !PreferanceLabelFontName = PreferanceLabelFontName
    !PreferanceLabelFontSize = PreferanceLabelFontSize
    !PreferanceLabelFontBold = PreferanceLabelFontBold
    !PreferanceLabelFontItalic = PreferanceLabelFontItalic
    


    .Update


    If chkLblResults.Value = 1 Then
        !PreferanceChkLblResults = True
    Else
        !PreferanceChkLblResults = False
    End If

    !txtInstitutionName = EncreptedWord(txtInstitutionName.Text)
    !InstitutionNameFontName = PreferanceInstitutionNameFontName
    !InstitutionNameFontSize = PreferanceInstitutionNameFontSize
    !InstitutionNameFontBold = PreferanceInstitutionNameFontBold
    !InstitutionNameFontItalic = PreferanceInstitutionNameFontItalic

    If chkResults.Value = 1 Then
        !PreferanceChkResults = True
    Else
        !PreferanceChkResults = False
    End If


    If chkLblComments.Value = 1 Then
        !PreferanceChkLblComments = True
    Else
        !PreferanceChkLblComments = False
    End If

    If chkComments.Value = 1 Then
        !PreferanceChkComments = True
    Else
        !PreferanceChkComments = False
    End If
    

    .Update
    
'    !PreferanceTxtLblResults = txtLblResults.Text
    !PreferanceTxtLblComments = txtLblComments.Text

    !TopicFontName = TopicFontName
    !TopicFontSize = TopicFontSize
    !TopicFontBold = TopicFontBold
    !TopicFontItalic = TopicFontItalic
    !ValueFontName = ValueFontName
    !ValueFontSize = ValueFontSize
    !ValueFontBold = ValueFontBold
    !ValueFontItalic = ValueFontItalic

    .Update
    .Close
End With

End Sub


Private Sub txtMessageDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtMessageDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0

End Sub


Private Sub txtResultsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtResultsDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

Private Sub txtUnitsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
Set temsource = txtUnitsDisplay
If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
KeyCode = 0
End Sub

