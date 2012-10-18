VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrintingPatientBill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Preferances"
   ClientHeight    =   10605
   ClientLeft      =   255
   ClientTop       =   225
   ClientWidth     =   15180
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   15180
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPhotoName 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnSaveExit 
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   9120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      Appearance      =   3
      Caption         =   "Save And Exit"
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
   Begin VB.CheckBox Check18 
      Caption         =   "Sex"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   11160
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Institution Details"
      TabPicture(0)   =   "frmPrintingPatientBill.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameInstitution"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Investigation"
      TabPicture(1)   =   "frmPrintingPatientBill.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameInvestigation"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Results"
      TabPicture(2)   =   "frmPrintingPatientBill.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frameResults"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Other Details"
      TabPicture(3)   =   "frmPrintingPatientBill.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameOther"
      Tab(3).ControlCount=   1
      Begin VB.Frame frameResults 
         Caption         =   "Results"
         Height          =   7455
         Left            =   -74880
         TabIndex        =   80
         Top             =   360
         Width           =   6735
         Begin VB.CheckBox chkParameters 
            Caption         =   "List of Parameters"
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   4680
            Width           =   3495
         End
         Begin VB.CheckBox chkResults 
            Caption         =   "List of Results"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   4920
            Width           =   3495
         End
         Begin VB.CheckBox chkUnits 
            Caption         =   "List of Units"
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   5160
            Width           =   3495
         End
         Begin VB.CheckBox chkTest 
            Caption         =   "Test"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   1560
            Width           =   3495
         End
         Begin VB.CheckBox chkSpeciman 
            Caption         =   "Speciman"
            Height          =   255
            Left            =   240
            TabIndex        =   94
            Top             =   1800
            Width           =   3495
         End
         Begin VB.TextBox txtLblTest 
            Height          =   285
            Left            =   2280
            TabIndex        =   93
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox txtLblSpeciman 
            Height          =   285
            Left            =   2280
            TabIndex        =   92
            Top             =   720
            Width           =   4335
         End
         Begin VB.TextBox txtLblParameters 
            Height          =   285
            Left            =   2280
            TabIndex        =   91
            Top             =   2760
            Width           =   4335
         End
         Begin VB.TextBox txtLblResults 
            Height          =   285
            Left            =   2280
            TabIndex        =   90
            Top             =   3120
            Width           =   4335
         End
         Begin VB.TextBox txtLblReferances 
            Height          =   285
            Left            =   2280
            TabIndex        =   89
            Top             =   3840
            Width           =   4335
         End
         Begin VB.CheckBox chkReferances 
            Caption         =   "List of Referance Ranges"
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   5400
            Width           =   3495
         End
         Begin VB.TextBox txtLblComments 
            Height          =   285
            Left            =   2280
            TabIndex        =   86
            Top             =   4200
            Width           =   4335
         End
         Begin VB.CheckBox Check3 
            Caption         =   "List of Referance Ranges"
            Height          =   255
            Left            =   480
            TabIndex        =   85
            Top             =   6600
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.CheckBox chkComments 
            Caption         =   "Comments"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   5640
            Width           =   3495
         End
         Begin VB.TextBox txtLblSpecimanNo 
            Height          =   285
            Left            =   2280
            TabIndex        =   83
            Top             =   1080
            Width           =   4335
         End
         Begin VB.CheckBox chkSpecimanNo 
            Caption         =   "Speciman No."
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   2040
            Width           =   3495
         End
         Begin VB.TextBox txtLblUnits 
            Height          =   285
            Left            =   2280
            TabIndex        =   81
            Top             =   3480
            Width           =   4335
         End
         Begin btButtonEx.ButtonEx bttnTopicFont 
            Height          =   255
            Left            =   4680
            TabIndex        =   106
            Top             =   1560
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            Appearance      =   3
            Caption         =   "Topic Font"
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
         Begin btButtonEx.ButtonEx bttnValueFont 
            Height          =   255
            Left            =   4680
            TabIndex        =   107
            Top             =   4800
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            Appearance      =   3
            Caption         =   "Value Font"
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
         Begin VB.CheckBox chkLblTest 
            Caption         =   "Test Label"
            Height          =   255
            Left            =   240
            TabIndex        =   96
            Top             =   360
            Width           =   3495
         End
         Begin VB.CheckBox chkLblSpeciman 
            Caption         =   "Speciman Label"
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   720
            Width           =   3495
         End
         Begin VB.CheckBox chkLblSpecimanNo 
            Caption         =   "Speciman No. Label"
            Height          =   255
            Left            =   240
            TabIndex        =   104
            Top             =   1080
            Width           =   3495
         End
         Begin VB.CheckBox chkLblParameters 
            Caption         =   "Parameter Label"
            Height          =   255
            Left            =   240
            TabIndex        =   102
            Top             =   2760
            Width           =   3495
         End
         Begin VB.CheckBox chkLblResults 
            Caption         =   "Results Label"
            Height          =   255
            Left            =   240
            TabIndex        =   103
            Top             =   3120
            Width           =   3495
         End
         Begin VB.CheckBox chkLblUnits 
            Caption         =   "Units Label"
            Height          =   255
            Left            =   240
            TabIndex        =   105
            Top             =   3480
            Width           =   3495
         End
         Begin VB.CheckBox chkLblReferances 
            Caption         =   "Ref. Range Label"
            Height          =   255
            Left            =   240
            TabIndex        =   101
            Top             =   3840
            Width           =   3495
         End
         Begin VB.CheckBox chkLblComments 
            Caption         =   "Comments Label"
            Height          =   255
            Left            =   240
            TabIndex        =   87
            Top             =   4200
            Width           =   3495
         End
      End
      Begin VB.Frame frameInvestigation 
         Caption         =   "Investigation"
         Height          =   7455
         Left            =   -74880
         TabIndex        =   53
         Top             =   480
         Width           =   6735
         Begin VB.CheckBox chkPatientName 
            Caption         =   "Patient Name "
            Height          =   255
            Left            =   4800
            TabIndex        =   69
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox chkPatientID 
            Caption         =   "Patient ID"
            Height          =   255
            Left            =   4800
            TabIndex        =   68
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox chkPatientAge 
            Caption         =   "Age"
            Height          =   255
            Left            =   4800
            TabIndex        =   67
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CheckBox chkPatientSex 
            Caption         =   "Sex"
            Height          =   255
            Left            =   4800
            TabIndex        =   66
            Top             =   2160
            Width           =   2055
         End
         Begin VB.CheckBox chkTime 
            Caption         =   "Time"
            Height          =   255
            Left            =   4800
            TabIndex        =   65
            Top             =   3240
            Width           =   2055
         End
         Begin VB.CheckBox chkRDoctorName 
            Caption         =   "Doctor's Name"
            Height          =   255
            Left            =   4800
            TabIndex        =   64
            Top             =   3840
            Width           =   2055
         End
         Begin VB.CheckBox chkRInstitutionName 
            Caption         =   "Institution Name"
            Height          =   255
            Left            =   4800
            TabIndex        =   63
            Top             =   4440
            Width           =   2055
         End
         Begin VB.CheckBox chkDate 
            Caption         =   "Date"
            Height          =   255
            Left            =   4800
            TabIndex        =   62
            Top             =   2640
            Width           =   2055
         End
         Begin VB.TextBox txtLblPatientName 
            Height          =   285
            Left            =   1920
            TabIndex        =   61
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtLblPatientID 
            Height          =   285
            Left            =   1920
            TabIndex        =   60
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtLblPatientAge 
            Height          =   285
            Left            =   1920
            TabIndex        =   59
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox txtLblPatientSex 
            Height          =   285
            Left            =   1920
            TabIndex        =   58
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtLblDate 
            Height          =   285
            Left            =   1920
            TabIndex        =   57
            Top             =   2640
            Width           =   2415
         End
         Begin VB.TextBox txtLblTime 
            Height          =   285
            Left            =   1920
            TabIndex        =   56
            Top             =   3240
            Width           =   2415
         End
         Begin VB.TextBox txtLblRDoctorName 
            Height          =   285
            Left            =   1920
            TabIndex        =   55
            Top             =   3840
            Width           =   2415
         End
         Begin VB.TextBox txtLblRInstitutionName 
            Height          =   285
            Left            =   1920
            TabIndex        =   54
            Top             =   4440
            Width           =   2415
         End
         Begin btButtonEx.ButtonEx bttnLabelFont 
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   4800
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
            Left            =   4800
            TabIndex        =   79
            Top             =   4800
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
         Begin VB.CheckBox chkLblPatientName 
            Caption         =   "Label Patient Name "
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox chkLblPatientID 
            Caption         =   "Label Patient ID"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox chkLblPatientAge 
            Caption         =   "Label Age"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CheckBox chkLblPatientSex 
            Caption         =   "Label Sex"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   2160
            Width           =   2055
         End
         Begin VB.CheckBox chkLblDate 
            Caption         =   "Label Date"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CheckBox chkLblTime 
            Caption         =   "Label Time"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   3240
            Width           =   2055
         End
         Begin VB.CheckBox chkLblRDoctorName 
            Caption         =   "Label Doctor's Name"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   3840
            Width           =   2055
         End
         Begin VB.CheckBox chkLblRInstitutionName 
            Caption         =   "Label Institution Name"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   4440
            Width           =   2055
         End
      End
      Begin VB.Frame FrameOther 
         Caption         =   "Other"
         Height          =   7455
         Left            =   -74880
         TabIndex        =   34
         Top             =   360
         Width           =   6735
         Begin VB.CheckBox checkHLine1 
            Caption         =   "1st Horizontal Line"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   1800
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.CheckBox ChkConfidential 
            Caption         =   "Confidential"
            Enabled         =   0   'False
            Height          =   495
            Left            =   360
            TabIndex        =   41
            Top             =   360
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin VB.TextBox txtConfidential 
            Height          =   285
            Left            =   2400
            TabIndex        =   40
            Top             =   480
            Width           =   3375
         End
         Begin VB.CheckBox chkReport 
            Caption         =   "Laborotary Report"
            Enabled         =   0   'False
            Height          =   495
            Left            =   360
            TabIndex        =   39
            Top             =   840
            Value           =   1  'Checked
            Width           =   3495
         End
         Begin VB.TextBox txtLabReport 
            Height          =   285
            Left            =   2400
            TabIndex        =   38
            Top             =   960
            Width           =   3375
         End
         Begin VB.CheckBox checkHLine2 
            Caption         =   "2nd Horizontal Line"
            Height          =   495
            Left            =   240
            TabIndex        =   37
            Top             =   2880
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox checkHLine3 
            Caption         =   "3rd Horizontal Line"
            Height          =   495
            Left            =   240
            TabIndex        =   36
            Top             =   3960
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox checkHLine4 
            Caption         =   "4th Horizontal Line"
            Height          =   495
            Left            =   240
            TabIndex        =   35
            Top             =   5280
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin btButtonEx.ButtonEx bttnHLineUP1 
            Height          =   375
            Left            =   3120
            TabIndex        =   42
            Top             =   1560
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
            Left            =   3120
            TabIndex        =   44
            Top             =   2040
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
            Left            =   3120
            TabIndex        =   45
            Top             =   2640
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
            Left            =   3120
            TabIndex        =   46
            Top             =   3120
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
            Left            =   3120
            TabIndex        =   47
            Top             =   3720
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
            Left            =   3120
            TabIndex        =   48
            Top             =   4200
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
            Left            =   3120
            TabIndex        =   49
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
            Left            =   3120
            TabIndex        =   50
            Top             =   5400
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
         Begin btButtonEx.ButtonEx bttnConfidentialFont 
            Height          =   255
            Left            =   5880
            TabIndex        =   51
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
         Begin btButtonEx.ButtonEx bttnLabReportFont 
            Height          =   255
            Left            =   5880
            TabIndex        =   52
            Top             =   960
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
      End
      Begin VB.Frame frameInstitution 
         Caption         =   "Institution Details"
         Height          =   7455
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   6735
         Begin VB.CheckBox chkLogo 
            Caption         =   "Logo"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1920
            Width           =   1935
         End
         Begin VB.CheckBox chkAdvertiestment1 
            Caption         =   "Advertiestment 1"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   3000
            Width           =   2415
         End
         Begin VB.CheckBox chkDoctorMLT1 
            Caption         =   "Doctor / Technician 1"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   4200
            Width           =   2415
         End
         Begin VB.CheckBox chkDoctorMLT2 
            Caption         =   "Doctor / Technician 2"
            Height          =   255
            Left            =   3480
            TabIndex        =   19
            Top             =   4200
            Width           =   2295
         End
         Begin VB.CheckBox chkDoctorMLT3 
            Caption         =   "Doctor / Technician 3"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   5400
            Width           =   2415
         End
         Begin VB.CheckBox chkDoctorMLT4 
            Caption         =   "Doctor / Technician 4"
            Height          =   255
            Left            =   3480
            TabIndex        =   17
            Top             =   5400
            Width           =   2295
         End
         Begin VB.CheckBox chkMessage 
            Caption         =   "Message"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   6600
            Width           =   2655
         End
         Begin VB.TextBox txtInstitutionName 
            Height          =   285
            Left            =   1560
            TabIndex        =   15
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox txtInstitutionAddress 
            Height          =   645
            Left            =   1560
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   600
            Width           =   4335
         End
         Begin VB.TextBox txtInstitutionContact 
            Height          =   495
            Left            =   1560
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   1320
            Width           =   4335
         End
         Begin VB.TextBox txtAdvertiestment1 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   3240
            Width           =   2415
         End
         Begin VB.CheckBox chkAdvertiestment2 
            Caption         =   "Advertiestment 2"
            Height          =   255
            Left            =   3480
            TabIndex        =   11
            Top             =   3000
            Width           =   2415
         End
         Begin VB.TextBox txtAdvertiestment2 
            Height          =   735
            Left            =   3480
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   3240
            Width           =   2415
         End
         Begin VB.TextBox txtDoctorMLT1 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   4440
            Width           =   2415
         End
         Begin VB.TextBox txtDoctorMLT2 
            Height          =   735
            Left            =   3480
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   4440
            Width           =   2415
         End
         Begin VB.TextBox txtDoctorMLT3 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   5640
            Width           =   2415
         End
         Begin VB.TextBox txtDoctorMLT4 
            Height          =   735
            Left            =   3480
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   5640
            Width           =   2415
         End
         Begin VB.TextBox txtMessage1 
            Height          =   495
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   6840
            Width           =   5775
         End
         Begin btButtonEx.ButtonEx bttnLogoLoad 
            Height          =   375
            Left            =   4560
            TabIndex        =   25
            Top             =   1920
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "Load"
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
         Begin btButtonEx.ButtonEx bttnLogoRemove 
            Height          =   375
            Left            =   4560
            TabIndex        =   26
            Top             =   2520
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "Remove"
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
         Begin btButtonEx.ButtonEx bttnFontInstitutionName 
            Height          =   255
            Left            =   6000
            TabIndex        =   28
            Top             =   240
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
            Left            =   6000
            TabIndex        =   29
            Top             =   600
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
         Begin btButtonEx.ButtonEx bttnFontInstitutionContact 
            Height          =   255
            Left            =   6000
            TabIndex        =   30
            Top             =   1320
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
            Left            =   2640
            TabIndex        =   31
            Top             =   3240
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
         Begin btButtonEx.ButtonEx bttnDoctorMLTFont 
            Height          =   255
            Left            =   2640
            TabIndex        =   32
            Top             =   5280
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
            Left            =   6000
            TabIndex        =   33
            Top             =   6840
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
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2655
         End
         Begin VB.CheckBox chkInstitutionAddress 
            Caption         =   "Address"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   2655
         End
         Begin VB.CheckBox chkInstitutionContact 
            Caption         =   "Contact"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Image ImageInstitutionLogo 
            BorderStyle     =   1  'Fixed Single
            Height          =   975
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmPrintingPatientBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim InsideX As Long
'Dim InsideY As Long
'Dim temsource As TextBox
'Dim VerticalExpansion As Double
'Dim HorizontalExpansion As Double
'Dim TemResponce As Byte
'
'
'
'
''Private Sub bttnAdvertiestmentFont_Click()
''CommonDialog1.FontName = PreferanceAdvertiestmentFontName
''CommonDialog1.FontSize = PreferanceAdvertiestmentFontSize
''CommonDialog1.FontBold = PreferanceAdvertiestmentFontBold
''CommonDialog1.FontItalic = PreferanceAdvertiestmentFontItalic
''CommonDialog1.Flags = cdlCFBoth
''CommonDialog1.ShowFont
''PreferanceAdvertiestmentFontName = CommonDialog1.FontName
''PreferanceAdvertiestmentFontSize = CommonDialog1.FontSize
''PreferanceAdvertiestmentFontBold = CommonDialog1.FontBold
''PreferanceAdvertiestmentFontItalic = CommonDialog1.FontItalic
''
''End Sub
'
'
'
'Private Sub bttnConfidentialFont_Click()
'CommonDialog1.FontName = LblConfidentialFontName
'CommonDialog1.FontSize = LblConfidentialFontSize
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'LblConfidentialFontName = CommonDialog1.FontName
'LblConfidentialFontSize = CommonDialog1.FontSize
'End Sub
'
'Private Sub bttnDoctorMLTFont_Click()
'CommonDialog1.FontName = PreferanceDoctorMLTFontName
'CommonDialog1.FontSize = PreferanceDoctorMLTFontSize
'CommonDialog1.FontBold = PreferanceDoctorMLTFontBold
'CommonDialog1.FontItalic = PreferanceDoctorMLTFontItalic
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'PreferanceDoctorMLTFontName = CommonDialog1.FontName
'PreferanceDoctorMLTFontSize = CommonDialog1.FontSize
'PreferanceDoctorMLTFontBold = CommonDialog1.FontBold
'PreferanceDoctorMLTFontItalic = CommonDialog1.FontItalic
'End Sub
'
'
'Private Sub bttnFontInstitutionAddress_Click()
'CommonDialog1.FontName = PreferanceInstitutionAddressFontName
'CommonDialog1.FontSize = PreferanceInstitutionAddressFontSize
'CommonDialog1.FontBold = PreferanceInstitutionAddressFontBold
'CommonDialog1.FontItalic = PreferanceInstitutionAddressFontItalic
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'PreferanceInstitutionAddressFontName = CommonDialog1.FontName
'PreferanceInstitutionAddressFontSize = CommonDialog1.FontSize
'PreferanceInstitutionAddressFontBold = CommonDialog1.FontBold
'PreferanceInstitutionAddressFontItalic = CommonDialog1.FontItalic
'End Sub
'
'Private Sub bttnFontInstitutionContact_Click()
'CommonDialog1.FontName = PreferanceInstitutionContactFontName
'CommonDialog1.FontSize = PreferanceInstitutionContactFontSize
'CommonDialog1.FontBold = PreferanceInstitutionContactFontBold
'CommonDialog1.FontItalic = PreferanceInstitutionContactFontItalic
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'PreferanceInstitutionContactFontName = CommonDialog1.FontName
'PreferanceInstitutionContactFontSize = CommonDialog1.FontSize
'PreferanceInstitutionContactFontBold = CommonDialog1.FontBold
'PreferanceInstitutionContactFontItalic = CommonDialog1.FontItalic
'End Sub
'
'
'
'Private Sub bttnLabelFont_Click()
'CommonDialog1.FontName = PreferanceLabelFontName
'CommonDialog1.FontSize = PreferanceLabelFontSize
'CommonDialog1.FontBold = PreferanceLabelFontBold
'CommonDialog1.FontItalic = PreferanceLabelFontItalic
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'PreferanceLabelFontName = CommonDialog1.FontName
'PreferanceLabelFontSize = CommonDialog1.FontSize
'PreferanceLabelFontBold = CommonDialog1.FontBold
'PreferanceLabelFontItalic = CommonDialog1.FontItalic
'End Sub
'
'
'Private Sub bttnLabReportFont_Click()
'CommonDialog1.FontName = LblReportFontName
'CommonDialog1.FontSize = LblReportFontSize
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'LblReportFontName = CommonDialog1.FontName
'LblReportFontSize = CommonDialog1.FontSize
'End Sub
'
'Private Sub bttnMessageFont_Click()
'CommonDialog1.FontName = PreferanceMessageFontName
'CommonDialog1.FontSize = PreferanceMessageFontSize
'CommonDialog1.FontBold = PreferanceMessageFontBold
'CommonDialog1.FontItalic = PreferanceMessageFontItalic
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'PreferanceMessageFontName = CommonDialog1.FontName
'PreferanceMessageFontSize = CommonDialog1.FontSize
'PreferanceMessageFontBold = CommonDialog1.FontBold
'PreferanceMessageFontItalic = CommonDialog1.FontItalic
'
'End Sub
'
'
'Private Sub bttnFontInstitutionName_Click()
'CommonDialog1.FontName = PreferanceInstitutionNameFontName
'CommonDialog1.FontSize = PreferanceInstitutionNameFontSize
'CommonDialog1.FontBold = PreferanceInstitutionNameFontBold
'CommonDialog1.FontItalic = PreferanceInstitutionNameFontItalic
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'PreferanceInstitutionNameFontName = CommonDialog1.FontName
'PreferanceInstitutionNameFontSize = CommonDialog1.FontSize
'PreferanceInstitutionNameFontBold = CommonDialog1.FontBold
'PreferanceInstitutionNameFontItalic = CommonDialog1.FontItalic
'End Sub
'
'
'Private Sub bttnTextFont_Click()
'CommonDialog1.FontName = PreferanceTextFontName
'CommonDialog1.FontSize = PreferanceTextFontSize
'CommonDialog1.FontBold = PreferanceTextFontBold
'CommonDialog1.FontItalic = PreferanceTextFontItalic
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'PreferanceTextFontName = CommonDialog1.FontName
'PreferanceTextFontSize = CommonDialog1.FontSize
'PreferanceTextFontBold = CommonDialog1.FontBold
'PreferanceTextFontItalic = CommonDialog1.FontItalic
'End Sub
'
'
'Private Sub bttnTopicFont_Click()
'CommonDialog1.FontName = TopicFontName
'CommonDialog1.FontSize = TopicFontSize
'CommonDialog1.FontBold = TopicFontBold
'CommonDialog1.FontItalic = TopicFontItalic
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'TopicFontName = CommonDialog1.FontName
'TopicFontSize = CommonDialog1.FontSize
'TopicFontBold = CommonDialog1.FontBold
'TopicFontItalic = CommonDialog1.FontItalic
'End Sub
'
'
'Private Sub bttnValueFont_Click()
'CommonDialog1.FontName = ValueFontName
'CommonDialog1.FontSize = ValueFontSize
'CommonDialog1.FontBold = ValueFontBold
'CommonDialog1.FontItalic = ValueFontItalic
'CommonDialog1.Flags = cdlCFBoth
'CommonDialog1.ShowFont
'ValueFontName = CommonDialog1.FontName
'ValueFontSize = CommonDialog1.FontSize
'ValueFontBold = CommonDialog1.FontBold
'ValueFontItalic = CommonDialog1.FontItalic
'End Sub
'
'Private Sub bttnLogoRemove_Click()
'ImageInstitutionLogo.Picture = LoadPicture()
'PreferanceInstitutionLogoFileName = Empty
'End Sub
'
'
'Private Sub bttnLogoLoad_Click()
'ImageInstitutionLogo.Stretch = True
'CommonDialog1.Filter = "BMP|*.BMP|JPG|*.JPG;JPE;JPEG|GIF|*.GIF|All Images|*.BMP;*.JPG;*.JPE;*.JPGE;*.GIF|All Files|*.*"
'CommonDialog1.ShowOpen
'On Error GoTo PhotoError:
'ImageInstitutionLogo.Picture = LoadPicture(CommonDialog1.FileName)
'PreferanceInstitutionLogoFileName = CommonDialog1.FileName
'Exit Sub
'PhotoError:
'If Err.Number = 481 Then
'    TemResponce = MsgBox("The Photo you choose is not suitable, try using a medium size BMP, JPG or GIF file", vbOKOnly, "Photo Error")
'ElseIf Err.Number = 53 Then
'    TemResponce = MsgBox("No photo exist to selected, try to select again correctly.", vbOKOnly, "Photo Error")
'Else
'    TemResponce = MsgBox("An unknown error has occured, try again," & Chr(13) & Err.Description, vbOKOnly, "Photo Error")
'End If
'
'End Sub
'
'
'
'Private Sub bttnSaveExit_Click()
'Call SavePreferances
'Call LoadPreferances
'Unload Me
'End Sub
'
'Private Sub CheckInstitutionName_Click()
'If CheckInstitutionName.Value = 1 Then
'    txtInstitutionNameDisplay.Visible = True
'Else
'    txtInstitutionNameDisplay.Visible = False
'End If
'End Sub
'
'Private Sub Form_Load()
'Call LoadPreferances
'Call SetPreferances
'
'SSTab1.Tab = 0
'frameInstitution.Visible = True
'frameInvestigation.Visible = False
'frameResults.Visible = False
'FrameOther.Visible = False
'
'
'End Sub
'
'
'
'
'Private Sub SSTab1_Click(PreviousTab As Integer)
'Select Case SSTab1.Tab
'Case 0: frameInstitution.Visible = True
'        frameInvestigation.Visible = False
'        frameResults.Visible = False
'        FrameOther.Visible = False
'Case 1: frameInvestigation.Visible = True
'        frameInstitution.Visible = False
'        frameResults.Visible = False
'        FrameOther.Visible = False
'Case 2: frameResults.Visible = True
'        frameInstitution.Visible = False
'        frameInvestigation.Visible = False
'        FrameOther.Visible = False
'Case 3: FrameOther.Visible = True
'        frameInstitution.Visible = False
'        frameInvestigation.Visible = False
'        frameResults.Visible = False
'End Select
'End Sub
'
'
'
'Private Sub bttnHLineUP1_Click()
'HLine1.Y1 = HLine1.Y1 - 10
'HLine1.Y2 = HLine1.Y2 - 10
'End Sub
'
'Private Sub bttnHLineUP2_Click()
'HLine2.Y1 = HLine2.Y1 - 10
'HLine2.Y2 = HLine2.Y2 - 10
'End Sub
'
'Private Sub bttnHLineUP3_Click()
'HLine3.Y1 = HLine3.Y1 - 10
'HLine3.Y2 = HLine3.Y2 - 10
'End Sub
'
'Private Sub bttnHLineUP4_Click()
'HLine4.Y1 = HLine4.Y1 - 10
'HLine4.Y2 = HLine4.Y2 - 10
'End Sub
'
'
'Private Sub bttnHLineDown1_Click()
'HLine1.Y1 = HLine1.Y1 + 10
'HLine1.Y2 = HLine1.Y2 + 10
'End Sub
'
'Private Sub bttnHLineDown2_Click()
'HLine2.Y1 = HLine2.Y1 + 10
'HLine2.Y2 = HLine2.Y2 + 10
'End Sub
'
'Private Sub bttnHLineDown3_Click()
'HLine3.Y1 = HLine3.Y1 + 10
'HLine3.Y2 = HLine3.Y2 + 10
'End Sub
'
'Private Sub bttnHLineDown4_Click()
'HLine4.Y1 = HLine4.Y1 + 10
'HLine4.Y2 = HLine4.Y2 + 10
'End Sub
'
'Private Sub txtCommentsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtCommentsDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'
'Private Sub txtLblCommentsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblCommentsDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'
'Private Sub txtLblConfidentialDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblConfidentialDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblLabReportDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblLabReportDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblParametersDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblParametersDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtlblpatientnamedisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblPatientNameDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtlblpatientagedisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblPatientAgeDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtlblpatientsexdisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblPatientSexDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtlblpatientiddisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblPatientIDDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblDateDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblDateDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblReferancesDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblReferancesDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblSpecimanDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblSpecimanDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblSpecimanNoDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblSpecimanNoDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblTestDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblTestDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLbltimeDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblTimeDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblRDoctorDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblRDoctorDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblRInstitutionDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblRInstitutionDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblresultsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblResultsDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLblUnitsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLblUnitsDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtParamatersDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtParamatersDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtpatientnamedisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtPatientNameDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtpatientagedisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtPatientAgeDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtpatientsexdisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtPatientSexDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtpatientiddisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtPatientIDDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtDateDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtDateDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtReferancesDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtReferancesDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtSpecimanDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtSpecimanDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtSpecimanNoDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtSpecimanNoDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtTestDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtTestDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txttimeDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtTimeDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtRDoctorDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtRDoctorDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'Private Sub txtRInstitutionDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtRInstitutionDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'KeyCode = 0
'End Sub
'
'
'Private Sub txtAddressDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtAddressDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtLogo_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtLogo
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtAdvertiestmentDisplay1_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtAdvertiestmentDisplay1
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtAdvertiestmentDisplay2_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtAdvertiestmentDisplay2
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtContactDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtContactDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtDoctorMLTDetailsDisplay1_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtDoctorMLTDetailsDisplay1
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtDoctorMLTDetailsDisplay2_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtDoctorMLTDetailsDisplay2
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtDoctorMLTDetailsDisplay3_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtDoctorMLTDetailsDisplay3
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtDoctorMLTDetailsDisplay4_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtDoctorMLTDetailsDisplay4
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtInstitutionNameDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtInstitutionNameDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtMessage1_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtMessage1
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'
'
'Private Sub SliderLeft_Change()
'Call SetPrintableArea
'Call SetLabels
'End Sub
'Private Sub Sliderright_Change()
'Call SetPrintableArea
'Call SetLabels
'End Sub
'Private Sub Slidertop_Change()
'Call SetPrintableArea
'Call SetLabels
'End Sub
'Private Sub Sliderbottom_Change()
'Call SetPrintableArea
'Call SetLabels
'End Sub
'
'
'Private Sub SetPrintableArea()
'Dim PreviousWidth As Long
'Dim PreviousHeight As Long
'Dim PreviousLeft As Long
'Dim PreviousTop As Long
'Dim CurrentWidth As Long
'Dim CurrentHeight As Long
'Dim CurrentLeft As Long
'Dim CurrentTop As Long
'
'PreviousWidth = FramePrintableArea.Width
'PreviousHeight = FramePrintableArea.Height
'
'FramePrintableArea.Top = FramePaperArea.Height * SliderTop.Value / 20
'FramePrintableArea.Left = FramePaperArea.Width * SliderLeft.Value / 20
'FramePrintableArea.Height = (FramePaperArea.Height) * (SliderBottom.Value - SliderTop.Value) / 20
'FramePrintableArea.Width = FramePaperArea.Width * (SliderRight.Value - SliderLeft.Value) / 20
'
'CurrentWidth = FramePrintableArea.Width
'CurrentHeight = FramePrintableArea.Height
'
'VerticalExpansion = PreviousHeight / CurrentHeight
'HorizontalExpansion = PreviousWidth / CurrentWidth
'
'End Sub
'
'Private Sub SetLabels()
'Set temsource = txtLogo
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtInstitutionNameDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtAddressDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtContactDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtAdvertiestmentDisplay1
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtAdvertiestmentDisplay2
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtDoctorMLTDetailsDisplay1
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtDoctorMLTDetailsDisplay2
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtDoctorMLTDetailsDisplay3
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtDoctorMLTDetailsDisplay4
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtMessageDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtPatientNameDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtPatientAgeDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtPatientIDDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtPatientSexDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtDateDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtTimeDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtRDoctorDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'
'Set temsource = txtRInstitutionDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblPatientNameDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblPatientAgeDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblPatientIDDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblPatientSexDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblDateDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblTimeDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblRDoctorDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'
'Set temsource = txtLblRInstitutionDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'
'Set temsource = txtLblTestDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'
'Set temsource = txtTestDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblSpecimanDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtSpecimanDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblSpecimanNoDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtSpecimanNoDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'
'Set temsource = txtLblParametersDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblResultsDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblUnitsDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblReferancesDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtParamatersDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtResultsDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtUnitsDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtReferancesDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblCommentsDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtCommentsDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblConfidentialDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'Set temsource = txtLblLabReportDisplay
'temsource.Height = temsource.Height / VerticalExpansion
'temsource.Top = temsource.Top / VerticalExpansion
'temsource.Width = temsource.Width / HorizontalExpansion
'temsource.Left = temsource.Left / HorizontalExpansion
'
'HLine1.Y1 = HLine1.Y1 / VerticalExpansion
'HLine1.Y2 = HLine1.Y2 / VerticalExpansion
'HLine2.Y1 = HLine2.Y1 / VerticalExpansion
'HLine2.Y2 = HLine2.Y2 / VerticalExpansion
'HLine3.Y1 = HLine3.Y1 / VerticalExpansion
'HLine3.Y2 = HLine3.Y2 / VerticalExpansion
'HLine4.Y1 = HLine4.Y1 / VerticalExpansion
'HLine4.Y2 = HLine4.Y2 / VerticalExpansion
'
'HLine1.X1 = 0
'HLine1.X2 = FramePrintableArea.Width
'HLine2.X1 = 0
'HLine2.X2 = FramePrintableArea.Width
'HLine3.X1 = 0
'HLine3.X2 = FramePrintableArea.Width
'HLine4.X1 = 0
'HLine4.X2 = FramePrintableArea.Width
'
'
'End Sub
'
'Private Sub LoadPreferances()
'Call OpenPreferances
'
'With DataEnvironment1.rscmmdPreferances
'    If .RecordCount = 0 Then Exit Sub
'    .MoveFirst
'
'        PreferanceLX = !LX
'        PreferanceRX = !RX
'        PreferanceTY = !TY
'        PreferanceBY = !By
'
'        If Not IsNull(!chkInstitutionName) Then PreferancechkInstitutionName = !chkInstitutionName
'        If Not IsNull(!txtInstitutionName) Then PreferancetxtInstitutionName = DecreptedWord(!txtInstitutionName)
'        If Not IsNull(!InstitutionNameFontName) Then PreferanceInstitutionNameFontName = !InstitutionNameFontName
'        If Not IsNull(!InstitutionNameFontSize) Then PreferanceInstitutionNameFontSize = !InstitutionNameFontSize
'        If Not IsNull(!InstitutionNameFontBold) Then PreferanceInstitutionNameFontBold = !InstitutionNameFontBold
'        If Not IsNull(!InstitutionNameFontItalic) Then PreferanceInstitutionNameFontItalic = !InstitutionNameFontItalic
'        If Not IsNull(!chkInstitutionAddress) Then PreferancechkInstitutionAddress = !chkInstitutionAddress
'        If Not IsNull(!chkinstitutionlogo) Then PreferancechkInstitutionLogo = !chkinstitutionlogo
'        If Not IsNull(!InstitutionLogoFileName) Then PreferanceInstitutionLogoFileName = !InstitutionLogoFileName
'        If Not IsNull(!txtInstitutionAddress) Then PreferancetxtInstitutionAddress = !txtInstitutionAddress
'        If Not IsNull(!InstitutionaddressFontName) Then PreferanceInstitutionAddressFontName = !InstitutionaddressFontName
'        If Not IsNull(!InstitutionAddressFontSize) Then PreferanceInstitutionAddressFontSize = !InstitutionAddressFontSize
'        If Not IsNull(!InstitutionAddressFontBole) Then PreferanceInstitutionAddressFontBold = !InstitutionAddressFontBole
'        If Not IsNull(!InstitutionAddressFontItalic) Then PreferanceInstitutionAddressFontItalic = !InstitutionAddressFontItalic
'        If Not IsNull(!chkInstitutionContact) Then PreferancechkInstitutionContact = !chkInstitutionContact
'        If Not IsNull(!txtInstitutionContact) Then PreferancetxtInstitutionContact = !txtInstitutionContact
'        If Not IsNull(!InstitutionContactFontName) Then PreferanceInstitutionContactFontName = !InstitutionContactFontName
'        If Not IsNull(!InstitutionContactFontSize) Then PreferanceInstitutionContactFontSize = !InstitutionContactFontSize
'        If Not IsNull(!InstitutionContactFontBold) Then PreferanceInstitutionContactFontBold = !InstitutionContactFontBold
'        If Not IsNull(!InstitutionContactFontItalic) Then PreferanceInstitutionContactFontItalic = !InstitutionContactFontItalic
'        If Not IsNull(!InstitutionLogoFileName) Then PreferanceInstitutionLogoFileName = !InstitutionLogoFileName
'        If Not IsNull(!chkAdvertiestment1) Then PreferancechkAdvertiestment1 = !chkAdvertiestment1
'        If Not IsNull(!txtAdvertiestment1) Then PreferancetxtAdvertiestment1 = !txtAdvertiestment1
'        If Not IsNull(!AdvertiestmentFontName) Then PreferanceAdvertiestmentFontName = !AdvertiestmentFontName
'        If Not IsNull(!AdvertiestmentFontSize) Then PreferanceAdvertiestmentFontSize = !AdvertiestmentFontSize
'        If Not IsNull(!AdvertiestmentFontBold) Then PreferanceAdvertiestmentFontBold = !AdvertiestmentFontBold
'        If Not IsNull(!AdvertiestmentFontItalic) Then PreferanceAdvertiestmentFontItalic = !AdvertiestmentFontItalic
'        If Not IsNull(!chkAdvertiestment2) Then PreferancechkAdvertiestment2 = !chkAdvertiestment2
'        If Not IsNull(!txtAdvertiestment2) Then PreferancetxtAdvertiestment2 = !txtAdvertiestment2
'        If Not IsNull(!chkDoctorMLT1) Then PreferancechkDoctorMLT1 = !chkDoctorMLT1
'        If Not IsNull(!txtDoctorMLT1) Then PreferancetxtDoctorMLT1 = !txtDoctorMLT1
'        If Not IsNull(!DoctorMLTFontName) Then PreferanceDoctorMLTFontName = !DoctorMLTFontName
'        If Not IsNull(!DoctorMLTFontSize) Then PreferanceDoctorMLTFontSize = !DoctorMLTFontSize
'        If Not IsNull(!DoctorMLTFontBold) Then PreferanceDoctorMLTFontBold = !DoctorMLTFontBold
'        If Not IsNull(!DoctorMLTFontItalic) Then PreferanceDoctorMLTFontItalic = !DoctorMLTFontItalic
'        If Not IsNull(!chkDoctorMLT2) Then PreferancechkDoctorMLT2 = !chkDoctorMLT2
'        If Not IsNull(!txtDoctorMLT2) Then PreferancetxtDoctorMLT2 = !txtDoctorMLT2
'        If Not IsNull(!chkDoctorMLT3) Then PreferancechkDoctorMLT3 = !chkDoctorMLT3
'        If Not IsNull(!txtDoctorMLT3) Then PreferancetxtDoctorMLT3 = !txtDoctorMLT3
'        If Not IsNull(!chkDoctorMLT4) Then PreferancechkDoctorMLT4 = !chkDoctorMLT4
'        If Not IsNull(!txtDoctorMLT4) Then PreferancetxtDoctorMLT4 = !txtDoctorMLT4
'        If Not IsNull(!Messagechk) Then PreferanceChkMessage = !Messagechk
'        If Not IsNull(!Messagetxt) Then PreferancetxtMessage = !Messagetxt
'        If Not IsNull(!MessageFontName) Then PreferanceMessageFontName = !MessageFontName
'        If Not IsNull(!MessageFontsize) Then PreferanceMessageFontSize = !MessageFontsize
'        If Not IsNull(!MessageFontbold) Then PreferanceMessageFontBold = !MessageFontbold
'        If Not IsNull(!MessageFontitalic) Then PreferanceMessageFontItalic = !MessageFontitalic
'
'        If Not IsNull(!InstitutionNameLX) Then PreferanceInstitutionNameLX = !InstitutionNameLX
'        If Not IsNull(!InstitutionNameRX) Then PreferanceInstitutionNameRX = !InstitutionNameRX
'        If Not IsNull(!InstitutionNameTY) Then PreferanceInstitutionNameTY = !InstitutionNameTY
'        If Not IsNull(!InstitutionNameBY) Then PreferanceInstitutionNameBY = !InstitutionNameBY
'
'        If Not IsNull(!InstitutionAddressLX) Then PreferanceInstitutionAddressLX = !InstitutionAddressLX
'        If Not IsNull(!InstitutionAddressRX) Then PreferanceInstitutionAddressRX = !InstitutionAddressRX
'        If Not IsNull(!InstitutionAddressTY) Then PreferanceInstitutionAddressTY = !InstitutionAddressTY
'        If Not IsNull(!InstitutionAddressBY) Then PreferanceInstitutionAddressBY = !InstitutionAddressBY
'        If Not IsNull(!InstitutionContactLX) Then PreferanceInstitutionContactLX = !InstitutionContactLX
'        If Not IsNull(!InstitutionContactRX) Then PreferanceInstitutionContactRX = !InstitutionContactRX
'        If Not IsNull(!InstitutionContactTY) Then PreferanceInstitutionContactTY = !InstitutionContactTY
'        If Not IsNull(!InstitutionContactBY) Then PreferanceInstitutionContactBY = !InstitutionContactBY
'
'        If Not IsNull(!InstitutionLogoLX) Then PreferanceInstitutionLogoLX = !InstitutionLogoLX
'        If Not IsNull(!InstitutionLogoRX) Then PreferanceInstitutionLogoRX = !InstitutionLogoRX
'        If Not IsNull(!InstitutionLogoTY) Then PreferanceInstitutionLogoTY = !InstitutionLogoTY
'        If Not IsNull(!InstitutionLogoBY) Then PreferanceInstitutionLogoBY = !InstitutionLogoBY
'
'        If Not IsNull(!AdvertiestmentLX1) Then PreferanceAdvertiestmentLX1 = !AdvertiestmentLX1
'        If Not IsNull(!AdvertiestmentRX1) Then PreferanceAdvertiestmentRX1 = !AdvertiestmentRX1
'        If Not IsNull(!AdvertiestmentTY1) Then PreferanceAdvertiestmentTY1 = !AdvertiestmentTY1
'        If Not IsNull(!AdvertiestmentBY1) Then PreferanceAdvertiestmentBY1 = !AdvertiestmentBY1
'        If Not IsNull(!AdvertiestmentLX2) Then PreferanceAdvertiestmentLX2 = !AdvertiestmentLX2
'        If Not IsNull(!AdvertiestmentRX2) Then PreferanceAdvertiestmentRX2 = !AdvertiestmentRX2
'        If Not IsNull(!AdvertiestmentTY2) Then PreferanceAdvertiestmentTY2 = !AdvertiestmentTY2
'        If Not IsNull(!AdvertiestmentBY2) Then PreferanceAdvertiestmentBY2 = !AdvertiestmentBY2
'        If Not IsNull(!DoctorMLTLX1) Then PreferanceDoctorMLTLX1 = !DoctorMLTLX1
'        If Not IsNull(!DoctorMLTRX1) Then PreferanceDoctorMLTRX1 = !DoctorMLTRX1
'        If Not IsNull(!DoctorMLTTY1) Then PreferanceDoctorMLTTY1 = !DoctorMLTTY1
'        If Not IsNull(!DoctorMLTBY1) Then PreferanceDoctorMLTBY1 = !DoctorMLTBY1
'        If Not IsNull(!DoctorMLTLX2) Then PreferanceDoctorMLTLX2 = !DoctorMLTLX2
'        If Not IsNull(!DoctorMLTRX2) Then PreferanceDoctorMLTRX2 = !DoctorMLTRX2
'        If Not IsNull(!DoctorMLTTY2) Then PreferanceDoctorMLTTY2 = !DoctorMLTTY2
'        If Not IsNull(!DoctorMLTBY2) Then PreferanceDoctorMLTBY2 = !DoctorMLTBY2
'        If Not IsNull(!DoctorMLTLX3) Then PreferanceDoctorMLTLX3 = !DoctorMLTLX3
'        If Not IsNull(!DoctorMLTRX3) Then PreferanceDoctorMLTRX3 = !DoctorMLTRX3
'        If Not IsNull(!DoctorMLTTY3) Then PreferanceDoctorMLTTY3 = !DoctorMLTTY3
'        If Not IsNull(!DoctorMLTBY3) Then PreferanceDoctorMLTBY3 = !DoctorMLTBY3
'        If Not IsNull(!DoctorMLTLX4) Then PreferanceDoctorMLTLX4 = !DoctorMLTLX4
'        If Not IsNull(!DoctorMLTRX4) Then PreferanceDoctorMLTRX4 = !DoctorMLTRX4
'        If Not IsNull(!DoctorMLTTY4) Then PreferanceDoctorMLTTY4 = !DoctorMLTTY4
'        If Not IsNull(!DoctorMLTBY4) Then PreferanceDoctorMLTBY4 = !DoctorMLTBY4
'        If Not IsNull(!messageLX) Then PreferanceMessageLX = !messageLX
'        If Not IsNull(!messageRX) Then PreferanceMessageRX = !messageRX
'        If Not IsNull(!messageTY) Then PreferanceMessageTY = !messageTY
'        If Not IsNull(!messageBY) Then PreferanceMessageBY = !messageBY
'
'        If Not IsNull(!sliderleftvalue) Then SliderLeft.Value = !sliderleftvalue
'        If Not IsNull(!sliderrightvalue) Then SliderRight.Value = !sliderrightvalue
'        If Not IsNull(!slidertopvalue) Then SliderTop.Value = !slidertopvalue
'        If Not IsNull(!sliderbottomvalue) Then SliderBottom.Value = !sliderbottomvalue
'
'        If Not IsNull(!PreferancechkPatientName) Then PreferancechkPatientName = !PreferancechkPatientName
'        If Not IsNull(!PreferancechkPatientAge) Then PreferancechkPatientAge = !PreferancechkPatientAge
'        If Not IsNull(!PreferancechkPatientID) Then PreferancechkPatientID = !PreferancechkPatientID
'        If Not IsNull(!PreferancechkPatientSex) Then PreferancechkPatientSex = !PreferancechkPatientSex
'        If Not IsNull(!PreferancechkDate) Then PreferancechkDate = !PreferancechkDate
'        If Not IsNull(!PreferancechkTime) Then PreferancechkTime = !PreferancechkTime
'        If Not IsNull(!PreferancechkRDoctorName) Then PreferancechkRDoctorName = !PreferancechkRDoctorName
'        If Not IsNull(!PreferancechkRInstitutionName) Then PreferancechkRInstitutionName = !PreferancechkRInstitutionName
'        If Not IsNull(!PreferancechkLblPatientName) Then PreferancechkLblPatientName = !PreferancechkLblPatientName
'        If Not IsNull(!PreferancechkLblPatientAge) Then PreferancechkLblPatientAge = !PreferancechkLblPatientAge
'        If Not IsNull(!PreferancechkLblPatientID) Then PreferancechkLblPatientID = !PreferancechkLblPatientID
'        If Not IsNull(!PreferancechkLblPatientSex) Then PreferancechkLblPatientSex = !PreferancechkLblPatientSex
'        If Not IsNull(!PreferancechkLblDate) Then PreferancechkLblDate = !PreferancechkLblDate
'        If Not IsNull(!PreferancechkLblTime) Then PreferancechkLblTime = !PreferancechkLblTime
'        If Not IsNull(!PreferancechkLblRDoctorName) Then PreferancechkLblRDoctorName = !PreferancechkLblRDoctorName
'        If Not IsNull(!PreferancechkLblRInstitutionName) Then PreferancechkLblRInstitutionName = !PreferancechkLblRInstitutionName
'        If Not IsNull(!PreferancetxtLblPatientName) Then PreferancetxtLblPatientName = !PreferancetxtLblPatientName
'        If Not IsNull(!PreferancetxtLblPatientAge) Then PreferancetxtLblPatientAge = !PreferancetxtLblPatientAge
'        If Not IsNull(!PreferancetxtLblPatientID) Then PreferancetxtLblPatientID = !PreferancetxtLblPatientID
'        If Not IsNull(!PreferancetxtLblPatientSex) Then PreferancetxtLblPatientSex = !PreferancetxtLblPatientSex
'        If Not IsNull(!PreferancetxtLblDate) Then PreferancetxtLblDate = !PreferancetxtLblDate
'        If Not IsNull(!PreferancetxtLblTime) Then PreferancetxtLblTime = !PreferancetxtLblTime
'        If Not IsNull(!PreferancetxtLblRDoctorName) Then PreferancetxtLblRDoctorName = !PreferancetxtLblRDoctorName
'        If Not IsNull(!PreferancetxtLblRInstitutionName) Then PreferancetxtLblRInstitutionName = !PreferancetxtLblRInstitutionName
'        If Not IsNull(!PreferanceLblPatientNameLX) Then PreferanceLblPatientNameLX = !PreferanceLblPatientNameLX
'        If Not IsNull(!PreferanceLblPatientAgeLX) Then PreferanceLblPatientAgeLX = !PreferanceLblPatientAgeLX
'        If Not IsNull(!PreferanceLblPatientIDLX) Then PreferanceLblPatientIDLX = !PreferanceLblPatientIDLX
'        If Not IsNull(!PreferanceLblPatientSexLX) Then PreferanceLblPatientSexLX = !PreferanceLblPatientSexLX
'        If Not IsNull(!PreferanceLblDateLX) Then PreferanceLblDateLX = !PreferanceLblDateLX
'        If Not IsNull(!PreferanceLblTimeLX) Then PreferanceLblTimeLX = !PreferanceLblTimeLX
'        If Not IsNull(!PreferanceLblRDoctorNameLX) Then PreferanceLblRDoctorNameLX = !PreferanceLblRDoctorNameLX
'        If Not IsNull(!PreferanceLblRInstitutionNameLX) Then PreferanceLblRInstitutionNameLX = !PreferanceLblRInstitutionNameLX
'        If Not IsNull(!PreferancePatientNameLX) Then PreferancePatientNameLX = !PreferancePatientNameLX
'        If Not IsNull(!PreferancePatientAgeLX) Then PreferancePatientAgeLX = !PreferancePatientAgeLX
'        If Not IsNull(!PreferancePatientIDLX) Then PreferancePatientIDLX = !PreferancePatientIDLX
'        If Not IsNull(!PreferancePatientSexLX) Then PreferancePatientSexLX = !PreferancePatientSexLX
'        If Not IsNull(!PreferanceDateLX) Then PreferanceDateLX = !PreferanceDateLX
'        If Not IsNull(!PreferanceTimeLX) Then PreferanceTimeLX = !PreferanceTimeLX
'        If Not IsNull(!PreferanceRDoctorNameLX) Then PreferanceRDoctorNameLX = !PreferanceRDoctorNameLX
'        If Not IsNull(!PreferanceRInstitutionNameLX) Then PreferanceRInstitutionNameLX = !PreferanceRInstitutionNameLX
'        If Not IsNull(!PreferanceLblPatientNameTY) Then PreferanceLblPatientNameTY = !PreferanceLblPatientNameTY
'        If Not IsNull(!PreferanceLblPatientAgeTY) Then PreferanceLblPatientAgeTY = !PreferanceLblPatientAgeTY
'        If Not IsNull(!PreferanceLblPatientIDTY) Then PreferanceLblPatientIDTY = !PreferanceLblPatientIDTY
'        If Not IsNull(!PreferanceLblPatientSexTY) Then PreferanceLblPatientSexTY = !PreferanceLblPatientSexTY
'        If Not IsNull(!PreferanceLblDateTY) Then PreferanceLblDateTY = !PreferanceLblDateTY
'        If Not IsNull(!PreferanceLblTimeTY) Then PreferanceLblTimeTY = !PreferanceLblTimeTY
'        If Not IsNull(!PreferanceLblRDoctorNameTY) Then PreferanceLblRDoctorNameTY = !PreferanceLblRDoctorNameTY
'        If Not IsNull(!PreferanceLblRInstitutionNameTY) Then PreferanceLblRInstitutionNameTY = !PreferanceLblRInstitutionNameTY
'        If Not IsNull(!PreferancePatientNameTY) Then PreferancePatientNameTY = !PreferancePatientNameTY
'        If Not IsNull(!PreferancePatientAgeTY) Then PreferancePatientAgeTY = !PreferancePatientAgeTY
'        If Not IsNull(!PreferancePatientIDTY) Then PreferancePatientIDTY = !PreferancePatientIDTY
'        If Not IsNull(!PreferancePatientSexTY) Then PreferancePatientSexTY = !PreferancePatientSexTY
'        If Not IsNull(!PreferanceDateTY) Then PreferanceDateTY = !PreferanceDateTY
'        If Not IsNull(!PreferanceTimeTY) Then PreferanceTimeTY = !PreferanceTimeTY
'        If Not IsNull(!PreferanceRDoctorNameTY) Then PreferanceRDoctorNameTY = !PreferanceRDoctorNameTY
'        If Not IsNull(!PreferanceRInstitutionNameTY) Then PreferanceRInstitutionNameTY = !PreferanceRInstitutionNameTY
'        If Not IsNull(!PreferanceLabelFontName) Then PreferanceLabelFontName = !PreferanceLabelFontName
'        If Not IsNull(!PreferanceLabelFontSize) Then PreferanceLabelFontSize = !PreferanceLabelFontSize
'        If Not IsNull(!PreferanceLabelFontBold) Then PreferanceLabelFontBold = !PreferanceLabelFontBold
'        If Not IsNull(!PreferanceLabelFontItalic) Then PreferanceLabelFontItalic = !PreferanceLabelFontItalic
'        If Not IsNull(!PreferanceTextFontName) Then PreferanceTextFontName = !PreferanceTextFontName
'        If Not IsNull(!PreferanceTextFontSize) Then PreferanceTextFontSize = !PreferanceTextFontSize
'        If Not IsNull(!PreferanceTextFontBold) Then PreferanceTextFontBold = !PreferanceTextFontBold
'        If Not IsNull(!PreferanceTextFontItalic) Then PreferanceTextFontItalic = !PreferanceTextFontItalic
'
'        If Not IsNull(!PreferanceChkLblTest) Then PreferanceChkLblTest = !PreferanceChkLblTest
'        If Not IsNull(!PreferanceChkLblSpeciman) Then PreferanceChkLblSpeciman = !PreferanceChkLblSpeciman
'        If Not IsNull(!PreferanceChkLblSpecimanNo) Then PreferanceChkLblSpecimanNo = !PreferanceChkLblSpecimanNo
'        If Not IsNull(!PreferanceTxtLblTest) Then PreferanceTxtLblTest = !PreferanceTxtLblTest
'        If Not IsNull(!PreferanceTxtLblSpeciman) Then PreferanceTxtLblSpeciman = !PreferanceTxtLblSpeciman
'        If Not IsNull(!PreferanceTxtLblSpecimanNo) Then PreferanceTxtLblSpecimanNo = !PreferanceTxtLblSpecimanNo
'        If Not IsNull(!PreferanceChkTest) Then PreferanceChkTest = !PreferanceChkTest
'        If Not IsNull(!PreferanceChkSpeciman) Then PreferanceChkSpeciman = !PreferanceChkSpeciman
'        If Not IsNull(!PreferanceChkSpecimanNo) Then PreferanceChkSpecimanNo = !PreferanceChkSpecimanNo
'        If Not IsNull(!PreferanceChkLblParameters) Then PreferanceChkLblParameters = !PreferanceChkLblParameters
'        If Not IsNull(!PreferanceChkLblResults) Then PreferanceChkLblResults = !PreferanceChkLblResults
'        If Not IsNull(!PreferanceChkLblUnits) Then PreferanceChkLblUnits = !PreferanceChkLblUnits
'        If Not IsNull(!PreferanceChkLblReferances) Then PreferanceChkLblReferances = !PreferanceChkLblReferances
'        If Not IsNull(!PreferanceChkLblComments) Then PreferanceChkLblComments = !PreferanceChkLblComments
'        If Not IsNull(!PreferanceTxtLblParameters) Then PreferanceTxtLblParameters = !PreferanceTxtLblParameters
'        If Not IsNull(!PreferanceTxtLblResults) Then PreferanceTxtLblResults = !PreferanceTxtLblResults
'        If Not IsNull(!PreferanceTxtLblUnits) Then PreferanceTxtLblUnits = !PreferanceTxtLblUnits
'        If Not IsNull(!PreferanceTxtLblReferances) Then PreferanceTxtLblReferances = !PreferanceTxtLblReferances
'        If Not IsNull(!PreferanceTxtLblComments) Then PreferanceTxtLblComments = !PreferanceTxtLblComments
'        If Not IsNull(!PreferanceChkParameters) Then PreferanceChkParameters = !PreferanceChkParameters
'        If Not IsNull(!PreferanceChkResults) Then PreferanceChkResults = !PreferanceChkResults
'        If Not IsNull(!PreferanceChkUnits) Then PreferanceChkUnits = !PreferanceChkUnits
'        If Not IsNull(!PreferanceChkReferances) Then PreferanceChkReferances = !PreferanceChkReferances
'        If Not IsNull(!PreferanceChkComments) Then PreferanceChkComments = !PreferanceChkComments
'        If Not IsNull(!LblTestLX) Then LblTestLX = !LblTestLX
'        If Not IsNull(!LblTestTY) Then LblTestTY = !LblTestTY
'        If Not IsNull(!LblSpecimanLX) Then LblSpecimanLX = !LblSpecimanLX
'        If Not IsNull(!LblSpecimanTY) Then LblSpecimanTY = !LblSpecimanTY
'        If Not IsNull(!LblSpecimanNoLX) Then LblSpecimanNoLX = !LblSpecimanNoLX
'        If Not IsNull(!LblSpecimanNoTY) Then LblSpecimanNoTY = !LblSpecimanNoTY
'        If Not IsNull(!TestLX) Then TestLX = !TestLX
'        If Not IsNull(!TestTY) Then TestTY = !TestTY
'        If Not IsNull(!SpecimanLX) Then SpecimanLX = !SpecimanLX
'        If Not IsNull(!SpecimanTY) Then SpecimanTY = !SpecimanTY
'        If Not IsNull(!SpecimanNoLX) Then SpecimanNoLX = !SpecimanNoLX
'        If Not IsNull(!SpecimanNoTY) Then SpecimanNoTY = !SpecimanNoTY
'        If Not IsNull(!LblParametersLX) Then LblParametersLX = !LblParametersLX
'        If Not IsNull(!LblParametersTY) Then LblParametersTY = !LblParametersTY
'        If Not IsNull(!LblResultsLX) Then LblResultsLX = !LblResultsLX
'        If Not IsNull(!LblResultsTY) Then LblResultsTY = !LblResultsTY
'        If Not IsNull(!LblUnitsLX) Then LblUnitsLX = !LblUnitsLX
'        If Not IsNull(!LblUnitsTY) Then LblUnitsTY = !LblUnitsTY
'        If Not IsNull(!LblReferancesLX) Then LblReferancesLX = !LblReferancesLX
'        If Not IsNull(!LblReferancesTY) Then LblReferancesTY = !LblReferancesTY
'        If Not IsNull(!LblCommentsLX) Then LblCommentsLX = !LblCommentsLX
'        If Not IsNull(!LblCommentsTY) Then LblCommentsTY = !LblCommentsTY
'        If Not IsNull(!ParametersLX) Then ParametersLX = !ParametersLX
'        If Not IsNull(!ParametersRX) Then ParametersRX = !ParametersRX
'        If Not IsNull(!ParametersTY) Then ParametersTY = !ParametersTY
'        If Not IsNull(!ParametersBY) Then ParametersBY = !ParametersBY
'        If Not IsNull(!ResultsLX) Then ResultsLX = !ResultsLX
'        If Not IsNull(!ResultsRX) Then ResultsRX = !ResultsRX
'        If Not IsNull(!ResultsTY) Then ResultsTY = !ResultsTY
'        If Not IsNull(!ResultsBY) Then ResultsBY = !ResultsBY
'        If Not IsNull(!UnitsLX) Then UnitsLX = !UnitsLX
'        If Not IsNull(!UnitsRX) Then UnitsRX = !UnitsRX
'        If Not IsNull(!UnitsTY) Then UnitsTY = !UnitsTY
'        If Not IsNull(!UnitsBY) Then UnitsBY = !UnitsBY
'        If Not IsNull(!ReferancesLX) Then ReferancesLX = !ReferancesLX
'        If Not IsNull(!ReferancesRX) Then ReferancesRX = !ReferancesRX
'        If Not IsNull(!ReferancesTY) Then ReferancesTY = !ReferancesTY
'        If Not IsNull(!ReferancesBY) Then ReferancesBY = !ReferancesBY
'        If Not IsNull(!CommentsLX) Then CommentsLX = !CommentsLX
'        If Not IsNull(!CommentsRX) Then CommentsRX = !CommentsRX
'        If Not IsNull(!CommentsTY) Then CommentsTY = !CommentsTY
'        If Not IsNull(!CommentsBY) Then CommentsBY = !CommentsBY
'        If Not IsNull(!TopicFontName) Then TopicFontName = !TopicFontName
'        If Not IsNull(!TopicFontSize) Then TopicFontSize = !TopicFontSize
'        If Not IsNull(!TopicFontBold) Then TopicFontBold = !TopicFontBold
'        If Not IsNull(!TopicFontItalic) Then TopicFontItalic = !TopicFontItalic
'        If Not IsNull(!ValueFontName) Then ValueFontName = !ValueFontName
'        If Not IsNull(!ValueFontSize) Then ValueFontSize = !ValueFontSize
'        If Not IsNull(!ValueFontBold) Then ValueFontBold = !ValueFontBold
'        If Not IsNull(!ValueFontItalic) Then ValueFontItalic = !ValueFontItalic
'
'        If Not IsNull(!InbetweenY) Then InbetweenY = !InbetweenY
'        If Not IsNull(!HLineY1) Then HLineY1 = !HLineY1
'        If Not IsNull(!HLineY2) Then HLineY2 = !HLineY2
'        If Not IsNull(!HLineY3) Then HLineY3 = !HLineY3
'        If Not IsNull(!HLineY4) Then HLineY4 = !HLineY4
'        If Not IsNull(!LbltxtConfidential) Then LbltxtConfidential = !LbltxtConfidential
'        If Not IsNull(!LblConfidentialFontName) Then LblConfidentialFontName = !LblConfidentialFontName
'        If Not IsNull(!LblConfidentialLX) Then LblConfidentialLX = !LblConfidentialLX
'        If Not IsNull(!LblConfidentialTY) Then LblConfidentialTY = !LblConfidentialTY
'        If Not IsNull(!LblConfidentialFontSize) Then LblConfidentialFontSize = !LblConfidentialFontSize
'        If Not IsNull(!LblConfidentialFontName) Then LblConfidentialFontName = !LblConfidentialFontName
'        If Not IsNull(!LbltxtLblReport) Then LbltxtLblReport = !LbltxtLblReport
'        If Not IsNull(!LblReportFontName) Then LblReportFontName = !LblReportFontName
'        If Not IsNull(!LblReportFontSize) Then LblReportFontSize = !LblReportFontSize
'        If Not IsNull(!LblReportLX) Then LblReportLX = !LblReportLX
'        If Not IsNull(!LblReportTY) Then LblReportTY = !LblReportTY
'        If Not IsNull(!chkHLine3) Then chkHLine3 = !chkHLine3
'        If Not IsNull(!chkHLine2) Then chkHLine2 = !chkHLine2
'        If Not IsNull(!chkHLine4) Then chkHLine4 = !chkHLine4
'
'Call ClosePreferances
'End With
'End Sub
'
'Private Sub SetPreferances()
'
''Exit Sub
'
'    FramePrintableArea.Left = PreferanceLX * FramePaperArea.Width
'    FramePrintableArea.Top = PreferanceTY * FramePaperArea.Height
'    FramePrintableArea.Width = (PreferanceRX - PreferanceLX) * FramePaperArea.Width
'    FramePrintableArea.Height = (PreferanceBY - PreferanceTY) * FramePaperArea.Height
'
'    Set temsource = txtLogo
'    temsource.Left = FramePrintableArea.Width * PreferanceInstitutionLogoLX
'    temsource.Top = FramePrintableArea.Height * PreferanceInstitutionLogoTY
'    temsource.Width = FramePrintableArea.Width * (PreferanceInstitutionLogoRX - PreferanceInstitutionLogoLX)
'    temsource.Height = FramePrintableArea.Height * (PreferanceInstitutionLogoBY - PreferanceInstitutionLogoTY)
'
'    Set temsource = txtInstitutionNameDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceInstitutionNameLX
'    temsource.Top = FramePrintableArea.Height * PreferanceInstitutionNameTY
'    temsource.Width = FramePrintableArea.Width * (PreferanceInstitutionNameRX - PreferanceInstitutionNameLX)
'    temsource.Height = FramePrintableArea.Height * (PreferanceInstitutionNameBY - PreferanceInstitutionNameTY)
'
'    Set temsource = txtContactDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceInstitutionContactLX
'    temsource.Top = FramePrintableArea.Height * PreferanceInstitutionContactTY
'    temsource.Width = FramePrintableArea.Width * (PreferanceInstitutionContactRX - PreferanceInstitutionContactLX)
'    temsource.Height = FramePrintableArea.Height * (PreferanceInstitutionContactBY - PreferanceInstitutionContactTY)
'
'    Set temsource = txtAddressDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceInstitutionAddressLX
'    temsource.Top = FramePrintableArea.Height * PreferanceInstitutionAddressTY
'    temsource.Width = FramePrintableArea.Width * (PreferanceInstitutionAddressRX - PreferanceInstitutionAddressLX)
'    temsource.Height = FramePrintableArea.Height * (PreferanceInstitutionAddressBY - PreferanceInstitutionAddressTY)
'
'    Set temsource = txtAdvertiestmentDisplay1
'    temsource.Left = FramePrintableArea.Width * PreferanceAdvertiestmentLX1
'    temsource.Top = FramePrintableArea.Height * PreferanceAdvertiestmentTY1
'    temsource.Width = FramePrintableArea.Width * (PreferanceAdvertiestmentRX1 - PreferanceAdvertiestmentLX1)
'    temsource.Height = FramePrintableArea.Height * (PreferanceAdvertiestmentBY1 - PreferanceAdvertiestmentTY1)
'
'    Set temsource = txtAdvertiestmentDisplay2
'    temsource.Left = FramePrintableArea.Width * PreferanceAdvertiestmentLX2
'    temsource.Top = FramePrintableArea.Height * PreferanceAdvertiestmentTY2
'    temsource.Width = FramePrintableArea.Width * (PreferanceAdvertiestmentRX2 - PreferanceAdvertiestmentLX2)
'    temsource.Height = FramePrintableArea.Height * (PreferanceAdvertiestmentBY2 - PreferanceAdvertiestmentTY2)
'
'    Set temsource = txtDoctorMLTDetailsDisplay1
'    temsource.Left = FramePrintableArea.Width * PreferanceDoctorMLTLX1
'    temsource.Top = FramePrintableArea.Height * PreferanceDoctorMLTTY1
'    temsource.Width = FramePrintableArea.Width * (PreferanceDoctorMLTRX1 - PreferanceDoctorMLTLX1)
'    temsource.Height = FramePrintableArea.Height * (PreferanceDoctorMLTBY1 - PreferanceDoctorMLTTY1)
'
'    Set temsource = txtDoctorMLTDetailsDisplay2
'    temsource.Left = FramePrintableArea.Width * PreferanceDoctorMLTLX2
'    temsource.Top = FramePrintableArea.Height * PreferanceDoctorMLTTY2
'    temsource.Width = FramePrintableArea.Width * (PreferanceDoctorMLTRX2 - PreferanceDoctorMLTLX2)
'    temsource.Height = FramePrintableArea.Height * (PreferanceDoctorMLTBY2 - PreferanceDoctorMLTTY2)
'
'    Set temsource = txtDoctorMLTDetailsDisplay3
'    temsource.Left = FramePrintableArea.Width * PreferanceDoctorMLTLX3
'    temsource.Top = FramePrintableArea.Height * PreferanceDoctorMLTTY3
'    temsource.Width = FramePrintableArea.Width * (PreferanceDoctorMLTRX3 - PreferanceDoctorMLTLX3)
'    temsource.Height = FramePrintableArea.Height * (PreferanceDoctorMLTBY3 - PreferanceDoctorMLTTY3)
'
'    Set temsource = txtDoctorMLTDetailsDisplay4
'    temsource.Left = FramePrintableArea.Width * PreferanceDoctorMLTLX4
'    temsource.Top = FramePrintableArea.Height * PreferanceDoctorMLTTY4
'    temsource.Width = FramePrintableArea.Width * (PreferanceDoctorMLTRX4 - PreferanceDoctorMLTLX4)
'    temsource.Height = FramePrintableArea.Height * (PreferanceDoctorMLTBY4 - PreferanceDoctorMLTTY4)
'
'
'    Set temsource = txtLblPatientNameDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceLblPatientNameLX
'    temsource.Top = FramePrintableArea.Height * PreferanceLblPatientNameTY
'
'    Set temsource = txtLblPatientAgeDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceLblPatientAgeLX
'    temsource.Top = FramePrintableArea.Height * PreferanceLblPatientAgeTY
'
'    Set temsource = txtLblPatientSexDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceLblPatientSexLX
'    temsource.Top = FramePrintableArea.Height * PreferanceLblPatientSexTY
'
'    Set temsource = txtLblPatientIDDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceLblPatientIDLX
'    temsource.Top = FramePrintableArea.Height * PreferanceLblPatientIDTY
'
'    Set temsource = txtLblDateDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceLblDateLX
'    temsource.Top = FramePrintableArea.Height * PreferanceLblDateTY
'
'    Set temsource = txtLblTimeDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceLblTimeLX
'    temsource.Top = FramePrintableArea.Height * PreferanceLblTimeTY
'
'    Set temsource = txtLblRDoctorDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceLblRDoctorNameLX
'    temsource.Top = FramePrintableArea.Height * PreferanceLblRDoctorNameTY
'
'    Set temsource = txtLblRInstitutionDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceLblRInstitutionNameLX
'    temsource.Top = FramePrintableArea.Height * PreferanceLblRInstitutionNameTY
'
'    Set temsource = txtPatientNameDisplay
'    temsource.Left = FramePrintableArea.Width * PreferancePatientNameLX
'    temsource.Top = FramePrintableArea.Height * PreferancePatientNameTY
'
'    Set temsource = txtPatientAgeDisplay
'    temsource.Left = FramePrintableArea.Width * PreferancePatientAgeLX
'    temsource.Top = FramePrintableArea.Height * PreferancePatientAgeTY
'
'    Set temsource = txtPatientSexDisplay
'    temsource.Left = FramePrintableArea.Width * PreferancePatientSexLX
'    temsource.Top = FramePrintableArea.Height * PreferancePatientSexTY
'
'    Set temsource = txtPatientIDDisplay
'    temsource.Left = FramePrintableArea.Width * PreferancePatientIDLX
'    temsource.Top = FramePrintableArea.Height * PreferancePatientIDTY
'
'    Set temsource = txtDateDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceDateLX
'    temsource.Top = FramePrintableArea.Height * PreferanceDateTY
'
'    Set temsource = txtTimeDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceTimeLX
'    temsource.Top = FramePrintableArea.Height * PreferanceTimeTY
'
'    Set temsource = txtRDoctorDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceRDoctorNameLX
'    temsource.Top = FramePrintableArea.Height * PreferanceRDoctorNameTY
'
'    Set temsource = txtRInstitutionDisplay
'    temsource.Left = FramePrintableArea.Width * PreferanceRInstitutionNameLX
'    temsource.Top = FramePrintableArea.Height * PreferanceRInstitutionNameTY
'
'
'    Set temsource = txtLblTestDisplay
'    temsource.Left = FramePrintableArea.Width * LblTestLX
'    temsource.Top = FramePrintableArea.Height * LblTestTY
'
'    Set temsource = txtTestDisplay
'    temsource.Left = FramePrintableArea.Width * TestLX
'    temsource.Top = FramePrintableArea.Height * TestTY
'
'    Set temsource = txtLblSpecimanDisplay
'    temsource.Left = FramePrintableArea.Width * LblSpecimanLX
'    temsource.Top = FramePrintableArea.Height * LblSpecimanTY
'
'    Set temsource = txtSpecimanDisplay
'    temsource.Left = FramePrintableArea.Width * SpecimanLX
'    temsource.Top = FramePrintableArea.Height * SpecimanTY
'
'    Set temsource = txtLblSpecimanNoDisplay
'    temsource.Left = FramePrintableArea.Width * LblSpecimanNoLX
'    temsource.Top = FramePrintableArea.Height * LblSpecimanNoTY
'
'    Set temsource = txtSpecimanNoDisplay
'    temsource.Left = FramePrintableArea.Width * SpecimanNoLX
'    temsource.Top = FramePrintableArea.Height * SpecimanNoTY
'
'    Set temsource = txtLblParametersDisplay
'    temsource.Left = FramePrintableArea.Width * LblParametersLX
'    temsource.Top = FramePrintableArea.Height * LblParametersTY
'
'    Set temsource = txtLblResultsDisplay
'    temsource.Left = FramePrintableArea.Width * LblResultsLX
'    temsource.Top = FramePrintableArea.Height * LblResultsTY
'
'    Set temsource = txtLblUnitsDisplay
'    temsource.Left = FramePrintableArea.Width * LblUnitsLX
'    temsource.Top = FramePrintableArea.Height * LblUnitsTY
'
'    Set temsource = txtLblReferancesDisplay
'    temsource.Left = FramePrintableArea.Width * LblReferancesLX
'    temsource.Top = FramePrintableArea.Height * LblReferancesTY
'
'    Set temsource = txtLblCommentsDisplay
'    temsource.Left = FramePrintableArea.Width * LblCommentsLX
'    temsource.Top = FramePrintableArea.Height * LblCommentsTY
'
'    Set temsource = txtLblConfidentialDisplay
'    temsource.Left = FramePrintableArea.Width * LblConfidentialLX
'    temsource.Top = FramePrintableArea.Height * LblConfidentialTY
'
'    Set temsource = txtLblLabReportDisplay
'    temsource.Left = FramePrintableArea.Width * LblReportLX
'    temsource.Top = FramePrintableArea.Height * LblReportTY
'
'
'
'    Set temsource = txtParamatersDisplay
'    temsource.Left = FramePrintableArea.Width * ParametersLX
'    temsource.Top = FramePrintableArea.Height * ParametersTY
'    temsource.Width = FramePrintableArea.Width * (ParametersRX - ParametersLX)
'    temsource.Height = FramePrintableArea.Height * (ParametersBY - ParametersTY)
'
'    Set temsource = txtResultsDisplay
'    temsource.Left = FramePrintableArea.Width * ResultsLX
'    temsource.Top = FramePrintableArea.Height * ResultsTY
'    temsource.Width = FramePrintableArea.Width * (ResultsRX - ResultsLX)
'    temsource.Height = FramePrintableArea.Height * (ResultsBY - ResultsTY)
'
'    Set temsource = txtUnitsDisplay
'    temsource.Left = FramePrintableArea.Width * UnitsLX
'    temsource.Top = FramePrintableArea.Height * UnitsTY
'    temsource.Width = FramePrintableArea.Width * (UnitsRX - UnitsLX)
'    temsource.Height = FramePrintableArea.Height * (UnitsBY - UnitsTY)
'
'    Set temsource = txtReferancesDisplay
'    temsource.Left = FramePrintableArea.Width * ReferancesLX
'    temsource.Top = FramePrintableArea.Height * ReferancesTY
'    temsource.Width = FramePrintableArea.Width * (ReferancesRX - ReferancesLX)
'    temsource.Height = FramePrintableArea.Height * (ReferancesBY - ReferancesTY)
'
'
'    Set temsource = txtCommentsDisplay
'    temsource.Left = FramePrintableArea.Width * CommentsLX
'    temsource.Top = FramePrintableArea.Height * CommentsTY
'    temsource.Width = FramePrintableArea.Width * (CommentsRX - CommentsLX)
'    temsource.Height = FramePrintableArea.Height * (CommentsBY - CommentsTY)
'
'    HLine1.Y1 = FramePrintableArea.Height * HLineY1
'    HLine1.Y2 = FramePrintableArea.Height * HLineY1
'    HLine1.X1 = 0
'    HLine1.X2 = FramePrintableArea.Width
'
'    HLine2.Y1 = FramePrintableArea.Height * HLineY2
'    HLine2.Y2 = FramePrintableArea.Height * HLineY2
'    HLine2.X1 = 0
'    HLine2.X2 = FramePrintableArea.Width
'
'    HLine3.Y1 = FramePrintableArea.Height * HLineY3
'    HLine3.Y2 = FramePrintableArea.Height * HLineY3
'    HLine3.X1 = 0
'    HLine3.X2 = FramePrintableArea.Width
'
'    HLine4.Y1 = FramePrintableArea.Height * HLineY4
'    HLine4.Y2 = FramePrintableArea.Height * HLineY4
'    HLine4.X1 = 0
'    HLine4.X2 = FramePrintableArea.Width
'
'
'    If PreferancechkInstitutionName = True Then
'        CheckInstitutionName.Value = 1
'    Else
'        CheckInstitutionName.Value = 0
'    End If
'    If PreferancechkInstitutionAddress = True Then
'        chkInstitutionAddress.Value = 1
'    Else
'        chkInstitutionAddress.Value = 0
'    End If
'    If PreferancechkInstitutionContact = True Then
'        chkInstitutionContact.Value = 1
'    Else
'        chkInstitutionContact.Value = 0
'    End If
'    If PreferancechkInstitutionLogo = True Then
'        chkLogo.Value = 1
'    Else
'        chkLogo.Value = 0
'    End If
'    If PreferancechkAdvertiestment1 = True Then
'        chkAdvertiestment1.Value = 1
'    Else
'        chkAdvertiestment1.Value = 0
'    End If
'    If PreferancechkAdvertiestment2 = True Then
'        chkAdvertiestment2.Value = 1
'    Else
'        chkAdvertiestment2.Value = 0
'    End If
'    If PreferancechkDoctorMLT1 = True Then
'        chkDoctorMLT1.Value = 1
'    Else
'        chkDoctorMLT1.Value = 0
'    End If
'    If PreferancechkDoctorMLT2 = True Then
'        chkDoctorMLT2.Value = 1
'    Else
'        chkDoctorMLT2.Value = 0
'    End If
'    If PreferancechkDoctorMLT3 = True Then
'        chkDoctorMLT3.Value = 1
'    Else
'        chkDoctorMLT3.Value = 0
'    End If
'    If PreferancechkDoctorMLT4 = True Then
'        chkDoctorMLT4.Value = 1
'    Else
'        chkDoctorMLT4.Value = 0
'    End If
'    If PreferanceChkMessage = True Then
'        chkMessage.Value = 1
'    Else
'        chkMessage.Value = 0
'    End If
'    If PreferancechkLblPatientName = True Then
'        chkLblPatientName.Value = 1
'    Else
'        chkLblPatientName.Value = 0
'    End If
'    If PreferancechkLblPatientAge = True Then
'        chkLblPatientAge.Value = 1
'    Else
'        chkLblPatientAge.Value = 0
'    End If
'    If PreferancechkLblPatientSex = True Then
'        chkLblPatientSex.Value = 1
'    Else
'        chkLblPatientSex.Value = 0
'    End If
'    If PreferancechkLblPatientID = True Then
'        chkLblPatientID.Value = 1
'    Else
'        chkLblPatientID.Value = 0
'    End If
'    If PreferancechkLblDate = True Then
'        chkLblDate.Value = 1
'    Else
'        chkLblDate.Value = 0
'    End If
'    If PreferancechkLblTime = True Then
'        chkLblTime.Value = 1
'    Else
'        chkLblTime.Value = 0
'    End If
'    If PreferancechkLblRDoctorName = True Then
'        chkLblRDoctorName.Value = 1
'    Else
'        chkLblRDoctorName.Value = 0
'    End If
'    If PreferancechkLblRInstitutionName = True Then
'        chkLblRInstitutionName.Value = 1
'    Else
'        chkLblRInstitutionName.Value = 0
'    End If
'    If PreferancechkPatientName = True Then
'        chkPatientName.Value = 1
'    Else
'        chkPatientName.Value = 0
'    End If
'    If PreferancechkPatientAge = True Then
'        chkPatientAge.Value = 1
'    Else
'        chkPatientAge.Value = 0
'    End If
'    If PreferancechkPatientSex = True Then
'        chkPatientSex.Value = 1
'    Else
'        chkPatientSex.Value = 0
'    End If
'    If PreferancechkPatientID = True Then
'        chkPatientID.Value = 1
'    Else
'        chkPatientID.Value = 0
'    End If
'    If PreferancechkDate = True Then
'        chkDate.Value = 1
'    Else
'        chkDate.Value = 0
'    End If
'    If PreferancechkTime = True Then
'        chkTime.Value = 1
'    Else
'        chkTime.Value = 0
'    End If
'    If PreferancechkRDoctorName = True Then
'        chkRDoctorName.Value = 1
'    Else
'        chkRDoctorName.Value = 0
'    End If
'    If PreferancechkRInstitutionName = True Then
'        chkRInstitutionName.Value = 1
'    Else
'        chkRInstitutionName.Value = 0
'    End If
'
'    If PreferanceChkLblTest = True Then
'        chkLblTest.Value = 1
'    Else
'        chkLblTest.Value = 0
'    End If
'
'    If PreferanceChkTest = True Then
'        chkTest.Value = 1
'    Else
'        chkTest.Value = 0
'    End If
'
'    If PreferanceChkLblSpeciman = True Then
'        chkLblSpeciman.Value = 1
'    Else
'        chkLblSpeciman.Value = 0
'    End If
'
'    If PreferanceChkSpeciman = True Then
'        chkSpeciman.Value = 1
'    Else
'        chkSpeciman.Value = 0
'    End If
'
'
'    If PreferanceChkLblSpecimanNo = True Then
'        chkLblSpecimanNo.Value = 1
'    Else
'        chkLblSpecimanNo.Value = 0
'    End If
'
'    If PreferanceChkSpecimanNo = True Then
'        chkSpecimanNo.Value = 1
'    Else
'        chkSpecimanNo.Value = 0
'    End If
'
'    If PreferanceChkLblParameters = True Then
'        chkLblParameters.Value = 1
'    Else
'        chkLblParameters.Value = 0
'    End If
'
'    If PreferanceChkParameters = True Then
'        chkParameters.Value = 1
'    Else
'        chkParameters.Value = 0
'    End If
'
'    If PreferanceChkLblResults = True Then
'        chkLblResults.Value = 1
'    Else
'        chkLblResults.Value = 0
'    End If
'
'    If PreferanceChkResults = True Then
'        chkResults.Value = 1
'    Else
'        chkResults.Value = 0
'    End If
'
'    If PreferanceChkLblUnits = True Then
'        chkLblUnits.Value = 1
'    Else
'        chkLblUnits.Value = 0
'    End If
'
'    If PreferanceChkUnits = True Then
'        chkUnits.Value = 1
'    Else
'        chkUnits.Value = 0
'    End If
'
'
'    If PreferanceChkLblReferances = True Then
'        chkLblReferances.Value = 1
'    Else
'        chkLblReferances.Value = 0
'    End If
'
'    If PreferanceChkReferances = True Then
'        chkReferances.Value = 1
'    Else
'        chkReferances.Value = 0
'    End If
'
'    If PreferanceChkLblComments = True Then
'        chkLblComments.Value = 1
'    Else
'        chkLblComments.Value = 0
'    End If
'
'    If PreferanceChkComments = True Then
'        chkComments.Value = 1
'    Else
'        chkComments.Value = 0
'    End If
'
'
'    If chkHLine2 = True Then
'        checkHLine2.Value = 1
'    Else
'        checkHLine2.Value = 0
'    End If
'
'    If chkHLine3 = True Then
'        checkHLine3.Value = 1
'    Else
'        checkHLine3.Value = 0
'    End If
'
'    If chkHLine4 = True Then
'        checkHLine4.Value = 1
'    Else
'        checkHLine4.Value = 0
'    End If
'
'    txtInstitutionName.Text = PreferancetxtInstitutionName
'    txtAdvertiestment1.Text = PreferancetxtAdvertiestment1
'    txtAdvertiestment2.Text = PreferancetxtAdvertiestment2
'    txtInstitutionAddress.Text = PreferancetxtInstitutionAddress
'    txtInstitutionContact.Text = PreferancetxtInstitutionContact
'    txtDoctorMLT1.Text = PreferancetxtDoctorMLT1
'    txtDoctorMLT2.Text = PreferancetxtDoctorMLT2
'    txtDoctorMLT3.Text = PreferancetxtDoctorMLT3
'    txtDoctorMLT4.Text = PreferancetxtDoctorMLT4
'    txtMessage1.Text = PreferancetxtMessage
'    txtLblPatientName.Text = PreferancetxtLblPatientName
'    txtLblPatientAge.Text = PreferancetxtLblPatientAge
'    txtLblPatientSex.Text = PreferancetxtLblPatientSex
'    txtLblPatientID.Text = PreferancetxtLblPatientID
'    txtLblDate.Text = PreferancetxtLblDate
'    txtLblTime.Text = PreferancetxtLblTime
'    txtLblRDoctorName.Text = PreferancetxtLblRDoctorName
'    txtLblRInstitutionName.Text = PreferancetxtLblRInstitutionName
'
'    txtLblTest.Text = PreferanceTxtLblTest
'    txtLblSpeciman.Text = PreferanceTxtLblSpeciman
'    txtLblSpecimanNo.Text = PreferanceTxtLblSpecimanNo
'    txtLblParameters.Text = PreferanceTxtLblParameters
'    txtLblResults.Text = PreferanceTxtLblResults
'    txtLblUnits.Text = PreferanceTxtLblUnits
'    txtLblReferances.Text = PreferanceTxtLblReferances
'    txtLblComments.Text = PreferanceTxtLblComments
'
'    txtConfidential.Text = LbltxtConfidential
'    txtLabReport.Text = LbltxtLblReport
'
''    txtLblTest.Text = PreferanceTxtLblTest
''    txtLblSpeciman.Text = PreferanceTxtLblSpeciman
''    txtLblSpecimanNo.Text = PreferanceTxtLblSpecimanNo
''    txtLblParameters.Text = PreferanceTxtLblParameters
''    txtLblResults.Text = PreferanceTxtLblResults
''    txtLblUnits.Text = PreferanceTxtLblUnits
''    txtLblReferances.Text = PreferanceTxtLblReferances
'
'    On Error Resume Next
'    ImageInstitutionLogo.Picture = LoadPicture(PreferanceInstitutionLogoFileName)
'
'End Sub
'
'Private Sub SavePreferances()
'
'Call OpenPreferances
'
'With DataEnvironment1.rscmmdPreferances
'
'    If .RecordCount = 0 Then
'        .AddNew
'    Else
'        .MoveFirst
'    End If
'
'    !LX = FramePrintableArea.Left / FramePaperArea.Width
'    !TY = FramePrintableArea.Top / FramePaperArea.Height
'    !RX = (FramePrintableArea.Left + FramePrintableArea.Width) / FramePaperArea.Width
'    !By = (FramePrintableArea.Top + FramePrintableArea.Height) / FramePaperArea.Height
'
'    Set temsource = txtLogo
'    !InstitutionLogoLX = temsource.Left / FramePrintableArea.Width
'    !InstitutionLogoTY = temsource.Top / FramePrintableArea.Height
'    !InstitutionLogoRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !InstitutionLogoBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtInstitutionNameDisplay
'    !InstitutionNameLX = temsource.Left / FramePrintableArea.Width
'    !InstitutionNameTY = temsource.Top / FramePrintableArea.Height
'    !InstitutionNameRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !InstitutionNameBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtContactDisplay
'    !InstitutionContactLX = temsource.Left / FramePrintableArea.Width
'    !InstitutionContactTY = temsource.Top / FramePrintableArea.Height
'    !InstitutionContactRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !InstitutionContactBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtAddressDisplay
'    !InstitutionAddressLX = temsource.Left / FramePrintableArea.Width
'    !InstitutionAddressTY = temsource.Top / FramePrintableArea.Height
'    !InstitutionAddressRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !InstitutionAddressBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtAdvertiestmentDisplay1
'    !AdvertiestmentLX1 = temsource.Left / FramePrintableArea.Width
'    !AdvertiestmentTY1 = temsource.Top / FramePrintableArea.Height
'    !AdvertiestmentRX1 = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !AdvertiestmentBY1 = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtAdvertiestmentDisplay2
'    !AdvertiestmentLX2 = temsource.Left / FramePrintableArea.Width
'    !AdvertiestmentTY2 = temsource.Top / FramePrintableArea.Height
'    !AdvertiestmentRX2 = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !AdvertiestmentBY2 = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtDoctorMLTDetailsDisplay1
'    !DoctorMLTLX1 = temsource.Left / FramePrintableArea.Width
'    !DoctorMLTTY1 = temsource.Top / FramePrintableArea.Height
'    !DoctorMLTRX1 = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !DoctorMLTBY1 = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtDoctorMLTDetailsDisplay2
'    !DoctorMLTLX2 = temsource.Left / FramePrintableArea.Width
'    !DoctorMLTTY2 = temsource.Top / FramePrintableArea.Height
'    !DoctorMLTRX2 = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !DoctorMLTBY2 = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtDoctorMLTDetailsDisplay3
'    !DoctorMLTLX3 = temsource.Left / FramePrintableArea.Width
'    !DoctorMLTTY3 = temsource.Top / FramePrintableArea.Height
'    !DoctorMLTRX3 = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !DoctorMLTBY3 = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtDoctorMLTDetailsDisplay4
'    !DoctorMLTLX4 = temsource.Left / FramePrintableArea.Width
'    !DoctorMLTTY4 = temsource.Top / FramePrintableArea.Height
'    !DoctorMLTRX4 = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !DoctorMLTBY4 = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtMessageDisplay
'    !messageLX = temsource.Left / FramePrintableArea.Width
'    !messageTY = temsource.Top / FramePrintableArea.Height
'    !messageRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !messageBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'
'    Set temsource = txtLblTestDisplay
'    !LblTestLX = temsource.Left / FramePrintableArea.Width
'    !LblTestTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblSpecimanDisplay
'    !LblSpecimanLX = temsource.Left / FramePrintableArea.Width
'    !LblSpecimanTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblSpecimanNoDisplay
'    !LblSpecimanNoLX = temsource.Left / FramePrintableArea.Width
'    !LblSpecimanNoTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtSpecimanNoDisplay
'    !SpecimanNoLX = temsource.Left / FramePrintableArea.Width
'    !SpecimanNoTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtTestDisplay
'    !TestLX = txtTestDisplay.Left / FramePrintableArea.Width
'    !TestTY = txtTestDisplay.Top / FramePrintableArea.Height
'
'    Set temsource = txtSpecimanDisplay
'    !SpecimanLX = temsource.Left / FramePrintableArea.Width
'    !SpecimanTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblParametersDisplay
'    !LblParametersLX = temsource.Left / FramePrintableArea.Width
'    !LblParametersTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblResultsDisplay
'    !LblResultsLX = temsource.Left / FramePrintableArea.Width
'    !LblResultsTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblResultsDisplay
'    !LblResultsLX = temsource.Left / FramePrintableArea.Width
'    !LblResultsTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblUnitsDisplay
'    !LblUnitsLX = temsource.Left / FramePrintableArea.Width
'    !LblUnitsTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblReferancesDisplay
'    !LblReferancesLX = temsource.Left / FramePrintableArea.Width
'    !LblReferancesTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblCommentsDisplay
'    !LblCommentsLX = temsource.Left / FramePrintableArea.Width
'    !LblCommentsTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtParamatersDisplay
'    !ParametersLX = temsource.Left / FramePrintableArea.Width
'    !ParametersTY = temsource.Top / FramePrintableArea.Height
'    !ParametersRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !ParametersBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtResultsDisplay
'    !ResultsLX = temsource.Left / FramePrintableArea.Width
'    !ResultsTY = temsource.Top / FramePrintableArea.Height
'    !ResultsRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !ResultsBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtUnitsDisplay
'    !UnitsLX = temsource.Left / FramePrintableArea.Width
'    !UnitsTY = temsource.Top / FramePrintableArea.Height
'    !UnitsRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !UnitsBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    .Update
'
'    Set temsource = txtReferancesDisplay
'    !ReferancesLX = temsource.Left / FramePrintableArea.Width
'    !ReferancesTY = temsource.Top / FramePrintableArea.Height
'    !ReferancesRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !ReferancesBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'    Set temsource = txtCommentsDisplay
'    !CommentsLX = temsource.Left / FramePrintableArea.Width
'    !CommentsTY = temsource.Top / FramePrintableArea.Height
'    !CommentsRX = (temsource.Left + temsource.Width) / FramePrintableArea.Width
'    !CommentsBY = (temsource.Top + temsource.Height) / FramePrintableArea.Height
'
'
'    Set temsource = txtPatientNameDisplay
'    !PreferancePatientNameLX = txtPatientNameDisplay.Left / FramePrintableArea.Width
'    !PreferancePatientNameTY = txtPatientNameDisplay.Top / FramePrintableArea.Height
'
'    Set temsource = txtPatientAgeDisplay
'    !PreferancePatientAgeLX = txtPatientAgeDisplay.Left / FramePrintableArea.Width
'    !PreferancePatientAgeTY = txtPatientAgeDisplay.Top / FramePrintableArea.Height
'
'    Set temsource = txtPatientSexDisplay
'    !PreferancePatientSexLX = txtPatientSexDisplay.Left / FramePrintableArea.Width
'    !PreferancePatientSexTY = txtPatientSexDisplay.Top / FramePrintableArea.Height
'
'    Set temsource = txtPatientIDDisplay
'    !PreferancePatientIDLX = txtPatientIDDisplay.Left / FramePrintableArea.Width
'    !PreferancePatientIDTY = txtPatientIDDisplay.Top / FramePrintableArea.Height
'
'    Set temsource = txtDateDisplay
'    !PreferanceDateLX = txtDateDisplay.Left / FramePrintableArea.Width
'    !PreferanceDateTY = txtDateDisplay.Top / FramePrintableArea.Height
'
'    .Update
'
'    Set temsource = txtTimeDisplay
'    !PreferanceTimeLX = txtTimeDisplay.Left / FramePrintableArea.Width
'    !PreferanceTimeTY = txtTimeDisplay.Top / FramePrintableArea.Height
'
'    Set temsource = txtRDoctorDisplay
'    !PreferanceRDoctorNameLX = txtRDoctorDisplay.Left / FramePrintableArea.Width
'    !PreferanceRDoctorNameTY = txtRDoctorDisplay.Top / FramePrintableArea.Height
'
'    Set temsource = txtRInstitutionDisplay
'    !PreferanceRInstitutionNameLX = temsource.Left / FramePrintableArea.Width
'    !PreferanceRInstitutionNameTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblPatientNameDisplay
'    !PreferanceLblPatientNameLX = temsource.Left / FramePrintableArea.Width
'    !PreferanceLblPatientNameTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblPatientAgeDisplay
'    !PreferanceLblPatientAgeLX = temsource.Left / FramePrintableArea.Width
'    !PreferanceLblPatientAgeTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblPatientSexDisplay
'    !PreferanceLblPatientSexLX = temsource.Left / FramePrintableArea.Width
'    !PreferanceLblPatientSexTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblPatientIDDisplay
'    !PreferanceLblPatientIDLX = temsource.Left / FramePrintableArea.Width
'    !PreferanceLblPatientIDTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblDateDisplay
'    !PreferanceLblDateLX = temsource.Left / FramePrintableArea.Width
'    !PreferanceLblDateTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblTimeDisplay
'    !PreferanceLblTimeLX = temsource.Left / FramePrintableArea.Width
'    !PreferanceLblTimeTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblRDoctorDisplay
'    !PreferanceLblRDoctorNameLX = temsource.Left / FramePrintableArea.Width
'    !PreferanceLblRDoctorNameTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblRInstitutionDisplay
'    !PreferanceLblRInstitutionNameLX = temsource.Left / FramePrintableArea.Width
'    !PreferanceLblRInstitutionNameTY = temsource.Top / FramePrintableArea.Height
'
'    .Update
'
'    Set temsource = txtLblConfidentialDisplay
'    !LblConfidentialLX = temsource.Left / FramePrintableArea.Width
'    !LblConfidentialTY = temsource.Top / FramePrintableArea.Height
'
'    Set temsource = txtLblLabReportDisplay
'    !LblReportLX = temsource.Left / FramePrintableArea.Width
'    !LblReportTY = temsource.Top / FramePrintableArea.Height
'
'    !HLineY1 = HLine1.Y1 / FramePrintableArea.Height
'    !HLineY2 = HLine2.Y1 / FramePrintableArea.Height
'    !HLineY3 = HLine3.Y1 / FramePrintableArea.Height
'    !HLineY4 = HLine4.Y1 / FramePrintableArea.Height
'
'    If checkHLine2.Value = 1 Then
'        !chkHLine2 = True
'    Else
'        !chkHLine2 = False
'    End If
'
'    If checkHLine3.Value = 1 Then
'        !chkHLine3 = True
'    Else
'        !chkHLine3 = False
'    End If
'
'    If checkHLine4.Value = 1 Then
'        !chkHLine4 = True
'    Else
'        !chkHLine4 = False
'    End If
'
'    !LbltxtConfidential = txtConfidential.Text
'    !LbltxtLblReport = txtLabReport.Text
'
'    .Update
'
'    !sliderleftvalue = SliderLeft.Value
'    !sliderrightvalue = SliderRight.Value
'    !slidertopvalue = SliderTop.Value
'    !sliderbottomvalue = SliderBottom.Value
'
'
'    ' ***************************
'
'    If CheckInstitutionName.Value = 1 Then
'        !chkInstitutionName = True
'    Else
'        !chkInstitutionName = False
'    End If
'
'    !txtInstitutionName = EncreptedWord(txtInstitutionName.Text)
'    !InstitutionNameFontName = PreferanceInstitutionNameFontName
'    !InstitutionNameFontSize = PreferanceInstitutionNameFontSize
'    !InstitutionNameFontBold = PreferanceInstitutionNameFontBold
'    !InstitutionNameFontItalic = PreferanceInstitutionNameFontItalic
'
'    If chkLogo.Value = 1 Then
'        !chkinstitutionlogo = True
'    Else
'        !chkinstitutionlogo = False
'    End If
'
'    If chkInstitutionAddress.Value = 1 Then
'        !chkInstitutionAddress = True
'    Else
'        !chkInstitutionAddress = False
'    End If
'
'    !txtInstitutionAddress = txtInstitutionAddress.Text
'    !InstitutionaddressFontName = PreferanceInstitutionAddressFontName
'    !InstitutionAddressFontSize = PreferanceInstitutionAddressFontSize
'    !InstitutionAddressFontBole = PreferanceInstitutionAddressFontBold
'    !InstitutionAddressFontItalic = PreferanceInstitutionAddressFontItalic
'
'
'    .Update
'
'    If chkInstitutionContact.Value = 1 Then
'        !chkInstitutionContact = True
'    Else
'        !chkInstitutionContact = False
'    End If
'
'    !txtInstitutionContact = txtInstitutionContact.Text
'    !InstitutionContactFontName = PreferanceInstitutionContactFontName
'    !InstitutionContactFontSize = PreferanceInstitutionContactFontSize
'    !InstitutionContactFontBold = PreferanceInstitutionContactFontBold
'    !InstitutionContactFontItalic = PreferanceInstitutionContactFontItalic
'    !InstitutionLogoFileName = PreferanceInstitutionLogoFileName
'
'    If chkLogo.Value = 1 Then
'        !chkinstitutionlogo = True
'    Else
'        !chkinstitutionlogo = False
'    End If
'
'
'    If chkAdvertiestment1.Value = 1 Then
'        !chkAdvertiestment1 = True
'    Else
'        !chkAdvertiestment1 = False
'    End If
'
'    !txtAdvertiestment1 = txtAdvertiestment1.Text
'    !AdvertiestmentFontName = PreferanceAdvertiestmentFontName
'    !AdvertiestmentFontSize = PreferanceAdvertiestmentFontSize
'    !AdvertiestmentFontBold = PreferanceAdvertiestmentFontBold
'    !AdvertiestmentFontItalic = PreferanceAdvertiestmentFontItalic
'
'    If chkAdvertiestment2.Value = 1 Then
'        !chkAdvertiestment2 = True
'    Else
'        !chkAdvertiestment2 = False
'    End If
'
'    !txtAdvertiestment2 = txtAdvertiestment2.Text
'
'    If chkDoctorMLT1.Value = 1 Then
'        !chkDoctorMLT1 = True
'    Else
'        !chkDoctorMLT1 = False
'    End If
'
'    !txtDoctorMLT1 = txtDoctorMLT1.Text
'    !DoctorMLTFontName = PreferanceDoctorMLTFontName
'    !DoctorMLTFontSize = PreferanceDoctorMLTFontSize
'    !DoctorMLTFontBold = PreferanceDoctorMLTFontBold
'    !DoctorMLTFontItalic = PreferanceDoctorMLTFontItalic
'
'    .Update
'
'
'    If chkDoctorMLT2.Value = 1 Then
'        !chkDoctorMLT2 = True
'    Else
'        !chkDoctorMLT2 = False
'    End If
'
'    !txtDoctorMLT2 = txtDoctorMLT2
'
'    If chkDoctorMLT3.Value = 1 Then
'        !chkDoctorMLT3 = True
'    Else
'        !chkDoctorMLT3 = False
'    End If
'
'    !txtDoctorMLT3 = txtDoctorMLT3
'
'    If chkDoctorMLT4.Value = 1 Then
'        !chkDoctorMLT4 = True
'    Else
'        !chkDoctorMLT4 = False
'    End If
'
'
'
'    !txtDoctorMLT4 = txtDoctorMLT4.Text
'
'
'    If chkMessage.Value = 1 Then
'        !Messagechk = True
'    Else
'        !Messagechk = False
'    End If
'
'
'    !Messagetxt = txtMessage1.Text
'    !MessageFontName = PreferanceMessageFontName
'    !MessageFontsize = PreferanceMessageFontSize
'    !MessageFontbold = PreferanceMessageFontBold
'    !MessageFontitalic = PreferanceMessageFontItalic
'
'    !InstitutionLogoFileName = PreferanceInstitutionLogoFileName
'
'    .Update
'
'
'    If chkLblPatientName.Value = 1 Then
'        !PreferancechkLblPatientName = True
'    Else
'        !PreferancechkLblPatientName = False
'    End If
'
'    If chkLblPatientAge.Value = 1 Then
'        !PreferancechkLblPatientAge = True
'    Else
'        !PreferancechkLblPatientAge = False
'    End If
'
'    If chkLblPatientSex.Value = 1 Then
'        !PreferancechkLblPatientSex = True
'    Else
'        !PreferancechkLblPatientSex = False
'    End If
'
'    If chkLblPatientID.Value = 1 Then
'        !PreferancechkLblPatientID = True
'    Else
'        !PreferancechkLblPatientID = False
'    End If
'
'    If chkLblDate.Value = 1 Then
'        !PreferancechkLblDate = True
'    Else
'        !PreferancechkLblDate = False
'    End If
'
'    If chkLblTime.Value = 1 Then
'        !PreferancechkLblTime = True
'    Else
'        !PreferancechkLblTime = False
'    End If
'
'    If chkLblRDoctorName.Value = 1 Then
'        !PreferancechkLblRDoctorName = True
'    Else
'        !PreferancechkLblRDoctorName = False
'    End If
'
'    If chkLblRInstitutionName.Value = 1 Then
'        !PreferancechkLblRInstitutionName = True
'    Else
'        !PreferancechkLblRInstitutionName = False
'    End If
'
'
'
'    If chkPatientName.Value = 1 Then
'        !PreferancechkPatientName = True
'    Else
'        !PreferancechkPatientName = False
'    End If
'
'    If chkPatientAge.Value = 1 Then
'        !PreferancechkPatientAge = True
'    Else
'        !PreferancechkPatientAge = False
'    End If
'
'    If chkPatientSex.Value = 1 Then
'        !PreferancechkPatientSex = True
'    Else
'        !PreferancechkPatientSex = False
'    End If
'
'    If chkPatientID.Value = 1 Then
'        !PreferancechkPatientID = True
'    Else
'        !PreferancechkPatientID = False
'    End If
'
'    If chkDate.Value = 1 Then
'        !PreferancechkDate = True
'    Else
'        !PreferancechkDate = False
'    End If
'
'    If chkTime.Value = 1 Then
'        !PreferancechkTime = True
'    Else
'        !PreferancechkTime = False
'    End If
'
'    If chkRDoctorName.Value = 1 Then
'        !PreferancechkRDoctorName = True
'    Else
'        !PreferancechkRDoctorName = False
'    End If
'
'    If chkRInstitutionName.Value = 1 Then
'        !PreferancechkRInstitutionName = True
'    Else
'        !PreferancechkRInstitutionName = False
'    End If
'
'    !PreferancetxtLblPatientName = txtLblPatientName.Text
'    !PreferancetxtLblPatientAge = txtLblPatientAge.Text
'    !PreferancetxtLblPatientID = txtLblPatientID.Text
'    !PreferancetxtLblPatientSex = txtLblPatientSex.Text
'    !PreferancetxtLblDate = txtLblDate.Text
'    !PreferancetxtLblTime = txtLblTime.Text
'    !PreferancetxtLblRDoctorName = txtLblRDoctorName.Text
'    !PreferancetxtLblRInstitutionName = txtLblRInstitutionName.Text
'
'    !PreferanceTextFontName = PreferanceTextFontName
'    !PreferanceTextFontSize = PreferanceTextFontSize
'    !PreferanceTextFontBold = PreferanceTextFontBold
'    !PreferanceTextFontItalic = PreferanceTextFontItalic
'    !PreferanceLabelFontName = PreferanceLabelFontName
'    !PreferanceLabelFontSize = PreferanceLabelFontSize
'    !PreferanceLabelFontBold = PreferanceLabelFontBold
'    !PreferanceLabelFontItalic = PreferanceLabelFontItalic
'
'
'
'    .Update
'
'
'
'    If chkLblTest.Value = 1 Then
'        !PreferanceChkLblTest = True
'    Else
'        !PreferanceChkLblTest = False
'    End If
'
'    If chkLblSpeciman.Value = 1 Then
'        !PreferanceChkLblSpeciman = True
'    Else
'        !PreferanceChkLblSpeciman = False
'    End If
'
'    If chkLblSpecimanNo.Value = 1 Then
'        !PreferanceChkLblSpecimanNo = True
'    Else
'        !PreferanceChkLblSpecimanNo = False
'    End If
'
'    If chkTest.Value = 1 Then
'        !PreferanceChkTest = True
'    Else
'        !PreferanceChkTest = False
'    End If
'
'    If chkSpeciman.Value = 1 Then
'        !PreferanceChkSpeciman = True
'    Else
'        !PreferanceChkSpeciman = False
'    End If
'
'    If chkSpecimanNo.Value = 1 Then
'        !PreferanceChkSpecimanNo = True
'    Else
'        !PreferanceChkSpecimanNo = False
'    End If
'
'    If chkLblParameters.Value = 1 Then
'        !PreferanceChkLblParameters = True
'    Else
'        !PreferanceChkLblParameters = False
'    End If
'
'    If chkLblResults.Value = 1 Then
'        !PreferanceChkLblResults = True
'    Else
'        !PreferanceChkLblResults = False
'    End If
'
'    If chkLblResults.Value = 1 Then
'        !PreferanceChkLblResults = True
'    Else
'        !PreferanceChkLblResults = False
'    End If
'
'    If chkLblUnits.Value = 1 Then
'        !PreferanceChkLblUnits = True
'    Else
'        !PreferanceChkLblUnits = False
'    End If
'
'    If chkLblReferances.Value = 1 Then
'        !PreferanceChkLblReferances = True
'    Else
'        !PreferanceChkLblReferances = False
'    End If
'
'
'    If chkParameters.Value = 1 Then
'        !PreferanceChkParameters = True
'    Else
'        !PreferanceChkParameters = False
'    End If
'
'    If chkResults.Value = 1 Then
'        !PreferanceChkResults = True
'    Else
'        !PreferanceChkResults = False
'    End If
'
'    If chkResults.Value = 1 Then
'        !PreferanceChkResults = True
'    Else
'        !PreferanceChkResults = False
'    End If
'
'    If chkUnits.Value = 1 Then
'        !PreferanceChkUnits = True
'    Else
'        !PreferanceChkUnits = False
'    End If
'
'    If chkReferances.Value = 1 Then
'        !PreferanceChkReferances = True
'    Else
'        !PreferanceChkReferances = False
'    End If
'
'    If chkLblComments.Value = 1 Then
'        !PreferanceChkLblComments = True
'    Else
'        !PreferanceChkLblComments = False
'    End If
'
'    If chkComments.Value = 1 Then
'        !PreferanceChkComments = True
'    Else
'        !PreferanceChkComments = False
'    End If
'
'
'    .Update
'
'    !PreferanceTxtLblTest = txtLblTest.Text
'    !PreferanceTxtLblSpeciman = txtLblSpeciman.Text
'    !PreferanceTxtLblSpecimanNo = txtLblSpecimanNo.Text
'    !PreferanceTxtLblParameters = txtLblParameters.Text
'    !PreferanceTxtLblResults = txtLblResults.Text
'    !PreferanceTxtLblUnits = txtLblUnits.Text
'    !PreferanceTxtLblReferances = txtLblReferances.Text
'    !PreferanceTxtLblComments = txtLblComments.Text
'
'    !TopicFontName = TopicFontName
'    !TopicFontSize = TopicFontSize
'    !TopicFontBold = TopicFontBold
'    !TopicFontItalic = TopicFontItalic
'    !ValueFontName = ValueFontName
'    !ValueFontSize = ValueFontSize
'    !ValueFontBold = ValueFontBold
'    !ValueFontItalic = ValueFontItalic
'
'    .Update
'
'
'    !LblConfidentialFontSize = LblConfidentialFontSize
'    !LblConfidentialFontName = LblConfidentialFontName
'    '!LbltxtLblReport = LbltxtLblReport
'    !LblReportFontName = LblReportFontName
'    !LblReportFontSize = LblReportFontSize
'
'    .Update
'
'End With
'
'End Sub
'
'
'Private Sub txtMessageDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtMessageDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'
'End Sub
'
'
'Private Sub txtResultsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtResultsDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
'
'Private Sub txtUnitsDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
'Set temsource = txtUnitsDisplay
'If KeyCode = vbKeyLeft Then temsource.Left = temsource.Left - 10
'If KeyCode = vbKeyRight Then temsource.Left = temsource.Left + 10
'If KeyCode = vbKeyUp Then temsource.Top = temsource.Top - 10
'If KeyCode = vbKeyDown Then temsource.Top = temsource.Top + 10
'If KeyCode = vbKeyPageUp Then temsource.Height = temsource.Height + 10
'If KeyCode = vbKeyPageDown Then temsource.Height = temsource.Height - 10
'If KeyCode = vbKeyHome Then temsource.Width = temsource.Width + 10
'If KeyCode = vbKeyEnd Then temsource.Width = temsource.Width - 10
'KeyCode = 0
'End Sub
