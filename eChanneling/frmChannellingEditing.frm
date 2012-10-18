VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChannellingEditing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Channeling Scheduling"
   ClientHeight    =   8250
   ClientLeft      =   300
   ClientTop       =   75
   ClientWidth     =   12150
   Icon            =   "frmChannellingEditing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   12150
   Begin VB.ListBox ListConsultantIDs 
      Height          =   1035
      Left            =   3600
      TabIndex        =   43
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox ListConsultants 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   240
      TabIndex        =   42
      Top             =   1440
      Width           =   3975
   End
   Begin VB.ListBox ListSpecialityIDs 
      Height          =   1035
      Left            =   3000
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox ListSpecialities 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   240
      TabIndex        =   40
      Top             =   120
      Width           =   3975
   End
   Begin VB.ListBox ListSecessionIDs 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2640
      TabIndex        =   33
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   10680
      TabIndex        =   14
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Close"
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
   Begin VB.Frame FrameSecessions 
      Height          =   7575
      Left            =   4320
      TabIndex        =   22
      Top             =   120
      Width           =   7695
      Begin TabDlg.SSTab SSTab2 
         Height          =   1935
         Left            =   2160
         TabIndex        =   44
         Top             =   2760
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3413
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Normal"
         TabPicture(0)   =   "frmChannellingEditing.frx":0442
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtAgentOtherFee"
         Tab(0).Control(1)=   "txtForeignOtherFee"
         Tab(0).Control(2)=   "txtLocalOtherFee"
         Tab(0).Control(3)=   "txtAgentHospitalFee"
         Tab(0).Control(4)=   "txtAgentDoctorFee"
         Tab(0).Control(5)=   "txtFogrignerHospitalFee"
         Tab(0).Control(6)=   "txtForeginerDoctorFee"
         Tab(0).Control(7)=   "txtLocalHospitalFee"
         Tab(0).Control(8)=   "txtLocalDoctorFee"
         Tab(0).Control(9)=   "Label6"
         Tab(0).Control(10)=   "Label36"
         Tab(0).Control(11)=   "Label35"
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "With Scanning"
         TabPicture(1)   =   "frmChannellingEditing.frx":045E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label7"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label8"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label9"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "txtSLocalDoctorFee"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "txtSLocalHospitalFee"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "txtSForeginerDoctorFee"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "txtSFogrignerHospitalFee"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txtSAgentDoctorFee"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "txtSAgentHospitalFee"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "txtSLocalOtherFee"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "txtSForeignOtherFee"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "txtSAgentOtherFee"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).ControlCount=   12
         Begin VB.TextBox txtSAgentOtherFee 
            Height          =   360
            Left            =   3600
            MaxLength       =   250
            TabIndex        =   65
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtSForeignOtherFee 
            Height          =   360
            Left            =   3600
            MaxLength       =   250
            TabIndex        =   64
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtSLocalOtherFee 
            Height          =   360
            Left            =   3600
            MaxLength       =   250
            TabIndex        =   63
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtSAgentHospitalFee 
            Height          =   360
            Left            =   1920
            MaxLength       =   250
            TabIndex        =   62
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtSAgentDoctorFee 
            Height          =   360
            Left            =   120
            MaxLength       =   250
            TabIndex        =   61
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtSFogrignerHospitalFee 
            Height          =   360
            Left            =   1920
            MaxLength       =   250
            TabIndex        =   60
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtSForeginerDoctorFee 
            Height          =   360
            Left            =   120
            MaxLength       =   250
            TabIndex        =   59
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtSLocalHospitalFee 
            Height          =   360
            Left            =   1920
            MaxLength       =   250
            TabIndex        =   58
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtSLocalDoctorFee 
            Height          =   360
            Left            =   120
            MaxLength       =   250
            TabIndex        =   57
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtAgentOtherFee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -71400
            MaxLength       =   250
            TabIndex        =   55
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtForeignOtherFee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -71400
            MaxLength       =   250
            TabIndex        =   54
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtLocalOtherFee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -71400
            MaxLength       =   250
            TabIndex        =   53
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtAgentHospitalFee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -73080
            MaxLength       =   250
            TabIndex        =   50
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtAgentDoctorFee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -74880
            MaxLength       =   250
            TabIndex        =   49
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtFogrignerHospitalFee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -73080
            MaxLength       =   250
            TabIndex        =   48
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtForeginerDoctorFee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -74880
            MaxLength       =   250
            TabIndex        =   47
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtLocalHospitalFee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -73080
            MaxLength       =   250
            TabIndex        =   46
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtLocalDoctorFee 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -74880
            MaxLength       =   250
            TabIndex        =   45
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Other Fee"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   68
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hospital Fee"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   67
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor Fee"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Other Fee"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -71400
            TabIndex        =   56
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hospital Fee"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -73080
            TabIndex        =   52
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor Fee"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74880
            TabIndex        =   51
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.TextBox txtRoomNo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   38
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox txtComments 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   2160
         TabIndex        =   10
         Top             =   5880
         Width           =   5055
      End
      Begin VB.ComboBox cmbSecession 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmChannellingEditing.frx":047A
         Left            =   2160
         List            =   "frmChannellingEditing.frx":0484
         TabIndex        =   6
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtMaximum 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox chkBypassOrder 
         Caption         =   "Can bypass order"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   7200
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox ChkCalculateTime 
         Caption         =   "Calculate Appointment Time"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   6960
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtUsualDuration 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2160
         MaxLength       =   250
         TabIndex        =   9
         Top             =   4920
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   360
         Left            =   2160
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   66519042
         CurrentDate     =   39401
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   7080
         Width           =   975
         _ExtentX        =   1720
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   7080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Change"
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
         Left            =   4920
         TabIndex        =   13
         Top             =   7080
         Width           =   975
         _ExtentX        =   1720
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
         BackStyle       =   0  'Transparent
         Caption         =   "Room No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   5400
         Width           =   4335
      End
      Begin VB.Label lblSecession 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Doctor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1410
         TabIndex        =   37
         Top             =   600
         Width           =   4275
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSecessionName 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblDoctor 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Doctor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1410
         TabIndex        =   35
         Top             =   240
         Width           =   4275
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Secession Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Usual Duration                                             minutes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   4920
         Width           =   5295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   5880
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Bookings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Local Patients"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Foreginers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3720
         Width           =   1095
      End
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   7200
      Width           =   975
      _ExtentX        =   1720
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
   Begin VB.ListBox ListSecessions 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   360
      TabIndex        =   2
      Top             =   6120
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Weekdays"
      TabPicture(0)   =   "frmChannellingEditing.frx":049A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "OptionSunday"
      Tab(0).Control(1)=   "OptionSaturday"
      Tab(0).Control(2)=   "OptionFriday"
      Tab(0).Control(3)=   "OptionThursday"
      Tab(0).Control(4)=   "OptionWednesday"
      Tab(0).Control(5)=   "OptionTuesday"
      Tab(0).Control(6)=   "OptionMonday"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Other Days"
      TabPicture(1)   =   "frmChannellingEditing.frx":04B6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "MonthView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.OptionButton OptionSunday 
         Caption         =   "Sunday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   21
         Top             =   2640
         Width           =   2415
      End
      Begin VB.OptionButton OptionSaturday 
         Caption         =   "Saturday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   20
         Top             =   2280
         Width           =   2415
      End
      Begin VB.OptionButton OptionFriday 
         Caption         =   "Friday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   19
         Top             =   1920
         Width           =   2415
      End
      Begin VB.OptionButton OptionThursday 
         Caption         =   "Thursday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   18
         Top             =   1560
         Width           =   2415
      End
      Begin VB.OptionButton OptionWednesday 
         Caption         =   "Wednesday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   17
         Top             =   1200
         Width           =   2415
      End
      Begin VB.OptionButton OptionTuesday 
         Caption         =   "Tuesday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   16
         Top             =   840
         Width           =   2415
      End
      Begin VB.OptionButton OptionMonday 
         Caption         =   "Monday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   15
         Top             =   480
         Value           =   -1  'True
         Width           =   2415
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2820
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   66519041
         CurrentDate     =   39461
      End
   End
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   7200
      Width           =   975
      _ExtentX        =   1720
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   7200
      Width           =   975
      _ExtentX        =   1720
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
End
Attribute VB_Name = "frmChannellingEditing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemWeekday As Long

Private Sub Setcolours()
    bttnAdd.BackColor = BttnBackColour
    bttnAdd.ForeColor = BttnForeColour
    bttnCancel.BackColor = BttnBackColour
    bttnCancel.ForeColor = BttnForeColour
    bttnChange.BackColor = BttnBackColour
    bttnChange.ForeColor = BttnForeColour
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
    bttnEdit.BackColor = BttnBackColour
    bttnEdit.ForeColor = BttnForeColour
    bttnSave.BackColor = BttnBackColour
    bttnSave.ForeColor = BttnForeColour
    bttnDelete.BackColor = BttnBackColour
    bttnDelete.ForeColor = BttnForeColour
    bttnChange.BackColor = BttnBackColour
    bttnChange.ForeColor = BttnForeColour
    bttnDelete.BackColor = BttnBackColour
    bttnDelete.ForeColor = BttnForeColour
    frmChannellingEditing.BackColor = FrameBackColour
    frmChannellingEditing.ForeColor = FrameForeColour
    FrameSecessions.BackColor = FrameBackColour
    FrameSecessions.ForeColor = FrameForeColour
    chkBypassOrder.BackColor = LblBackColour
    chkBypassOrder.ForeColor = LblForeColour
    ChkCalculateTime.BackColor = LblBackColour
    ChkCalculateTime.ForeColor = LblForeColour
    Label1.BackColor = LblBackColour
    Label1.ForeColor = LblForeColour
    Label10.BackColor = LblBackColour
    Label10.ForeColor = LblForeColour
    Label16.BackColor = LblBackColour
    Label16.ForeColor = LblForeColour
    Label2.BackColor = LblBackColour
    Label2.ForeColor = LblForeColour
    Label3.BackColor = LblBackColour
    Label3.ForeColor = LblForeColour
    Label4.BackColor = LblBackColour
    Label4.ForeColor = LblForeColour
    Label4.BackColor = LblBackColour
    Label4.ForeColor = LblForeColour
    Label5.BackColor = LblBackColour
    Label5.ForeColor = LblForeColour
End Sub

Private Sub BeforeAddEdit()
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    bttnDelete.Enabled = True
    ListSecessions.Enabled = True
    SSTab1.Enabled = True
        
    ListSecessions.Enabled = True
    ListSpecialities.Enabled = True
    ListConsultants.Enabled = True
    
    FrameSecessions.Enabled = False
    
    bttnChange.Visible = False
    bttnSave.Visible = False
    bttnCancel.Visible = False
End Sub

Private Sub AfterAdd()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    ListSecessions.Enabled = False
    SSTab1.Enabled = False
    
    ListSecessions.Enabled = False
    ListSpecialities.Enabled = False
    ListConsultants.Enabled = False
    
    FrameSecessions.Enabled = True
    
    bttnChange.Visible = False
    bttnSave.Visible = True
    bttnCancel.Visible = True
End Sub

Private Sub AfterEdit()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    ListSecessions.Enabled = False
    SSTab1.Enabled = False
    
    ListSecessions.Enabled = False
    ListSpecialities.Enabled = False
    ListConsultants.Enabled = False
    
    FrameSecessions.Enabled = True
    
    bttnChange.Visible = True
    bttnSave.Visible = False
    bttnCancel.Visible = True
End Sub


Private Sub AfterDelete()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    ListSecessions.Enabled = False
    SSTab1.Enabled = False
    
    
    
    FrameSecessions.Enabled = False
    
    bttnChange.Visible = True
    bttnSave.Visible = False
    bttnCancel.Visible = True
End Sub

Private Sub SearchSecessions()
If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
ListSecessionIDs.Clear
ListSecessions.Clear
    With DataEnvironment1.rssqlTem4
        If SSTab1.Tab = 1 Then
                If .State = 1 Then .Close
                .Source = "Select * from tblfacilitysecession where hospitalfacility_ID =  10  and staff_ID = " & ListConsultantIDs.Text & " and AlteredDate = '" & MonthView1.Value & "' order by StartingTime"
                .Open
                    If .RecordCount <> 0 Then
                        While .EOF = False
                            ListSecessionIDs.AddItem !facilitysecession_ID
                            ListSecessions.AddItem FindSecessionFromID(!facilitysecession_ID)
                            .MoveNext
                        Wend
                    End If
                    .Close
        Else
                If .State = 1 Then .Close
                .Source = "Select * from tblfacilitysecession where hospitalfacility_ID = 10 and staff_ID = " & ListConsultantIDs.Text & " and SecessionWeekday = " & TemWeekday & " order by StartingTime"
                .Open
                If .RecordCount <> 0 Then
                    While .EOF = False
                        ListSecessionIDs.AddItem !facilitysecession_ID
                        ListSecessions.AddItem FindSecessionFromID(!facilitysecession_ID)
                        .MoveNext
                    Wend
                End If
                .Close
        End If
    End With
End Sub

Private Sub LocateSecession()
On Error Resume Next
If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
If Not IsNumeric(ListSecessionIDs.Text) Then Exit Sub
With DataEnvironment1.rssqlTem15
    If .State = 1 Then .Close
    .Source = "SELECT * from tblfacilitysecession where facilitysecession_ID = " & ListSecessionIDs.Text
    .Open
    If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!SecessionName) Then
            cmbSecession.Text = !SecessionName
            cmbSecession.Text = !SecessionName
        End If
    lblDoctor.Caption = ListConsultants.Text
    If SSTab1.Tab = 0 Then
        lblSecessionName.Caption = "Weekday"
        lblSecession.Caption = WeekdayName(TemWeekday)
    Else
        lblSecessionName.Caption = "Date"
        lblSecession.Caption = Format(MonthView1.Value, DefaultLongDate)
    End If
        If Not IsNull(!startingtime) Then dtpStart.Value = !startingtime
        If Not IsNull(!usualduration) Then txtUsualDuration.Text = !usualduration
        If Not IsNull(!Maximum) Then txtMaximum.Text = !Maximum
        
        If Not IsNull(!LocaldoctorFee) Then txtLocalDoctorFee.Text = Format(!LocaldoctorFee, "0.00")
        If Not IsNull(!LocalHospitalFee) Then txtLocalHospitalFee.Text = Format(!LocalHospitalFee, "0.00")
        If Not IsNull(!LocalOtherFee) Then txtLocalOtherFee.Text = Format(!LocalOtherFee, "0.00")
        
        If Not IsNull(!Foreigndoctorfee) Then txtForeginerDoctorFee.Text = Format(!Foreigndoctorfee, "0.00")
        If Not IsNull(!ForeignHospitalFee) Then txtFogrignerHospitalFee.Text = Format(!ForeignHospitalFee, "0.00")
        If Not IsNull(!ForeignOtherFee) Then txtForeignOtherFee.Text = Format(!ForeignOtherFee, "0.00")
        
        If Not IsNull(!AgentDoctorFee) Then txtAgentDoctorFee.Text = Format(!AgentDoctorFee, "0.00")
        If Not IsNull(!AgentHospitalFee) Then txtAgentHospitalFee.Text = Format(!AgentHospitalFee, "0.00")
        If Not IsNull(!AgentOtherFee) Then txtAgentOtherFee.Text = Format(!AgentOtherFee, "0.00")
        
        If Not IsNull(!SLocaldoctorFee) Then txtSLocalDoctorFee.Text = Format(!SLocaldoctorFee, "0.00")
        If Not IsNull(!SLocalHospitalFee) Then txtSLocalHospitalFee.Text = Format(!SLocalHospitalFee, "0.00")
        If Not IsNull(!SLocalOtherFee) Then txtSLocalOtherFee.Text = Format(!SLocalOtherFee, "0.00")
        
        If Not IsNull(!SForeigndoctorfee) Then txtSForeginerDoctorFee.Text = Format(!SForeigndoctorfee, "0.00")
        If Not IsNull(!SForeignHospitalFee) Then txtSFogrignerHospitalFee.Text = Format(!SForeignHospitalFee, "0.00")
        If Not IsNull(!SForeignOtherFee) Then txtSForeignOtherFee.Text = Format(!SForeignOtherFee, "0.00")
        
        If Not IsNull(!SAgentDoctorFee) Then txtSAgentDoctorFee.Text = Format(!SAgentDoctorFee, "0.00")
        If Not IsNull(!SAgentHospitalFee) Then txtSAgentHospitalFee.Text = Format(!SAgentHospitalFee, "0.00")
        If Not IsNull(!SAgentOtherFee) Then txtSAgentOtherFee.Text = Format(!SAgentOtherFee, "0.00")
        
        
        If !CanByPassOrder = True Then chkBypassOrder.Value = 1
        If !calculateappointment = True Then ChkCalculateTime.Value = 1
        If Not IsNull(!Comments) Then txtComments.Text = !Comments
        If Not IsNull(!roomno) Then txtRoomNo.Text = !roomno
        .Close
End With
End Sub

Private Sub bttnAdd_Click()
    Dim TemResponce As Long
    If Not IsNumeric(ListConsultantIDs.Text) Then
        TemResponce = MsgBox("You have not selected a doctor to add the channeling details", vbCritical, "No Doctor")
        ListConsultants.SetFocus
        Exit Sub
    End If
    lblDoctor.Caption = ListConsultants.Text
    If SSTab1.Tab = 0 Then
        lblSecessionName.Caption = "Weekday"
        lblSecession.Caption = WeekdayName(TemWeekday)
    Else
        lblSecessionName.Caption = "Date"
        lblSecession.Caption = Format(MonthView1.Value, DefaultLongDate)
    End If
    Call AfterAdd
End Sub

Private Sub bttnAdd_KeyPress(ByVal KeyAscii As Integer)
    If KeyAscii = 13 Then cmbSecession.SetFocus
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
End Sub

Private Sub bttnChange_Click()
    If CanAdd = False Then Exit Sub
    With DataEnvironment1.rssqlTem3
        If .State = 1 Then .Close
        .Source = "Select * from tblfacilitysecession where facilitysecession_ID = " & ListSecessionIDs.Text
        .Open
        If .RecordCount = 0 Then
            .Close
            ClearValues
            BeforeAddEdit
            Exit Sub
        End If
                    !Staff_ID = ListConsultantIDs.Text
                    !HospitalFacility_ID = 10
                    !SecessionName = cmbSecession.Text
                    
                    If SSTab1.Tab = 0 Then
                        !SecessionWeekday = TemWeekday
                    Else
                        !fulldayleave = False
                        !altereddate = MonthView1.Value
                    End If
                    !startingtime = TimeSerial(Hour(dtpStart.Value), Minute(dtpStart.Value), 0)
                    !usualduration = Val(txtUsualDuration.Text)
                    !Maximum = Val(txtMaximum.Text)
                    
                    !LocaldoctorFee = Val(txtLocalDoctorFee.Text)
                    !LocalHospitalFee = Val(txtLocalHospitalFee.Text)
                    !LocalOtherFee = Val(txtLocalOtherFee.Text)
                    
                    !Foreigndoctorfee = Val(txtForeginerDoctorFee.Text)
                    !ForeignHospitalFee = Val(txtFogrignerHospitalFee.Text)
                    !ForeignOtherFee = Val(txtForeignOtherFee.Text)
                    
                    !AgentDoctorFee = Val(txtAgentDoctorFee.Text)
                    !AgentHospitalFee = Val(txtAgentHospitalFee.Text)
                    !AgentOtherFee = Val(txtAgentOtherFee.Text)
                    
                    !SLocaldoctorFee = Val(txtSLocalDoctorFee.Text)
                    !SLocalHospitalFee = Val(txtSLocalHospitalFee.Text)
                    !SLocalOtherFee = Val(txtSLocalOtherFee.Text)
                    
                    !SForeigndoctorfee = Val(txtSForeginerDoctorFee.Text)
                    !SForeignHospitalFee = Val(txtSFogrignerHospitalFee.Text)
                    !SForeignOtherFee = Val(txtSForeignOtherFee.Text)
                    
                    !SAgentDoctorFee = Val(txtSAgentDoctorFee.Text)
                    !SAgentHospitalFee = Val(txtSAgentHospitalFee.Text)
                    !SAgentOtherFee = Val(txtSAgentOtherFee.Text)
                    
                    If chkBypassOrder.Value = 1 Then
                        !CanByPassOrder = True
                    Else
                        !CanByPassOrder = False
                    End If
                    If ChkCalculateTime.Value = 1 Then
                        !calculateappointment = True
                    Else
                        !calculateappointment = False
                    End If
                    !Comments = txtComments.Text
                    !roomno = txtRoomNo.Text
                    .Update
                    .Close
    End With
    BeforeAddEdit
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDelete_Click()
Dim TemResponce As Long
    If Not IsNumeric(ListConsultantIDs.Text) Then
        TemResponce = MsgBox("You have not selected a doctor to delete", vbCritical, "No Doctor")
        ListConsultants.SetFocus
        Exit Sub
    End If
    If ListSecessions.ListIndex < 0 Then
        TemResponce = MsgBox("You have not selected a secession to delete", vbCritical, "No Doctor")
        ListSecessions.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(ListSecessionIDs.Text) Then Exit Sub
    
TemResponce = MsgBox("Are you sure you want to delete this secession", vbQuestion + vbYesNo, "Delete?")
If TemResponce = vbNo Then Exit Sub
With DataEnvironment1.rssqlTem15
    If .State = 1 Then .Close
    .Source = "Select * from tblfacilitysecession where facilitysecession_ID = " & ListSecessionIDs.Text
    .Open
    If .RecordCount = 0 Then
        TemResponce = MsgBox("The selected secession is not available to delete", vbCritical, "No Doctor")
        ListSecessions.SetFocus
        Exit Sub
    End If
    .Delete adAffectCurrent
    .Close
End With
Call SearchSecessions

End Sub

Private Sub bttnEdit_Click()
    Dim TemResponce As Long
    If Not IsNumeric(ListConsultantIDs.Text) Then
        TemResponce = MsgBox("You have not selected a doctor to add the channeling details", vbCritical, "No Doctor")
         ListConsultants.SetFocus
        Exit Sub
    End If
    If ListSecessions.ListIndex < 0 Then
        TemResponce = MsgBox("You have not selected a secession to edit the channeling details", vbCritical, "No Doctor")
        ListSecessions.SetFocus
        Exit Sub
    End If
    
    lblDoctor.Caption = ListConsultants.Text
    If SSTab1.Tab = 0 Then
        lblSecessionName.Caption = "Weekday"
        lblSecession.Caption = WeekdayName(TemWeekday)
    Else
        lblSecessionName.Caption = "Date"
        lblSecession.Caption = Format(MonthView1.Value, DefaultLongDate)
    End If
    Call AfterEdit

End Sub

Private Sub bttnSave_Click()
    If CanAdd = False Then Exit Sub

    With DataEnvironment1.rssqlTem13
        If .State = 1 Then .Close
        .Source = "Select * from tblfacilitysecession"
        .Open
        .AddNew
                    !Staff_ID = ListConsultantIDs.Text
                    !HospitalFacility_ID = 10
                    !SecessionName = cmbSecession.Text
                    If SSTab1.Tab = 0 Then
                        !SecessionWeekday = TemWeekday
                    Else
                        !fulldayleave = False
                        !altereddate = MonthView1.Value
                    End If
                    !startingtime = TimeSerial(Hour(dtpStart.Value), Minute(dtpStart.Value), 0)
                    !usualduration = Val(txtUsualDuration.Text)
                    !Maximum = Val(txtMaximum.Text)
                    
                    !LocaldoctorFee = Val(txtLocalDoctorFee.Text)
                    !LocalHospitalFee = Val(txtLocalHospitalFee.Text)
                    !LocalOtherFee = Val(txtLocalOtherFee.Text)
                    
                    !Foreigndoctorfee = Val(txtForeginerDoctorFee.Text)
                    !ForeignHospitalFee = Val(txtFogrignerHospitalFee.Text)
                    !ForeignOtherFee = Val(txtForeignOtherFee.Text)
                    
                    !AgentDoctorFee = Val(txtAgentDoctorFee.Text)
                    !AgentHospitalFee = Val(txtAgentHospitalFee.Text)
                    !AgentOtherFee = Val(txtAgentOtherFee.Text)
                    
                    !SLocaldoctorFee = Val(txtSLocalDoctorFee.Text)
                    !SLocalHospitalFee = Val(txtSLocalHospitalFee.Text)
                    !SLocalOtherFee = Val(txtSLocalOtherFee.Text)
                    
                    !SForeigndoctorfee = Val(txtSForeginerDoctorFee.Text)
                    !SForeignHospitalFee = Val(txtSFogrignerHospitalFee.Text)
                    !SForeignOtherFee = Val(txtSForeignOtherFee.Text)
                    
                    !SAgentDoctorFee = Val(txtSAgentDoctorFee.Text)
                    !SAgentHospitalFee = Val(txtSAgentHospitalFee.Text)
                    !SAgentOtherFee = Val(txtSAgentOtherFee.Text)
                    
                    
                    If chkBypassOrder.Value = 1 Then
                        !CanByPassOrder = True
                    Else
                        !CanByPassOrder = False
                    End If
                    If ChkCalculateTime.Value = 1 Then
                        !calculateappointment = True
                    Else
                        !calculateappointment = False
                    End If
                    !Comments = txtComments.Text
                    !roomno = txtRoomNo.Text
                    .Update
                    .Close
    End With
    BeforeAddEdit
End Sub

Private Sub bttnSave_KeyPress(ByVal KeyAscii As Integer)
If KeyAscii = 13 Then bttnClose.SetFocus

End Sub

Private Sub cmbSecession_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpStart.SetFocus
End Sub

Private Sub ClearValues()
    cmbSecession.Text = Empty
    dtpStart.Value = Date
    txtUsualDuration.Text = Empty
    txtMaximum.Text = Empty
    
    txtLocalDoctorFee.Text = Empty
    txtLocalHospitalFee.Text = Empty
    txtLocalOtherFee.Text = Empty
    
    txtForeginerDoctorFee.Text = Empty
    txtFogrignerHospitalFee.Text = Empty
    txtForeignOtherFee.Text = Empty
    
    txtAgentDoctorFee.Text = Empty
    txtAgentHospitalFee.Text = Empty
    txtAgentOtherFee.Text = Empty
    
    txtSLocalDoctorFee.Text = Empty
    txtSLocalHospitalFee.Text = Empty
    txtSLocalOtherFee.Text = Empty
    
    txtSForeginerDoctorFee.Text = Empty
    txtSFogrignerHospitalFee.Text = Empty
    txtSForeignOtherFee.Text = Empty
    
    txtSAgentDoctorFee.Text = Empty
    txtSAgentHospitalFee.Text = Empty
    txtSAgentOtherFee.Text = Empty
    
    
    chkBypassOrder.Value = 0
    ChkCalculateTime.Value = 0
    txtComments.Text = Empty
    txtRoomNo.Text = Empty
End Sub

Private Sub dtpStart_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtMaximum.SetFocus
End Sub

Private Sub Form_Load()
Call FillSpeciality
TemWeekday = vbMonday
OptionMonday.Value = True
BeforeAddEdit
SSTab1.Tab = 0
MonthView1.Value = Date
dtpStart.Value = TimeSerial(0, 0, 0)
Call Setcolours
End Sub

Private Sub FormatGridSpeciality()
    ListSpecialities.Clear
    ListSpecialityIDs.Clear
End Sub

Private Sub FormatGridConsultants()
    ListConsultants.Clear
    ListConsultantIDs.Clear
End Sub

Private Sub FillSpeciality()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblspeciality order by speciality "
    .Open
    If NoAllNames = False Then
        ListSpecialities.AddItem "All"
        ListSpecialityIDs.AddItem "All"
    End If
    If .RecordCount <> 0 Then
        While Not .EOF
            ListSpecialities.AddItem !Speciality
            ListSpecialityIDs.AddItem !speciality_ID
            .MoveNext
        Wend
    End If
    .Close
End With
End Sub


Private Sub ListAllConsultants()
Call FormatGridConsultants
With DataEnvironment1.rssqlTem1
    If .State = 1 Then .Close
    If SurnameFirst = True Then
        .Source = "SELECT  tbldoctor.*  FROM  tbldoctor  order by doctorlistedname"
    Else
        .Source = "SELECT  tbldoctor.*  FROM  tbldoctor  order by doctorname"
    End If
    .Open
    If .RecordCount = 0 Then Exit Sub
    While Not .EOF
            If SurnameFirst = True Then
                ListConsultants.AddItem !doctorlistedname
            Else
                ListConsultants.AddItem !doctorname
            End If
        ListConsultantIDs.AddItem !Doctor_ID
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub ListSelectedConsultants()
    Call FormatGridConsultants
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        If SurnameFirst = True Then
            .Source = "SELECT tbldoctor.* FROM tbldoctor where  doctorspeciality_ID = " & Val(ListSpecialityIDs.Text) & " order by doctorlistedname"
        Else
            .Source = "SELECT tbldoctor.* FROM tbldoctor where  doctorspeciality_ID = " & Val(ListSpecialityIDs.Text) & " order by doctorname"
        End If
        .Open
        If .RecordCount = 0 Then Exit Sub
        While Not .EOF
            
            If SurnameFirst = True Then
                ListConsultants.AddItem !doctorlistedname
            Else
                ListConsultants.AddItem !doctorname
            End If
            
            ListConsultantIDs.AddItem !Doctor_ID
            .MoveNext
        Wend
        .Close
    End With
End Sub


Private Sub ListConsultants_Click()
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    If ListConsultantIDs.ListIndex < 0 Then Exit Sub
    If ListConsultants.ListIndex < 0 Then Exit Sub
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    lblDoctor.Caption = Empty
    lblSecession.Caption = Empty
    Call ClearValues
    ListSecessionIDs.Clear
    ListSecessions.Clear
    Call SearchSecessions
End Sub

Private Sub ListSpecialities_Click()
    ListSpecialityIDs.ListIndex = ListSpecialities.ListIndex
    ListConsultantIDs.Clear
    ListConsultants.Clear
    If ListSpecialities.Text = "All" Then
        ListAllConsultants
    ElseIf ListSpecialities.Text <> "All" And IsNumeric(ListSpecialityIDs.Text) = True Then
        ListSelectedConsultants
    Else
        FormatGridConsultants
    End If
End Sub

Private Sub ListSpecialities_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    ListConsultants.SetFocus
    KeyCode = Empty
Else

End If
End Sub

Private Sub ListSecessions_Click()
    ListSecessionIDs.ListIndex = ListSecessions.ListIndex
    If Not IsNumeric(ListSecessionIDs.Text) Then Exit Sub
    Call LocateSecession
End Sub

Private Sub ListSecessions_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then bttnAdd.SetFocus
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    SearchSecessions
End Sub

Private Sub OptionMonday_Click()
    If OptionMonday.Value = True Then TemWeekday = vbMonday: SearchSecessions
End Sub

Private Sub OptionMonday_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ListSecessions.SetFocus
End Sub

Private Sub OptionTuesday_Click()
    If OptionTuesday.Value = True Then TemWeekday = vbTuesday: SearchSecessions
End Sub
Private Sub OptionWednesday_Click()
    If OptionWednesday.Value = True Then TemWeekday = vbWednesday: SearchSecessions
End Sub
Private Sub OptionThursday_Click()
    If OptionThursday.Value = True Then TemWeekday = vbThursday: SearchSecessions
End Sub
Private Sub OptionFriday_Click()
    If OptionFriday.Value = True Then TemWeekday = vbFriday: SearchSecessions
End Sub
Private Sub Optionsaturday_Click()
    If OptionSaturday.Value = True Then TemWeekday = vbSaturday: SearchSecessions
End Sub
Private Sub Optionsunday_Click()
    If OptionSunday.Value = True Then TemWeekday = vbSunday: SearchSecessions
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    ListSecessionIDs.Clear
    ListSecessions.Clear
    If ListConsultantIDs.ListIndex < 0 Or Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    Call LocateSecession
End Sub

Private Function CanAdd() As Boolean
    CanAdd = False
    Dim TemResponce As Integer
    If Not IsNumeric(ListConsultantIDs.Text) Then
        TemResponce = MsgBox("You have not selected a doctor", vbCritical, "Doctor?")
        ListConsultants.SetFocus
        Exit Function
    End If
    If dtpStart.Value = TimeSerial(0, 0, 0) Then
        TemResponce = MsgBox("You have not enterd an starting time for the secession", vbCritical, "Starting time?")
        dtpStart.SetFocus
        Exit Function
    End If
    If Trim(cmbSecession.Text) = "" Then
        TemResponce = MsgBox("You have not enterd a name for the secession", vbCritical, "Secession name?")
        cmbSecession.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtLocalDoctorFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Doctor fee for local patients", vbCritical, "No doctor charge")
        txtLocalDoctorFee.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtLocalHospitalFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Hospital fee for local patients", vbCritical, "No doctor charge")
        txtLocalHospitalFee.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtForeginerDoctorFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Doctor fee for foreign patients", vbCritical, "No doctor charge")
        txtForeginerDoctorFee.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtFogrignerHospitalFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Hospital fee for foreign patients", vbCritical, "No doctor charge")
        txtFogrignerHospitalFee.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtAgentDoctorFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Doctor fee for patients booking through agents", vbCritical, "No doctor charge")
        txtAgentDoctorFee.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtAgentHospitalFee.Text) Then
        TemResponce = MsgBox("You have not entered a valied Hospital fee for patients booking through agents", vbCritical, "No doctor charge")
        txtAgentHospitalFee.SetFocus
        Exit Function
    End If
    
    CanAdd = True
End Function

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ListSecessions.SetFocus

End Sub

Private Sub txtAgentDoctorFee_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAgentHospitalFee.SetFocus

End Sub

Private Sub txtAgentHospitalFee_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtUsualDuration.SetFocus

End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then bttnSave.SetFocus

End Sub

Private Sub txtFogrignerHospitalFee_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAgentDoctorFee.SetFocus

End Sub

Private Sub txtForeginerDoctorFee_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtFogrignerHospitalFee.SetFocus

End Sub

Private Sub txtLocalDoctorFee_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtLocalHospitalFee.SetFocus

End Sub

Private Sub txtLocalHospitalFee_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtForeginerDoctorFee.SetFocus

End Sub

Private Sub txtMaximum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtLocalDoctorFee.SetFocus

End Sub

Private Sub txtRoomNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtComments.SetFocus

End Sub

Private Sub txtUsualDuration_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRoomNo.SetFocus

End Sub
