VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChannelingDoctors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff & Facilities"
   ClientHeight    =   8625
   ClientLeft      =   3720
   ClientTop       =   1755
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChannelingDoctors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11295
   Begin btButtonEx.ButtonEx bttnChange 
      Height          =   495
      Left            =   5400
      TabIndex        =   86
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&hange"
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
      Left            =   7920
      TabIndex        =   88
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Ca&ncel"
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
      Height          =   495
      Left            =   5400
      TabIndex        =   87
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackStyle       =   1
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   4680
      TabIndex        =   91
      Top             =   120
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483630
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmChannelingDoctors.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "framFacility"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dates"
      TabPicture(1)   =   "frmChannelingDoctors.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTabDates"
      Tab(1).ControlCount=   1
      Begin VB.Frame framFacility 
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   120
         TabIndex        =   92
         Top             =   360
         Width           =   6135
         Begin VB.TextBox txtInstitutionForeignerFee 
            Height          =   375
            Left            =   2280
            TabIndex        =   154
            Top             =   3360
            Width           =   3735
         End
         Begin VB.TextBox txtUsualDuration 
            Height          =   360
            Left            =   2280
            MaxLength       =   250
            TabIndex        =   8
            Top             =   4200
            Width           =   855
         End
         Begin VB.TextBox txtDoctorForeignerFee 
            Height          =   375
            Left            =   2280
            TabIndex        =   7
            Top             =   1920
            Width           =   3735
         End
         Begin VB.TextBox txtComments 
            Height          =   1335
            Left            =   1920
            TabIndex        =   9
            Top             =   5640
            Width           =   4095
         End
         Begin VB.TextBox txtDoctorLocalFee 
            Height          =   375
            Left            =   2280
            TabIndex        =   5
            Top             =   1320
            Width           =   3735
         End
         Begin VB.TextBox txtInstitutionLocalFee 
            Height          =   375
            Left            =   2280
            TabIndex        =   6
            Top             =   2880
            Width           =   3735
         End
         Begin MSDataListLib.DataCombo DataComboDoctor 
            Bindings        =   "frmChannelingDoctors.frx":047A
            Height          =   360
            Left            =   1920
            TabIndex        =   4
            Top             =   360
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   "DoctorListedName"
            BoundColumn     =   "Doctor_ID"
            Text            =   ""
            Object.DataMember      =   "sqlDoctor"
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Local"
            Height          =   255
            Left            =   960
            TabIndex        =   153
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Forigner"
            Height          =   255
            Left            =   960
            TabIndex        =   152
            Top             =   3480
            Width           =   2055
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "Usual Duration                             minutes"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   4200
            Width           =   4935
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            Height          =   255
            Left            =   360
            TabIndex        =   98
            Top             =   5640
            Width           =   2055
         End
         Begin VB.Label LblDoctorStaff 
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor Name"
            Height          =   375
            Left            =   240
            TabIndex        =   97
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblDoctorStaffFee 
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor Fee"
            Height          =   375
            Left            =   240
            TabIndex        =   96
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Institution Fee"
            Height          =   255
            Left            =   240
            TabIndex        =   95
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Forigner"
            Height          =   255
            Left            =   960
            TabIndex        =   94
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label lblInstitutionFee 
            BackStyle       =   0  'Transparent
            Caption         =   "Local"
            Height          =   255
            Left            =   960
            TabIndex        =   93
            Top             =   1440
            Width           =   2055
         End
      End
      Begin TabDlg.SSTab SSTabDates 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   100
         Top             =   360
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   12515
         _Version        =   393216
         Tabs            =   8
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Monday"
         TabPicture(0)   =   "frmChannelingDoctors.frx":0499
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label37"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label36"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label35"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label34"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label33"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label16"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "dtpMondayEveningEnd"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "dtpMondayEveningStart"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "dtpMondayMorningEnd"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "dtpMondayMorningStart"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "chkMondayFullLeave"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtMondayMorningMax"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "chkMondayMorning"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkMondayEvening"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txtMondayEveningMax"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "Tuesday"
         TabPicture(1)   =   "frmChannelingDoctors.frx":04B5
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label20"
         Tab(1).Control(1)=   "Label18"
         Tab(1).Control(2)=   "Label15"
         Tab(1).Control(3)=   "Label14"
         Tab(1).Control(4)=   "Label13"
         Tab(1).Control(5)=   "Label12"
         Tab(1).Control(6)=   "dtpTuesdayEveningEnd"
         Tab(1).Control(7)=   "dtpTuesdayEveningStart"
         Tab(1).Control(8)=   "dtpTuesdayMorningEnd"
         Tab(1).Control(9)=   "dtpTuesdayMorningStart"
         Tab(1).Control(10)=   "txtTuesdayEveningMax"
         Tab(1).Control(11)=   "chkTuesdayEvening"
         Tab(1).Control(12)=   "chkTuesdayMorning"
         Tab(1).Control(13)=   "txtTuesdayMorningMax"
         Tab(1).Control(14)=   "chkTuesdayFullLeave"
         Tab(1).ControlCount=   15
         TabCaption(2)   =   "Wednesday"
         TabPicture(2)   =   "frmChannelingDoctors.frx":04D1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label26"
         Tab(2).Control(1)=   "Label25"
         Tab(2).Control(2)=   "Label24"
         Tab(2).Control(3)=   "Label23"
         Tab(2).Control(4)=   "Label22"
         Tab(2).Control(5)=   "Label21"
         Tab(2).Control(6)=   "dtpWednesdayEveningEnd"
         Tab(2).Control(7)=   "dtpWednesdayEveningStart"
         Tab(2).Control(8)=   "dtpWednesdayMorningEnd"
         Tab(2).Control(9)=   "dtpWednesdayMorningStart"
         Tab(2).Control(10)=   "txtWednesdayEveningMax"
         Tab(2).Control(11)=   "chkWednesdayEvening"
         Tab(2).Control(12)=   "chkWednesdayMorning"
         Tab(2).Control(13)=   "txtWednesdayMorningMax"
         Tab(2).Control(14)=   "chkWednesdayFullLeave"
         Tab(2).ControlCount=   15
         TabCaption(3)   =   "Thursday"
         TabPicture(3)   =   "frmChannelingDoctors.frx":04ED
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label32"
         Tab(3).Control(1)=   "Label31"
         Tab(3).Control(2)=   "Label30"
         Tab(3).Control(3)=   "Label29"
         Tab(3).Control(4)=   "Label28"
         Tab(3).Control(5)=   "Label27"
         Tab(3).Control(6)=   "dtpThursdayEveningEnd"
         Tab(3).Control(7)=   "dtpThursdayEveningStart"
         Tab(3).Control(8)=   "dtpThursdayMorningEnd"
         Tab(3).Control(9)=   "dtpThursdayMorningStart"
         Tab(3).Control(10)=   "txtThursdayEveningMax"
         Tab(3).Control(11)=   "chkThursdayEvening"
         Tab(3).Control(12)=   "chkThursdayMorning"
         Tab(3).Control(13)=   "txtThursdayMorningMax"
         Tab(3).Control(14)=   "chkThursdayFullLeave"
         Tab(3).ControlCount=   15
         TabCaption(4)   =   "Friday"
         TabPicture(4)   =   "frmChannelingDoctors.frx":0509
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label43"
         Tab(4).Control(1)=   "Label42"
         Tab(4).Control(2)=   "Label41"
         Tab(4).Control(3)=   "Label40"
         Tab(4).Control(4)=   "Label39"
         Tab(4).Control(5)=   "Label38"
         Tab(4).Control(6)=   "dtpFridayEveningEnd"
         Tab(4).Control(7)=   "dtpFridayEveningStart"
         Tab(4).Control(8)=   "dtpFridayMorningEnd"
         Tab(4).Control(9)=   "dtpFridayMorningStart"
         Tab(4).Control(10)=   "txtFridayEveningMax"
         Tab(4).Control(11)=   "chkFridayEvening"
         Tab(4).Control(12)=   "chkFridayMorning"
         Tab(4).Control(13)=   "txtFridayMorningMax"
         Tab(4).Control(14)=   "chkFridayFullLeave"
         Tab(4).ControlCount=   15
         TabCaption(5)   =   "Saturday"
         TabPicture(5)   =   "frmChannelingDoctors.frx":0525
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Label49"
         Tab(5).Control(1)=   "Label48"
         Tab(5).Control(2)=   "Label47"
         Tab(5).Control(3)=   "Label46"
         Tab(5).Control(4)=   "Label45"
         Tab(5).Control(5)=   "Label44"
         Tab(5).Control(6)=   "dtpSaturdayEveningEnd"
         Tab(5).Control(7)=   "dtpSaturdayEveningStart"
         Tab(5).Control(8)=   "dtpSaturdayMorningEnd"
         Tab(5).Control(9)=   "dtpSaturdayMorningStart"
         Tab(5).Control(10)=   "chkSaturdayFullLeave"
         Tab(5).Control(11)=   "txtSaturdayMorningMax"
         Tab(5).Control(12)=   "chkSaturdayMorning"
         Tab(5).Control(13)=   "chkSaturdayEvening"
         Tab(5).Control(14)=   "txtSaturdayEveningMax"
         Tab(5).ControlCount=   15
         TabCaption(6)   =   "Sunday"
         TabPicture(6)   =   "frmChannelingDoctors.frx":0541
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Label55"
         Tab(6).Control(1)=   "Label54"
         Tab(6).Control(2)=   "Label53"
         Tab(6).Control(3)=   "Label52"
         Tab(6).Control(4)=   "Label51"
         Tab(6).Control(5)=   "Label50"
         Tab(6).Control(6)=   "dtpSundayEveningEnd"
         Tab(6).Control(7)=   "dtpSundayEveningStart"
         Tab(6).Control(8)=   "dtpSundayMorningEnd"
         Tab(6).Control(9)=   "dtpSundayMorningStart"
         Tab(6).Control(10)=   "chkSundayFullLeave"
         Tab(6).Control(11)=   "txtSundayMorningMax"
         Tab(6).Control(12)=   "chkSundayMorning"
         Tab(6).Control(13)=   "chkSundayEvening"
         Tab(6).Control(14)=   "txtSundayEveningMax"
         Tab(6).ControlCount=   15
         TabCaption(7)   =   "Other Leave"
         TabPicture(7)   =   "frmChannelingDoctors.frx":055D
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Label19"
         Tab(7).Control(1)=   "Label4"
         Tab(7).Control(2)=   "Label5"
         Tab(7).Control(3)=   "Label8"
         Tab(7).Control(4)=   "Label9"
         Tab(7).Control(5)=   "Label10"
         Tab(7).Control(6)=   "Label11"
         Tab(7).Control(7)=   "Label17"
         Tab(7).Control(8)=   "dtpEveningEnding"
         Tab(7).Control(9)=   "dtpEveningStarting"
         Tab(7).Control(10)=   "dtpMorningEnding"
         Tab(7).Control(11)=   "dtpMorningStarting"
         Tab(7).Control(12)=   "bttnLeaveDelete"
         Tab(7).Control(13)=   "bttnAddLeave"
         Tab(7).Control(14)=   "dtpLeaveDate"
         Tab(7).Control(15)=   "Grid2"
         Tab(7).Control(16)=   "chkFullDayLeave"
         Tab(7).Control(17)=   "txtMNo"
         Tab(7).Control(18)=   "txtLeaveComments"
         Tab(7).Control(19)=   "chkMoning"
         Tab(7).Control(20)=   "chkEvening"
         Tab(7).Control(21)=   "txtENo"
         Tab(7).ControlCount=   22
         Begin VB.TextBox txtENo 
            Height          =   360
            Left            =   -69840
            MaxLength       =   250
            TabIndex        =   82
            Top             =   3240
            Width           =   855
         End
         Begin VB.CheckBox chkEvening 
            Caption         =   "Practcing in Evening Secession"
            Height          =   375
            Left            =   -74640
            TabIndex        =   79
            Top             =   2760
            Width           =   4215
         End
         Begin VB.CheckBox chkMoning 
            Caption         =   "Practcing in Morning Secession"
            Height          =   375
            Left            =   -74640
            TabIndex        =   75
            Top             =   1800
            Width           =   4215
         End
         Begin VB.TextBox txtLeaveComments 
            Height          =   720
            Left            =   -74640
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   83
            Top             =   3840
            Width           =   5655
         End
         Begin VB.TextBox txtMNo 
            Height          =   360
            Left            =   -69840
            MaxLength       =   250
            TabIndex        =   78
            Top             =   2280
            Width           =   855
         End
         Begin VB.CheckBox chkFullDayLeave 
            Caption         =   "NOT available full day"
            Height          =   375
            Left            =   -74640
            TabIndex        =   74
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox txtMondayEveningMax 
            Height          =   360
            Left            =   4080
            MaxLength       =   250
            TabIndex        =   18
            Top             =   5520
            Width           =   1455
         End
         Begin VB.CheckBox chkMondayEvening 
            Caption         =   "Practcing in Evening Secession"
            Height          =   375
            Left            =   600
            TabIndex        =   15
            Top             =   4080
            Width           =   4215
         End
         Begin VB.CheckBox chkMondayMorning 
            Caption         =   "Practcing in Morning Secession"
            Height          =   255
            Left            =   600
            TabIndex        =   11
            Top             =   1920
            Width           =   4215
         End
         Begin VB.TextBox txtMondayMorningMax 
            Height          =   360
            Left            =   4080
            MaxLength       =   250
            TabIndex        =   14
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CheckBox chkMondayFullLeave 
            Caption         =   "NOT available full day"
            Height          =   375
            Left            =   600
            TabIndex        =   10
            Top             =   840
            Width           =   2655
         End
         Begin VB.CheckBox chkTuesdayFullLeave 
            Caption         =   "NOT available full day"
            Height          =   375
            Left            =   -74400
            TabIndex        =   19
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox txtTuesdayMorningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   23
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CheckBox chkTuesdayMorning 
            Caption         =   "Practcing in Morning Secession"
            Height          =   255
            Left            =   -74400
            TabIndex        =   20
            Top             =   1920
            Width           =   4215
         End
         Begin VB.CheckBox chkTuesdayEvening 
            Caption         =   "Practcing in Evening Secession"
            Height          =   375
            Left            =   -74400
            TabIndex        =   24
            Top             =   4080
            Width           =   4215
         End
         Begin VB.TextBox txtTuesdayEveningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   27
            Top             =   5520
            Width           =   1455
         End
         Begin VB.CheckBox chkWednesdayFullLeave 
            Caption         =   "NOT available full day"
            Height          =   375
            Left            =   -74400
            TabIndex        =   28
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox txtWednesdayMorningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   32
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CheckBox chkWednesdayMorning 
            Caption         =   "Practcing in Morning Secession"
            Height          =   495
            Left            =   -74400
            TabIndex        =   29
            Top             =   1800
            Width           =   3135
         End
         Begin VB.CheckBox chkWednesdayEvening 
            Caption         =   "Practcing in Evening Secession"
            Height          =   375
            Left            =   -74400
            TabIndex        =   33
            Top             =   4080
            Width           =   4215
         End
         Begin VB.TextBox txtWednesdayEveningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   36
            Top             =   5520
            Width           =   1455
         End
         Begin VB.CheckBox chkThursdayFullLeave 
            Caption         =   "NOT available full day"
            Height          =   375
            Left            =   -74400
            TabIndex        =   37
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox txtThursdayMorningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   41
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CheckBox chkThursdayMorning 
            Caption         =   "Practcing in Morning Secession"
            Height          =   495
            Left            =   -74400
            TabIndex        =   38
            Top             =   1800
            Width           =   4215
         End
         Begin VB.CheckBox chkThursdayEvening 
            Caption         =   "Practcing in Evening Secession"
            Height          =   375
            Left            =   -74400
            TabIndex        =   42
            Top             =   4080
            Width           =   4215
         End
         Begin VB.TextBox txtThursdayEveningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   45
            Top             =   5520
            Width           =   1455
         End
         Begin VB.CheckBox chkFridayFullLeave 
            Caption         =   "NOT available full day"
            Height          =   375
            Left            =   -74400
            TabIndex        =   46
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox txtFridayMorningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   50
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CheckBox chkFridayMorning 
            Caption         =   "Practcing in Morning Secession"
            Height          =   495
            Left            =   -74400
            TabIndex        =   47
            Top             =   1800
            Width           =   4215
         End
         Begin VB.CheckBox chkFridayEvening 
            Caption         =   "Practcing in Evening Secession"
            Height          =   375
            Left            =   -74400
            TabIndex        =   51
            Top             =   4080
            Width           =   4215
         End
         Begin VB.TextBox txtFridayEveningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   54
            Top             =   5520
            Width           =   1455
         End
         Begin VB.TextBox txtSaturdayEveningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   63
            Top             =   5520
            Width           =   1455
         End
         Begin VB.CheckBox chkSaturdayEvening 
            Caption         =   "Practcing in Evening Secession"
            Height          =   375
            Left            =   -74400
            TabIndex        =   60
            Top             =   4080
            Width           =   4215
         End
         Begin VB.CheckBox chkSaturdayMorning 
            Caption         =   "Practcing in Morning Secession"
            Height          =   495
            Left            =   -74400
            TabIndex        =   56
            Top             =   1800
            Width           =   4215
         End
         Begin VB.TextBox txtSaturdayMorningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   59
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CheckBox chkSaturdayFullLeave 
            Caption         =   "NOT available full day"
            Height          =   375
            Left            =   -74400
            TabIndex        =   55
            Top             =   840
            Width           =   2295
         End
         Begin VB.TextBox txtSundayEveningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   72
            Top             =   5520
            Width           =   1455
         End
         Begin VB.CheckBox chkSundayEvening 
            Caption         =   "Practcing in Evening Secession"
            Height          =   375
            Left            =   -74400
            TabIndex        =   69
            Top             =   4080
            Width           =   3255
         End
         Begin VB.CheckBox chkSundayMorning 
            Caption         =   "Practcing in Morning Secession"
            Height          =   495
            Left            =   -74400
            TabIndex        =   65
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox txtSundayMorningMax 
            Height          =   360
            Left            =   -70920
            MaxLength       =   250
            TabIndex        =   68
            Top             =   3240
            Width           =   1455
         End
         Begin VB.CheckBox chkSundayFullLeave 
            Caption         =   "NOT available full day"
            Height          =   375
            Left            =   -74400
            TabIndex        =   64
            Top             =   840
            Width           =   2295
         End
         Begin MSFlexGridLib.MSFlexGrid Grid2 
            Height          =   1815
            Left            =   -74640
            TabIndex        =   101
            Top             =   5160
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   3201
            _Version        =   393216
         End
         Begin MSComCtl2.DTPicker dtpLeaveDate 
            Height          =   375
            Left            =   -72360
            TabIndex        =   73
            Top             =   720
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55050241
            CurrentDate     =   39401
         End
         Begin btButtonEx.ButtonEx bttnAddLeave 
            Height          =   255
            Left            =   -71400
            TabIndex        =   84
            Top             =   4800
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
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
         Begin btButtonEx.ButtonEx bttnLeaveDelete 
            Height          =   255
            Left            =   -70200
            TabIndex        =   85
            Top             =   4800
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            Appearance      =   3
            Enabled         =   0   'False
            Caption         =   "Delete"
            Enabled         =   0   'False
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
         Begin MSComCtl2.DTPicker dtpMorningStarting 
            Height          =   375
            Left            =   -74280
            TabIndex        =   76
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpMorningEnding 
            Height          =   375
            Left            =   -72240
            TabIndex        =   77
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpEveningStarting 
            Height          =   375
            Left            =   -74280
            TabIndex        =   80
            Top             =   3240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpEveningEnding 
            Height          =   375
            Left            =   -72240
            TabIndex        =   81
            Top             =   3240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpMondayMorningStart 
            Height          =   360
            Left            =   4080
            TabIndex        =   12
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpMondayMorningEnd 
            Height          =   360
            Left            =   4080
            TabIndex        =   13
            Top             =   2760
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpMondayEveningStart 
            Height          =   360
            Left            =   4080
            TabIndex        =   16
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpMondayEveningEnd 
            Height          =   360
            Left            =   4080
            TabIndex        =   17
            Top             =   5040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpTuesdayMorningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   21
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpTuesdayMorningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   22
            Top             =   2760
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpTuesdayEveningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   25
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpTuesdayEveningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   26
            Top             =   5040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpWednesdayMorningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   30
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpWednesdayMorningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   31
            Top             =   2760
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpWednesdayEveningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   34
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpWednesdayEveningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   35
            Top             =   5040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpThursdayMorningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   39
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpThursdayMorningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   40
            Top             =   2760
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpThursdayEveningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   43
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpThursdayEveningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   44
            Top             =   5040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpFridayMorningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   48
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpFridayMorningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   49
            Top             =   2760
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpFridayEveningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   52
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpFridayEveningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   53
            Top             =   5040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpSaturdayMorningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   57
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpSaturdayMorningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   58
            Top             =   2760
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpSaturdayEveningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   61
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpSaturdayEveningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   62
            Top             =   5040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpSundayMorningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   66
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpSundayMorningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   67
            Top             =   2760
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpSundayEveningStart 
            Height          =   360
            Left            =   -70920
            TabIndex        =   70
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin MSComCtl2.DTPicker dtpSundayEveningEnd 
            Height          =   360
            Left            =   -70920
            TabIndex        =   71
            Top             =   5040
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39401
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Changed Dates"
            Height          =   255
            Left            =   -74640
            TabIndex        =   151
            Top             =   4800
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72600
            TabIndex        =   150
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -74760
            TabIndex        =   149
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -70560
            TabIndex        =   148
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72600
            TabIndex        =   147
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -74760
            TabIndex        =   146
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            Height          =   255
            Left            =   -74640
            TabIndex        =   145
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -70560
            TabIndex        =   144
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   2280
            TabIndex        =   143
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   2280
            TabIndex        =   142
            Top             =   4680
            Width           =   855
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   2280
            TabIndex        =   141
            Top             =   5640
            Width           =   855
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   2280
            TabIndex        =   140
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   2280
            TabIndex        =   139
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   2280
            TabIndex        =   138
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   137
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   136
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   135
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   134
            Top             =   5640
            Width           =   855
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   133
            Top             =   4680
            Width           =   855
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   132
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   131
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   130
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   129
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   128
            Top             =   5640
            Width           =   855
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   127
            Top             =   4680
            Width           =   855
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   126
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   125
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   124
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   123
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   122
            Top             =   5640
            Width           =   855
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   121
            Top             =   4680
            Width           =   855
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   120
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   119
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   118
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   117
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   116
            Top             =   5640
            Width           =   855
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   115
            Top             =   4680
            Width           =   855
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   114
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   113
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   112
            Top             =   4680
            Width           =   855
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   111
            Top             =   5640
            Width           =   855
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   110
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   109
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   108
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label50 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   107
            Top             =   5160
            Width           =   855
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   106
            Top             =   4680
            Width           =   855
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   105
            Top             =   5640
            Width           =   855
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   255
            Left            =   -72720
            TabIndex        =   104
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   255
            Left            =   -72720
            TabIndex        =   103
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. No."
            Height          =   255
            Left            =   -72720
            TabIndex        =   102
            Top             =   3240
            Width           =   855
         End
      End
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9960
      TabIndex        =   89
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   7215
      Left            =   480
      TabIndex        =   90
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   12726
      _Version        =   393216
      ScrollTrack     =   -1  'True
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
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
End
Attribute VB_Name = "frmChannelingDoctors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemStaffFacility As Long
    Dim FromGrid As Boolean
    Dim CatogeryID  As Integer
    Dim TemDoctorID  As Integer

Private Sub SetColour()
Select Case ColourScheme
    Case 1:
        BttnBackColour = 5341695
        BttnForeColour = 1314458
        FrmBackColour = 11066623
        FrmForeColour = 1314458
        FrameBackColour = 11066623
        FrameForeColour = 1314458
        TxtBackColour = 9881851
        TxtForeColour = 1314458
        LblBackColour = 11066623
        LblForeColour = 1314458
        GridBackColor = 9881855
        GridBackColorBkg = 10474239
        GridBackColorFixed = 8566015
        GridBackColorSel = 5341695
        GridForeColor = 1314458
        GridForeColorFixed = 11944
        GridForeColorSel = 3014824
    Case 2:
        BttnBackColour = 14803300
        BttnForeColour = 5539362
        FrmBackColour = 16766120
        FrmForeColour = 5539362
        FrameBackColour = 16766120
        FrameForeColour = 5539362
        TxtBackColour = 16760450
        TxtForeColour = 5539362
        LblBackColour = 16766120
        LblForeColour = 5539362
        GridBackColor = 16760450
        GridBackColorBkg = 16771260
        GridBackColorFixed = 16105620
        GridBackColorSel = 16737380
        GridForeColor = 5539362
        GridForeColorFixed = 5539362
        GridForeColorSel = 16765588
    Case 3:
        BttnBackColour = 51455
        BttnForeColour = 942490
        FrmBackColour = 11070719
        FrmForeColour = 942490
        FrameBackColour = 11070719
        FrameForeColour = 942490
        TxtBackColour = 11528439
        TxtForeColour = 1314458
        LblBackColour = 11070719
        LblForeColour = 942490
        GridBackColor = 16760450
        GridBackColorBkg = 16771260
        GridBackColorFixed = 16105620
        GridBackColorSel = 16737380
        GridForeColor = 5539362
        GridForeColorFixed = 5539362
        GridForeColorSel = 16765588
    End Select
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
    bttnAdd.BackColor = BttnBackColour
    bttnAdd.ForeColor = BttnForeColour
    bttnAddLeave.BackColor = BttnBackColour
    bttnAddLeave.ForeColor = BttnForeColour
    bttnCancel.BackColor = BttnBackColour
    bttnCancel.ForeColor = BttnForeColour
    bttnChange.BackColor = BttnBackColour
    bttnChange.ForeColor = BttnForeColour
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
    bttnDelete.BackColor = BttnBackColour
    bttnDelete.ForeColor = BttnForeColour
    bttnEdit.BackColor = BttnBackColour
    bttnEdit.ForeColor = BttnForeColour
    bttnLeaveDelete.BackColor = BttnBackColour
    bttnLeaveDelete.ForeColor = BttnForeColour
    bttnSave.BackColor = BttnBackColour
    bttnSave.ForeColor = BttnForeColour
    grid1.BackColor = GridBackColor
    grid1.ForeColor = GridForeColor
    grid1.BackColorBkg = GridBackColorBkg
    grid1.BackColorFixed = GridBackColorFixed
    grid1.BackColorSel = GridBackColorSel
    grid1.ForeColor = GridForeColor
    grid1.ForeColorFixed = GridForeColorFixed
    grid1.ForeColorSel = GridForeColorSel
    grid1.ForeColor = GridForeColor
    Label1.BackColor = LblBackColour
    Label1.ForeColor = LblForeColour
    LblDoctorStaff.BackColor = LblBackColour
    LblDoctorStaff.ForeColor = LblForeColour
    lblInstitutionFee.BackColor = LblBackColour
    lblInstitutionFee.ForeColor = LblForeColour
End Sub

Private Sub chkFullDayLeave_Click()
    If chkFullDayLeave.Value = 1 Then
        chkMoning.Value = 0
        chkEvening.Value = 0
    End If
End Sub
Private Sub chkMoning_Click()
    If chkMoning.Value = 1 Then
        chkFullDayLeave.Value = 0
    ElseIf chkMoning.Value = 0 And chkEvening.Value = 0 Then
        chkFullDayLeave.Value = 1
    End If
End Sub
Private Sub chkEvening_Click()
    If chkEvening.Value = 1 Then
        chkFullDayLeave.Value = 0
    ElseIf chkMoning.Value = 0 And chkEvening.Value = 0 Then
        chkFullDayLeave.Value = 1
    End If
End Sub
Private Sub chkMondayFullLeave_Click()
    If chkMondayFullLeave.Value = 1 Then
        chkMondayEvening.Value = 0
        chkMondayMorning.Value = 0
    End If
End Sub
Private Sub chkMondayMorning_Click()
    If chkMondayMorning.Value = 1 Then
        chkMondayFullLeave.Value = 0
    ElseIf chkMondayMorning.Value = 0 And chkMondayEvening.Value = 0 Then
        chkMondayFullLeave.Value = 1
    End If
End Sub
Private Sub chkMondayEvening_Click()
    If chkMondayEvening.Value = 1 Then
        chkMondayFullLeave.Value = 0
    ElseIf chkMondayMorning.Value = 0 And chkMondayEvening.Value = 0 Then
        chkMondayFullLeave.Value = 1
    End If
End Sub
Private Sub chktuesdayFullLeave_Click()
    If chkTuesdayFullLeave.Value = 1 Then
        chkTuesdayEvening.Value = 0
        chkTuesdayMorning.Value = 0
    End If
End Sub
Private Sub chktuesdayMorning_Click()
    If chkTuesdayMorning.Value = 1 Then
        chkTuesdayFullLeave.Value = 0
    ElseIf chkTuesdayMorning.Value = 0 And chkTuesdayEvening.Value = 0 Then
        chkTuesdayFullLeave.Value = 1
    End If
End Sub
Private Sub chktuesdayEvening_Click()
    If chkTuesdayEvening.Value = 1 Then
        chkTuesdayFullLeave.Value = 0
    ElseIf chkTuesdayMorning.Value = 0 And chkTuesdayEvening.Value = 0 Then
        chkTuesdayFullLeave.Value = 1
    End If
End Sub
Private Sub chkwednesdayFullLeave_Click()
    If chkWednesdayFullLeave.Value = 1 Then
        chkWednesdayEvening.Value = 0
        chkWednesdayMorning.Value = 0
    End If
End Sub
Private Sub chkwednesdayMorning_Click()
    If chkWednesdayMorning.Value = 1 Then
        chkWednesdayFullLeave.Value = 0
    ElseIf chkWednesdayMorning.Value = 0 And chkWednesdayEvening.Value = 0 Then
        chkWednesdayFullLeave.Value = 1
    End If
End Sub
Private Sub chkwednesdayEvening_Click()
    If chkWednesdayEvening.Value = 1 Then
        chkWednesdayFullLeave.Value = 0
    ElseIf chkWednesdayMorning.Value = 0 And chkWednesdayEvening.Value = 0 Then
        chkWednesdayFullLeave.Value = 1
    End If
End Sub
Private Sub chkthursdayFullLeave_Click()
    If chkThursdayFullLeave.Value = 1 Then
        chkThursdayEvening.Value = 0
        chkThursdayMorning.Value = 0
    End If
End Sub
Private Sub chkthursdayMorning_Click()
    If chkThursdayMorning.Value = 1 Then
        chkThursdayFullLeave.Value = 0
    ElseIf chkThursdayMorning.Value = 0 And chkThursdayEvening.Value = 0 Then
        chkThursdayFullLeave.Value = 1
    End If
End Sub
Private Sub chkthursdayEvening_Click()
    If chkThursdayEvening.Value = 1 Then
        chkThursdayFullLeave.Value = 0
    ElseIf chkThursdayMorning.Value = 0 And chkThursdayEvening.Value = 0 Then
        chkThursdayFullLeave.Value = 1
    End If
End Sub
Private Sub chkfridayFullLeave_Click()
    If chkFridayFullLeave.Value = 1 Then
        chkFridayEvening.Value = 0
        chkFridayMorning.Value = 0
    End If
End Sub
Private Sub chkfridayMorning_Click()
    If chkFridayMorning.Value = 1 Then
        chkFridayFullLeave.Value = 0
    ElseIf chkFridayMorning.Value = 0 And chkFridayEvening.Value = 0 Then
        chkFridayFullLeave.Value = 1
    End If
End Sub
Private Sub chkfridayEvening_Click()
    If chkFridayEvening.Value = 1 Then
        chkFridayFullLeave.Value = 0
    ElseIf chkFridayMorning.Value = 0 And chkFridayEvening.Value = 0 Then
        chkFridayFullLeave.Value = 1
    End If
End Sub
Private Sub chksaturdayFullLeave_Click()
    If chkSaturdayFullLeave.Value = 1 Then
        chkSaturdayEvening.Value = 0
        chkSaturdayMorning.Value = 0
    End If
End Sub
Private Sub chksaturdayMorning_Click()
    If chkSaturdayMorning.Value = 1 Then
        chkSaturdayFullLeave.Value = 0
    ElseIf chkSaturdayMorning.Value = 0 And chkSaturdayEvening.Value = 0 Then
        chkSaturdayFullLeave.Value = 1
    End If
End Sub
Private Sub chksaturdayEvening_Click()
    If chkSaturdayEvening.Value = 1 Then
        chkSaturdayFullLeave.Value = 0
    ElseIf chkSaturdayMorning.Value = 0 And chkSaturdayEvening.Value = 0 Then
        chkSaturdayFullLeave.Value = 1
    End If
End Sub
Private Sub chksundayFullLeave_Click()
    If chkSundayFullLeave.Value = 1 Then
        chkSundayEvening.Value = 0
        chkSundayMorning.Value = 0
    End If
End Sub
Private Sub chksundayMorning_Click()
    If chkSundayMorning.Value = 1 Then
        chkSundayFullLeave.Value = 0
    ElseIf chkSundayMorning.Value = 0 And chkSundayEvening.Value = 0 Then
        chkSundayFullLeave.Value = 1
    End If
End Sub
Private Sub chksundayEvening_Click()
    If chkSundayEvening.Value = 1 Then
        chkSundayFullLeave.Value = 0
    ElseIf chkSundayMorning.Value = 0 And chkSundayEvening.Value = 0 Then
        chkSundayFullLeave.Value = 1
    End If
End Sub


Private Sub Form_Load()
    Call BeforeAddEdit
    Call SetColour
    Call FormatGrid
    Call FillGrid
    Call PrepareForDoctor
    SSTab1.Tab = 0
    dtpLeaveDate.Format = dtpCustom
    dtpLeaveDate.CustomFormat = "dd MMM yyyy"
    dtpLeaveDate.MinDate = Date
End Sub


Private Sub ClearValues()
    DataComboDoctor.Text = Empty
    txtComments.Text = Empty
    txtInstitutionForeignerFee.Text = Empty
    txtInstitutionLocalFee.Text = Empty
    txtDoctorLocalFee.Text = Empty
    txtDoctorForeignerFee.Text = Empty
    
    txtLeaveComments.Text = Empty
        
    chkMondayFullLeave.Value = Empty
    chkMondayMorning.Value = Empty
    dtpMondayMorningStart.Value = 0
    dtpMondayMorningEnd.Value = 0
    txtMondayMorningMax.Text = Empty
    chkMondayEvening.Value = Empty
    dtpMondayEveningStart.Value = 0
    dtpMondayEveningEnd.Value = 0
    txtMondayEveningMax.Text = Empty
    
    chkTuesdayFullLeave.Value = Empty
    chkTuesdayMorning.Value = Empty
    dtpTuesdayMorningStart.Value = 0
    dtpTuesdayMorningEnd.Value = 0
    txtTuesdayMorningMax.Text = Empty
    chkTuesdayEvening.Value = Empty
    dtpTuesdayEveningStart.Value = 0
    dtpTuesdayEveningEnd.Value = 0
    txtTuesdayEveningMax.Text = Empty
    
    chkWednesdayFullLeave.Value = Empty
    chkWednesdayMorning.Value = Empty
    dtpWednesdayMorningStart.Value = 0
    dtpWednesdayMorningEnd.Value = 0
    txtWednesdayMorningMax.Text = Empty
    chkWednesdayEvening.Value = Empty
    dtpWednesdayEveningStart.Value = 0
    dtpWednesdayEveningEnd.Value = 0
    txtWednesdayEveningMax.Text = Empty
    
    chkThursdayFullLeave.Value = Empty
    chkThursdayMorning.Value = Empty
    dtpThursdayMorningStart.Value = 0
    dtpThursdayMorningEnd.Value = 0
    txtThursdayMorningMax.Text = Empty
    chkThursdayEvening.Value = Empty
    dtpThursdayEveningStart.Value = 0
    dtpThursdayEveningEnd.Value = 0
    txtThursdayEveningMax.Text = Empty
    
    chkFridayFullLeave.Value = Empty
    chkFridayMorning.Value = Empty
    dtpFridayMorningStart.Value = 0
    dtpFridayMorningEnd.Value = 0
    txtFridayMorningMax.Text = Empty
    chkFridayEvening.Value = Empty
    dtpFridayEveningStart.Value = 0
    dtpFridayEveningEnd.Value = 0
    txtFridayEveningMax.Text = Empty
    
    chkSaturdayFullLeave.Value = Empty
    chkSaturdayMorning.Value = Empty
    dtpSaturdayMorningStart.Value = 0
    dtpSaturdayMorningEnd.Value = 0
    txtSaturdayMorningMax.Text = Empty
    chkSaturdayEvening.Value = Empty
    dtpSaturdayEveningStart.Value = 0
    dtpSaturdayEveningEnd.Value = 0
    txtSaturdayEveningMax.Text = Empty
    
    chkSundayFullLeave.Value = Empty
    chkSundayMorning.Value = Empty
    dtpSundayMorningStart.Value = 0
    dtpSundayMorningEnd.Value = 0
    txtSundayMorningMax.Text = Empty
    chkSundayEvening.Value = Empty
    dtpSundayEveningStart.Value = 0
    dtpSundayEveningEnd.Value = 0
    txtSundayEveningMax.Text = Empty
    
    FormatLeaveGrid
End Sub

Private Sub BeforeAddEdit()
    Call ClearValues
    bttnAdd.Enabled = True
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    grid1.Enabled = True
    framFacility.Enabled = False
    SSTabDates.Enabled = False
    bttnSave.Visible = False
    bttnChange.Visible = False
    bttnCancel.Visible = False
    txtSearch.Text = Empty
    SSTab1.Tab = 0
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub AfterAdd()
    Call ClearValues
    SSTabDates.Enabled = True
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    grid1.Enabled = False
    framFacility.Enabled = True
    bttnAddLeave.Enabled = False
    bttnLeaveDelete.Enabled = False
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    SSTab1.Tab = 0
    DataComboDoctor.SetFocus
End Sub


Private Sub AfterEdit()
    bttnAdd.Enabled = False
    bttnEdit.Enabled = False
    bttnDelete.Enabled = False
    grid1.Enabled = False
    framFacility.Enabled = True
    SSTabDates.Enabled = True
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    bttnAddLeave.Enabled = True
    bttnLeaveDelete.Enabled = True
    DataComboDoctor.SetFocus
End Sub

Private Sub FormatGrid()
    Dim BorderMargin As Integer
    BorderMargin = 100
    With grid1
        .Clear
        .Cols = 3
        .Rows = 1
        .Row = 0
        .ColWidth(0) = 600
        .ColWidth(2) = 1
        .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + BorderMargin)
        .Col = 0
        .CellAlignment = 4
        .Text = "No."
        .Col = 1
        .CellAlignment = 4
        .Text = "Doctor"
    End With
End Sub


Private Sub FillGrid()
    Dim NowRow As Long
    With DataEnvironment1.rssqlFacilityStaff
        If .State = 1 Then .Close
        .Source = "SELECT tblFacilitystaff.* from tblFacilitystaff where hospitalfacility_ID = 10 order by facilitystaff "
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        NowRow = 0
        While .EOF = False
            NowRow = NowRow + 1
            grid1.Rows = NowRow + 1
            grid1.Row = NowRow
            grid1.Col = 0
            grid1.CellAlignment = 1
            grid1.Text = NowRow
            grid1.Col = 1
            grid1.CellAlignment = 1
            grid1.Text = FindDoctorFromID(!staff_ID)
            grid1.Col = 2
            grid1.Text = !FacilityStaff_ID
            .MoveNext
        Wend
    End With
End Sub

Private Sub bttnAdd_Click()
    Call AfterAdd
End Sub


Private Sub bttnChange_Click()
    Dim TemResponce  As Integer
    If Not IsNumeric(DataComboDoctor.BoundText) Then
        TemResponce = MsgBox("You have not selected a name of a doctor", vbCritical, "Doctor?")
        DataComboDoctor.SetFocus
        Exit Sub
    End If
    Call EditData
    Call BeforeAddEdit
    Call ClearValues
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDelete_Click()
    Dim TemResponce  As Integer
    grid1.Col = 2
    If Not IsNumeric(grid1.Text) Then Exit Sub
    grid1.Col = 1
    TemResponce = MsgBox("Are you sure you want to remove " & grid1.Text & " from the Facilities list that the hospital provide", vbCritical + vbYesNo, "Remove?")
    If TemResponce = vbNo Then Exit Sub
    grid1.Col = 2
    With DataEnvironment1.rssqlFacilityStaff
        If .State = 1 Then .Close
        .Source = "Select tblfacilitystaff.* from tblfacilitystaff where facilitystaff_ID = " & grid1.Text
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .Delete adAffectCurrent
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    Call BeforeAddEdit
End Sub

Private Sub bttnEdit_Click()
    grid1.Col = 2
    TemStaffFacility = grid1.Text
    Call AfterEdit
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
End Sub


Private Sub bttnSave_Click()
    Dim TemResponce  As Integer
    If Not IsNumeric(DataComboDoctor.BoundText) Then
        TemResponce = MsgBox("You have not selected a name of a doctor", vbCritical, "Doctor?")
        DataComboDoctor.SetFocus
        Exit Sub
    End If
    Call SaveData
    Call BeforeAddEdit
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub GetData()
    With DataEnvironment1.rssqlFacilityStaff
        If .State = 1 Then .Close
        .Source = "SELECT tblFacilitystaff.* from tblFacilitystaff where (Facilitystaff_ID = " & TemStaffFacility & ")"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        ClearValues
        DataComboDoctor.BoundText = !staff_ID
        
        TemDoctorID = !staff_ID
        
        If Not IsNull(!usualpersonalFee) Then txtDoctorLocalFee.Text = Format(!usualpersonalFee, "0.00")
        If Not IsNull(!UsualInstitutionFee) Then txtInstitutionLocalFee.Text = Format(!UsualInstitutionFee, "0.00")
        If Not IsNull(!foreignerpersonalfee) Then txtDoctorForeignerFee.Text = !foreignerpersonalfee
        If Not IsNull(!ForeignerinstitutionFee) Then txtInstitutionForeignerFee.Text = !ForeignerinstitutionFee
        If Not IsNull(!FacilityStaffComment) Then txtComments.Text = !FacilityStaffComment
        If Not IsNull(!usualduration) Then txtUsualDuration.Text = !usualduration
        
        If !FullDayLeaveMonday = True Then chkMondayFullLeave.Value = 1
        If !FacilityMondayM = True Then chkMondayMorning.Value = 1
        If Not IsNull(!FacilityMondayMStarting) Then dtpMondayMorningStart.Value = !FacilityMondayMStarting
        If Not IsNull(!FacilityMondayMEnding) Then dtpMondayMorningEnd.Value = !FacilityMondayMEnding
        If Not IsNull(!FacilityMondayMNo) Then txtMondayMorningMax.Text = !FacilityMondayMNo
        If !FacilityMondayE = True Then chkMondayEvening.Value = 1
        If Not IsNull(!FacilityMondayEStarting) Then dtpMondayEveningStart.Value = !FacilityMondayEStarting
        If Not IsNull(!FacilityMondayEEnding) Then dtpMondayEveningEnd.Value = !FacilityMondayEEnding
        If Not IsNull(!FacilityMondayENo) Then txtMondayEveningMax.Text = !FacilityMondayENo

        If !FullDayLeaveTuesday = True Then chkTuesdayFullLeave.Value = 1
        If !FacilityTuesdayM = True Then chkTuesdayMorning.Value = 1
        If Not IsNull(!FacilitytuesdayMStarting) Then dtpTuesdayMorningStart.Value = !FacilitytuesdayMStarting
        If Not IsNull(!FacilitytuesdayMEnding) Then dtpTuesdayMorningEnd.Value = !FacilitytuesdayMEnding
        If Not IsNull(!FacilitytuesdayMNo) Then txtTuesdayMorningMax.Text = !FacilitytuesdayMNo
        If !FacilityTuesdayE = True Then chkTuesdayEvening.Value = 1
        If Not IsNull(!FacilitytuesdayEStarting) Then dtpTuesdayEveningStart.Value = !FacilitytuesdayEStarting
        If Not IsNull(!FacilitytuesdayEEnding) Then dtpTuesdayEveningEnd.Value = !FacilitytuesdayEEnding
        If Not IsNull(!FacilitytuesdayENo) Then txtTuesdayEveningMax.Text = !FacilitytuesdayENo
        
        If !FullDayLeaveWednesday = True Then chkWednesdayFullLeave.Value = 1
        If !FacilityWednesdayM = True Then chkWednesdayMorning.Value = 1
        If Not IsNull(!FacilitywednesdayMStarting) Then dtpWednesdayMorningStart.Value = !FacilitywednesdayMStarting
        If Not IsNull(!FacilitywednesdayMEnding) Then dtpWednesdayMorningEnd.Value = !FacilitywednesdayMEnding
        If Not IsNull(!FacilitywednesdayMNo) Then txtWednesdayMorningMax.Text = !FacilitywednesdayMNo
        If !FacilityWednesdayE = True Then chkWednesdayEvening.Value = 1
        If Not IsNull(!FacilitywednesdayEStarting) Then dtpWednesdayEveningStart.Value = !FacilitywednesdayEStarting
        If Not IsNull(!FacilitywednesdayEEnding) Then dtpWednesdayEveningEnd.Value = !FacilitywednesdayEEnding
        If Not IsNull(!FacilitywednesdayENo) Then txtWednesdayEveningMax.Text = !FacilitywednesdayENo
        
        If !FullDayLeaveThursday = True Then chkThursdayFullLeave.Value = 1
        If !FacilityThursdayM = True Then chkThursdayMorning.Value = 1
        If Not IsNull(!FacilitythursdayMStarting) Then dtpThursdayMorningStart.Value = !FacilitythursdayMStarting
        If Not IsNull(!FacilitythursdayMEnding) Then dtpThursdayMorningEnd.Value = !FacilitythursdayMEnding
        If Not IsNull(!FacilitythursdayMNo) Then txtThursdayMorningMax.Text = !FacilitythursdayMNo
        If !FacilityThursdayE = True Then chkThursdayEvening.Value = 1
        If Not IsNull(!FacilitythursdayEStarting) Then dtpThursdayEveningStart.Value = !FacilitythursdayEStarting
        If Not IsNull(!FacilitythursdayEEnding) Then dtpThursdayEveningEnd.Value = !FacilitythursdayEEnding
        If Not IsNull(!FacilitythursdayENo) Then txtThursdayEveningMax.Text = !FacilitythursdayENo
        
        If !FullDayLeaveFriday = True Then chkFridayFullLeave.Value = 1
        If !FacilityFridayM = True Then chkFridayMorning.Value = 1
        If Not IsNull(!FacilityfridayMStarting) Then dtpFridayMorningStart.Value = !FacilityfridayMStarting
        If Not IsNull(!FacilityfridayMEnding) Then dtpFridayMorningEnd.Value = !FacilityfridayMEnding
        If Not IsNull(!FacilityfridayMNo) Then txtFridayMorningMax.Text = !FacilityfridayMNo
        If !FacilityFridayE = True Then chkFridayEvening.Value = 1
        If Not IsNull(!FacilityfridayEStarting) Then dtpFridayEveningStart.Value = !FacilityfridayEStarting
        If Not IsNull(!FacilityfridayEEnding) Then dtpFridayEveningEnd.Value = !FacilityfridayEEnding
        If Not IsNull(!FacilityfridayENo) Then txtFridayEveningMax.Text = !FacilityfridayENo
        
        If !FullDayLeaveSaturday = True Then chkSaturdayFullLeave.Value = 1
        If !FacilitySaturdayM = True Then chkSaturdayMorning.Value = 1
        If Not IsNull(!FacilitysaturdayMStarting) Then dtpSaturdayMorningStart.Value = !FacilitysaturdayMStarting
        If Not IsNull(!FacilitysaturdayMEnding) Then dtpSaturdayMorningEnd.Value = !FacilitysaturdayMEnding
        If Not IsNull(!FacilitysaturdayMNo) Then txtSaturdayMorningMax.Text = !FacilitysaturdayMNo
        If !FacilitySaturdayE = True Then chkSaturdayEvening.Value = 1
        If Not IsNull(!FacilitysaturdayEStarting) Then dtpSaturdayEveningStart.Value = !FacilitysaturdayEStarting
        If Not IsNull(!FacilitysaturdayEEnding) Then dtpSaturdayEveningEnd.Value = !FacilitysaturdayEEnding
        If Not IsNull(!FacilitysaturdayENo) Then txtSaturdayEveningMax.Text = !FacilitysaturdayENo

        If !FullDayLeaveSunday = True Then chkSundayFullLeave.Value = 1
        If !FacilitySundayM = True Then chkSundayMorning.Value = 1
        If Not IsNull(!FacilitysundayMStarting) Then dtpSundayMorningStart.Value = !FacilitysundayMStarting
        If Not IsNull(!FacilitysundayMEnding) Then dtpSundayMorningEnd.Value = !FacilitysundayMEnding
        If Not IsNull(!FacilitysundayMNo) Then txtSundayMorningMax.Text = !FacilitysundayMNo
        If !FacilitySundayE = True Then chkSundayEvening.Value = 1
        If Not IsNull(!FacilitysundayEStarting) Then dtpSundayEveningStart.Value = !FacilitysundayEStarting
        If Not IsNull(!FacilitysundayEEnding) Then dtpSundayEveningEnd.Value = !FacilitysundayEEnding
        If Not IsNull(!FacilitysundayENo) Then txtSundayEveningMax.Text = !FacilitysundayENo
       
        .Close
    End With
    Call FormatLeaveGrid
    Call FillLeaveGrid
End Sub

Private Sub SaveData()

With DataEnvironment1.rssqlFacilityStaff
    If .State = 1 Then .Close
    .Source = "SELECT tblFacilitystaff.* from tblFacilitystaff"
    If .State = 0 Then .Open
    .AddNew
    !TwoSecessions = True
    !facilitystaff = "Channeling - " & DataComboDoctor.Text
    !HospitalFacility_id = 10
    !staff_ID = DataComboDoctor.BoundText
    
    !usualpersonalFee = Val(txtDoctorLocalFee.Text)
    !UsualInstitutionFee = Val(txtInstitutionLocalFee.Text)
    
    !foreignerpersonalfee = Val(txtDoctorForeignerFee.Text)
    !ForeignerinstitutionFee = Val(txtInstitutionForeignerFee.Text)
    
    !FacilityStaffComment = txtComments.Text
    !usualduration = Val(txtUsualDuration.Text)
    
        
'    .Update
    
        If chkMondayFullLeave.Value = 1 Then
            !FullDayLeaveMonday = True
        Else
            !FullDayLeaveMonday = False
        End If
        If chkMondayMorning.Value = 1 Then
            !FacilityMondayM = True
        Else
            !FacilityMondayM = False
        End If
        !FacilityMondayMStarting = dtpMondayMorningStart.Value
        !FacilityMondayMEnding = dtpMondayMorningEnd.Value
        !FacilityMondayMNo = Val(txtMondayMorningMax.Text)
        If chkMondayEvening.Value = 1 Then
            !FacilityMondayE = True
        Else
            !FacilityMondayE = False
        End If
        !FacilityMondayEStarting = dtpMondayEveningStart.Value
        !FacilityMondayEEnding = dtpMondayEveningEnd.Value
        !FacilityMondayENo = Val(txtMondayEveningMax.Text)
'    .Update
    
    
        If chkTuesdayFullLeave.Value = 1 Then
            !FullDayLeaveTuesday = True
        Else
            !FullDayLeaveTuesday = False
        End If
        If chkTuesdayMorning.Value = 1 Then
            !FacilityTuesdayM = True
        Else
            !FacilityTuesdayM = False
        End If
        !FacilitytuesdayMStarting = dtpTuesdayMorningStart.Value
        !FacilitytuesdayMEnding = dtpTuesdayMorningEnd.Value
        !FacilitytuesdayMNo = Val(txtTuesdayMorningMax.Text)
        If chkTuesdayEvening.Value = 1 Then
            !FacilityTuesdayE = True
        Else
            !FacilityTuesdayE = False
        End If
        !FacilitytuesdayEStarting = dtpTuesdayEveningStart.Value
        !FacilitytuesdayEEnding = dtpTuesdayEveningEnd.Value
        !FacilitytuesdayENo = Val(txtTuesdayEveningMax.Text)
    
'    .Update
    
        If chkWednesdayFullLeave.Value = 1 Then
            !FullDayLeaveWednesday = True
        Else
            !FullDayLeaveWednesday = False
        End If
        If chkWednesdayMorning.Value = 1 Then
            !FacilityWednesdayM = True
        Else
            !FacilityWednesdayM = False
        End If
        !FacilitywednesdayMStarting = dtpWednesdayMorningStart.Value
        !FacilitywednesdayMEnding = dtpWednesdayMorningEnd.Value
        !FacilitywednesdayMNo = Val(txtWednesdayMorningMax.Text)
        If chkWednesdayEvening.Value = 1 Then
            !FacilityWednesdayE = True
        Else
            !FacilityWednesdayE = False
        End If
        !FacilitywednesdayEStarting = dtpWednesdayEveningStart.Value
        !FacilitywednesdayEEnding = dtpWednesdayEveningEnd.Value
        !FacilitywednesdayENo = Val(txtWednesdayEveningMax.Text)
'    .Update
    
        If chkThursdayFullLeave.Value = 1 Then
            !FullDayLeaveThursday = True
        Else
            !FullDayLeaveThursday = False
        End If
        If chkThursdayMorning.Value = 1 Then
            !FacilityThursdayM = True
        Else
            !FacilityThursdayM = False
        End If
        !FacilitythursdayMStarting = dtpThursdayMorningStart.Value
        !FacilitythursdayMEnding = dtpThursdayMorningEnd.Value
        !FacilitythursdayMNo = Val(txtThursdayMorningMax.Text)
        If chkThursdayEvening.Value = 1 Then
            !FacilityThursdayE = True
        Else
            !FacilityThursdayE = False
        End If
        !FacilitythursdayEStarting = dtpThursdayEveningStart.Value
        !FacilitythursdayEEnding = dtpThursdayEveningEnd.Value
        !FacilitythursdayENo = Val(txtThursdayEveningMax.Text)
'    .Update
    
        If chkFridayFullLeave.Value = 1 Then
            !FullDayLeaveFriday = True
        Else
            !FullDayLeaveFriday = False
        End If
        If chkFridayMorning.Value = 1 Then
            !FacilityFridayM = True
        Else
            !FacilityFridayM = False
        End If
        !FacilityfridayMStarting = dtpFridayMorningStart.Value
        !FacilityfridayMEnding = dtpFridayMorningEnd.Value
        !FacilityfridayMNo = Val(txtFridayMorningMax.Text)
        If chkFridayEvening.Value = 1 Then
            !FacilityFridayE = True
        Else
            !FacilityFridayE = False
        End If
        !FacilityfridayEStarting = dtpFridayEveningStart.Value
        !FacilityfridayEEnding = dtpFridayEveningEnd.Value
        !FacilityfridayENo = Val(txtFridayEveningMax.Text)
'    .Update
    
        If chkSaturdayFullLeave.Value = 1 Then
            !FullDayLeaveSaturday = True
        Else
            !FullDayLeaveSaturday = False
        End If
        If chkSaturdayMorning.Value = 1 Then
            !FacilitySaturdayM = True
        Else
            !FacilitySaturdayM = False
        End If
        !FacilitysaturdayMStarting = dtpSaturdayMorningStart.Value
        !FacilitysaturdayMEnding = dtpSaturdayMorningEnd.Value
        !FacilitysaturdayMNo = Val(txtSaturdayMorningMax.Text)
        If chkSaturdayEvening.Value = 1 Then
            !FacilitySaturdayE = True
        Else
            !FacilitySaturdayE = False
        End If
        !FacilitysaturdayEStarting = dtpSaturdayEveningStart.Value
        !FacilitysaturdayEEnding = dtpSaturdayEveningEnd.Value
        !FacilitysaturdayENo = Val(txtSaturdayEveningMax.Text)
'    .Update
    
        If chkSundayFullLeave.Value = 1 Then
            !FullDayLeaveSunday = True
        Else
            !FullDayLeaveSunday = False
        End If
        If chkSundayMorning.Value = 1 Then
            !FacilitySundayM = True
        Else
            !FacilitySundayM = False
        End If
        !FacilitysundayMStarting = dtpSundayMorningStart.Value
        !FacilitysundayMEnding = dtpSundayMorningEnd.Value
        !FacilitysundayMNo = Val(txtSundayMorningMax.Text)
        If chkSundayEvening.Value = 1 Then
            !FacilitySundayE = True
        Else
            !FacilitySundayE = False
        End If
        !FacilitysundayEStarting = dtpSundayEveningStart.Value
        !FacilitysundayEEnding = dtpSundayEveningEnd.Value
        !FacilitysundayENo = Val(txtSundayEveningMax.Text)
    
    .Update
    
    .Close
End With
'Call Clearvalues
'Call BeforeAddEdit
'Call FormatGrid
'Call FillGrid
End Sub

Private Sub EditData()
    With DataEnvironment1.rssqlFacilityStaff
        If .State = 1 Then .Close
        .Source = "SELECT tblFacilitystaff.* from tblFacilitystaff where facilitystaff_ID =" & TemStaffFacility
        If .State = 0 Then .Open

    !TwoSecessions = True
    !facilitystaff = "Channeling - " & DataComboDoctor.Text
    !HospitalFacility_id = 10
    !staff_ID = DataComboDoctor.BoundText
    
    !usualpersonalFee = Val(txtDoctorLocalFee.Text)
    !UsualInstitutionFee = Val(txtInstitutionLocalFee.Text)
    
    !foreignerpersonalfee = Val(txtDoctorForeignerFee.Text)
    !ForeignerinstitutionFee = Val(txtInstitutionForeignerFee.Text)
    
    !FacilityStaffComment = txtComments.Text
    !usualduration = Val(txtUsualDuration.Text)
    
        
'    .Update
    
        If chkMondayFullLeave.Value = 1 Then
            !FullDayLeaveMonday = True
        Else
            !FullDayLeaveMonday = False
        End If
        If chkMondayMorning.Value = 1 Then
            !FacilityMondayM = True
        Else
            !FacilityMondayM = False
        End If
        !FacilityMondayMStarting = dtpMondayMorningStart.Value
        !FacilityMondayMEnding = dtpMondayMorningEnd.Value
        !FacilityMondayMNo = Val(txtMondayMorningMax.Text)
        If chkMondayEvening.Value = 1 Then
            !FacilityMondayE = True
        Else
            !FacilityMondayE = False
        End If
        !FacilityMondayEStarting = dtpMondayEveningStart.Value
        !FacilityMondayEEnding = dtpMondayEveningEnd.Value
        !FacilityMondayENo = Val(txtMondayEveningMax.Text)
'    .Update
    
    
        If chkTuesdayFullLeave.Value = 1 Then
            !FullDayLeaveTuesday = True
        Else
            !FullDayLeaveTuesday = False
        End If
        If chkTuesdayMorning.Value = 1 Then
            !FacilityTuesdayM = True
        Else
            !FacilityTuesdayM = False
        End If
        !FacilitytuesdayMStarting = dtpTuesdayMorningStart.Value
        !FacilitytuesdayMEnding = dtpTuesdayMorningEnd.Value
        !FacilitytuesdayMNo = Val(txtTuesdayMorningMax.Text)
        If chkTuesdayEvening.Value = 1 Then
            !FacilityTuesdayE = True
        Else
            !FacilityTuesdayE = False
        End If
        !FacilitytuesdayEStarting = dtpTuesdayEveningStart.Value
        !FacilitytuesdayEEnding = dtpTuesdayEveningEnd.Value
        !FacilitytuesdayENo = Val(txtTuesdayEveningMax.Text)
    
'    .Update
    
        If chkWednesdayFullLeave.Value = 1 Then
            !FullDayLeaveWednesday = True
        Else
            !FullDayLeaveWednesday = False
        End If
        If chkWednesdayMorning.Value = 1 Then
            !FacilityWednesdayM = True
        Else
            !FacilityWednesdayM = False
        End If
        !FacilitywednesdayMStarting = dtpWednesdayMorningStart.Value
        !FacilitywednesdayMEnding = dtpWednesdayMorningEnd.Value
        !FacilitywednesdayMNo = Val(txtWednesdayMorningMax.Text)
        If chkWednesdayEvening.Value = 1 Then
            !FacilityWednesdayE = True
        Else
            !FacilityWednesdayE = False
        End If
        !FacilitywednesdayEStarting = dtpWednesdayEveningStart.Value
        !FacilitywednesdayEEnding = dtpWednesdayEveningEnd.Value
        !FacilitywednesdayENo = Val(txtWednesdayEveningMax.Text)
'    .Update
    
        If chkThursdayFullLeave.Value = 1 Then
            !FullDayLeaveThursday = True
        Else
            !FullDayLeaveThursday = False
        End If
        If chkThursdayMorning.Value = 1 Then
            !FacilityThursdayM = True
        Else
            !FacilityThursdayM = False
        End If
        !FacilitythursdayMStarting = dtpThursdayMorningStart.Value
        !FacilitythursdayMEnding = dtpThursdayMorningEnd.Value
        !FacilitythursdayMNo = Val(txtThursdayMorningMax.Text)
        If chkThursdayEvening.Value = 1 Then
            !FacilityThursdayE = True
        Else
            !FacilityThursdayE = False
        End If
        !FacilitythursdayEStarting = dtpThursdayEveningStart.Value
        !FacilitythursdayEEnding = dtpThursdayEveningEnd.Value
        !FacilitythursdayENo = Val(txtThursdayEveningMax.Text)
'    .Update
    
        If chkFridayFullLeave.Value = 1 Then
            !FullDayLeaveFriday = True
        Else
            !FullDayLeaveFriday = False
        End If
        If chkFridayMorning.Value = 1 Then
            !FacilityFridayM = True
        Else
            !FacilityFridayM = False
        End If
        !FacilityfridayMStarting = dtpFridayMorningStart.Value
        !FacilityfridayMEnding = dtpFridayMorningEnd.Value
        !FacilityfridayMNo = Val(txtFridayMorningMax.Text)
        If chkFridayEvening.Value = 1 Then
            !FacilityFridayE = True
        Else
            !FacilityFridayE = False
        End If
        !FacilityfridayEStarting = dtpFridayEveningStart.Value
        !FacilityfridayEEnding = dtpFridayEveningEnd.Value
        !FacilityfridayENo = Val(txtFridayEveningMax.Text)
'    .Update
    
        If chkSaturdayFullLeave.Value = 1 Then
            !FullDayLeaveSaturday = True
        Else
            !FullDayLeaveSaturday = False
        End If
        If chkSaturdayMorning.Value = 1 Then
            !FacilitySaturdayM = True
        Else
            !FacilitySaturdayM = False
        End If
        !FacilitysaturdayMStarting = dtpSaturdayMorningStart.Value
        !FacilitysaturdayMEnding = dtpSaturdayMorningEnd.Value
        !FacilitysaturdayMNo = Val(txtSaturdayMorningMax.Text)
        If chkSaturdayEvening.Value = 1 Then
            !FacilitySaturdayE = True
        Else
            !FacilitySaturdayE = False
        End If
        !FacilitysaturdayEStarting = dtpSaturdayEveningStart.Value
        !FacilitysaturdayEEnding = dtpSaturdayEveningEnd.Value
        !FacilitysaturdayENo = Val(txtSaturdayEveningMax.Text)
'    .Update
    
        If chkSundayFullLeave.Value = 1 Then
            !FullDayLeaveSunday = True
        Else
            !FullDayLeaveSunday = False
        End If
        If chkSundayMorning.Value = 1 Then
            !FacilitySundayM = True
        Else
            !FacilitySundayM = False
        End If
        !FacilitysundayMStarting = dtpSundayMorningStart.Value
        !FacilitysundayMEnding = dtpSundayMorningEnd.Value
        !FacilitysundayMNo = Val(txtSundayMorningMax.Text)
        If chkSundayEvening.Value = 1 Then
            !FacilitySundayE = True
        Else
            !FacilitySundayE = False
        End If
        !FacilitysundayEStarting = dtpSundayEveningStart.Value
        !FacilitysundayEEnding = dtpSundayEveningEnd.Value
        !FacilitysundayENo = Val(txtSundayEveningMax.Text)
    
            .Update
        .Close
    End With
'    Call Clearvalues
'    Call BeforeAddEdit
'    Call FormatGrid
'    Call FillGrid
End Sub

Private Sub Grid1_Click()
    FromGrid = True
    With grid1
        .Col = 1
        txtSearch.Text = .Text
        If .Row < 1 Then Exit Sub
        .Col = 2
        If Not IsNumeric(.Text) Then Exit Sub
        TemStaffFacility = Val(.Text)
        Call GetData
        .Col = 0
        .ColSel = .Cols - 1
        txtSearch.SetFocus
        SendKeys "{home}+{end}"
    FromGrid = False
    bttnAdd.Enabled = True
    bttnEdit.Enabled = True
    bttnDelete.Enabled = True
    End With
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then Grid1_Click
End Sub

Private Sub PrepareForDoctor()
On Error Resume Next
    With DataEnvironment1.rssqlDoctor
        If .State = 1 Then .Close
        .Source = "SELECT tbldoctor.* from tbldoctor order by doctorListedName"
        If .State = 0 Then .Open
    End With
    DataComboDoctor.RowMember = "sqlDoctor"
    DataComboDoctor.ListField = "DoctorListedName"
    DataComboDoctor.BoundColumn = "Doctor_ID"
End Sub


Private Sub TwoSecessions()
    chkMondayEvening.Enabled = True
    chkMondayMorning.Enabled = True
    dtpMondayEveningEnd.Enabled = True
    dtpMondayEveningStart.Enabled = True
    dtpMondayMorningEnd.Enabled = True
    dtpMondayMorningStart.Enabled = True
    txtMondayEveningMax.Enabled = True
    txtMondayMorningMax.Enabled = True
    chkTuesdayEvening.Enabled = True
    chkTuesdayMorning.Enabled = True
    dtpTuesdayEveningEnd.Enabled = True
    dtpTuesdayEveningStart.Enabled = True
    dtpTuesdayMorningEnd.Enabled = True
    dtpTuesdayMorningStart.Enabled = True
    txtTuesdayEveningMax.Enabled = True
    txtTuesdayMorningMax.Enabled = True
    chkWednesdayEvening.Enabled = True
    chkWednesdayMorning.Enabled = True
    dtpWednesdayEveningEnd.Enabled = True
    dtpWednesdayEveningStart.Enabled = True
    dtpWednesdayMorningEnd.Enabled = True
    dtpWednesdayMorningStart.Enabled = True
    txtWednesdayEveningMax.Enabled = True
    txtWednesdayMorningMax.Enabled = True
    chkThursdayEvening.Enabled = True
    chkThursdayMorning.Enabled = True
    dtpThursdayEveningEnd.Enabled = True
    dtpThursdayEveningStart.Enabled = True
    dtpThursdayMorningEnd.Enabled = True
    dtpThursdayMorningStart.Enabled = True
    txtThursdayEveningMax.Enabled = True
    txtThursdayMorningMax.Enabled = True
    chkFridayEvening.Enabled = True
    chkFridayMorning.Enabled = True
    dtpFridayEveningEnd.Enabled = True
    dtpFridayEveningStart.Enabled = True
    dtpFridayMorningEnd.Enabled = True
    dtpFridayMorningStart.Enabled = True
    txtFridayEveningMax.Enabled = True
    txtFridayMorningMax.Enabled = True
    chkSaturdayEvening.Enabled = True
    chkSaturdayMorning.Enabled = True
    dtpSaturdayEveningEnd.Enabled = True
    dtpSaturdayEveningStart.Enabled = True
    dtpSaturdayMorningEnd.Enabled = True
    dtpSaturdayMorningStart.Enabled = True
    txtSaturdayEveningMax.Enabled = True
    txtSaturdayMorningMax.Enabled = True
    chkSundayEvening.Enabled = True
    chkSundayMorning.Enabled = True
    dtpSundayEveningEnd.Enabled = True
    dtpSundayEveningStart.Enabled = True
    dtpSundayMorningEnd.Enabled = True
    dtpSundayMorningStart.Enabled = True
    txtSundayEveningMax.Enabled = True
    txtSundayMorningMax.Enabled = True
    
    chkMoning.Enabled = True
    chkEvening.Enabled = True
    txtMNo.Enabled = True
    txtENo.Enabled = True
    dtpMorningStarting.Enabled = True
    dtpMorningEnding.Enabled = True
    dtpEveningStarting.Enabled = True
    dtpEveningEnding.Enabled = True
        
End Sub

Private Sub OneSecession()
    chkMondayEvening.Enabled = False
    chkMondayMorning.Enabled = False
    dtpMondayEveningEnd.Enabled = False
    dtpMondayEveningStart.Enabled = False
    dtpMondayMorningEnd.Enabled = False
    dtpMondayMorningStart.Enabled = False
    txtMondayEveningMax.Enabled = False
    txtMondayMorningMax.Enabled = False
    chkTuesdayEvening.Enabled = False
    chkTuesdayMorning.Enabled = False
    dtpTuesdayEveningEnd.Enabled = False
    dtpTuesdayEveningStart.Enabled = False
    dtpTuesdayMorningEnd.Enabled = False
    dtpTuesdayMorningStart.Enabled = False
    txtTuesdayEveningMax.Enabled = False
    txtTuesdayMorningMax.Enabled = False
    chkWednesdayEvening.Enabled = False
    chkWednesdayMorning.Enabled = False
    dtpWednesdayEveningEnd.Enabled = False
    dtpWednesdayEveningStart.Enabled = False
    dtpWednesdayMorningEnd.Enabled = False
    dtpWednesdayMorningStart.Enabled = False
    txtWednesdayEveningMax.Enabled = False
    txtWednesdayMorningMax.Enabled = False
    chkThursdayEvening.Enabled = False
    chkThursdayMorning.Enabled = False
    dtpThursdayEveningEnd.Enabled = False
    dtpThursdayEveningStart.Enabled = False
    dtpThursdayMorningEnd.Enabled = False
    dtpThursdayMorningStart.Enabled = False
    txtThursdayEveningMax.Enabled = False
    txtThursdayMorningMax.Enabled = False
    chkFridayEvening.Enabled = False
    chkFridayMorning.Enabled = False
    dtpFridayEveningEnd.Enabled = False
    dtpFridayEveningStart.Enabled = False
    dtpFridayMorningEnd.Enabled = False
    dtpFridayMorningStart.Enabled = False
    txtFridayEveningMax.Enabled = False
    txtFridayMorningMax.Enabled = False
    chkSaturdayEvening.Enabled = False
    chkSaturdayMorning.Enabled = False
    dtpSaturdayEveningEnd.Enabled = False
    dtpSaturdayEveningStart.Enabled = False
    dtpSaturdayMorningEnd.Enabled = False
    dtpSaturdayMorningStart.Enabled = False
    txtSaturdayEveningMax.Enabled = False
    txtSaturdayMorningMax.Enabled = False
    chkSundayEvening.Enabled = False
    chkSundayMorning.Enabled = False
    dtpSundayEveningEnd.Enabled = False
    dtpSundayEveningStart.Enabled = False
    dtpSundayMorningEnd.Enabled = False
    dtpSundayMorningStart.Enabled = False
    txtSundayEveningMax.Enabled = False
    txtSundayMorningMax.Enabled = False
    
    chkMoning.Enabled = False
    chkEvening.Enabled = False
    txtMNo.Enabled = False
    txtENo.Enabled = False
    dtpMorningStarting.Enabled = False
    dtpMorningEnding.Enabled = False
    dtpEveningStarting.Enabled = False
    dtpEveningEnding.Enabled = False
    
    
End Sub

Private Sub txtSearch_Change()
    If FromGrid = True Then Exit Sub
    Dim TemFRows As Long
    Dim TemNowRow As Long
    Dim TemArray As Long
    Dim SearchSuccess As Boolean
    Dim TemLength As Single
    TemFRows = grid1.Rows
    grid1.Col = 1
    If Len(txtSearch.Text) = 0 Then GoTo MeasureSuccess
    SearchSuccess = False
    For TemArray = 1 To (TemFRows - 1)
        grid1.Row = TemArray
        If Len(txtSearch.Text) > Len(grid1.Text) Then
            GoTo FinishLoop
        Else
            TemLength = Len(txtSearch.Text)
        End If
        
        If UCase(Left((grid1.Text), TemLength)) = UCase(txtSearch.Text) Then
            SearchSuccess = True
            Exit For
        Else
            SearchSuccess = False
        End If
FinishLoop:
    Next
MeasureSuccess:
    If SearchSuccess = True Then
        grid1.TopRow = TemArray
        grid1.Row = TemArray
        grid1.Col = 0
        grid1.ColSel = (grid1.Cols - 1)
        bttnEdit.Enabled = True
        bttnDelete.Enabled = True
        bttnAdd.Enabled = True
        grid1.Col = 2
        TemStaffFacility = grid1.Text
        Call GetData
        grid1.Col = 0
        grid1.ColSel = grid1.Cols - 1
    Else
        grid1.TopRow = 1
        grid1.Row = 0
        grid1.Col = 0
        grid1.ColSel = 0
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
        bttnDelete.Enabled = False
    End If
End Sub

Private Sub FormatLeaveGrid()
    Dim BorderMargin As Integer
    BorderMargin = 150
    With Grid2
        .Clear
        .Cols = 4
        .Rows = 1
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
        .ColWidth(3) = 1
        .ColWidth(2) = .Width - (.ColWidth(0) + .ColWidth(1) + .ColWidth(3) + BorderMargin)
        .Row = 0
        .Col = 0
        .CellAlignment = 4
        .Text = "No."
        .Col = 1
        .CellAlignment = 4
        .Text = "Date"
        .Col = 2
        .CellAlignment = 4
        .Text = "Comments"
    End With
End Sub

Private Sub FillLeaveGrid()
    Dim NowRow As Long
    Dim TemText As String
    
    grid1.Col = 2
    If Not IsNumeric(grid1.Text) Then Exit Sub
    With DataEnvironment1.rssqlFacilityStaffLeave
        If .State = 1 Then .Close
        .Source = "Select tblFacilityStaffLeave.* from tblFacilityStaffLeave where (FacilityStaff_ID = " & TemStaffFacility & ") and (FacilityStaffLeaveDate >= #" & Date & "#) order by FacilityStaff_ID , FacilityStaffLeaveDate "
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        NowRow = 0
        Grid2.Rows = 1
        Grid2.Col = 0
        Grid2.WordWrap = True
        Grid2.ColSel = Grid2.Cols - 1
        
        While Not .EOF
            NowRow = NowRow + 1
            Grid2.Rows = NowRow + 1
            Grid2.Row = NowRow
            Grid2.Col = 0
            Grid2.CellAlignment = 7
            Grid2.Text = NowRow
            Grid2.Col = 1
            Grid2.CellAlignment = 4
            Grid2.Text = Format(!facilitystaffleavedate, "dd mmm yyyy")
            Grid2.Col = 2
            Grid2.Text = !facilitystaffleavecomments
            Grid2.Col = 3
            
            
            Grid2.Text = !FacilityStaffLeave_ID
            
            .MoveNext
        Wend
    End With
End Sub


Private Sub bttnAddLeave_Click()
    Dim temnumber As Long
    Dim TemDate As Date
    Dim TemResponce  As Integer
    Grid2.Col = 1
    For temnumber = 1 To Grid2.Rows - 1
        Grid2.Row = temnumber
        TemDate = Grid2.Text
        If dtpLeaveDate.Value = TemDate Then
            TemResponce = MsgBox("The date is already added", vbInformation, "Alredy Added")
            dtpLeaveDate.SetFocus
            Exit Sub
        End If
    Next
    With DataEnvironment1.rssqlFacilityStaffLeave
        If .State = 0 Then .Open
        .AddNew
        !facilitystaffleavedate = dtpLeaveDate.Value
        If chkFullDayLeave.Value = 1 Then
            !FullDayLeave = True
        Else
            !FullDayLeave = False
        End If
        If chkMoning.Value = 1 Then
            !Morning = True
        Else
            !Morning = False
        End If
        If chkEvening.Value = 1 Then
            !Evening = True
        Else
            !Evening = False
        End If
        !morningStarting = dtpMorningStarting.Value
        !morningending = dtpMorningEnding.Value
        !eveningstarting = dtpEveningStarting.Value
        !eveningending = dtpEveningEnding.Value
        !morningMax = Val(txtMNo.Text)
        !eveningmax = Val(txtENo.Text)
        !facilitystaffleavecomments = txtLeaveComments.Text
        !FacilityStaff_ID = TemStaffFacility
        .Update
        .Close
    End With
    Call FormatLeaveGrid
    Call FillLeaveGrid
End Sub

Private Sub bttnLeaveDelete_Click()
    If Grid2.Rows <= 1 Then Exit Sub
    Grid2.Col = 3
    If Not IsNumeric(Grid2.Text) Then Exit Sub
    With DataEnvironment1.rssqlFacilityStaffLeave
        If .State = 1 Then .Close
        .Source = "select tblfacilitystaffleave.* from tblfacilitystaffleave where (facilitystaffleave_id = " & Grid2.Text & ")"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .Delete adAffectCurrent
        .Close
    End With
    FormatLeaveGrid
    FillLeaveGrid
    Grid2.Col = 0
    Grid2.ColSel = Grid2.Cols - 1
End Sub


Private Sub Grid2_Click()
    If Grid2.Cols <= 1 Then Exit Sub
        bttnDelete.Enabled = True
        Grid2.Col = 0
        Grid2.ColSel = Grid2.Cols - 1
End Sub


Private Sub Setcolours()


Select Case ColourScheme

Case 1:

BttnBackColour = 5341695
BttnForeColour = 1314458
FrmBackColour = 11066623
FrmForeColour = 1314458
FrameBackColour = 11066623
FrameForeColour = 1314458
TxtBackColour = 9881851
TxtForeColour = 1314458
LblBackColour = 11066623
LblForeColour = 1314458



GridBackColor = 9881855
GridBackColorBkg = 10474239
GridBackColorFixed = 8566015
GridBackColorSel = 5341695

GridForeColor = 1314458
GridForeColorFixed = 11944
GridForeColorSel = 3014824

'GridCellBackColor = 5853695
'GridCellForeColor = 658120


Case 2:

BttnBackColour = 14803300
BttnForeColour = 5539362
FrmBackColour = 16766120
FrmForeColour = 5539362
FrameBackColour = 16766120
FrameForeColour = 5539362
TxtBackColour = 16760450
TxtForeColour = 5539362
LblBackColour = 16766120
LblForeColour = 5539362

GridBackColor = 16760450
GridBackColorBkg = 16771260
GridBackColorFixed = 16105620
GridBackColorSel = 16737380

GridForeColor = 5539362
GridForeColorFixed = 5539362
GridForeColorSel = 16765588


Case 3:

BttnBackColour = 51455
BttnForeColour = 942490
FrmBackColour = 11070719
FrmForeColour = 942490
FrameBackColour = 11070719
FrameForeColour = 942490
TxtBackColour = 11528439
TxtForeColour = 1314458
LblBackColour = 11070719
LblForeColour = 942490

GridBackColor = 16760450
GridBackColorBkg = 16771260
GridBackColorFixed = 16105620
GridBackColorSel = 16737380

GridForeColor = 5539362
GridForeColorFixed = 5539362
GridForeColorSel = 16765588

End Select

bttnAdd.BackColor = BttnBackColour
bttnAdd.ForeColor = BttnForeColour

bttnAddLeave.BackColor = BttnBackColour
bttnAddLeave.ForeColor = BttnForeColour


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

frmStaffFacilities.BackColor = FrmBackColour
frmStaffFacilities.ForeColor = FrmForeColour

SSTab1.BackColor = FrameBackColour
SSTab1.ForeColor = FrameForeColour

SSTabDates.BackColor = FrameBackColour
SSTabDates.ForeColor = FrameForeColour


framFacility.BackColor = FrameBackColour
framFacility.ForeColor = FrameForeColour

'FrameOfficial.BackColor = FrameBackColour
'FrameOfficial.ForeColor = FrameForeColour

'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour

'CheckFriday.BackColor = LblBackColour
'CheckFriday.ForeColor = LblForeColour
'CheckMonday.BackColor = LblBackColour
'CheckMonday.ForeColor = LblForeColour
'CheckSunday.BackColor = LblBackColour
'CheckSunday.ForeColor = LblForeColour
'CheckThursday.BackColor = LblBackColour
'CheckThursday.ForeColor = LblForeColour
'CheckTuesday.BackColor = LblBackColour
'CheckTuesday.ForeColor = LblForeColour
'CheckWednesday.BackColor = LblBackColour
'CheckWednesday.ForeColor = LblForeColour
'CheckSaturday.BackColor = LblBackColour
'CheckSaturday.ForeColor = LblForeColour
''CheckFriday.BackColor = LblBackColour
''CheckFriday.ForeColor = LblForeColour
'
'
'DataComboDoctor.BackColor = TxtBackColour
'DataComboDoctor.ForeColor = TxtForeColour
'
'DataComboFacility.BackColor = TxtBackColour
'DataComboFacility.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboSex.ForeColor = TxtForeColour
'
'DataComboSpeciality.BackColor = TxtBackColour
'DataComboSpeciality.ForeColor = TxtForeColour
'
'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour




grid1.BackColor = GridBackColor
grid1.ForeColor = GridForeColor

grid1.BackColorBkg = GridBackColorBkg
grid1.BackColorFixed = GridBackColorFixed
grid1.BackColorSel = GridBackColorSel

grid1.ForeColor = GridForeColor
grid1.ForeColorFixed = GridForeColorFixed
grid1.ForeColorSel = GridForeColorSel

'grid1.ForeColor = Grid



Label1.BackColor = LblBackColour
Label1.ForeColor = LblForeColour

'Label10.BackColor = LblBackColour
'Label10.ForeColor = LblForeColour
'Label11.BackColor = LblBackColour
'Label11.ForeColor = LblForeColour
'Label12.BackColor = LblBackColour
'Label12.ForeColor = LblForeColour
'Label13.BackColor = LblBackColour
'Label13.ForeColor = LblForeColour
'Label14.BackColor = LblBackColour
'Label14.ForeColor = LblForeColour
'Label15.BackColor = LblBackColour
'Label15.ForeColor = LblForeColour
Label16.BackColor = LblBackColour
Label16.ForeColor = LblForeColour
Label2.BackColor = LblBackColour
Label2.ForeColor = LblForeColour
'Label18.BackColor = LblBackColour
'Label18.ForeColor = LblForeColour
Label3.BackColor = LblBackColour
Label3.ForeColor = LblForeColour
'Label20.BackColor = LblBackColour
'Label20.ForeColor = LblForeColour
'Label21.BackColor = LblBackColour
'Label21.ForeColor = LblForeColour
'Label4.BackColor = LblBackColour
'Label4.ForeColor = LblForeColour
'Label23.BackColor = LblBackColour
'Label23.ForeColor = LblForeColour
'Label24.BackColor = LblBackColour
'Label24.ForeColor = LblForeColour
'Label25.BackColor = LblBackColour
'Label25.ForeColor = LblForeColour
'Label26.BackColor = LblBackColour
'Label26.ForeColor = LblForeColour
'Label27.BackColor = LblBackColour
'Label27.ForeColor = LblForeColour
txtUsualDuration.BackColor = LblBackColour
txtUsualDuration.ForeColor = LblForeColour
DataComboDoctor.BackColor = LblBackColour
DataComboDoctor.ForeColor = LblForeColour
Label6.BackColor = LblBackColour
Label6.ForeColor = LblForeColour
Label7.BackColor = LblBackColour
Label7.ForeColor = LblForeColour

Label17.BackColor = LblBackColour
Label17.ForeColor = LblForeColour



Label19.BackColor = LblBackColour
Label19.ForeColor = LblForeColour




LblDoctorStaff.BackColor = LblBackColour
LblDoctorStaff.ForeColor = LblForeColour

Label56.BackColor = LblBackColour
Label56.ForeColor = LblForeColour







lblDoctorStaffFee.BackColor = LblBackColour
lblDoctorStaffFee.ForeColor = LblForeColour

lblInstitutionFee.BackColor = LblBackColour
lblInstitutionFee.ForeColor = LblForeColour


txtComments.BackColor = TxtBackColour
txtComments.ForeColor = TxtForeColour

txtDoctorLocalFee.BackColor = TxtBackColour
txtDoctorLocalFee.ForeColor = TxtForeColour

txtInstitutionLocalFee.BackColor = TxtBackColour
txtInstitutionLocalFee.ForeColor = TxtForeColour
txtLeaveComments.BackColor = TxtBackColour
txtLeaveComments.ForeColor = TxtForeColour
'txtMaximumPerDay.BackColor = TxtBackColour
'txtMaximumPerDay.ForeColor = TxtForeColour
'txtName.BackColor = TxtBackColour
'txtListedName.ForeColor = TxtForeColour

'txtOfficialTel.BackColor = TxtBackColour
'txtOfficialTel.ForeColor = TxtForeColour
'txtOfficialWebsite.BackColor = TxtBackColour
'txtOfficialWebsite.ForeColor = TxtForeColour
'
'txtPrivateAddress.BackColor = TxtBackColour
'txtPrivateAddress.ForeColor = TxtForeColour
'txtPrivateEmail.BackColor = TxtBackColour
'txtPrivateEmail.ForeColor = TxtForeColour
'txtPrivateFax.BackColor = TxtBackColour
'txtPrivateFax.ForeColor = TxtForeColour
'txtPrivateMobile.BackColor = TxtBackColour
'txtPrivateMobile.ForeColor = TxtForeColour
'txtPrivateTel.BackColor = TxtBackColour
'txtPrivateTel.ForeColor = TxtForeColour
'
'
'txtQualifications.BackColor = TxtBackColour
'txtQualifications.ForeColor = TxtForeColour
'txtRegistation.BackColor = TxtBackColour
'txtRegistation.ForeColor = TxtForeColour
'txtSearch.BackColor = TxtBackColour
'txtSearch.ForeColor = TxtForeColour
'txtTel.ForeColor = TxtForeColour
'txtTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour




End Sub

