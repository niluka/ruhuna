VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPreferances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferances"
   ClientHeight    =   9270
   ClientLeft      =   4440
   ClientTop       =   1680
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7320
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Program Preferances"
      TabPicture(0)   =   "frmPreferances.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmTest"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameColour"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Printing Preferances"
      TabPicture(1)   =   "frmPreferances.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameAddForm"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "FramePrinting"
      Tab(1).Control(4)=   "bttnPrintingPositions"
      Tab(1).Control(5)=   "Frame6"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Owners Choices"
      TabPicture(2)   =   "frmPreferances.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "SliderIncomeDeflation"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Institution Preferances"
      TabPicture(3)   =   "frmPreferances.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "framInstutions"
      Tab(3).Control(1)=   "Frame8"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame5 
         Caption         =   "Booking Days in Advance"
         Height          =   975
         Left            =   240
         TabIndex        =   72
         Top             =   4920
         Width           =   3495
         Begin MSComctlLib.Slider SliderAdvancedBookingDays 
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            SelStart        =   3
            Value           =   3
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "10"
            Height          =   255
            Left            =   3120
            TabIndex        =   76
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            Height          =   255
            Left            =   1440
            TabIndex        =   75
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   600
            Width           =   135
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Need Agent Referance No"
         Height          =   975
         Left            =   240
         TabIndex        =   69
         Top             =   1680
         Width           =   3495
         Begin VB.OptionButton OptionNoNeedOfAgentReferanceNo 
            Caption         =   "No"
            Height          =   240
            Left            =   240
            TabIndex        =   71
            Top             =   600
            Width           =   2655
         End
         Begin VB.OptionButton OptionNeedAgentReferanceNo 
            Caption         =   "Yes"
            Height          =   240
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.Frame frameAddForm 
         Caption         =   "Add New Form"
         Height          =   3375
         Left            =   -67440
         TabIndex        =   43
         Top             =   720
         Width           =   6255
         Begin VB.TextBox txtFormRight 
            Height          =   360
            Left            =   3720
            TabIndex        =   60
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtFormLeft 
            Height          =   360
            Left            =   840
            TabIndex        =   59
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtFormBottom 
            Height          =   360
            Left            =   3720
            TabIndex        =   54
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtFormTop 
            Height          =   360
            Left            =   840
            TabIndex        =   53
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtFormName 
            Height          =   360
            Left            =   1800
            TabIndex        =   50
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox txtFormHeight 
            Height          =   360
            Left            =   840
            TabIndex        =   45
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtFormWidth 
            Height          =   360
            Left            =   3720
            TabIndex        =   44
            Top             =   840
            Width           =   1335
         End
         Begin btButtonEx.ButtonEx bttnAddForm 
            Height          =   495
            Left            =   1800
            TabIndex        =   52
            Top             =   2760
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            Appearance      =   3
            Caption         =   "Add Form"
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
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   5160
            TabIndex        =   64
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   2280
            TabIndex        =   63
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
            Height          =   375
            Left            =   3120
            TabIndex        =   62
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   5160
            TabIndex        =   58
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   2280
            TabIndex        =   57
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Bottom"
            Height          =   375
            Left            =   3120
            TabIndex        =   56
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Top"
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Form Name"
            Height          =   375
            Left            =   480
            TabIndex        =   51
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Height"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            Height          =   375
            Left            =   3120
            TabIndex        =   48
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   2280
            TabIndex        =   47
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   255
            Left            =   5160
            TabIndex        =   46
            Top             =   840
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Report Printer"
         Height          =   2655
         Left            =   -74760
         TabIndex        =   39
         Top             =   3360
         Width           =   7095
         Begin VB.ListBox ListReportPrinterPapers 
            Height          =   1500
            Left            =   1080
            TabIndex        =   67
            Top             =   840
            Width           =   5895
         End
         Begin VB.ListBox ListReportPrinterPapers1 
            Height          =   300
            Left            =   1080
            TabIndex        =   68
            Top             =   840
            Width           =   5895
         End
         Begin VB.ComboBox ComboReportPrinter 
            Height          =   360
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   360
            Width           =   5895
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Paper"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Printer"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Bill Printing"
         Height          =   2535
         Left            =   -74760
         TabIndex        =   35
         Top             =   720
         Width           =   7095
         Begin VB.ListBox ListBillPrinterPapers 
            Height          =   1500
            Left            =   1080
            TabIndex        =   65
            Top             =   840
            Width           =   5895
         End
         Begin VB.ListBox ListBillPrinterPapers1 
            Height          =   300
            Left            =   1080
            TabIndex        =   66
            Top             =   840
            Width           =   5895
         End
         Begin VB.ComboBox ComboBillPrinter 
            Height          =   360
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   360
            Width           =   5895
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Paper"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Printer"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ask before Adding"
         Height          =   975
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   3495
         Begin VB.OptionButton OptionAskBeforeAdding 
            Caption         =   "Yes"
            Height          =   240
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.OptionButton OptionDoNotAskBeforeAdding 
            Caption         =   "No"
            Height          =   240
            Left            =   240
            TabIndex        =   33
            Top             =   600
            Width           =   2655
         End
      End
      Begin VB.Frame FramePrinting 
         Caption         =   "Printing On"
         Height          =   975
         Left            =   -74760
         TabIndex        =   28
         Top             =   6120
         Width           =   2535
         Begin VB.OptionButton OptionBlankPaper 
            Caption         =   "Blank Papers"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton OptionPrintedPaper 
            Caption         =   "Printed Forms"
            Height          =   240
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame framInstutions 
         Height          =   4815
         Left            =   -72000
         TabIndex        =   8
         Top             =   1080
         Width           =   8655
         Begin VB.TextBox txtFax 
            Height          =   375
            Left            =   2040
            TabIndex        =   19
            Top             =   3240
            Width           =   6375
         End
         Begin VB.TextBox txtWebsite02 
            Height          =   375
            Left            =   5280
            TabIndex        =   18
            Top             =   4200
            Width           =   3135
         End
         Begin VB.TextBox txtwbsite01 
            Height          =   375
            Left            =   2040
            TabIndex        =   17
            Top             =   4200
            Width           =   3135
         End
         Begin VB.TextBox txtEmail02 
            Height          =   375
            Left            =   5280
            TabIndex        =   16
            Top             =   3720
            Width           =   3135
         End
         Begin VB.TextBox txtEmail01 
            Height          =   375
            Left            =   2040
            TabIndex        =   15
            Top             =   3720
            Width           =   3135
         End
         Begin VB.TextBox txtTelephone02 
            Height          =   375
            Left            =   5280
            TabIndex        =   14
            Top             =   2760
            Width           =   3135
         End
         Begin VB.TextBox txtTelephone01 
            Height          =   375
            Left            =   2040
            TabIndex        =   13
            Top             =   2760
            Width           =   3135
         End
         Begin VB.TextBox txtAddress01 
            Height          =   975
            Left            =   2040
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   1680
            Width           =   6375
         End
         Begin VB.TextBox txtRegistration 
            Height          =   375
            Left            =   2040
            TabIndex        =   11
            Top             =   1200
            Width           =   6375
         End
         Begin VB.TextBox txtDiscription 
            Height          =   375
            Left            =   2040
            TabIndex        =   10
            Top             =   720
            Width           =   6375
         End
         Begin VB.TextBox txtInsname 
            Height          =   375
            Left            =   2040
            TabIndex        =   9
            Top             =   240
            Width           =   6375
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "&Fax"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   3120
            Width           =   2295
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "&Website"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   4200
            Width           =   2295
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "&Email "
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   3720
            Width           =   2295
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Tele&phone No"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "&Address"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "&Registration No"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "&Discription"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Institution &Name"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame FrameColour 
         Caption         =   "Colour Scheme"
         Height          =   1455
         Left            =   7440
         TabIndex        =   3
         Top             =   3960
         Width           =   3615
         Begin VB.OptionButton OptionSunny 
            Caption         =   "Sunny"
            Height          =   240
            Left            =   240
            TabIndex        =   7
            Top             =   1095
            Width           =   1215
         End
         Begin VB.OptionButton OptionAqua 
            Caption         =   "Aqua"
            Height          =   240
            Left            =   240
            TabIndex        =   6
            Top             =   840
            Width           =   2535
         End
         Begin VB.OptionButton OptionEnergy 
            Caption         =   "Energy"
            Height          =   240
            Left            =   240
            TabIndex        =   5
            Top             =   600
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton OptionNoColour 
            Caption         =   "No Colour Scheme"
            Height          =   240
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   2295
         End
      End
      Begin MSComctlLib.Slider SliderIncomeDeflation 
         Height          =   435
         Left            =   -74640
         TabIndex        =   2
         Top             =   1320
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   767
         _Version        =   393216
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin btButtonEx.ButtonEx bttnPrintingPositions 
         Height          =   495
         Left            =   -72120
         TabIndex        =   31
         Top             =   6240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Appearance      =   3
         Caption         =   "Printing Positions"
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
      Begin VB.Frame frmTest 
         Height          =   7815
         Left            =   120
         TabIndex        =   77
         Top             =   360
         Width           =   14655
         Begin VB.Frame Frame24 
            Caption         =   "Hospital Details in Reports"
            Height          =   975
            Left            =   11040
            TabIndex        =   129
            Top             =   240
            Width           =   3495
            Begin VB.OptionButton OptionHospitalDetailsNo 
               Caption         =   "No"
               Height          =   240
               Left            =   240
               TabIndex        =   131
               Top             =   600
               Width           =   2535
            End
            Begin VB.OptionButton OptionHospitalDetailsYes 
               Caption         =   "Yes"
               Height          =   240
               Left            =   240
               TabIndex        =   130
               Top             =   360
               Value           =   -1  'True
               Width           =   2775
            End
         End
         Begin VB.Frame Frame23 
            Caption         =   "Absent / Present"
            Height          =   975
            Left            =   7320
            TabIndex        =   126
            Top             =   6720
            Width           =   3615
            Begin VB.OptionButton OptionAllowAbsent 
               Caption         =   "Allow to mark absent"
               Height          =   240
               Left            =   240
               TabIndex        =   128
               Top             =   360
               Value           =   -1  'True
               Width           =   3255
            End
            Begin VB.OptionButton OptionDoNotAllowAbsent 
               Caption         =   "Do not allow"
               Height          =   240
               Left            =   240
               TabIndex        =   127
               Top             =   600
               Width           =   3255
            End
         End
         Begin VB.Frame Frame22 
            Caption         =   "After adding a patient"
            Height          =   1455
            Left            =   7320
            TabIndex        =   121
            Top             =   5160
            Width           =   3615
            Begin VB.OptionButton OptionAfterAddSpeciality 
               Caption         =   "Focus on Speciality"
               Height          =   240
               Left            =   240
               TabIndex        =   125
               Top             =   360
               Width           =   3255
            End
            Begin VB.OptionButton OptionAfterAddConsultant 
               Caption         =   "Focus on Consultants"
               Height          =   240
               Left            =   240
               TabIndex        =   124
               Top             =   600
               Value           =   -1  'True
               Width           =   3255
            End
            Begin VB.OptionButton OptionAfterAddDates 
               Caption         =   "Focus on dates"
               Height          =   240
               Left            =   240
               TabIndex        =   123
               Top             =   840
               Width           =   3135
            End
            Begin VB.OptionButton OptionAfterAddPatient 
               Caption         =   "Focus on Adding Patients"
               Height          =   240
               Left            =   240
               TabIndex        =   122
               Top             =   1095
               Width           =   3255
            End
         End
         Begin VB.Frame Frame21 
            Caption         =   "Database"
            Height          =   1575
            Left            =   7320
            TabIndex        =   118
            Top             =   240
            Width           =   3615
            Begin VB.TextBox txtDatabase 
               Height          =   360
               Left            =   120
               TabIndex        =   119
               Top             =   240
               Width           =   3375
            End
            Begin btButtonEx.ButtonEx bttnSelectDatabasePath 
               Height          =   375
               Left            =   360
               TabIndex        =   120
               Top             =   840
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   661
               Appearance      =   3
               Caption         =   "Select Database"
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
         Begin VB.Frame Frame20 
            Caption         =   "Absent Patients' Fee"
            Height          =   975
            Left            =   3720
            TabIndex        =   115
            Top             =   5640
            Width           =   3495
            Begin VB.OptionButton Option2 
               Caption         =   "Do not Pay to doctor"
               Height          =   240
               Left            =   240
               TabIndex        =   117
               Top             =   600
               Width           =   2175
            End
            Begin VB.OptionButton OptionPayToDoctor 
               Caption         =   "Pay to doctor"
               Height          =   240
               Left            =   240
               TabIndex        =   116
               Top             =   360
               Value           =   -1  'True
               Width           =   2535
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "Default Backup Path"
            Height          =   1575
            Left            =   7320
            TabIndex        =   112
            Top             =   1920
            Width           =   3615
            Begin VB.TextBox txtPath 
               Height          =   360
               Left            =   120
               TabIndex        =   113
               Top             =   360
               Width           =   3375
            End
            Begin btButtonEx.ButtonEx bttnSelectBackupPath 
               Height          =   375
               Left            =   360
               TabIndex        =   114
               Top             =   960
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   661
               Appearance      =   3
               Caption         =   "Select Backup Path"
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
         Begin VB.Frame Frame18 
            Caption         =   "Reprints"
            Height          =   975
            Left            =   120
            TabIndex        =   109
            Top             =   6720
            Width           =   3495
            Begin VB.OptionButton OptionAllowReprints 
               Caption         =   "Allow"
               Height          =   240
               Left            =   240
               TabIndex        =   111
               Top             =   360
               Value           =   -1  'True
               Width           =   2535
            End
            Begin VB.OptionButton OptionDoNotAllowReprints 
               Caption         =   "Do not allow"
               Height          =   240
               Left            =   240
               TabIndex        =   110
               Top             =   600
               Width           =   2775
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "After adding a patient"
            Height          =   975
            Left            =   3720
            TabIndex        =   106
            Top             =   6720
            Width           =   3495
            Begin VB.OptionButton OptionDoNotClearAgentDetails 
               Caption         =   "Do not clear"
               Height          =   240
               Left            =   240
               TabIndex        =   108
               Top             =   600
               Width           =   2175
            End
            Begin VB.OptionButton OptionClearAgentDetails 
               Caption         =   "Clear Patient Details"
               Height          =   240
               Left            =   240
               TabIndex        =   107
               Top             =   240
               Value           =   -1  'True
               Width           =   2415
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Cash/Credit/Agent Selection"
            Height          =   975
            Left            =   120
            TabIndex        =   103
            Top             =   5640
            Width           =   3495
            Begin VB.OptionButton OptionChangeToCash 
               Caption         =   "Change to cash"
               Height          =   240
               Left            =   120
               TabIndex        =   105
               Top             =   360
               Value           =   -1  'True
               Width           =   3015
            End
            Begin VB.OptionButton OptionDoNotChangeToCash 
               Caption         =   "Remain in same selection"
               Height          =   240
               Left            =   120
               TabIndex        =   104
               Top             =   600
               Width           =   2775
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Selection"
            Height          =   975
            Left            =   120
            TabIndex        =   100
            Top             =   3480
            Width           =   3495
            Begin VB.OptionButton OptionCanSelectAgent 
               Caption         =   "Agent"
               Height          =   240
               Left            =   120
               TabIndex        =   102
               Top             =   360
               Value           =   -1  'True
               Width           =   3255
            End
            Begin VB.OptionButton OptionCanNotSelectAgent 
               Caption         =   "Agent Code only"
               Height          =   240
               Left            =   120
               TabIndex        =   101
               Top             =   600
               Width           =   3255
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Name Listing"
            Height          =   975
            Left            =   3720
            TabIndex        =   97
            Top             =   4560
            Width           =   3495
            Begin VB.OptionButton OptionNoAllNames 
               Caption         =   "Don't"
               Height          =   240
               Left            =   120
               TabIndex        =   99
               Top             =   600
               Width           =   2175
            End
            Begin VB.OptionButton OptionAllNames 
               Caption         =   "List All Names"
               Height          =   240
               Left            =   120
               TabIndex        =   98
               Top             =   360
               Value           =   -1  'True
               Width           =   2415
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Name Listing"
            Height          =   975
            Left            =   120
            TabIndex        =   94
            Top             =   2400
            Width           =   3495
            Begin VB.OptionButton OptionSurnameFirst 
               Caption         =   "Surname First"
               Height          =   240
               Left            =   240
               TabIndex        =   96
               Top             =   360
               Value           =   -1  'True
               Width           =   2655
            End
            Begin VB.OptionButton OptionFirstNameFirst 
               Caption         =   "First Name First"
               Height          =   240
               Left            =   240
               TabIndex        =   95
               Top             =   600
               Width           =   2295
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Agent Name For Credit Bookings"
            Height          =   975
            Left            =   3720
            TabIndex        =   91
            Top             =   3480
            Width           =   3495
            Begin VB.OptionButton OptionAgentNameForCreditBookings 
               Caption         =   "No"
               Height          =   240
               Left            =   240
               TabIndex        =   93
               Top             =   600
               Width           =   1935
            End
            Begin VB.OptionButton OptionAgentNameForCreditBookingsYes 
               Caption         =   "Yes"
               Height          =   240
               Left            =   240
               TabIndex        =   92
               Top             =   360
               Value           =   -1  'True
               Width           =   2775
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Patient Names Capitalization"
            Height          =   975
            Left            =   3720
            TabIndex        =   88
            Top             =   2400
            Width           =   3495
            Begin VB.OptionButton OptionAutomaticCapitalizationYes 
               Caption         =   "Yes"
               Height          =   240
               Left            =   240
               TabIndex        =   90
               Top             =   360
               Value           =   -1  'True
               Width           =   2535
            End
            Begin VB.OptionButton OptionAutomaticCapitalizationNo 
               Caption         =   "No"
               Height          =   240
               Left            =   240
               TabIndex        =   89
               Top             =   600
               Width           =   2655
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Add Foreigner as a suffix"
            Height          =   975
            Left            =   3720
            TabIndex        =   85
            Top             =   1320
            Width           =   3495
            Begin VB.OptionButton OptionForeignerSuffixNo 
               Caption         =   "No"
               Height          =   240
               Left            =   240
               TabIndex        =   87
               Top             =   600
               Width           =   2415
            End
            Begin VB.OptionButton OptionForeignerSuffixYes 
               Caption         =   "Yes"
               Height          =   240
               Left            =   240
               TabIndex        =   86
               Top             =   360
               Value           =   -1  'True
               Width           =   2895
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Allow Change of Patient Names"
            Height          =   975
            Left            =   3720
            TabIndex        =   82
            Top             =   240
            Width           =   3495
            Begin VB.OptionButton OptionAllowNameChange 
               Caption         =   "Yes"
               Height          =   240
               Left            =   240
               TabIndex        =   84
               Top             =   360
               Value           =   -1  'True
               Width           =   2775
            End
            Begin VB.OptionButton OptionDoNotAllowChangeOfNames 
               Caption         =   "No"
               Height          =   240
               Left            =   240
               TabIndex        =   83
               Top             =   600
               Width           =   2535
            End
         End
      End
      Begin VB.Frame Frame6 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   78
         Top             =   360
         Width           =   14655
      End
      Begin VB.Frame Frame7 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   79
         Top             =   360
         Width           =   14655
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Income Deflation"
            Height          =   255
            Left            =   360
            TabIndex        =   80
            Top             =   480
            Width           =   2415
         End
      End
      Begin VB.Frame Frame8 
         Height          =   7815
         Left            =   -74880
         TabIndex        =   81
         Top             =   360
         Width           =   14655
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   13080
      TabIndex        =   1
      Top             =   8640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Save / Exit"
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
Attribute VB_Name = "frmPreferances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim NumForms As Long, I As Long
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
    Dim SuppliedWord As String
    Dim FSys As New Scripting.FileSystemObject
    Private cSetPrinter As New cSetDfltPrinter
    
    Private Const BIF_RETURNONLYFSDIRS = 1
    Private Const BIF_DONTGOBELOWDOMAIN = 2
    Private Const MAX_PATH = 260
    
    Private Declare Function SHBrowseForFolder Lib "shell32" _
                                      (lpbi As BrowseInfo) As Long
    
    Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                      (ByVal pidList As Long, _
                                      ByVal lpBuffer As String) As Long
    
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                      (ByVal lpString1 As String, ByVal _
                                      lpString2 As String) As Long
    
    Private Type BrowseInfo
       hWndOwner      As Long
       pIDLRoot       As Long
       pszDisplayName As Long
       lpszTitle      As Long
       ulFlags        As Long
       lpfnCallback   As Long
       lparam         As Long
       iImage         As Long
    End Type


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
    
    bttnAddForm.BackColor = BttnBackColour
    bttnAddForm.ForeColor = BttnForeColour
    bttnPrintingPositions.BackColor = BttnBackColour
    bttnPrintingPositions.ForeColor = BttnForeColour
    bttnSelectDatabasePath.BackColor = BttnBackColour
    bttnSelectDatabasePath.ForeColor = BttnForeColour
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
    bttnSelectBackupPath.BackColor = BttnBackColour
    bttnSelectBackupPath.ForeColor = BttnForeColour
    frmTest.BackColor = FrameBackColour
    frmTest.ForeColor = FrameForeColour
    frmPreferances.BackColor = FrameBackColour
    frmPreferances.ForeColor = FrameForeColour
    frameAddForm.BackColor = FrameBackColour
    frameAddForm.ForeColor = FrameForeColour
    FrameColour.BackColor = FrameBackColour
    FrameColour.ForeColor = FrameForeColour
    FramePrinting.BackColor = FrameBackColour
    FramePrinting.ForeColor = FrameForeColour
    framInstutions.BackColor = FrameBackColour
    framInstutions.ForeColor = FrameForeColour
    Frame1.BackColor = FrameBackColour
    Frame1.ForeColor = FrameForeColour
    Frame2.BackColor = FrameBackColour
    Frame2.ForeColor = FrameForeColour
    Frame3.BackColor = FrameBackColour
    Frame3.ForeColor = FrameForeColour
    Frame4.BackColor = FrameBackColour
    Frame4.ForeColor = FrameForeColour
    Frame5.BackColor = FrameBackColour
    Frame5.ForeColor = FrameForeColour
    Frame6.BackColor = FrameBackColour
    Frame6.ForeColor = FrameForeColour
    Frame7.BackColor = FrameBackColour
    Frame7.ForeColor = FrameForeColour
    Frame8.BackColor = FrameBackColour
    Frame8.ForeColor = FrameForeColour
    
    Frame9.BackColor = FrameBackColour
    Frame9.ForeColor = FrameForeColour
    Frame10.BackColor = FrameBackColour
    Frame10.ForeColor = FrameForeColour
    Frame11.BackColor = FrameBackColour
    Frame11.ForeColor = FrameForeColour
    Frame12.BackColor = FrameBackColour
    Frame12.ForeColor = FrameForeColour
    Frame13.BackColor = FrameBackColour
    Frame13.ForeColor = FrameForeColour
    Frame14.BackColor = FrameBackColour
    Frame14.ForeColor = FrameForeColour
    Frame15.BackColor = FrameBackColour
    Frame15.ForeColor = FrameForeColour
    Frame16.BackColor = FrameBackColour
    Frame16.ForeColor = FrameForeColour
    
    Frame17.BackColor = FrameBackColour
    Frame17.ForeColor = FrameForeColour
    Frame18.BackColor = FrameBackColour
    Frame18.ForeColor = FrameForeColour
    Frame19.BackColor = FrameBackColour
    Frame19.ForeColor = FrameForeColour
    
    Frame21.BackColor = FrameBackColour
    Frame21.ForeColor = FrameForeColour
    
    Frame20.BackColor = FrameBackColour
    Frame20.ForeColor = FrameForeColour
    
    OptionPayToDoctor.BackColor = FrameBackColour
    OptionPayToDoctor.ForeColor = FrameForeColour
    
    Option2.BackColor = FrameBackColour
    Option2.ForeColor = FrameForeColour
    
    
    
    OptionNoColour.BackColor = FrameBackColour
    OptionNoColour.ForeColor = FrameForeColour
    OptionAqua.BackColor = FrameBackColour
    OptionAqua.ForeColor = FrameForeColour
    OptionAskBeforeAdding.BackColor = FrameBackColour
    OptionAskBeforeAdding.ForeColor = FrameForeColour
    OptionBlankPaper.BackColor = FrameBackColour
    OptionBlankPaper.ForeColor = FrameForeColour
    OptionDoNotAskBeforeAdding.BackColor = FrameBackColour
    OptionDoNotAskBeforeAdding.ForeColor = FrameForeColour
    OptionEnergy.BackColor = FrameBackColour
    OptionEnergy.ForeColor = FrameForeColour
    OptionNeedAgentReferanceNo.BackColor = FrameBackColour
    OptionNeedAgentReferanceNo.ForeColor = FrameForeColour
    OptionNoNeedOfAgentReferanceNo.BackColor = FrameBackColour
    OptionNoNeedOfAgentReferanceNo.ForeColor = FrameForeColour
    OptionPrintedPaper.BackColor = FrameBackColour
    OptionPrintedPaper.ForeColor = FrameForeColour
    OptionSunny.BackColor = FrameBackColour
    OptionSunny.ForeColor = FrameForeColour
    OptionAutomaticCapitalizationNo.BackColor = FrameBackColour
    OptionAutomaticCapitalizationNo.ForeColor = FrameForeColour
    OptionAutomaticCapitalizationYes.BackColor = FrameBackColour
    OptionAutomaticCapitalizationYes.ForeColor = FrameForeColour
    OptionAllowNameChange.BackColor = FrameBackColour
    OptionAllowNameChange.ForeColor = FrameForeColour
    OptionDoNotAllowChangeOfNames.BackColor = FrameBackColour
    OptionDoNotAllowChangeOfNames.ForeColor = FrameForeColour
    
    
    OptionSurnameFirst.BackColor = FrameBackColour
    OptionSurnameFirst.ForeColor = FrameForeColour
    OptionFirstNameFirst.BackColor = FrameBackColour
    OptionFirstNameFirst.ForeColor = FrameForeColour
    OptionCanSelectAgent.BackColor = FrameBackColour
    OptionCanSelectAgent.ForeColor = FrameForeColour
    OptionCanNotSelectAgent.BackColor = FrameBackColour
    OptionCanNotSelectAgent.ForeColor = FrameForeColour
    OptionForeignerSuffixYes.BackColor = FrameBackColour
    OptionForeignerSuffixYes.ForeColor = FrameForeColour
    OptionAgentNameForCreditBookingsYes.BackColor = FrameBackColour
    OptionAgentNameForCreditBookingsYes.ForeColor = FrameForeColour
    OptionAgentNameForCreditBookings.BackColor = FrameBackColour
    OptionAgentNameForCreditBookings.ForeColor = FrameForeColour
    OptionAllNames.BackColor = FrameBackColour
    OptionAllNames.ForeColor = FrameForeColour
    OptionNoAllNames.BackColor = FrameBackColour
    OptionNoAllNames.ForeColor = FrameForeColour
    OptionChangeToCash.BackColor = FrameBackColour
    OptionChangeToCash.ForeColor = FrameForeColour
    OptionDoNotChangeToCash.BackColor = FrameBackColour
    OptionDoNotChangeToCash.ForeColor = FrameForeColour
    OptionClearAgentDetails.BackColor = FrameBackColour
    OptionClearAgentDetails.ForeColor = FrameForeColour
    OptionDoNotClearAgentDetails.BackColor = FrameBackColour
    OptionDoNotClearAgentDetails.ForeColor = FrameForeColour
    OptionAllowReprints.BackColor = FrameBackColour
    OptionAllowReprints.ForeColor = FrameForeColour
    OptionDoNotAllowReprints.BackColor = FrameBackColour
    OptionDoNotAllowReprints.ForeColor = FrameForeColour
    OptionForeignerSuffixNo.BackColor = FrameBackColour
    OptionForeignerSuffixNo.ForeColor = FrameForeColour


End Sub

Private Sub bttnSelectBackupPath_Click()
         Dim lpIDList As Long
         Dim sBuffer As String
         Dim szTitle As String
         Dim tBrowseInfo As BrowseInfo
         szTitle = "Select Backup Directory"
         With tBrowseInfo
            .hWndOwner = Me.hwnd
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
         End With
         lpIDList = SHBrowseForFolder(tBrowseInfo)
         If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            txtPath.Text = sBuffer
         End If
End Sub

Private Sub Form_Load()
    Dim ingRet As Long
    Dim TabPrinter(2) As Long
    TabPrinter(0) = 48
    TabPrinter(1) = 78
    ingRet = SendMessage(ListBillPrinterPapers.hwnd, LB_SETTABSTOPS, 2, TabPrinter(0))
    ingRet = SendMessage(ListReportPrinterPapers.hwnd, LB_SETTABSTOPS, 2, TabPrinter(0))
    If UserAuthority = AuthorityAdministrator Then
        txtInsname.Locked = False
    Else
        txtInsname.Locked = True
    End If
    Call PopulatePrinters
    Call PopulateBillPrinterPapers
    Call PopulateReportPrinterPapers
    Call SetPreferances
    Call Setcolours
    SSTab1.Tab = 0
End Sub

Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        ComboBillPrinter.AddItem MyPrinter.DeviceName
        ComboReportPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub PopulateBillPrinterPapers()
    ListBillPrinterPapers.Clear
    ListBillPrinterPapers1.Clear
    SetPrinter = False
    cSetPrinter.SetPrinterAsDefault (BillPrinterName)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
            .cx = BillPaperHeight
            .cy = BillPaperWidth
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For I = 0 To NumForms - 1
            With aFI1(I)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                ListBillPrinterPapers1.AddItem PtrCtoVbString(.pName)
                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next I
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub PopulateReportPrinterPapers()
    ListReportPrinterPapers.Clear
    ListReportPrinterPapers1.Clear
    SetPrinter = False
    cSetPrinter.SetPrinterAsDefault (ReportPaperName)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
            .cx = ReportPaperWidth
            .cy = ReportPaperHeight
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For I = 0 To NumForms - 1
            With aFI1(I)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboReportPrinterPapers.AddItem FormItem
                ListReportPrinterPapers1.AddItem PtrCtoVbString(.pName)
                ListReportPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next I
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub bttnAddForm_Click()
    Dim TemResponce As Long
    If Trim(txtFormName.Text) = "" Then
        TemResponce = MsgBox("You have not enter a valid name for the form", vbCritical, "No name")
        txtFormName.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtFormHeight.Text) Then
        TemResponce = MsgBox("You have not entered a valid height in millimeters for the height of the form", vbCritical, "No Height")
        txtFormHeight.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtFormWidth.Text) Then
        TemResponce = MsgBox("You have not entered a valid width in millimeters for the width of the form", vbCritical, "No Width")
        txtFormWidth.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtFormTop.Text) Then
        TemResponce = MsgBox("You have not entered a valid top margin in millimeters for the height of the form", vbCritical, "No Top Margin")
        txtFormTop.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtFormBottom.Text) Then
        TemResponce = MsgBox("You have not entered a valid bottom margin in millimeters for the width of the form", vbCritical, "No Bottom Margin")
        txtFormBottom.SetFocus
        Exit Sub
    End If
    
     If Not IsNumeric(txtFormRight.Text) Then
        TemResponce = MsgBox("You have not entered a valid right margin in millimeters for the height of the form", vbCritical, "No Right Margin")
        txtFormRight.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtFormLeft.Text) Then
        TemResponce = MsgBox("You have not entered a valid left margin in millimeters for the width of the form", vbCritical, "No Left Margin")
        txtFormLeft.SetFocus
        Exit Sub
    End If
   
    
    Dim TemFormName As String
    Dim PrinterHandle As Long   ' Handle to printer
    
    If OpenPrinter(Printer.DeviceName, PrinterHandle, 0&) Then
        TemFormName = AddMyNewForm(PrinterHandle, Trim(txtFormName.Text), Val(txtFormHeight.Text) * 1000, Val(txtFormWidth.Text) * 1000, Val(txtFormBottom.Text) * 1000, Val(txtFormTop.Text) * 1000, Val(txtFormLeft.Text) * 1000, Val(txtFormRight.Text) * 1000)
        
        If TemFormName <> "none" Then
            TemResponce = MsgBox("The new form was added", vbInformation, "Added")
            Call PopulatePrinters
            Call PopulateBillPrinterPapers
            Call PopulateReportPrinterPapers
        End If
    End If
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub SetPreferances()
    Dim TemResponce As Integer
    
    If FSys.FileExists(DatabasePath) = True Then
        txtDatabase.Text = DatabasePath
        Call GetInstitutionDetails
    Else
        txtDatabase.Text = "You have not selected a valid database"
        txtDatabase.ForeColor = vbYellow
        txtDatabase.BackColor = vbRed
    End If
    
    SliderIncomeDeflation.Value = IncomeDeflation
    Select Case ColourScheme
        Case NoColourScheme: OptionNoColour.Value = True
        Case ColourAqua: OptionAqua.Value = True
        Case ColourEnergy: OptionEnergy.Value = True
        Case ColourSunny: OptionSunny.Value = True
    End Select
    If AskBeforeAdding = True Then OptionAskBeforeAdding.Value = True
    If AgentEssential = True Then
        OptionNeedAgentReferanceNo.Value = True
    Else
        OptionNoNeedOfAgentReferanceNo.Value = True
    End If
            
    OptionBlankPaper.Value = PrintingOnBlankPaper
    OptionPrintedPaper.Value = PrintingOnPrintedPaper
    
    OptionNoAllNames.Value = NoAllNames
    OptionSurnameFirst.Value = SurnameFirst
        
    SliderAdvancedBookingDays.Value = AdvanceBookingDays
    
    OptionForeignerSuffixYes.Value = AddForeignerSuffix
    OptionAllowNameChange.Value = AllowNameChange
    OptionAutomaticCapitalizationYes.Value = AutomaticCapitalization
    OptionAgentNameForCreditBookingsYes.Value = AgentNameForCreditBookings
    OptionCanSelectAgent.Value = CanSelectAgent
    OptionChangeToCash.Value = ChangeToCash
    OptionClearAgentDetails.Value = ClearAgentDetails
    OptionAllowReprints.Value = AllowReprint
    txtPath.Text = BackUpPath
    OptionPayToDoctor.Value = PayToDoctor
    OptionHospitalDetailsYes.Value = HospitalDetails
    
    On Error GoTo ErrBillPrinter
    ComboBillPrinter.Text = BillPrinterName
    
    On Error GoTo ErrBillPrinterPaper
    ListBillPrinterPapers1.Text = BillPaperName
    ListBillPrinterPapers.ListIndex = ListBillPrinterPapers1.ListIndex
    
    On Error GoTo ErrReportPrinter
    ComboReportPrinter.Text = ReportPrinterName
    
    On Error GoTo ErrReportPrinterPaper
    ListReportPrinterPapers1.Text = ReportPaperName
    ListReportPrinterPapers.ListIndex = ListReportPrinterPapers1.ListIndex
    
    Exit Sub
    
    
ErrBillPrinter:
    If Err.Number = 383 Then
        TemResponce = MsgBox("The Bill printer you have selected is not available now. Please select another printer", vbCritical, "New Bill Printer")
        If ComboBillPrinter.ListCount <> 0 Then ComboBillPrinter.ListIndex = 0
    Else
        TemResponce = MsgBox("An unknown error occured. Please contact Lakmedipro (077 3177874) with following details" & vbNewLine & Err.Description & vbNewLine & Err.Number & vbNewLine & Me.Caption, vbCritical, "Unknown Error")
    End If
    Exit Sub

ErrBillPrinterPaper:
    If Err.Number = 383 Then
        TemResponce = MsgBox("The Bill printer paper you have selected is not available now. Please select another printer", vbCritical, "New Bill Printer Paper")
        If ListBillPrinterPapers.ListCount <> 0 Then ListBillPrinterPapers.ListIndex = 0
    Else
        TemResponce = MsgBox("An unknown error occured. Please contact Lakmedipro (077 3177874) with following details" & vbNewLine & Err.Description & vbNewLine & Err.Number & vbNewLine & Me.Caption, vbCritical, "Unknown Error")
    End If
    Exit Sub


ErrReportPrinter:
    If Err.Number = 383 Then
        TemResponce = MsgBox("The Report printer you have selected is not available now. Please select another printer", vbCritical, "New Bill Printer")
        If ComboReportPrinter.ListCount <> 0 Then ComboReportPrinter.ListIndex = 0
    Else
        TemResponce = MsgBox("An unknown error occured. Please contact Lakmedipro (077 3177874) with following details" & vbNewLine & Err.Description & vbNewLine & Err.Number & vbNewLine & Me.Caption, vbCritical, "Unknown Error")
    End If
    Exit Sub
    
ErrReportPrinterPaper:
    If Err.Number = 383 Then
        TemResponce = MsgBox("The Report printer paper you have selected is not available now. Please select another printer", vbCritical, "New Bill Printer Paper")
        If ListReportPrinterPapers.ListCount <> 0 Then ListReportPrinterPapers.ListIndex = 0
    Else
        TemResponce = MsgBox("An unknown error occured. Please contact Lakmedipro (077 3177874) with following details" & vbNewLine & Err.Description & vbNewLine & Err.Number & vbNewLine & Me.Caption, vbCritical, "Unknown Error")
    End If
    Exit Sub
    
    Exit Sub
    
    
    
End Sub

Private Sub SetDefaults()
    IncomeDeflation = 5
    ColourScheme = NoColourScheme
    AdvanceBookingDays = 3
    AskBeforeAdding = True
End Sub

Private Sub SavePreferancesToFile()

    If OptionNoColour.Value = True Then
        SaveSetting App.EXEName, "Options", "ColourScheme", NoColourScheme
    ElseIf OptionAqua.Value = True Then
        SaveSetting App.EXEName, "Options", "ColourScheme", ColourAqua
    ElseIf OptionEnergy.Value = True Then
        SaveSetting App.EXEName, "Options", "ColourScheme", ColourEnergy
    ElseIf OptionSunny.Value = True Then
        SaveSetting App.EXEName, "Options", "ColourScheme", ColourSunny
    End If
    
    SaveSetting App.EXEName, "Options", "BackupPath", txtPath.Text
    SaveSetting App.EXEName, "Options", "BillPrinterName", ComboBillPrinter.Text
    SaveSetting App.EXEName, "Options", "BillPaperName", ListBillPrinterPapers1.Text
    SaveSetting App.EXEName, "Options", "ReportPrinterName", ComboReportPrinter.Text
    SaveSetting App.EXEName, "Options", "ReportPaperName", ListReportPrinterPapers1.Text
    SaveSetting App.EXEName, "Options", "IncomeDeflation", SliderIncomeDeflation.Value
    SaveSetting App.EXEName, "Options", "AdvanceBookingDays", SliderAdvancedBookingDays.Value
    SaveSetting App.EXEName, "Options", "PrintingOnBlankPaper", OptionBlankPaper.Value
    SaveSetting App.EXEName, "Options", "PrintingOnPrintedPaper", OptionPrintedPaper.Value
    SaveSetting App.EXEName, "Options", "AskBeforeAdding", OptionAskBeforeAdding.Value
    SaveSetting App.EXEName, "Options", "agentessential", OptionNeedAgentReferanceNo.Value
    SaveSetting App.EXEName, "Options", "BillPrinterName", ComboBillPrinter.Text
    SaveSetting App.EXEName, "Options", "BillPaperName", ListBillPrinterPapers1.Text         ' Mid(ComboBillPrinterPapers.Text, 1, InStr(1, ComboBillPrinterPapers.Text, " -") - 1)                                          '   ComboBillPrinterPapers.Text
    SaveSetting App.EXEName, "Options", "ReportPrinterName", ComboReportPrinter.Text
    SaveSetting App.EXEName, "Options", "ReportPaperName", ListReportPrinterPapers1.Text   ' Mid(ComboReportPrinterPapers.Text, 1, InStr(1, ComboReportPrinterPapers.Text, " -") - 1)                                          '   ComboBillPrinterPapers.Text
    SaveSetting App.EXEName, "Options", "AllowNameChange", OptionAllowNameChange.Value
    SaveSetting App.EXEName, "Options", "AddForeignerSuffix", OptionForeignerSuffixYes.Value
    SaveSetting App.EXEName, "Options", "AutomaticCapitalization", OptionAutomaticCapitalizationYes.Value
    SaveSetting App.EXEName, "Options", "AgentNameForCreditBookings", OptionAgentNameForCreditBookingsYes.Value
    SaveSetting App.EXEName, "Options", "NoAllNames", OptionNoAllNames.Value
    SaveSetting App.EXEName, "Options", "SurnameFirst", OptionSurnameFirst.Value
    SaveSetting App.EXEName, "Options", "CanSelectAgent", OptionCanSelectAgent.Value
    SaveSetting App.EXEName, "Options", "ChangeToCash", OptionChangeToCash.Value
    SaveSetting App.EXEName, "Options", "ClearAgentDetails", OptionClearAgentDetails.Value
    SaveSetting App.EXEName, "Options", "AllowReprint", OptionAllowReprints.Value
    SaveSetting App.EXEName, "Options", "BackUpPath", txtPath.Text
    SaveSetting App.EXEName, "Options", "PayToDoctor", OptionPayToDoctor.Value
    SaveSetting App.EXEName, "Options", "AllowAbsent", OptionAllowAbsent.Value
    SaveSetting App.EXEName, "Options", "AfterAddSpeciality", OptionAfterAddSpeciality.Value
    SaveSetting App.EXEName, "Options", "AfterAddConsultant", OptionAfterAddConsultant.Value
    SaveSetting App.EXEName, "Options", "AfterAddDates", OptionAfterAddDates.Value
    SaveSetting App.EXEName, "options", "AfterAddPatient", OptionAfterAddPatient.Value
    SaveSetting App.EXEName, "Options", "HospitalDetails", OptionHospitalDetailsYes.Value
    
    
    Call SaveInstitutionDetails
End Sub

Private Sub SavePreferancesToMemory()
    If OptionNoColour.Value = True Then
        PreferanceColourScheme = NoColourScheme
    ElseIf OptionAqua.Value = True Then
        PreferanceColourScheme = ColourAqua
    ElseIf OptionEnergy.Value = True Then
        PreferanceColourScheme = ColourEnergy
    ElseIf OptionSunny.Value = True Then
        PreferanceColourScheme = ColourSunny
    End If
    BillPrinterName = ComboBillPrinter.Text
    BillPaperName = ListBillPrinterPapers1.Text
    ReportPrinterName = ComboReportPrinter.Text
    ReportPaperName = ListReportPrinterPapers1.Text
    IncomeDeflation = SliderIncomeDeflation.Value
    AdvanceBookingDays = SliderAdvancedBookingDays.Value
    PrintingOnBlankPaper = OptionBlankPaper.Value
    PrintingOnPrintedPaper = OptionPrintedPaper.Value
    AskBeforeAdding = OptionAskBeforeAdding.Value
    AgentEssential = OptionNeedAgentReferanceNo.Value
    BillPrinterName = ComboBillPrinter.Text
    BillPaperName = ListBillPrinterPapers1.Text
    ReportPrinterName = ComboReportPrinter.Text
    ReportPaperName = ListReportPrinterPapers1.Text
    AllowNameChange = OptionAllowNameChange.Value
    AddForeignerSuffix = OptionForeignerSuffixYes.Value
    AutomaticCapitalization = OptionAutomaticCapitalizationYes.Value
    AgentNameForCreditBookings = OptionAgentNameForCreditBookingsYes.Value
    NoAllNames = OptionNoAllNames.Value
    SurnameFirst = OptionSurnameFirst.Value
    CanSelectAgent = OptionCanSelectAgent.Value
    ChangeToCash = OptionChangeToCash.Value
    ClearAgentDetails = OptionClearAgentDetails.Value
    AllowReprint = OptionAllowReprints.Value
    BackUpPath = txtPath.Text
    PayToDoctor = OptionPayToDoctor.Value
    AllowAbsent = OptionAllowAbsent.Value
    AfterAddSpeciality = OptionAfterAddSpeciality.Value
    AfterAddConsultant = OptionAfterAddConsultant.Value
    AfterAddDates = OptionAfterAddDates.Value
    AfterAddPatient = OptionAfterAddPatient.Value
    HospitalDetails = OptionHospitalDetailsYes.Value
    
    Call SaveInstitutionDetails
End Sub

Private Sub bttnPrintingPositions_Click()
    frmPrintingPositions.Show
    frmPrintingPositions.ZOrder 0
End Sub

Private Sub bttnSelectDatabasePath_Click()
    
    CommonDialog1.FileName = GetSetting(App.EXEName, "Options", "DatabaseLocation", App.Path & "\hospital.mdb")
    
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "mdb"
    CommonDialog1.Filter = "Lakmedipro Database|hospital.mdb"
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        txtDatabase.Text = CommonDialog1.FileName
        SaveSetting App.EXEName, "Options", "DatabaseLocation", txtDatabase.Text
        DatabasePath = txtDatabase.Text
    Else
        MsgBox "You have not selected valid database. The program may not function", vbCritical, "No database"
    End If
End Sub

Private Sub ComboBillPrinter_Change()
    cSetPrinter.SetPrinterAsDefault (ComboBillPrinter.Text)
    Call PopulateBillPrinterPapers
End Sub

Private Sub ComboReportPrinter_Change()
    cSetPrinter.SetPrinterAsDefault (ComboReportPrinter.Text)
    Call PopulateReportPrinterPapers
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim TemResponce As Integer
If FSys.FileExists(txtDatabase.Text) = False Then
    MsgBox "You have not selected a valid database", vbCritical, "Database?"
    Cancel = True
    txtDatabase.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SavePreferancesToFile
    Call SavePreferancesToMemory
End Sub

Private Sub ListBillPrinterPapers_Click()
    ListBillPrinterPapers1.ListIndex = ListBillPrinterPapers.ListIndex
End Sub

Private Sub ListReportPrinterPapers_Click()
    ListReportPrinterPapers1.ListIndex = ListReportPrinterPapers.ListIndex
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    Dim TemResponce  As Integer
    If SSTab1.Tab = 2 Then
        If UserAuthority <> AuthorityOwner And UserAuthority <> AuthorityAdministrator Then SSTab1.Tab = PreviousTab
        If UserAuthority = AuthorityUser Then
            TemResponce = MsgBox("Only Owners are allowed alter these settings", vbInformation, "No Authority")
            Exit Sub
        ElseIf UserAuthority = AuthorityHumanResources Then
            TemResponce = MsgBox("Only Owners are allowed alter these settings", vbInformation, "No Authority")
            Exit Sub
        ElseIf UserAuthority = AuthorityUser Then
            TemResponce = MsgBox("Only Owners are allowed alter these settings", vbInformation, "No Authority")
            Exit Sub
        ElseIf UserAuthority = AuthorityAccount Then
            TemResponce = MsgBox("Only Owners are allowed alter these settings", vbInformation, "No Authority")
            Exit Sub
        ElseIf UserAuthority = AuthorityOwnerCOvered Then
            TemResponce = MsgBox("You are not allowed alter these settings", vbInformation, "No Authority")
            Exit Sub
        Else
                'SSTab1.Tab = 2
                Exit Sub
        End If
        
    
    End If
End Sub

Private Sub SaveInstitutionDetails()
Dim TemResponce As Long

If txtInsname.Text = "" Then
    TemResponce = MsgBox("You have not entered an Institution Name", vbCritical, "? Institution Name")
    txtInsname.SetFocus
    Exit Sub
End If

'On Error GoTo ErrorHandler

With DataEnvironment1.rscmmdInstitutionDetails
    
    If .State = 0 Then .Open
    

    SuppliedWord = txtInsname.Text
    !InstitutionName = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtDiscription.Text
    !InstitutionDescription = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtRegistration.Text
    !InstitutionRegistation = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtAddress01.Text
    !InstitutionAddress = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtTelephone01.Text
    !institutiontelephone1 = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtTelephone02.Text
    !InstitutionTelephone2 = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtFax.Text
    !InstitutionFax = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtEmail01.Text
    !InstitutionEmail = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtEmail02.Text
    !InstitutionEmail2 = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtwbsite01.Text
    !InstitutionWebSite1 = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtWebsite02.Text
    !InstitutionWebSite2 = EncreptedWord(SuppliedWord)
    
    
    .Update

    If .State = 1 Then .Close
    
End With

Exit Sub

ErrorHandler:
    MsgBox ("An Error Occured during Updating" & vbNewLine & Me.Caption & vbNewLine & Err.Number & vbNewLine & Err.Description)
    DataEnvironment1.rscmmdInstitutionDetails.CancelUpdate

End Sub

Private Sub GetInstitutionDetails()

With DataEnvironment1.rscmmdInstitutionDetails
    If .State = 0 Then .Open
    
    
    SuppliedWord = !InstitutionName
    txtInsname.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionDescription
    txtDiscription.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionRegistation
    txtRegistration.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionAddress
    txtAddress01.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !institutiontelephone1
    txtTelephone01.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionTelephone2
    txtTelephone02.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionFax
    txtFax.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionEmail
    txtEmail01.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionEmail2
    txtEmail02.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionWebSite1
    txtwbsite01.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionWebSite2
    txtWebsite02.Text = DecreptedWord(SuppliedWord)
    
    If .State = 1 Then .Close
    
End With

End Sub
