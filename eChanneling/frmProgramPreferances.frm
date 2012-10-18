VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProgramPreferances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferances"
   ClientHeight    =   6675
   ClientLeft      =   4440
   ClientTop       =   1680
   ClientWidth     =   11265
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
   ScaleHeight     =   6675
   ScaleWidth      =   11265
   Begin TabDlg.SSTab SSTab2 
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   10398
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Program"
      TabPicture(0)   =   "frmProgramPreferances.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame16"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame18"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FrameColour"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame22"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame17"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame14"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame24"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame10"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame11"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame7"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Practice"
      TabPicture(1)   =   "frmProgramPreferances.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame20"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame23"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame12"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "frmProgramPreferances.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame21"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame19"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame7 
         Caption         =   "Print check box"
         Height          =   975
         Left            =   -67680
         TabIndex        =   71
         Top             =   3600
         Width           =   3495
         Begin VB.OptionButton OptionDisplayPrintChkBox 
            Caption         =   "Display"
            Height          =   240
            Left            =   240
            TabIndex        =   73
            Top             =   360
            Width           =   2775
         End
         Begin VB.OptionButton OptionDoNotDisplaychkPrint 
            Caption         =   "Do not display"
            Height          =   240
            Left            =   240
            TabIndex        =   72
            Top             =   600
            Value           =   -1  'True
            Width           =   2655
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Agent Name For Credit Bookings"
         Height          =   975
         Left            =   3720
         TabIndex        =   65
         Top             =   2520
         Width           =   3615
         Begin VB.OptionButton OptionAgentNameForCreditBookingsYes 
            Caption         =   "Yes"
            Height          =   240
            Left            =   240
            TabIndex        =   67
            Top             =   360
            Width           =   2775
         End
         Begin VB.OptionButton OptionAgentNameForCreditBookings 
            Caption         =   "No"
            Height          =   240
            Left            =   240
            TabIndex        =   66
            Top             =   600
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Patient Names Capitalization"
         Height          =   975
         Left            =   -67680
         TabIndex        =   62
         Top             =   2520
         Width           =   3495
         Begin VB.OptionButton OptionAutomaticCapitalizationNo 
            Caption         =   "No"
            Height          =   240
            Left            =   240
            TabIndex        =   64
            Top             =   600
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton OptionAutomaticCapitalizationYes 
            Caption         =   "Yes"
            Height          =   240
            Left            =   240
            TabIndex        =   63
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Add Foreigner as a suffix"
         Height          =   975
         Left            =   -71280
         TabIndex        =   59
         Top             =   4680
         Width           =   3495
         Begin VB.OptionButton OptionForeignerSuffixYes 
            Caption         =   "Yes"
            Height          =   240
            Left            =   240
            TabIndex        =   61
            Top             =   360
            Width           =   2895
         End
         Begin VB.OptionButton OptionForeignerSuffixNo 
            Caption         =   "No"
            Height          =   240
            Left            =   240
            TabIndex        =   60
            Top             =   600
            Value           =   -1  'True
            Width           =   2415
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Hospital Details in Reports"
         Height          =   975
         Left            =   -67680
         TabIndex        =   56
         Top             =   1440
         Width           =   3495
         Begin VB.OptionButton OptionHospitalDetailsYes 
            Caption         =   "Yes"
            Height          =   240
            Left            =   240
            TabIndex        =   58
            Top             =   360
            Width           =   2775
         End
         Begin VB.OptionButton OptionHospitalDetailsNo 
            Caption         =   "No"
            Height          =   240
            Left            =   240
            TabIndex        =   57
            Top             =   600
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Name Listing"
         Height          =   975
         Left            =   -74880
         TabIndex        =   53
         Top             =   4680
         Width           =   3495
         Begin VB.OptionButton OptionAllNames 
            Caption         =   "List All Names"
            Height          =   240
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton OptionNoAllNames 
            Caption         =   "Don't"
            Height          =   240
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   2175
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Absent / Present"
         Height          =   975
         Left            =   3720
         TabIndex        =   50
         Top             =   1440
         Width           =   3615
         Begin VB.OptionButton OptionDoNotAllowAbsent 
            Caption         =   "Do not allow"
            Height          =   240
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Value           =   -1  'True
            Width           =   3255
         End
         Begin VB.OptionButton OptionAllowAbsent 
            Caption         =   "Allow to mark absent"
            Height          =   240
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "After adding a patient"
         Height          =   975
         Left            =   -67680
         TabIndex        =   47
         Top             =   360
         Width           =   3495
         Begin VB.OptionButton OptionClearAgentDetails 
            Caption         =   "Clear Patient Details"
            Height          =   240
            Left            =   240
            TabIndex        =   49
            Top             =   240
            Width           =   2415
         End
         Begin VB.OptionButton OptionDoNotClearAgentDetails 
            Caption         =   "Do not clear"
            Height          =   240
            Left            =   240
            TabIndex        =   48
            Top             =   600
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Absent Patients' Fee"
         Height          =   975
         Left            =   3720
         TabIndex        =   44
         Top             =   360
         Width           =   3615
         Begin VB.OptionButton OptionDoNotPayDoctor 
            Caption         =   "Do not Pay to doctor"
            Height          =   240
            Left            =   240
            TabIndex        =   46
            Top             =   600
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton OptionPayToDoctor 
            Caption         =   "Pay to doctor"
            Height          =   240
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "After adding a patient"
         Height          =   1455
         Left            =   -71280
         TabIndex        =   39
         Top             =   3000
         Width           =   3495
         Begin VB.OptionButton OptionAfterAddPatient 
            Caption         =   "Focus on Adding Patients"
            Height          =   240
            Left            =   240
            TabIndex        =   43
            Top             =   1095
            Width           =   2895
         End
         Begin VB.OptionButton OptionAfterAddDates 
            Caption         =   "Focus on dates"
            Height          =   240
            Left            =   240
            TabIndex        =   42
            Top             =   840
            Width           =   2775
         End
         Begin VB.OptionButton OptionAfterAddConsultant 
            Caption         =   "Focus on Consultants"
            Height          =   240
            Left            =   240
            TabIndex        =   41
            Top             =   600
            Value           =   -1  'True
            Width           =   2895
         End
         Begin VB.OptionButton OptionAfterAddSpeciality 
            Caption         =   "Focus on Speciality"
            Height          =   240
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame FrameColour 
         Caption         =   "Colour Scheme"
         Height          =   1455
         Left            =   -71280
         TabIndex        =   34
         Top             =   1440
         Width           =   3495
         Begin VB.OptionButton OptionNoColour 
            Caption         =   "No Colour Scheme"
            Height          =   240
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton OptionEnergy 
            Caption         =   "Energy"
            Height          =   240
            Left            =   240
            TabIndex        =   37
            Top             =   600
            Width           =   2535
         End
         Begin VB.OptionButton OptionAqua 
            Caption         =   "Aqua"
            Height          =   240
            Left            =   240
            TabIndex        =   36
            Top             =   840
            Width           =   2535
         End
         Begin VB.OptionButton OptionSunny 
            Caption         =   "Sunny"
            Height          =   240
            Left            =   240
            TabIndex        =   35
            Top             =   1095
            Width           =   1215
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Default Backup Path"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   31
         Top             =   2040
         Width           =   6015
         Begin VB.TextBox txtPath 
            Height          =   360
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   5775
         End
         Begin btButtonEx.ButtonEx bttnSelectBackupPath 
            Height          =   375
            Left            =   1440
            TabIndex        =   33
            Top             =   960
            Width           =   3135
            _ExtentX        =   5530
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
      Begin VB.Frame Frame21 
         Caption         =   "Database"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox txtDatabase 
            Height          =   360
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   5775
         End
         Begin btButtonEx.ButtonEx bttnSelectDatabasePath 
            Height          =   375
            Left            =   1440
            TabIndex        =   30
            Top             =   960
            Width           =   3135
            _ExtentX        =   5530
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
      Begin VB.Frame Frame18 
         Caption         =   "Reprints"
         Height          =   975
         Left            =   -71280
         TabIndex        =   25
         Top             =   360
         Width           =   3495
         Begin VB.OptionButton OptionDoNotAllowReprints 
            Caption         =   "Do not allow"
            Height          =   240
            Left            =   240
            TabIndex        =   27
            Top             =   600
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.OptionButton OptionAllowReprints 
            Caption         =   "Allow"
            Height          =   240
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Selection"
         Height          =   975
         Left            =   -74880
         TabIndex        =   22
         Top             =   3600
         Width           =   3495
         Begin VB.OptionButton OptionCanSelectAgent 
            Caption         =   "Agent"
            Height          =   240
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   3255
         End
         Begin VB.OptionButton OptionCanNotSelectAgent 
            Caption         =   "Agent Code only"
            Height          =   240
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Value           =   -1  'True
            Width           =   3255
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Cash/Credit/Agent Selection"
         Height          =   975
         Left            =   -74880
         TabIndex        =   19
         Top             =   2520
         Width           =   3495
         Begin VB.OptionButton OptionChangeToCash 
            Caption         =   "Change to cash"
            Height          =   240
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   3015
         End
         Begin VB.OptionButton OptionDoNotChangeToCash 
            Caption         =   "Remain in same selection"
            Height          =   240
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Value           =   -1  'True
            Width           =   2775
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Booking Days in Advance"
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   3495
         Begin MSComctlLib.Slider SliderAdvancedBookingDays 
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            SelStart        =   3
            Value           =   3
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            Height          =   255
            Left            =   1440
            TabIndex        =   17
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "10"
            Height          =   255
            Left            =   3120
            TabIndex        =   16
            Top             =   600
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Allow Change of Patient Names"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   3495
         Begin VB.OptionButton OptionDoNotAllowChangeOfNames 
            Caption         =   "No"
            Height          =   240
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton OptionAllowNameChange 
            Caption         =   "Yes"
            Height          =   240
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Name Listing"
         Height          =   975
         Left            =   -74880
         TabIndex        =   8
         Top             =   1440
         Width           =   3495
         Begin VB.OptionButton OptionFirstNameFirst 
            Caption         =   "First Name First"
            Height          =   240
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton OptionSurnameFirst 
            Caption         =   "Surname First"
            Height          =   240
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Need Agent Referance No"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3495
         Begin VB.OptionButton OptionNeedAgentReferanceNo 
            Caption         =   "Yes"
            Height          =   240
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   2535
         End
         Begin VB.OptionButton OptionNoNeedOfAgentReferanceNo 
            Caption         =   "No"
            Height          =   240
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Value           =   -1  'True
            Width           =   2655
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ask before Adding"
         Height          =   975
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   3495
         Begin VB.OptionButton OptionDoNotAskBeforeAdding 
            Caption         =   "No"
            Height          =   240
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton OptionAskBeforeAdding 
            Caption         =   "Yes"
            Height          =   240
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "One Chit for Agent Bookings"
         Height          =   5415
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   10695
         Begin VB.Frame Frame33 
            Caption         =   "Agent Bill Numbers"
            Height          =   855
            Left            =   7440
            TabIndex        =   102
            Top             =   3600
            Width           =   3255
            Begin VB.OptionButton OptionAgentBillNumber 
               Caption         =   "Yes"
               Height          =   240
               Left            =   240
               TabIndex        =   104
               Top             =   240
               Width           =   2775
            End
            Begin VB.OptionButton OptionNoAgentBillNumber 
               Caption         =   "No"
               Height          =   240
               Left            =   240
               TabIndex        =   103
               Top             =   480
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin VB.Frame Frame32 
            Caption         =   "Doctor Payments"
            Height          =   855
            Left            =   7440
            TabIndex        =   99
            Top             =   2760
            Width           =   3255
            Begin VB.OptionButton OptionPaymentSummery 
               Caption         =   "Summery Report"
               Height          =   240
               Left            =   240
               TabIndex        =   101
               Top             =   480
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton OptionDetailDoctorPayment 
               Caption         =   "Detailed Report"
               Height          =   240
               Left            =   240
               TabIndex        =   100
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.Frame Frame31 
            Caption         =   "Requre Agent  Booking Validation"
            Height          =   855
            Left            =   7440
            TabIndex        =   96
            Top             =   1920
            Width           =   3255
            Begin VB.OptionButton OptionAgentBookingValidation 
               Caption         =   "Yes"
               Height          =   240
               Left            =   240
               TabIndex        =   98
               Top             =   240
               Width           =   2775
            End
            Begin VB.OptionButton OptionNoAgentBookingValidation 
               Caption         =   "No"
               Height          =   240
               Left            =   240
               TabIndex        =   97
               Top             =   480
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin VB.Frame Frame30 
            Caption         =   "Check Login"
            Height          =   855
            Left            =   7440
            TabIndex        =   93
            Top             =   1080
            Width           =   3255
            Begin VB.OptionButton OptionNoCheckLogin 
               Caption         =   "No"
               Height          =   240
               Left            =   240
               TabIndex        =   95
               Top             =   480
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton OptionCheckLogin 
               Caption         =   "Yes"
               Height          =   240
               Left            =   240
               TabIndex        =   94
               Top             =   240
               Width           =   2775
            End
         End
         Begin VB.Frame Frame29 
            Caption         =   "One print for agent bookings"
            Height          =   975
            Left            =   3600
            TabIndex        =   90
            Top             =   4320
            Width           =   3615
            Begin VB.OptionButton OptionOnePrintForAgents 
               Caption         =   "Yes"
               Height          =   240
               Left            =   240
               TabIndex        =   92
               Top             =   360
               Width           =   3255
            End
            Begin VB.OptionButton OptionTwoPrintsForAgents 
               Caption         =   "No"
               Height          =   240
               Left            =   240
               TabIndex        =   91
               Top             =   600
               Value           =   -1  'True
               Width           =   3255
            End
         End
         Begin VB.Frame Frame28 
            Caption         =   "Patient Counts in Reports"
            Height          =   975
            Left            =   0
            TabIndex        =   87
            Top             =   4320
            Width           =   3495
            Begin VB.OptionButton OptionDetailCount 
               Caption         =   "Detail Count"
               Height          =   240
               Left            =   240
               TabIndex        =   89
               Top             =   360
               Width           =   2775
            End
            Begin VB.OptionButton OptionOneCount 
               Caption         =   "One Count"
               Height          =   240
               Left            =   240
               TabIndex        =   88
               Top             =   600
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin VB.Frame Frame27 
            Caption         =   "Allow partial refunds / cancellations"
            Height          =   975
            Left            =   3600
            TabIndex        =   84
            Top             =   3240
            Width           =   3615
            Begin VB.OptionButton OptionPartialRepaymentsNo 
               Caption         =   "No"
               Height          =   240
               Left            =   240
               TabIndex        =   86
               Top             =   600
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton OptionPartialRepayments 
               Caption         =   "Yes"
               Height          =   240
               Left            =   240
               TabIndex        =   85
               Top             =   360
               Width           =   2775
            End
         End
         Begin VB.Frame Frame25 
            Caption         =   "Agent Bookings"
            Height          =   975
            Left            =   0
            TabIndex        =   78
            Top             =   3240
            Width           =   3495
            Begin VB.OptionButton OptionAgentCashAndCredit 
               Caption         =   "Cash and Credit"
               Height          =   240
               Left            =   240
               TabIndex        =   80
               Top             =   600
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton OptionAgentCash 
               Caption         =   "Cash only"
               Height          =   240
               Left            =   240
               TabIndex        =   79
               Top             =   360
               Width           =   2775
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Payment Methods Allowed"
            Height          =   1095
            Left            =   7440
            TabIndex        =   74
            Top             =   0
            Width           =   3255
            Begin VB.CheckBox chkAgent 
               Caption         =   "Agent"
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   480
               Width           =   2775
            End
            Begin VB.CheckBox chkCredit 
               Caption         =   "Credit"
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   720
               Width           =   2775
            End
            Begin VB.CheckBox chkCash 
               Caption         =   "Cash"
               Height          =   255
               Left            =   120
               TabIndex        =   75
               Top             =   240
               Width           =   2775
            End
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   5415
         Left            =   -74880
         TabIndex        =   69
         Top             =   360
         Width           =   10695
         Begin VB.Frame Frame26 
            Caption         =   "Date Display"
            Height          =   975
            Left            =   7200
            TabIndex        =   81
            Top             =   4320
            Width           =   3495
            Begin VB.OptionButton OptionInternationalDateFormat 
               Caption         =   "YYYY / MM / DD"
               Height          =   240
               Left            =   240
               TabIndex        =   83
               Top             =   600
               Value           =   -1  'True
               Width           =   2655
            End
            Begin VB.OptionButton OptionEnglishDateFormat 
               Caption         =   "DD / MM / YYYY"
               Height          =   240
               Left            =   240
               TabIndex        =   82
               Top             =   360
               Width           =   2775
            End
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   5415
         Left            =   -74880
         TabIndex        =   70
         Top             =   360
         Width           =   10695
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9120
      TabIndex        =   0
      Top             =   6120
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
Attribute VB_Name = "frmProgramPreferances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
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
    
    bttnSelectDatabasePath.BackColor = BttnBackColour
    bttnSelectDatabasePath.ForeColor = BttnForeColour
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
    bttnSelectBackupPath.BackColor = BttnBackColour
    bttnSelectBackupPath.ForeColor = BttnForeColour
    FrameColour.BackColor = FrameBackColour
    FrameColour.ForeColor = FrameForeColour
    Me.BackColor = FrameBackColour
    Me.ForeColor = FrameForeColour
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
    Frame23.BackColor = FrameBackColour
    Frame23.ForeColor = FrameForeColour
    Frame22.BackColor = FrameBackColour
    Frame22.ForeColor = FrameForeColour
    Frame24.BackColor = FrameBackColour
    Frame24.ForeColor = FrameForeColour
    OptionDisplayPrintChkBox.BackColor = FrameBackColour
    OptionDisplayPrintChkBox.ForeColor = FrameForeColour
    OptionDoNotDisplaychkPrint.BackColor = FrameBackColour
    OptionDoNotDisplaychkPrint.ForeColor = FrameForeColour
    OptionPayToDoctor.BackColor = FrameBackColour
    OptionPayToDoctor.ForeColor = FrameForeColour
    OptionAllowAbsent.BackColor = FrameBackColour
    OptionAllowAbsent.ForeColor = FrameForeColour
    OptionDoNotAllowAbsent.BackColor = FrameBackColour
    OptionDoNotAllowAbsent.ForeColor = FrameForeColour
    OptionAfterAddConsultant.BackColor = FrameBackColour
    OptionAfterAddConsultant.ForeColor = FrameForeColour
    OptionAfterAddDates.BackColor = FrameBackColour
    OptionAfterAddDates.ForeColor = FrameForeColour
    OptionAfterAddPatient.BackColor = FrameBackColour
    OptionAfterAddPatient.ForeColor = FrameForeColour
    OptionAfterAddSpeciality.BackColor = FrameBackColour
    OptionAfterAddSpeciality.ForeColor = FrameForeColour
    OptionHospitalDetailsNo.BackColor = FrameBackColour
    OptionHospitalDetailsNo.ForeColor = FrameForeColour
    OptionHospitalDetailsYes.BackColor = FrameBackColour
    OptionHospitalDetailsYes.ForeColor = FrameForeColour
    OptionDoNotPayDoctor.BackColor = FrameBackColour
    OptionDoNotPayDoctor.ForeColor = FrameForeColour
    OptionNoColour.BackColor = FrameBackColour
    OptionNoColour.ForeColor = FrameForeColour
    OptionAqua.BackColor = FrameBackColour
    OptionAqua.ForeColor = FrameForeColour
    OptionAskBeforeAdding.BackColor = FrameBackColour
    OptionAskBeforeAdding.ForeColor = FrameForeColour
    OptionDoNotAskBeforeAdding.BackColor = FrameBackColour
    OptionDoNotAskBeforeAdding.ForeColor = FrameForeColour
    OptionEnergy.BackColor = FrameBackColour
    OptionEnergy.ForeColor = FrameForeColour
    OptionNeedAgentReferanceNo.BackColor = FrameBackColour
    OptionNeedAgentReferanceNo.ForeColor = FrameForeColour
    OptionNoNeedOfAgentReferanceNo.BackColor = FrameBackColour
    OptionNoNeedOfAgentReferanceNo.ForeColor = FrameForeColour
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
    Call SetPreferances
    Call Setcolours
    SSTab2.Tab = 0
    SSTab2.TabVisible(2) = False
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub SetPreferances()
    Dim TemResponce As Integer
    If FSys.FileExists(DatabasePath) = True Then
        txtDatabase.Text = DatabasePath
    Else
        txtDatabase.Text = "You have not selected a valid database"
        txtDatabase.ForeColor = vbYellow
        txtDatabase.BackColor = vbRed
    End If
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
    OptionDisplayPrintChkBox.Value = DisplayPrintChkBox
    OptionAfterAddConsultant.Value = AfterAddConsultant
    OptionAfterAddSpeciality.Value = AfterAddSpeciality
    OptionAfterAddDates.Value = AfterAddDates
    OptionAfterAddPatient.Value = AfterAddPatient
    chkCredit.Value = PaymentCredit
    chkCash.Value = PaymentCash
    chkAgent.Value = PaymentAgent
    OptionAllowAbsent.Value = AllowAbsent
    OptionAgentCash.Value = AgentCashOnly
    OptionEnglishDateFormat.Value = EnglishDateFormat
    OptionPartialRepayments.Value = PartialRepayments
    OptionDetailCount.Value = DetailedCount
    OptionOnePrintForAgents.Value = OnePrintForAgents
    OptionCheckLogin.Value = CheckLogin
    OptionAgentBookingValidation.Value = AgentBookingValidation
    OptionDetailDoctorPayment.Value = DoctorPaymentDetailedReport
    OptionAgentBillNumber.Value = AgentBillNumber
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
    SaveSetting App.EXEName, "Options", "AdvanceBookingDays", SliderAdvancedBookingDays.Value
    SaveSetting App.EXEName, "Options", "AskBeforeAdding", OptionAskBeforeAdding.Value
    SaveSetting App.EXEName, "Options", "agentessential", OptionNeedAgentReferanceNo.Value
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
    SaveSetting App.EXEName, "Options", "DisplayPrintChkBox", OptionDisplayPrintChkBox.Value
    SaveSetting App.EXEName, "Options", "PaymentCredit", chkCredit.Value
    SaveSetting App.EXEName, "Options", "PaymentCash", chkCash.Value
    SaveSetting App.EXEName, "Options", "PaymentAgent", chkAgent.Value
    SaveSetting App.EXEName, "Options", "AgentCashOnly", OptionAgentCash.Value
    SaveSetting App.EXEName, "Options", "EnglishDateFormat", OptionEnglishDateFormat.Value
    SaveSetting App.EXEName, "Options", "PartialRepayments", OptionPartialRepayments.Value
    SaveSetting App.EXEName, "Options", "DetailedCount", OptionDetailCount.Value
    SaveSetting App.EXEName, "Options", "OnePrintForAgents", OptionOnePrintForAgents.Value
    SaveSetting App.EXEName, "Options", "CheckLogin", OptionCheckLogin.Value
    SaveSetting App.EXEName, "Options", "AgentBookingValidation", OptionAgentBookingValidation.Value
    SaveSetting App.EXEName, "Options", "DoctorPaymentDetailedReport", OptionDetailDoctorPayment.Value
    SaveSetting App.EXEName, "Options", "AgentBillNumber", OptionAgentBillNumber.Value
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
    AdvanceBookingDays = SliderAdvancedBookingDays.Value
    AskBeforeAdding = OptionAskBeforeAdding.Value
    AgentEssential = OptionNeedAgentReferanceNo.Value
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
    DisplayPrintChkBox = OptionDisplayPrintChkBox.Value
    PaymentCash = chkCash.Value
    PaymentCredit = chkCredit.Value
    PaymentAgent = chkAgent.Value
    AgentCashOnly = OptionAgentCash.Value
    PartialRepayments = OptionPartialRepayments.Value
    EnglishDateFormat = OptionEnglishDateFormat.Value
    DetailedCount = OptionDetailCount.Value
    OnePrintForAgents = OptionOnePrintForAgents.Value
    CheckLogin = OptionCheckLogin.Value
    AgentBookingValidation = OptionAgentBookingValidation.Value
    DoctorPaymentDetailedReport = OptionDetailDoctorPayment.Value
    AgentBillNumber = OptionAgentBillNumber.Value
    
    If EnglishDateFormat = True Then
        DefaultLongDate = "DD MMMM YYYY"
        DefaultShortDate = "dd/mm/yy"
    Else
        DefaultLongDate = "YYYY MMMM DD"
        DefaultShortDate = "yy/mm/dd"
    End If
End Sub

Private Sub bttnSelectDatabasePath_Click()
'    CommonDialog1.FileName = GetSetting(App.EXEName, "Options", "DatabaseLocation", App.Path & "\hospital.mdb")
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
'    CommonDialog1.DefaultExt = "mdb"
'    CommonDialog1.Filter = "Lakmedipro Database|hospital.mdb"
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        txtDatabase.Text = CommonDialog1.FileName
        SaveSetting App.EXEName, "Options", "DatabaseLocation", txtDatabase.Text
        DatabasePath = txtDatabase.Text
    Else
        MsgBox "You have not selected valid database. The program may not function", vbCritical, "No database"
        SSTab2.Tab = 2
        bttnSelectDatabasePath.SetFocus
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim TemResponce As Integer
'If FSys.FileExists(txtDatabase.Text) = False Then
'    MsgBox "You have not selected a valid database", vbCritical, "Database?"
'    Cancel = True
'    SSTab2.Tab = 2
'    txtDatabase.SetFocus
'    SendKeys "{home}+{end}"
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SavePreferancesToFile
    Call SavePreferancesToMemory
End Sub
