VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBoooking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Booking"
   ClientHeight    =   8970
   ClientLeft      =   7725
   ClientTop       =   1755
   ClientWidth     =   7575
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBoookingA.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   53
      Top             =   120
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   14631
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Patient"
      TabPicture(0)   =   "frmBoookingA.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePatient"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Booking"
      TabPicture(1)   =   "frmBoookingA.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameBooking"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Payment"
      TabPicture(2)   =   "frmBoookingA.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FramePayment"
      Tab(2).ControlCount=   1
      Begin VB.Frame FrameBooking 
         Caption         =   "Booking"
         Height          =   7815
         Left            =   -74880
         TabIndex        =   54
         Top             =   360
         Width           =   7095
         Begin MSComCtl2.DTPicker DTPickerAppointment 
            Height          =   375
            Left            =   1680
            TabIndex        =   18
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MM yyyy"
            Format          =   55050243
            CurrentDate     =   39422
         End
         Begin VB.Frame FrameLanguage 
            Caption         =   "Print out Language"
            Height          =   1095
            Left            =   4200
            TabIndex        =   131
            Top             =   2880
            Width           =   2775
            Begin VB.OptionButton OptionSinhala 
               Caption         =   "Sinhala"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Value           =   -1  'True
               Width           =   2535
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Tamil"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   480
               Width           =   2535
            End
            Begin VB.OptionButton Option1 
               Caption         =   "English"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   720
               Width           =   2535
            End
         End
         Begin VB.Frame FrameSecession 
            Caption         =   "Secession"
            Height          =   1095
            Left            =   4200
            TabIndex        =   129
            Top             =   1680
            Width           =   2775
            Begin VB.OptionButton OptionNotRelevent 
               Caption         =   "No&t Relevent"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   720
               Width           =   2055
            End
            Begin VB.OptionButton OptionNoPreferance 
               Caption         =   "No Preferance"
               Height          =   255
               Left            =   120
               TabIndex        =   130
               Top             =   720
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.OptionButton OptionEvening 
               Caption         =   "E&vening Secession"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   480
               Value           =   -1  'True
               Width           =   2055
            End
            Begin VB.OptionButton OptionMorning 
               Caption         =   "&Morning Secession"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.TextBox txtTotalFee 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   3000
            Width           =   1815
         End
         Begin VB.TextBox txtPersonalFee 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   1680
            TabIndex        =   21
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox txtInstitutionFee 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   1680
            TabIndex        =   22
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox txtOtherFee 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   1680
            TabIndex        =   23
            Top             =   2520
            Width           =   1815
         End
         Begin VB.TextBox txtNumber 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1680
            TabIndex        =   25
            Top             =   3600
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo DataComboFacility 
            Bindings        =   "frmBoookingA.frx":0496
            Height          =   360
            Left            =   1680
            TabIndex        =   15
            Top             =   360
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   635
            _Version        =   393216
            ListField       =   "HospitalFacility"
            BoundColumn     =   "HospitalFacility_ID"
            Text            =   ""
            Object.DataMember      =   "sqlHospitalFacility"
         End
         Begin MSDataListLib.DataCombo DataComboDoctorStaff 
            Bindings        =   "frmBoookingA.frx":04B5
            Height          =   360
            Left            =   1680
            TabIndex        =   16
            Top             =   840
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   635
            _Version        =   393216
            ListField       =   "FacilityStaff"
            BoundColumn     =   "FacilityStaff_ID"
            Text            =   ""
            Object.DataMember      =   "sqlBookingFacility"
         End
         Begin btButtonEx.ButtonEx BttnAdd 
            Height          =   375
            Left            =   720
            TabIndex        =   32
            Top             =   4200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "A&dd"
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
         Begin btButtonEx.ButtonEx bttnRemove 
            Height          =   375
            Left            =   3960
            TabIndex        =   33
            Top             =   4200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "Re&move"
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
         Begin MSFlexGridLib.MSFlexGrid PatientGrid 
            Height          =   2655
            Left            =   240
            TabIndex        =   34
            Top             =   4680
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   4683
            _Version        =   393216
         End
         Begin btButtonEx.ButtonEx bttnList 
            Height          =   375
            Left            =   6240
            TabIndex        =   17
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "L&ist"
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
         Begin MSComCtl2.DTPicker DTPickerTime 
            Height          =   375
            Left            =   5280
            TabIndex        =   20
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55050242
            CurrentDate     =   39415
         End
         Begin VB.CheckBox chkAppTime 
            Caption         =   "&Approximate time"
            Height          =   255
            Left            =   3480
            TabIndex        =   19
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   "Total &Fee"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   128
            Top             =   3000
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "&Facility"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblDoctorStaff 
            Caption         =   "&Doctor"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label6 
            Caption         =   "D&ate"
            Height          =   255
            Left            =   240
            TabIndex        =   59
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label lblPersonalFee 
            Caption         =   "D&octor Fee"
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   1800
            Width           =   3255
         End
         Begin VB.Label lblInstitutionFee 
            Caption         =   "&Institution Fee"
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label lblOtherFee 
            Caption         =   "&Other Fee"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label lblNumber 
            Caption         =   "&Number"
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   3600
            Width           =   1935
         End
      End
      Begin VB.Frame FramePatient 
         Caption         =   "Patient Details"
         Height          =   7815
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Width           =   7095
         Begin VB.TextBox txtAge 
            Height          =   375
            Left            =   4200
            TabIndex        =   9
            Top             =   4200
            Width           =   2055
         End
         Begin VB.TextBox txtNotes 
            Height          =   705
            Left            =   2040
            TabIndex        =   14
            Top             =   6960
            Width           =   4215
         End
         Begin VB.TextBox txtEmail 
            Height          =   345
            Left            =   2040
            TabIndex        =   13
            Top             =   6480
            Width           =   4215
         End
         Begin VB.TextBox txtFax 
            Height          =   345
            Left            =   2040
            TabIndex        =   12
            Top             =   6000
            Width           =   4215
         End
         Begin VB.TextBox txtTelephone 
            Height          =   345
            Left            =   2040
            TabIndex        =   11
            Top             =   5520
            Width           =   4215
         End
         Begin VB.TextBox txtAddress 
            Height          =   735
            Left            =   2040
            TabIndex        =   10
            Top             =   4680
            Width           =   4215
         End
         Begin VB.TextBox txtNIC 
            Height          =   345
            Left            =   2040
            MaxLength       =   12
            TabIndex        =   7
            Top             =   3720
            Width           =   1695
         End
         Begin VB.TextBox txtSurname 
            Height          =   345
            Left            =   2040
            TabIndex        =   2
            Top             =   1320
            Width           =   4215
         End
         Begin VB.TextBox txtOtherName 
            Height          =   345
            Left            =   2040
            TabIndex        =   1
            Top             =   840
            Width           =   4215
         End
         Begin VB.TextBox txtFirstName 
            Height          =   345
            Left            =   2040
            TabIndex        =   0
            Top             =   360
            Width           =   4215
         End
         Begin MSDataListLib.DataCombo DataComboTitle 
            Bindings        =   "frmBoookingA.frx":04D4
            Height          =   360
            Left            =   2040
            TabIndex        =   3
            Top             =   1800
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   "Title"
            BoundColumn     =   "Title_ID"
            Text            =   ""
            Object.DataMember      =   "sqlTitle"
         End
         Begin MSDataListLib.DataCombo DataComboSex 
            Bindings        =   "frmBoookingA.frx":04F3
            Height          =   360
            Left            =   2040
            TabIndex        =   5
            Top             =   2760
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   "Sex"
            BoundColumn     =   "Sex_ID"
            Text            =   ""
            Object.DataMember      =   "sqlSex"
         End
         Begin MSDataListLib.DataCombo DataComboMarietal 
            Bindings        =   "frmBoookingA.frx":0512
            Height          =   360
            Left            =   2040
            TabIndex        =   4
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   "Marietal"
            BoundColumn     =   "Marietal_ID"
            Text            =   ""
            Object.DataMember      =   "sqlMarietal"
         End
         Begin MSDataListLib.DataCombo DataComboRace 
            Bindings        =   "frmBoookingA.frx":0531
            Height          =   360
            Left            =   2040
            TabIndex        =   6
            Top             =   3240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   "Race"
            BoundColumn     =   "Race_ID"
            Text            =   ""
            Object.DataMember      =   "sqlRace"
         End
         Begin MSComCtl2.DTPicker DTPickerDOB 
            Height          =   375
            Left            =   2040
            TabIndex        =   8
            Top             =   4200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   55050241
            CurrentDate     =   39413
         End
         Begin VB.Label Label18 
            Caption         =   "(Age)"
            Height          =   255
            Left            =   3600
            TabIndex        =   77
            Top             =   4200
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "D&ate of Birth"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   4200
            Width           =   3015
         End
         Begin VB.Image ImagePatient 
            BorderStyle     =   1  'Fixed Single
            Height          =   2295
            Left            =   3840
            Top             =   1800
            Width           =   2415
         End
         Begin VB.Label Label16 
            Caption         =   "&Notes"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   6960
            Width           =   3135
         End
         Begin VB.Label Label15 
            Caption         =   "Ema&il"
            Height          =   255
            Left            =   240
            TabIndex        =   74
            Top             =   6480
            Width           =   2415
         End
         Begin VB.Label Label14 
            Caption         =   "Fa&x"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   6000
            Width           =   3135
         End
         Begin VB.Label Label13 
            Caption         =   "Tele&phone"
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   5520
            Width           =   2415
         End
         Begin VB.Label Label12 
            Caption         =   "&Address"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   4680
            Width           =   3135
         End
         Begin VB.Label Label11 
            Caption         =   "N&IC No."
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   3720
            Width           =   2415
         End
         Begin VB.Label Label19 
            Caption         =   "&Race"
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   3360
            Width           =   2655
         End
         Begin VB.Label Label20 
            Caption         =   "&Marietal"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   2400
            Width           =   2415
         End
         Begin VB.Label Label21 
            Caption         =   "S&ex"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   2880
            Width           =   2655
         End
         Begin VB.Label Label22 
            Caption         =   "&Title"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   1800
            Width           =   2415
         End
         Begin VB.Label Label23 
            Caption         =   "&Surname"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   1320
            Width           =   3135
         End
         Begin VB.Label Label24 
            Caption         =   "&Other Names"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label Label25 
            Caption         =   "&First Name"
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame FramePayment 
         Caption         =   "Payment"
         Height          =   7815
         Left            =   -74880
         TabIndex        =   78
         Top             =   360
         Width           =   7095
         Begin VB.Frame Frame1 
            Caption         =   "Printing"
            Height          =   615
            Left            =   240
            TabIndex        =   124
            Top             =   7080
            Width           =   6735
            Begin VB.OptionButton OptionPrintOne 
               Caption         =   "Print"
               Height          =   255
               Left            =   4080
               TabIndex        =   43
               Top             =   240
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton OptionPrintSeperately 
               Caption         =   "Seperate Prints"
               Height          =   255
               Left            =   2160
               TabIndex        =   125
               Top             =   240
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.OptionButton OptionDoNotPrint 
               Caption         =   "Do not print"
               Height          =   255
               Left            =   720
               TabIndex        =   42
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.TextBox txtNetTotal 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   3960
            Width           =   2055
         End
         Begin VB.TextBox txtDiscount 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1680
            TabIndex        =   38
            Top             =   3480
            Width           =   2055
         End
         Begin VB.TextBox txtGrossTotal 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   3000
            Width           =   2055
         End
         Begin MSFlexGridLib.MSFlexGrid GridPayment 
            Height          =   2415
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   4260
            _Version        =   393216
         End
         Begin btButtonEx.ButtonEx bttnPay 
            Height          =   615
            Left            =   5760
            TabIndex        =   41
            Top             =   5520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
            Appearance      =   3
            Caption         =   "Settle &Payment"
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
         Begin VB.Frame FramePaymentMethod 
            Caption         =   "Payment Method"
            Height          =   1695
            Left            =   4200
            TabIndex        =   82
            Top             =   2880
            Width           =   2655
            Begin VB.OptionButton OptionAgent 
               Caption         =   "A&gent"
               Height          =   255
               Left            =   120
               TabIndex        =   92
               Top             =   1320
               Width           =   1695
            End
            Begin VB.OptionButton OptionCreditCard 
               Caption         =   "Cred&it Card"
               Height          =   255
               Left            =   120
               TabIndex        =   86
               Top             =   1080
               Width           =   1695
            End
            Begin VB.OptionButton OptionCheque 
               Caption         =   "Che&que"
               Height          =   255
               Left            =   120
               TabIndex        =   85
               Top             =   840
               Width           =   1215
            End
            Begin VB.OptionButton OptionCredit 
               Caption         =   "C&redit"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton OptionCash 
               Caption         =   "&Cash"
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.Frame FrameAgent 
            Caption         =   "Agent"
            Height          =   2535
            Left            =   240
            TabIndex        =   91
            Top             =   4560
            Width           =   5415
            Begin MSDataListLib.DataCombo DataComboAgent 
               Bindings        =   "frmBoookingA.frx":0550
               Height          =   360
               Left            =   1920
               TabIndex        =   40
               Top             =   240
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   635
               _Version        =   393216
               Style           =   2
               ListField       =   "InstitutionName"
               BoundColumn     =   "Institution_ID"
               Text            =   ""
               Object.DataMember      =   "sqlInstitutions"
            End
            Begin VB.Label lblAgentAmount 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   2280
               TabIndex        =   138
               Top             =   960
               Width           =   2895
            End
            Begin VB.Label Label31 
               Caption         =   "A&mount           : (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   139
               Top             =   960
               Width           =   2775
            End
            Begin VB.Label txtAgentBalance 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2280
               TabIndex        =   95
               Top             =   1560
               Width           =   2895
            End
            Begin VB.Label Label30 
               Caption         =   "Agent &Balance : (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   94
               Top             =   1560
               Width           =   2655
            End
            Begin VB.Label Label29 
               Caption         =   "&Agent              :"
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   360
               Width           =   2775
            End
         End
         Begin VB.Frame FrameCash 
            Caption         =   "Cash"
            Height          =   2535
            Left            =   240
            TabIndex        =   90
            Top             =   4560
            Width           =   5415
            Begin VB.TextBox txtCashPaid 
               Height          =   375
               Left            =   3600
               TabIndex        =   101
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label lblCashDue 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   3600
               TabIndex        =   100
               Top             =   840
               Width           =   1695
            End
            Begin VB.Label lblCashBalance 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   3600
               TabIndex        =   99
               Top             =   1440
               Width           =   1695
            End
            Begin VB.Label Label35 
               Caption         =   "Amount  :                                    (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   98
               Top             =   840
               Width           =   3615
            End
            Begin VB.Label Label34 
               Caption         =   "Change   :                                   (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   97
               Top             =   1440
               Width           =   3615
            End
            Begin VB.Label Label33 
               Caption         =   "Cash      :                                    (Rs.)"
               Height          =   255
               Left            =   120
               TabIndex        =   96
               Top             =   360
               Width           =   3495
            End
         End
         Begin VB.Frame FrameCredit 
            Caption         =   "Credit"
            Height          =   2535
            Left            =   240
            TabIndex        =   89
            Top             =   4560
            Width           =   5415
            Begin VB.TextBox txtPaidForCredit 
               Height          =   360
               Left            =   3480
               TabIndex        =   107
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label lblTotalCredit 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   3480
               TabIndex        =   105
               Top             =   1200
               Width           =   1815
            End
            Begin VB.Label Label39 
               Caption         =   "Current Credit   :                      (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   103
               Top             =   1200
               Width           =   3495
            End
            Begin VB.Label Label36 
               Caption         =   "Cash Paid          :                     ( Rs.)"
               Height          =   255
               Left            =   120
               TabIndex        =   102
               Top             =   240
               Width           =   3855
            End
            Begin VB.Label lblThisTimeCredit 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   3480
               TabIndex        =   106
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label40 
               Caption         =   "This time Credit :                      (Rs.)"
               Height          =   375
               Left            =   120
               TabIndex        =   104
               Top             =   720
               Width           =   3375
            End
         End
         Begin VB.Frame FrameCheque 
            Caption         =   "Cheque"
            Height          =   2535
            Left            =   240
            TabIndex        =   88
            Top             =   4560
            Width           =   5415
            Begin VB.TextBox txtBranch 
               Height          =   360
               Left            =   2280
               TabIndex        =   134
               Top             =   600
               Width           =   3015
            End
            Begin VB.TextBox txtChequeNo 
               Height          =   360
               Left            =   2280
               TabIndex        =   118
               Top             =   1560
               Width           =   3015
            End
            Begin MSDataListLib.DataCombo DataComboBank 
               Bindings        =   "frmBoookingA.frx":056F
               Height          =   360
               Left            =   2280
               TabIndex        =   114
               Top             =   240
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   635
               _Version        =   393216
               ListField       =   "BankName"
               BoundColumn     =   "Bank_ID"
               Text            =   ""
               Object.DataMember      =   "sqlBank"
            End
            Begin MSComCtl2.DTPicker DTPickerChequeDate 
               Height          =   375
               Left            =   2280
               TabIndex        =   119
               Top             =   2040
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   55050243
               CurrentDate     =   39414
            End
            Begin VB.Label Label8 
               Caption         =   "Branch     :"
               Height          =   375
               Left            =   120
               TabIndex        =   135
               Top             =   600
               Width           =   2295
            End
            Begin VB.Label lblChequeAmount 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Rs. 0.00"
               Height          =   375
               Left            =   2280
               TabIndex        =   121
               Top             =   1080
               Width           =   3015
            End
            Begin VB.Label Label47 
               Caption         =   "Amount    :           (Rs.)"
               Height          =   255
               Left            =   120
               TabIndex        =   120
               Top             =   1080
               Width           =   2415
            End
            Begin VB.Label Label46 
               Caption         =   "Date        :"
               Height          =   255
               Left            =   120
               TabIndex        =   117
               Top             =   2160
               Width           =   3855
            End
            Begin VB.Label Label45 
               Caption         =   "No.          :"
               Height          =   375
               Left            =   120
               TabIndex        =   116
               Top             =   1560
               Width           =   2535
            End
            Begin VB.Label Label43 
               Caption         =   "Bank        :"
               Height          =   255
               Left            =   120
               TabIndex        =   115
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.Frame FrameCreditCard 
            Caption         =   "Credit Card"
            Height          =   2535
            Left            =   240
            TabIndex        =   87
            Top             =   4560
            Width           =   5415
            Begin VB.TextBox txtCardNumber 
               Height          =   360
               Left            =   2400
               TabIndex        =   132
               Top             =   1560
               Width           =   2895
            End
            Begin VB.TextBox txtAuthorizationCode 
               Height          =   360
               Left            =   2400
               TabIndex        =   113
               Top             =   2040
               Width           =   2895
            End
            Begin VB.OptionButton OptionABC 
               Caption         =   "ABC"
               Height          =   255
               Left            =   3720
               TabIndex        =   111
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton OptionAmEx 
               Caption         =   "AmEX"
               Height          =   240
               Left            =   2640
               TabIndex        =   110
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton OptionMaster 
               Caption         =   "MASTER"
               Height          =   255
               Left            =   1320
               TabIndex        =   109
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton OptionVISA 
               Caption         =   "VISA"
               Height          =   255
               Left            =   240
               TabIndex        =   108
               Top             =   240
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker DTPickerCardExpiary 
               Height          =   375
               Left            =   2400
               TabIndex        =   136
               Top             =   1080
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "MMM / yyyy"
               Format          =   55050243
               CurrentDate     =   39423
            End
            Begin VB.Label Label9 
               Caption         =   "Expiary                 :"
               Height          =   255
               Left            =   120
               TabIndex        =   137
               Top             =   1080
               Width           =   2535
            End
            Begin VB.Label Label7 
               Caption         =   "Card Number        :"
               Height          =   255
               Left            =   120
               TabIndex        =   133
               Top             =   1560
               Width           =   2415
            End
            Begin VB.Label lblAmount 
               Alignment       =   1  'Right Justify
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0.00"
               Height          =   375
               Left            =   2400
               TabIndex        =   123
               Top             =   600
               Width           =   2895
            End
            Begin VB.Label Label49 
               Caption         =   "Amount                : (Rs.)"
               Height          =   255
               Left            =   120
               TabIndex        =   122
               Top             =   600
               Width           =   2415
            End
            Begin VB.Label Label44 
               Caption         =   "Authorization Code:"
               Height          =   255
               Left            =   120
               TabIndex        =   112
               Top             =   2040
               Width           =   2415
            End
         End
         Begin VB.Label Label28 
            Caption         =   "&Net Total (Rs.)"
            Height          =   375
            Left            =   120
            TabIndex        =   81
            Top             =   3960
            Width           =   2415
         End
         Begin VB.Label Label27 
            Caption         =   "&Discount (Rs.)"
            Height          =   375
            Left            =   120
            TabIndex        =   80
            Top             =   3480
            Width           =   2415
         End
         Begin VB.Label Label26 
            Caption         =   "&Gross Total (Rs.)"
            Height          =   375
            Left            =   120
            TabIndex        =   79
            Top             =   3000
            Width           =   2415
         End
      End
   End
   Begin VB.TextBox txtPhoto 
      Height          =   360
      Left            =   2040
      TabIndex        =   127
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   6240
      TabIndex        =   35
      Top             =   8520
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx bttnSearchPatient 
      Height          =   375
      Left            =   4920
      TabIndex        =   140
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Search"
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
   Begin VB.Frame frameSearchPatient 
      Caption         =   "Search Patient"
      Height          =   8175
      Left            =   600
      TabIndex        =   44
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txtSearchID 
         Height          =   345
         Left            =   1320
         TabIndex        =   47
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox txtSearchFirstName 
         Height          =   345
         Left            =   1320
         TabIndex        =   46
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtSearchSurname 
         Height          =   345
         Left            =   1320
         TabIndex        =   45
         Top             =   1200
         Width           =   4575
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   5295
         Left            =   120
         TabIndex        =   48
         Top             =   2160
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   9340
         _Version        =   393216
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
      End
      Begin btButtonEx.ButtonEx bttnSearch 
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Search"
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
      Begin btButtonEx.ButtonEx bttnSelect 
         Height          =   375
         Left            =   2160
         TabIndex        =   126
         Top             =   7560
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Select"
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
      Begin VB.Label Label2 
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Surname"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmBoooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemPatientID As Long
    Dim GrossTotal As Double
    Dim Discount As Double
    Dim NetTotal As Double
    Dim TemCatogery  As Integer
    Dim TemDailyMaximum As Integer
    Dim TemStaffFacilityID As Long
    Dim TemHospitalFacilityID As Long
    Dim TemstaffID As Long
    Dim TemPatientFacilityID As Long
    Dim TemBillID As Long
    Dim TemAppointmentTime As Long
    Dim FullPaid As Boolean
    Dim TemPatientCredit As Double
    Dim TemPatientMaxCredit As Double
    Dim BlackListedPatient As Boolean
    Dim TemAgentCredit As Double
    Dim TemAgentMaxCredit As Double
    Dim PatientAdded As Boolean
    Dim TemBillFinished As Boolean

Private Sub bttnAdd_Click()
    Dim TemResponce  As Integer
    Dim TemNum As Long

    frmPatientFacilityList.DTPickerAppointment.Value = DTPickerAppointment.Value

    If Not IsNumeric(DataComboFacility.BoundText) Then
        TemResponce = MsgBox("You have not selected a facility", vbCritical, "No Facility selected")
        DataComboFacility.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then
        Select Case TemCatogery
        Case Doctor:
            TemResponce = MsgBox("You have not selected a doctor", vbCritical, "No Doctor selected")
        Case Staff:
            TemResponce = MsgBox("You have not selected a staff member", vbCritical, "No Staff Member selected")
        Case Investigation:
            TemResponce = MsgBox("You have not selected an investigation", vbCritical, "No Investigation selected")
        End Select
        DataComboDoctorStaff.SetFocus
        Exit Sub
    End If

    If OptionMorning.Value = False And OptionEvening.Value = False And OptionNotRelevent.Value = False Then
        TemResponce = MsgBox("You must select the secession as Morning or Evening", vbCritical, "No secession selected")
        Exit Sub
    End If

    If FacilityAvailable = False Then Exit Sub


    If IsNumeric(txtNumber.Text) And Trim(txtNumber.Text) <> "" Then
            If Val(txtNumber.Text) > 5000 Then
                TemResponce = MsgBox("More than 500 patients can't be seen by a single doctor per a day", vbCritical, "Error")
                txtNumber.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            ElseIf Val(txtNumber.Text) > frmPatientFacilityList.GridList1.Rows Then
                frmPatientFacilityList.GridList1.Rows = Val(txtNumber.Text) + 1
                frmPatientFacilityList.GridList1.Row = Val(txtNumber.Text)
            Else
                If frmPatientFacilityList.GridList1.Rows < Val(txtNumber.Text) + 1 Then frmPatientFacilityList.GridList1.Rows = Val(txtNumber.Text) + 1
                frmPatientFacilityList.GridList1.Row = Val(txtNumber.Text)
            End If
            frmPatientFacilityList.GridList1.Col = 1
            If Trim(frmPatientFacilityList.GridList1.Text) = "" Then
                GoTo AddToDatabase
            Else
                TemResponce = MsgBox("There is already another patient with the same serial number. Please try another number", vbCritical, "Serial Not Available")
                frmPatientFacilityList.GridList1.Col = 1
                TemNum = 1
                Do While TemNum < frmPatientFacilityList.GridList1.Rows
                    frmPatientFacilityList.GridList1.Row = TemNum
                    If frmPatientFacilityList.GridList1.Text = "" Then Exit Do
                TemNum = TemNum + 1
                Loop
                txtNumber.Text = TemNum
                txtNumber.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            End If
    Else
                TemNum = frmPatientFacilityList.GridList1.Rows + 1
                frmPatientFacilityList.GridList1.Rows = TemNum
    
                frmPatientFacilityList.GridList1.Col = 1
    
                Do While TemNum >= 1
                    frmPatientFacilityList.GridList1.Row = TemNum - 1
                    If frmPatientFacilityList.GridList1.Text = "" Then txtNumber.Text = TemNum - 1
                TemNum = TemNum - 1
                Loop
    
                DTPickerTime.Value = AppoitmentTime
                
                GoTo AddToDatabase
    End If

Exit Sub

AddToDatabase:

    With DataEnvironment1.rssqlPatientFacility
        If .State = 1 Then .Close
        .Source = "SELECT tblpatientfacility.* from tblpatientfacility"
        If .State = 0 Then .Open
        .AddNew
        !User_ID = UserID
        !patientid = TemPatientID
        !HospitalFacility_id = TemHospitalFacilityID
        !FacilityStaff_ID = TemStaffFacilityID
        !FacilityCatogery = TemCatogery
        !PatientBill_ID = TemBillID
        !staff_ID = TemstaffID
        !BookingDate = Date
        !AppointmentDate = DTPickerAppointment.Value
        If OptionMorning.Value = True Then !secession = MorningSecession
        If OptionEvening.Value = True Then !secession = EveningSecession
        If OptionNoPreferance.Value = True Then !secession = NoSecessionPreferance
        If OptionNotRelevent.Value = True Then !secession = NoReleventSecession
        !DaySerial = txtNumber.Text
        !Personalfee = Val(txtPersonalFee.Text)
        !institutionfee = Val(txtInstitutionFee.Text)
        !otherfee = Val(txtOtherFee.Text)
        !totalfee = Val(txtPersonalFee.Text) + Val(txtInstitutionFee.Text) + Val(txtOtherFee)
'        !PatientFacility_ID = TemPatientFacilityID
        If chkAppTime.Value = 1 Then !appointmenttime = DTPickerTime.Value
        .Update
        .Close
    End With
    
    Call frmPatientFacilityList.FormatPatientFacilityList
    Call frmPatientFacilityList.FillPatientFacilityList
    Call FormatPatientGrid
    Call FillPatientGrid
    
    txtNumber.Text = Empty
    txtPersonalFee.Text = Empty
    txtInstitutionFee.Text = Empty
    txtOtherFee.Text = Empty
    txtTotalFee.Text = Empty
    DataComboFacility.Text = Empty
    DataComboDoctorStaff.Text = Empty
    DTPickerAppointment = Date
    DTPickerTime.Value = TimeSerial(0, 0, 0)
    
End Sub

Private Function AppoitmentTime() As Date
Dim TemHour As Long
Dim TemMinute As Long
    If TimeSerial(DTPickerTime.Hour, DTPickerTime.Minute, DTPickerTime.Second) <> TimeSerial(0, 0, 0) Then AppoitmentTime = DTPickerTime.Value: Exit Function
    AppoitmentTime = 0
    If chkAppTime.Value = 0 Then AppoitmentTime = 0: Exit Function
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblfacilitystaff.* from tblfacilitystaff where facilitystaff_ID = " & TemStaffFacilityID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Function
        Select Case Weekday(DTPickerAppointment.Value)
            Case vbMonday:
                If OptionMorning.Value = True Then
                    TemHour = Hour(!FacilityMondayMStarting)
                    TemMinute = Minute(!FacilityMondayMStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                ElseIf OptionEvening.Value = True Then
                    TemHour = Hour(!FacilityMondayEStarting)
                    TemMinute = Minute(!FacilityMondayEStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                End If
            Case vbTuesday:
                If OptionMorning.Value = True Then
                    TemHour = Hour(!FacilitytuesdayMStarting)
                    TemMinute = Minute(!FacilitytuesdayMStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                ElseIf OptionEvening.Value = True Then
                    TemHour = Hour(!FacilitytuesdayEStarting)
                    TemMinute = Minute(!FacilitytuesdayEStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                End If
            Case vbWednesday:
                If OptionMorning.Value = True Then
                    TemHour = Hour(!FacilitywednesdayMStarting)
                    TemMinute = Minute(!FacilitywednesdayMStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                ElseIf OptionEvening.Value = True Then
                    TemHour = Hour(!FacilitywednesdayEStarting)
                    TemMinute = Minute(!FacilitywednesdayEStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                End If
            Case vbThursday:
                If OptionMorning.Value = True Then
                    TemHour = Hour(!FacilitythursdayMStarting)
                    TemMinute = Minute(!FacilitythursdayMStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                ElseIf OptionEvening.Value = True Then
                    TemHour = Hour(!FacilitythursdayEStarting)
                    TemMinute = Minute(!FacilitythursdayEStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                End If
            Case vbFriday:
                If OptionMorning.Value = True Then
                    TemHour = Hour(!FacilityfridayMStarting)
                    TemMinute = Minute(!FacilityfridayMStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                ElseIf OptionEvening.Value = True Then
                    TemHour = Hour(!FacilityfridayEStarting)
                    TemMinute = Minute(!FacilityfridayEStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                End If
            Case vbSaturday:
                If OptionMorning.Value = True Then
                    TemHour = Hour(!FacilitysaturdayMStarting)
                    TemMinute = Minute(!FacilitysaturdayMStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                ElseIf OptionEvening.Value = True Then
                    TemHour = Hour(!FacilitysaturdayEStarting)
                    TemMinute = Minute(!FacilitysaturdayEStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                End If
            Case vbSunday:
                If OptionMorning.Value = True Then
                    TemHour = Hour(!FacilitysundayMStarting)
                    TemMinute = Minute(!FacilitysundayMStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                ElseIf OptionEvening.Value = True Then
                    TemHour = Hour(!FacilitysundayEStarting)
                    TemMinute = Minute(!FacilitysundayEStarting)
                    AppoitmentTime = TimeSerial(TemHour, (TemMinute + (!usualduration) * Val(txtNumber.Text)), 0)
                End If
        End Select
        .Close
    End With
End Function

Private Function FacilityAvailable() As Boolean
    Dim TemResponce  As Integer
    FacilityAvailable = False
    With DataEnvironment1.rssqlTem6
        If .State = 1 Then .Close
        .Source = "SELECT * from tblfacilitystaff where FacilityStaff_ID = " & TemStaffFacilityID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Function
        Select Case Weekday(DTPickerAppointment.Value)
        Case vbMonday:
            If !FullDayLeaveMonday = True Then
                TemResponce = MsgBox("This facility is not available on Mondays", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilityMondayM = False And OptionMorning.Value = True Then
                TemResponce = MsgBox("This facility is not available on Monday Morning", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilityMondayE = False And OptionEvening.Value = True Then
                TemResponce = MsgBox("This facility is not available on Monday Evening", vbInformation, "Not Available")
                Exit Function
            End If
        Case vbTuesday:
            If !FullDayLeaveTuesday = True Then
                TemResponce = MsgBox("This facility is not available on Tuesdays", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilityTuesdayM = False And OptionMorning.Value = True Then
                TemResponce = MsgBox("This facility is not available on Tuesday Morning", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilityTuesdayE = False And OptionEvening.Value = True Then
                TemResponce = MsgBox("This facility is not available on Tuesday Evening", vbInformation, "Not Available")
                Exit Function
            End If
        Case vbWednesday:
            If !FullDayLeaveWednesday = True Then
                TemResponce = MsgBox("This facility is not available on Wednesdays", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilityWednesdayM = False And OptionMorning.Value = True Then
                TemResponce = MsgBox("This facility is not available on Wednesday Morning", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilityWednesdayE = False And OptionEvening.Value = True Then
                TemResponce = MsgBox("This facility is not available on Wednesday Evening", vbInformation, "Not Available")
                Exit Function
            End If
        Case vbThursday:
            If !FullDayLeaveThursday = True Then
                TemResponce = MsgBox("This facility is not available on Thursdays", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilityThursdayM = False And OptionMorning.Value = True Then
                TemResponce = MsgBox("This facility is not available on Thursday Morning", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilityThursdayE = False And OptionEvening.Value = True Then
                TemResponce = MsgBox("This facility is not available on Thursday Evening", vbInformation, "Not Available")
                Exit Function
            End If
        Case vbFriday:
            If !FullDayLeaveFriday = True Then
                TemResponce = MsgBox("This facility is not available on Fridays", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilityFridayM = False And OptionMorning.Value = True Then
                TemResponce = MsgBox("This facility is not available on Friday Morning", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilityFridayE = False And OptionEvening.Value = True Then
                TemResponce = MsgBox("This facility is not available on Friday Evening", vbInformation, "Not Available")
                Exit Function
            End If
        Case vbSaturday:
            If !FullDayLeaveSaturday = True Then
                TemResponce = MsgBox("This facility is not available on Saturdays", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilitySaturdayM = False And OptionMorning.Value = True Then
                TemResponce = MsgBox("This facility is not available on Saturday Morning", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilitySaturdayE = False And OptionEvening.Value = True Then
                TemResponce = MsgBox("This facility is not available on Saturday Evening", vbInformation, "Not Available")
                Exit Function
            End If
        Case vbSunday:
            If !FullDayLeaveSunday = True Then
                TemResponce = MsgBox("This facility is not available on Sundays", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilitySundayM = False And OptionMorning.Value = True Then
                TemResponce = MsgBox("This facility is not available on Sunday Morning", vbInformation, "Not Available")
                Exit Function
            End If
            If !FacilitySundayE = False And OptionEvening.Value = True Then
                TemResponce = MsgBox("This facility is not available on Sunday Evening", vbInformation, "Not Available")
                Exit Function
            End If
        End Select
    .Close
    .Source = "SELECT * from tblfacilitystaffleave where (FacilityStaff_ID = " & TemStaffFacilityID & ") and (FacilityStaffLeaveDate = #" & DTPickerAppointment.Value & "#)"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FacilityAvailable = True: Exit Function
    If !FullDayLeave = True Then
        TemResponce = MsgBox("This facility is not available on " & DTPickerAppointment.Value, vbInformation, "Not Available")
        Exit Function
    End If
    If !Morning = False And OptionMorning.Value = True Then
        TemResponce = MsgBox("This facility is not available on " & DTPickerAppointment.Value & " morning", vbInformation, "Not Available")
        Exit Function
    End If
    If !Evening = False And OptionEvening.Value = True Then
        TemResponce = MsgBox("This facility is not available on " & DTPickerAppointment.Value & " evening", vbInformation, "Not Available")
        Exit Function
    End If
    .Close
    End With
    FacilityAvailable = True
End Function


Private Sub FormatPaymentGrid()
    With GridPayment
        .Clear
        .Rows = 1
        .Cols = 7
        
        .ColWidth(0) = 600
        .Col = 0
        .CellAlignment = 4
        .Text = "No."
        
        .ColWidth(1) = 2000
        .Col = 1
        .CellAlignment = 4
        .Text = "Facility"
        
        .ColWidth(2) = 3500
        .Col = 2
        .CellAlignment = 4
        .Text = "Description"
        
        .ColWidth(3) = 1200
        .Col = 3
        .CellAlignment = 4
        .Text = "Personal Fee"
        
        .ColWidth(4) = 1200
        .Col = 4
        .CellAlignment = 4
        .Text = "Institution Fee"
        
        .ColWidth(5) = 1200
        .Col = 5
        .CellAlignment = 4
        .Text = "Other Fee"
        
        .ColWidth(6) = 1600
        .Col = 6
        .CellAlignment = 1
        .Text = "Total Fee"
    End With

End Sub



Private Sub FormatPatientGrid()
    Call FormatPaymentGrid
    With PatientGrid
        .Clear
        .Rows = 1
        .Cols = 14
        
        .ColWidth(0) = 600
        .Col = 0
        .CellAlignment = 4
        .Text = "No."
        
        .ColWidth(1) = 2000
        .Col = 1
        .CellAlignment = 4
        .Text = "Facility"
        
        .ColWidth(2) = 1
        .Col = 2
        .CellAlignment = 4
        .Text = "" ' "FacilityID"
       
        .ColWidth(3) = 3500
        .Col = 3
        .CellAlignment = 4
        .Text = "Description"
        
        .ColWidth(4) = 1
        .Col = 4
        .CellAlignment = 4
        .Text = "" '"Personal ID"
        
        .ColWidth(5) = 1200
        .Col = 5
        .CellAlignment = 4
        .Text = "Date"
        
        .ColWidth(6) = 1
        .Col = 6
        .CellAlignment = 4
        .Text = "" '"Personal Fee"
        
        .ColWidth(7) = 1
        .Col = 7
        .CellAlignment = 4
        .Text = "" ' "Institution Fee"
        
        .ColWidth(8) = 1
        .Col = 8
        .CellAlignment = 4
        .Text = "" '"Other Fee"
        
        .ColWidth(9) = 1
        .Col = 9
        .CellAlignment = 4
        .Text = "" ' "Total Fee"
        
        .ColWidth(10) = 1000
        .Col = 10
        .CellAlignment = 4
        .Text = "Day No."
        
        .ColWidth(11) = 1
        .Col = 11
        .CellAlignment = 4
        .Text = "" ' "PatientFacilityID"
        
        .ColWidth(12) = 2500
        .Col = 12
        .CellAlignment = 4
        .Text = "Approximate time"
        
        .ColWidth(13) = 1
        .Col = 13
        .Text = "" 'catogery
        
    End With
    
    bttnRemove.Enabled = False
    
End Sub

Private Sub FillPatientGrid()


    Dim NowRow As Long
    Dim TemNum As Long
    
    GrossTotal = 0
    
    With DataEnvironment1.rssqlTem5
        If .State = 1 Then .Close
        .Source = "select tblpatientfacility.* from tblpatientfacility where (patientbill_ID = " & TemBillID & ") order by PatientFacility_ID"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        NowRow = 0
        While Not .EOF
        NowRow = NowRow + 1
        
        PatientGrid.Rows = NowRow + 1
        PatientGrid.Row = NowRow
        GridPayment.Rows = NowRow + 1
        GridPayment.Row = NowRow
        
        PatientGrid.Col = 0
        PatientGrid.CellAlignment = 7
        PatientGrid.Text = NowRow
        GridPayment.Col = 0
        GridPayment.CellAlignment = 7
        GridPayment.Text = NowRow

        PatientGrid.Col = 1
        PatientGrid.CellAlignment = 1
        PatientGrid.Text = FindHospitalFacilityFromID(!HospitalFacility_id)
        GridPayment.Col = 1
        GridPayment.CellAlignment = 1
        GridPayment.Text = PatientGrid.Text

        PatientGrid.Col = 2
        PatientGrid.CellAlignment = 0
        PatientGrid.Text = !HospitalFacility_id

        PatientGrid.Col = 3
        PatientGrid.CellAlignment = 1

        Select Case !FacilityCatogery
            Case Doctor:
                PatientGrid.Text = FindDoctorFromID(!staff_ID)
                PatientGrid.Col = 13
                PatientGrid.Text = Doctor
            Case Staff:
                PatientGrid.Text = FindStaffFromID(!staff_ID)
                PatientGrid.Col = 13
                PatientGrid.Text = Staff
            Case Investigation:
                PatientGrid.Text = FindInvestigationFromID(!staff_ID)
                PatientGrid.Col = 13
                PatientGrid.Text = Investigation
            Case Other:
        End Select

        Select Case !secession
            Case MorningSecession:
                PatientGrid.Text = PatientGrid.Text & "(Morning)"
            Case EveningSecession:
                PatientGrid.Text = PatientGrid.Text & "(Evening)"
            Case NoReleventSecession:

            Case NoSecessionPreferance:

        End Select

        GridPayment.Col = 2
        GridPayment.CellAlignment = 1
        GridPayment.Text = PatientGrid.Text

        PatientGrid.Col = 4
        PatientGrid.CellAlignment = 4
        Select Case TemCatogery
            Case Doctor:
                PatientGrid.Text = !staff_ID
            Case Staff:
                PatientGrid.Text = !staff_ID
            Case Investigation:
                PatientGrid.Text = !staff_ID
            Case Other:
        End Select

        PatientGrid.Col = 5
        PatientGrid.CellAlignment = 1
        PatientGrid.Text = Format(!BookingDate, "dd/MMM/yyyy")

        GridPayment.Col = 3
        GridPayment.CellAlignment = 7
        GridPayment.Text = Format(!Personalfee, "#0.00")

        GridPayment.Col = 4
        GridPayment.CellAlignment = 7
        GridPayment.Text = Format(!institutionfee, "#0.00")

        GridPayment.Col = 5
        GridPayment.CellAlignment = 7
        GridPayment.Text = Format(!otherfee, "#0.00")

        GridPayment.Col = 6
        GridPayment.CellAlignment = 7
        GrossTotal = GrossTotal + (!Personalfee + !institutionfee + !otherfee)
        GridPayment.Text = Format((!Personalfee + !institutionfee + !otherfee), "#0.00")

        PatientGrid.Col = 10
        PatientGrid.CellAlignment = 4
        PatientGrid.Text = !DaySerial

        PatientGrid.Col = 11
        PatientGrid.CellAlignment = 4
        PatientGrid.Text = !patientfacility_ID
        
        PatientGrid.Col = 12
        PatientGrid.CellAlignment = 4
        If IsNull(!appointmenttime) Or !appointmenttime = 0 Then
            PatientGrid.Text = "Not relevent"
        Else
            PatientGrid.Text = !appointmenttime
        End If

            .MoveNext
        Wend
        .Close
    PatientGrid.Col = 0
    End With
    
    txtGrossTotal.Text = Format(GrossTotal, "#0.00")
'    PatientGrid.Row = 1
'    PatientGrid.Col = 1



End Sub

Private Sub bttnList_Click()
    frmPatientFacilityList.Show
    frmPatientFacilityList.ZOrder 0
    frmPatientFacilityList.Left = MDIFrmReception.Width / 2
    frmPatientFacilityList.Top = 0
End Sub

Private Sub bttnPay_Click()
    
    If CanSettlePayment = False Then Exit Sub
    Call PrintBill
    Call UpdateBill
    Call UpdatePatientFacilities
    Call PrepareToBook
    
End Sub

Private Sub UpdatePatientFacilities()
    With DataEnvironment1.rssqlPatientFacility
        If .State = 1 Then .Close
        .Source = "SELECT tblpatientfacility.* from tblpatientfacility where Patientbill_ID = " & TemBillID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        While Not .EOF
            If OptionCash.Value = True Or OptionCreditCard.Value = True Or OptionAgent.Value = True Or OptionCheque.Value = True Then
                !fullypaid = True
            Else
                !fullypaid = False
            End If
            If OptionCash.Value = True Then
                !paymentmode = "Cash"
                !paymentmethod_ID = 1
                !resultsuccess = True
            ElseIf OptionAgent.Value = True Then
                !paymentmode = "Agent"
                !paymentmethod_ID = 2
                !resultsuccess = True
                !agent_ID = DataComboAgent.BoundText
            ElseIf OptionCreditCard.Value = True Then
                !paymentmode = "Credit Card"
                !paymentmethod_ID = 3
                !resultsuccess = True
            ElseIf OptionCredit.Value = True Then
                !paymentmode = "Credit"
                !paymentmethod_ID = 4
            ElseIf OptionCheque.Value = True Then
                !paymentmode = "Cheque"
                !paymentmethod_ID = 5
                !resultsuccess = True
            Else
                !paymentmode = "Other"
                !paymentmethod_ID = 6
                !resultauccess = False
            End If
            .Update
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub PrintBill()
    If OptionPrintOne.Value = True Then
        PrintTabBill
    ElseIf OptionPrintSeperately.Value = True Then
        PrintSeperateBills
    Else
        Exit Sub
    End If
End Sub

Private Sub PrintTabBill()
    Dim TemRows As Long

With Printer
    
    .Font = "Bernard MT Condensed"
    Printer.Print
    .FontSize = 14
    Printer.Print Tab(2); InstitutionName
    .FontSize = 9
    Printer.Print Tab(3); InstitutionAddress
    Printer.Print Tab(3); InstitutionTelephone
    
    .FontName = "Arial"
    .FontSize = 10
    Printer.Print
    
    Dim TemTab1 As Long
    Dim TemTab2 As Long
    Dim TemTab3 As Long
    Dim TemTab4 As Long
    Dim TemTab5 As Long
    Dim TemTab6 As Long
    
    TemTab1 = 2
    TemTab2 = 6
    TemTab3 = 16
    TemTab4 = 25
    TemTab5 = 36
    
    Printer.Print Tab(TemTab1); "Name";
    Printer.Print Tab(TemTab3); (DataComboTitle.Text & " " & txtFirstName.Text & " " & txtSurname.Text)
    
    If Val(txtAge.Text) <> 0 Then
        Printer.Print Tab(TemTab1); "Age";
        Printer.Print Tab(TemTab3); txtAge.Text
    End If
    
    If DataComboSex.Text <> "" Then
        Printer.Print Tab(TemTab1); "Sex";
        Printer.Print Tab(TemTab3); DataComboSex.Text
    End If
    
        Printer.Print Tab(TemTab1); "ID";
        Printer.Print Tab(TemTab3); TemPatientID
    
    Printer.Print
    
            For TemRows = 1 To (GridPayment.Rows - 1)
                PatientGrid.Row = TemRows
                GridPayment.Row = TemRows
                
                Printer.Print Tab(TemTab1); Format(TemRows, "00");
                
                PatientGrid.Col = 1
                Printer.Print Tab(TemTab2); PatientGrid.Text;
                
                PatientGrid.Col = 3
                Printer.Print Tab(TemTab4); PatientGrid.Text;
                
                
                Printer.Print Tab(TemTab2); "Fee";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Rs. " & Format(GridPayment.Text, "000.00")
                
                Printer.Print Tab(TemTab2); "Date";
                
                PatientGrid.Col = 5
                Printer.Print Tab(TemTab4); Format(PatientGrid.Text, "dd mmmm yyyy")
                
                Printer.Print Tab(TemTab2); "Day Serial";
                
                PatientGrid.Col = 10
                Printer.Print Tab(TemTab4); Format(PatientGrid.Text, "00")
                
                
                If IsDate(PatientGrid.Text) Then
                    If PatientGrid.Text <> TimeSerial(0, 0, 0) Then
                    Printer.Print Tab(TemTab2); "App. Time";
                    Printer.Print Tab(TemTab4); PatientGrid.Text
                    End If
                End If
                Printer.Print
            Next
            
                Printer.Print Tab(TemTab2); "Gross Total";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Rs. " & Format(Val(txtGrossTotal.Text), "00.00")
            
            
                Printer.Print Tab(TemTab2); "Discount";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Rs. " & Format(Val(txtDiscount.Text), "00.00")
            
                Printer.Print Tab(TemTab2); "Net Total";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Rs. " & Format(Val(txtNetTotal.Text), "00.00")
            
            Printer.Print
            
            If OptionCash.Value = True Then
                Printer.Print Tab(TemTab2); "Paid as";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Cash"
                Printer.Print Tab(TemTab2); "Payment";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Rs. " & Format(Val(lblCashDue.Caption), "00.00")
            ElseIf OptionCredit.Value = True Then
                Printer.Print Tab(TemTab2); "Paid as";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Credit"
                Printer.Print Tab(TemTab2); "Payment";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Rs. " & Format(txtPaidForCredit.Text, "00.00")
                
            ElseIf OptionCreditCard.Value = True Then
                Printer.Print Tab(TemTab2); "Paid as";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Credit Card"
                Printer.Print Tab(TemTab2); "Payment";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Rs. " & Format(lblAmount.Caption, "00.00")
                
            ElseIf OptionCheque.Value = True Then
                Printer.Print Tab(TemTab2); "Paid as";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Cheque"
                Printer.Print Tab(TemTab2); "Payment";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Rs. " & Format(lblChequeAmount.Caption, "00.00")
                
            ElseIf OptionAgent.Value = True Then
                Printer.Print Tab(TemTab2); "Paid as";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Agent"
                Printer.Print Tab(TemTab2); "Payment";
                GridPayment.Col = 6
                Printer.Print Tab(TemTab4); "Rs. " & Format(lblAmount.Caption, "00.00")
                
            End If
            
            Printer.Print
            Printer.Print
            Printer.Print Tab(TemTab2); "--------------------"
            Printer.Print Tab(TemTab2); UserName
            Printer.Print Tab(TemTab2); Format(Date, "dd mmmm yyyy")
            
    .EndDoc
End With
End Sub

Private Sub PrintOneBill()
    Dim TemFullText As String
    Dim TemText As String
    Dim TemRows As Long
    
    If GridPayment.Rows < 1 Or PatientGrid.Rows < 1 Then

    Else
        
        For TemRows = 1 To (GridPayment.Rows - 1)
            PatientGrid.Row = TemRows
            Call PrintingResults
            If PreferancechkPatientName = True Then Call PrintPatientName(DataComboTitle.Text & " " & txtFirstName.Text & " " & txtSurname.Text)
            If PreferancechkPatientAge = True Then Call PrintPatientAge(txtAge.Text)
            If PreferancechkPatientSex = True Then Call PrintPatientSex(DataComboSex.Text)
            If PreferancechkPatientID = True Then Call PrintPatientID(TemPatientID)
            If PreferanceChkLblComments = True Then Call PrintLblComments(PreferanceTxtLblComments)
            Call PrintLines
            TemFullText = ""
            PatientGrid.Col = 1
            TemText = Format(TemRows, "00") & "  Facility            :   " & PatientGrid.Text
            TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
            PatientGrid.Col = 13
            Select Case Val(PatientGrid.Text)
                Case Doctor:
                    PatientGrid.Col = 3
                    TemText = "      Doctor             :   " & PatientGrid.Text
                    
                Case Staff:
                    PatientGrid.Col = 3
                    TemText = "      Staff Member  :   " & PatientGrid.Text
                Case Investigation:
                    PatientGrid.Col = 3
                    TemText = "      Investigation   :   " & PatientGrid.Text
            End Select
            TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
            
            PatientGrid.Col = 5
            TemText = "      Date                :   " & Format(PatientGrid.Text, "dd/mmmm/yyyy")
            TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
            
            PatientGrid.Col = 10
            If Val(PatientGrid.Text) <> Investigation Then
                    TemText = "      Day No.          :   " & PatientGrid.Text
            TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
            End If
            PatientGrid.Col = 12
            If IsDate(PatientGrid.Text) Then
                If PatientGrid.Text <> TimeSerial(0, 0, 0) Then
                    TemText = "      App. Time       :   " & PatientGrid.Text
                End If
            TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
            End If
            PatientGrid.Col = 13
            TemFullText = TemFullText & Chr(13) & Chr(10)
            Select Case Val(PatientGrid.Text)
            Case Doctor:
                GridPayment.Col = 3
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Doctor Fee        : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
                GridPayment.Col = 4
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Institution Fee   : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
                GridPayment.Col = 5
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Other Fee          : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
                GridPayment.Col = 6
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Total Fee           : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
            Case Staff:
                GridPayment.Col = 3
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Doctor Fee       : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
                GridPayment.Col = 4
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Institution Fee  : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
                GridPayment.Col = 5
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Other Fee        : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
                GridPayment.Col = 6
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Total Fee        : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
            
            Case Investigation:
                GridPayment.Col = 3
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Investigation Fee  : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
                GridPayment.Col = 4
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Institution Fee  : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
                GridPayment.Col = 5
                If Val(GridPayment.Text) <> 0 Then
                    TemText = "      Other Fee        : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
                GridPayment.Col = 6
                If Val(GridPayment.Text) <> Val(GridPayment.TextMatrix(TemRows, 3)) Then
                    TemText = "      Total Fee        : Rs. " & Format(Val(GridPayment.Text), "##000.00")
                    TemFullText = TemFullText & TemText & Chr(13) & Chr(10)
                End If
            End Select
            Call PrintResultstList(TemFullText)
            Printer.EndDoc
        Next
    End If
    
End Sub


Private Sub PrintSeperateBills()


End Sub

Private Sub UpdateBill()
    With DataEnvironment1.rssqlPatientBill
    If .State = 1 Then .Close
    .Source = "SELECT tblpatientbill.* from tblpatientbill where patientbill_ID = " & TemBillID
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    !BillSuccess = True
    !Date = Date
    !GrossTotal = Val(txtGrossTotal.Text)
    !Discount = Val(txtDiscount.Text)
    !NetTotal = Val(txtNetTotal.Text)
    !patient_ID = TemPatientID
    If OptionCash.Value = True Then
        !paymentmethod = "Cash"
        !Cash = Val(txtNetTotal.Text)
    ElseIf OptionCredit.Value = True Then
        !paymentmethod = "Credit"
        !CreditCash = Val(txtPaidForCredit.Text)
        !Credit = Val(lblThisTimeCredit.Caption)
        UpdatePatientCredit
    ElseIf OptionCheque.Value = True Then
        !paymentmethod = "Cheque"
        !ChequeAmount = Val(lblChequeAmount.Caption)
        !ChequeDate = DTPickerChequeDate.Value
        !Bank_Id = DataComboBank.BoundText
        !Branch = txtBranch.Text
        !ChequeNo = txtChequeNo.Text
    ElseIf OptionCreditCard.Value = True Then
        !paymentmethod = "CreditCard"
        !CreditCardAmount = Val(lblAmount.Caption)
        If OptionVISA.Value = True Then
            !CreditCard = "VISA"
        ElseIf OptionMaster.Value = True Then
            !CreditCard = "MASTER"
        ElseIf OptionAmEx.Value = True Then
            !CreditCard = "AMEX"
        ElseIf OptionABC.Value = True Then
            !CreditCard = "ABC"
        Else
            !CreditCard = "OTHER"
        End If
        !CreditCardNo = txtCardNumber.Text
        !ExpiaryDate = DTPickerCardExpiary.Value
    ElseIf OptionAgent.Value = True Then
        !paymentmethod = "Agent"
        !AgentAmount = lblAgentAmount.Caption
        !agent_ID = DataComboAgent.BoundText
        UpdateAgentCredit
    End If
    !BillSuccess = True
    .Update
    .Close
    End With
End Sub

Private Sub UpdatePatientCredit()
    With DataEnvironment1.rssqlTem7
        If .State = 1 Then .Close
        .Source = "SELECT tblpatientmaindetails.* from tblpatientmaindetails where patient_ID =" & TemPatientID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        !Credit = !Credit - Val(txtNetTotal.Text)
        .Update
        .Close
    End With
End Sub

Private Sub UpdateAgentCredit()
    With DataEnvironment1.rssqlTem7
        If .State = 1 Then .Close
        .Source = "SELECT tblinstitutions.* from tblinstitutions where institution_ID =" & DataComboAgent.BoundText
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        !InstitutionCredit = !InstitutionCredit - Val(txtNetTotal.Text)
        .Update
        .Close
    End With
End Sub

Private Function CanSettlePayment() As Boolean
    Dim TemResponce  As Integer
    CanSettlePayment = False
    If OptionCash.Value = False And OptionCredit.Value = False And OptionCheque.Value = False And OptionCreditCard.Value = False And OptionAgent.Value = False Then
        TemResponce = MsgBox("You have not selected a payment method", vbInformation, "Payment Method")
        OptionCash.SetFocus
        Exit Function
    End If
    If OptionCash.Value = True Then
        If Val(txtCashPaid.Text) < Val(txtNetTotal.Text) Then
            TemResponce = MsgBox("When you pay cash, you have to pay the full amount", vbInformation, "Payment not enough")
            txtCashPaid.SetFocus
            Exit Function
        End If
    ElseIf OptionCredit.Value = True Then
         If BlackListedPatient = True Then
            TemResponce = MsgBox("This patient is a black listed patient. Please try another payment method or discuss with the management to remove from the black list", vbInformation, "Black Listed")
            OptionCash.SetFocus
            Exit Function
         End If
         If (0 - TemPatientMaxCredit) < (TemPatientCredit - Val(txtNetTotal.Text)) Then
            TemResponce = MsgBox("This trasaction will exceed the credit limit of the patient. If you want to proceed,", vbInformation, "Credit Limit")
            txtPaidForCredit.SetFocus
            Exit Function
         End If
    ElseIf OptionCheque.Value = True Then
        If Not IsNumeric(DataComboBank.BoundText) Then
            TemResponce = MsgBox("You have not selected the Bank of the check you accept", vbInformation, "Bank?")
            DataComboBank.SetFocus
            Exit Function
        End If
        If Trim(txtChequeNo.Text) = "" Then
            TemResponce = MsgBox("You have not entered the cheque number that you are going to accept", vbInformation, "Cheque No. ?")
            txtChequeNo.SetFocus
            Exit Function
        End If
    ElseIf OptionCreditCard.Value = True Then
        If OptionVISA.Value = False And OptionMaster.Value = False And OptionAmEx.Value = False And OptionABC.Value = False Then
            TemResponce = MsgBox("You have not selected the credit card you are accepting", vbInformation, "Credit card")
            OptionVISA.SetFocus
            Exit Function
        End If
        If Trim(txtCardNumber.Text) = "" Then
            TemResponce = MsgBox("You have not entered the credit card number", vbInformation, "Credit Card Number")
            txtCardNumber.SetFocus
            Exit Function
        End If
        If Trim(txtAuthorizationCode.Text) = "" Then
            TemResponce = MsgBox("You have not entered the authorization code", vbInformation, "Authorization Code")
            txtAuthorizationCode.SetFocus
            Exit Function
        End If
        If DTPickerCardExpiary.Value < Date Then
            TemResponce = MsgBox("You have going to accept an expired Card, check on that", vbInformation, "Expiary")
            DTPickerCardExpiary.SetFocus
            Exit Function
        End If
    ElseIf OptionAgent.Value = True Then
        If Not IsNumeric(DataComboAgent.BoundText) Then
            TemResponce = MsgBox("You have not selected an agent", vbInformation, "Agent")
            DataComboAgent.SetFocus
            Exit Function
        End If
        If TemAgentCredit - Val(txtNetTotal.Text) < (0 - TemAgentMaxCredit) Then
            TemResponce = MsgBox("This bill will lead to increase the credit limit of the agent. If you want to proceed, increase the credit limit or adviced the agent to settle cash", vbInformation, "Credit Limit")
            DataComboAgent.SetFocus
            Exit Function
        End If
    End If
        If OptionPrintOne.Value = False And OptionDoNotPrint.Value = False Then
            TemResponce = MsgBox("You have not selected weather to print or not?", vbInformation, "Printing")
            OptionPrintOne.SetFocus
            Exit Function
        End If
    CanSettlePayment = True
End Function

Private Sub bttnRemove_Click()
    If PatientGrid.Row < 1 Then Exit Sub
    PatientGrid.Col = 11
    If Not IsNumeric(PatientGrid.Text) Then Exit Sub
    With DataEnvironment1.rssqlPatientFacility
        If .State = 1 Then .Close
        .Source = "SELECT tblpatientfacility.* from tblpatientfacility where PatientFacility_ID = " & Val(PatientGrid.Text)
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .Delete adAffectCurrent
        .Close
    End With
    Call FormatPatientGrid
    Call FillPatientGrid
    Call frmPatientFacilityList.FormatPatientFacilityList
    Call frmPatientFacilityList.FillPatientFacilityList
End Sub

Private Sub bttnSearchPatient_Click()
    Call ClearSearchValues
    Call PrepareToSearchPatient
End Sub

Private Sub bttnSelect_Click()
    PatientAdded = True
    Call TemperaryBill
    Call PrepareToBook
End Sub

Private Sub TemperaryBill()
    If TemBillFinished = True Then Exit Sub
    With DataEnvironment1.rssqlPatientBill
        If .State = 1 Then .Close
        .Source = "Select tblPatientbill.* from tblpatientbill"
        If .State = 0 Then .Open
        .AddNew
        !patient_ID = TemPatientID
        !BillSuccess = False
        .Update
        TemBillID = !PatientBill_ID
        .Close
    End With
    TemBillFinished = True
End Sub
Private Sub PrepareForTwoSecessions()
    FrameSecession.Enabled = True
    OptionNotRelevent.Value = False
    OptionNotRelevent.Enabled = False
    OptionMorning.Enabled = True
    OptionEvening.Enabled = True
    OptionNoPreferance.Enabled = True
End Sub

Private Sub PrepareForOneSecession()
    FrameSecession.Enabled = False
    OptionMorning.Enabled = False
    OptionEvening.Enabled = False
    OptionNoPreferance.Enabled = False
    OptionNotRelevent.Enabled = True
    OptionNotRelevent.Value = True
End Sub

Private Sub DataComboAgent_Click(Area As Integer)
    Dim TemResponce  As Integer
    If Not IsNumeric(DataComboAgent.BoundText) Then Exit Sub
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblinstitutions.* from tblinstitutions where Institution_ID = " & DataComboAgent.BoundText
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!InstitutionCredit) Then
            TemAgentCredit = !InstitutionCredit
        Else
            TemAgentCredit = Empty
        End If
        txtAgentBalance.Caption = Format(TemAgentCredit, "#0.00")
        If Not IsNull(!InstitutionMaxCredit) Then
            TemAgentMaxCredit = !InstitutionMaxCredit
        Else
            TemAgentMaxCredit = 0
        End If
                
                
        If (0 - TemAgentMaxCredit) > TemAgentCredit Then
            TemResponce = MsgBox("This agent has already exceeded the credit limit, Increase the credit limit or ask the agent to pay some credit", vbInformation, "Exceed Credit Limit")
            DataComboAgent.Text = Empty
            DataComboAgent.SetFocus
        End If
        If !InstitutionBlackListed = True Then
            TemResponce = MsgBox("This agent is black listed, Select another agent or discuss with the management to remove from the Black List", vbInformation, "Black Listed Patient")
            DataComboAgent.Text = Empty
            DataComboAgent.SetFocus
        End If
        .Close
    End With
End Sub

Private Sub DataComboDoctorStaff_Click(Area As Integer)
    
On Error Resume Next
    
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    frmPatientFacilityList.DataComboDoctorStaff.BoundText = DataComboDoctorStaff.BoundText
    With DataEnvironment1.rssqlTem2
        If .State = 1 Then .Close
        .Source = "SELECT tblfacilitystaff.* from tblfacilitystaff where (HospitalFacility_ID = " & DataComboFacility.BoundText & ") and (Staff_ID = " & DataComboDoctorStaff.BoundText & ")"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!FacilityStaff_ID) Then TemStaffFacilityID = !FacilityStaff_ID
        If Not IsNull(!staff_ID) Then TemstaffID = !staff_ID
        If !TwoSecessions = True Then
            PrepareForTwoSecessions
        Else
            PrepareForOneSecession
        End If
        Select Case TemCatogery
        Case Doctor:
            LblDoctorStaff.Caption = "Doctor :"
            lblInstitutionFee.Caption = "Institution Fee"
            lblPersonalFee.Caption = "Doctor Fee"
            lblOtherFee.Caption = !OtherChargeName
            txtPersonalFee.Text = Format(!usualpersonalFee, "#0.00")
            txtInstitutionFee.Text = Format(!UsualInstitutionFee, "#0.00")
            txtOtherFee.Text = Format(!UsualOtherCharge, "#0.00")
            frmPatientFacilityList.LblDoctorStaff.Caption = "Doctor :"
        Case Staff:
            LblDoctorStaff.Caption = "Staff Member"
            lblInstitutionFee.Caption = "Institution Fee"
            lblPersonalFee.Caption = "Staff Fee"
            lblOtherFee.Caption = !OtherChargeName
            txtPersonalFee.Text = Format(!usualpersonalFee, "#0.00")
            txtInstitutionFee.Text = Format(!UsualInstitutionFee, "#0.00")
            txtOtherFee.Text = Format(!UsualOtherCharge, "#0.00")
            frmPatientFacilityList.LblDoctorStaff.Caption = "Staff Member :"

        Case Investigation:
            LblDoctorStaff.Caption = "Investigation"
            lblInstitutionFee.Caption = "Institution Fee"
            lblPersonalFee.Caption = "Investigation Fee"
            lblOtherFee.Caption = !OtherChargeName
            txtPersonalFee.Text = Format(!usualpersonalFee, "#0.00")
            txtInstitutionFee.Text = Format(!UsualInstitutionFee, "#0.00")
            txtOtherFee.Text = Format(!UsualOtherCharge, "#0.00")
            frmPatientFacilityList.LblDoctorStaff.Caption = "Investigation :"

        Case Other:
'            lblDoctorStaff.Caption = "Doctor"
'            lblInstitutionFee.Caption = "Institution Fee"
'            lblPersonalFee.Caption = "Doctor Fee"
'            lblOtherFee.Caption = !OtherChargeName
'            txtPersonalFee.Text = Format(!UsualPersonalFee, "#0.00")
'            txtInstitutionFee.Text = Format(!UsualInstitutionFee, "#0.00")
'            txtOtherFee.Text = Format(!UsualOtherCharge, "#0.00")
        End Select
        .Close
    
    End With
'    Call frmPatientFacilityList.FormatPatientFacilityList
'    Call frmPatientFacilityList.FillPatientFacilityList
    Call CalculateFacilityFee
End Sub


Private Sub CalculateFacilityFee()
    txtTotalFee = Val(txtPersonalFee.Text) + Val(txtInstitutionFee.Text) + Val(txtOtherFee.Text)
    txtTotalFee = Format(Val(txtTotalFee.Text), "#0.00")
End Sub



Private Sub DataComboFacility_Click(Area As Integer)

On Error Resume Next

    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    frmPatientFacilityList.DataComboFacility.BoundText = DataComboFacility.BoundText
    TemHospitalFacilityID = DataComboFacility.BoundText
    DataComboDoctorStaff.Text = Empty
    txtPersonalFee.Text = Empty
    txtInstitutionFee.Text = Empty
    txtOtherFee.Text = Empty
    txtTotalFee.Text = Empty
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = "SELECT tblhospitalfacility.* from tblhospitalfacility where hospitalfacility_ID = " & DataComboFacility.BoundText
        If .State = 0 Then .Open
        TemCatogery = !PersonCatogery
        .Close
    End With
    With DataComboDoctorStaff
        .RowMember = Empty
        .ListField = Empty
        .BoundColumn = Empty
    End With
    With DataEnvironment1.rssqlBookingFacility
        If .State = 1 Then .Close
        Select Case TemCatogery
            Case Doctor:
                .Source = "SELECT tblfacilitystaff.* , tbldoctor.* FROM tblfacilitystaff left join tbldoctor on tblfacilitystaff.staff_ID = tbldoctor.doctor_ID where HospitalFacility_ID = " & DataComboFacility.BoundText & " order by doctorname"
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                DataComboDoctorStaff.RowMember = "sqlBookingFacility"
                DataComboDoctorStaff.ListField = "doctorname"
                DataComboDoctorStaff.BoundColumn = "doctor_ID"
            Case Staff:
                .Source = "SELECT tblfacilitystaff.* , tblstaff.* FROM tblfacilitystaff left join tblstaff on tblfacilitystaff.staff_ID = tblstaff.staff_ID where HospitalFacility_ID = " & DataComboFacility.BoundText & " order by staffname"
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                DataComboDoctorStaff.RowMember = "sqlBookingFacility"
                DataComboDoctorStaff.ListField = "staffname"
                DataComboDoctorStaff.BoundColumn = "tblstaff.Staff_ID"
            Case Investigation:
                .Source = "SELECT tblfacilitystaff.* , tblinvestigations.* FROM tblfacilitystaff left join tblinvestigations on tblfacilitystaff.staff_ID = tblinvestigations.investigation_ID where HospitalFacility_ID = " & DataComboFacility.BoundText & " order by investigation"
                If .State = 0 Then .Open
                If .RecordCount = 0 Then Exit Sub
                DataComboDoctorStaff.RowMember = "sqlBookingFacility"
                DataComboDoctorStaff.ListField = "investigation"
                DataComboDoctorStaff.BoundColumn = "investigation_ID"
            Case Other:
        End Select
        .Close
    End With
    Call frmPatientFacilityList.FormatPatientFacilityList
    Call frmPatientFacilityList.FillPatientFacilityList

End Sub


Private Sub DTPickerAppointment_Change()
    Dim TemResponce  As Integer
    If DTPickerAppointment.Value < Date Then
        TemResponce = MsgBox("You can't book for the past. Today is " & Format(Date, "dd/mmm/yyyy") & " and you are going to book for " & Format(DTPickerAppointment.Value, "dd/mmm/yyyy") & " , Please change the date", vbInformation, "Can't book for the past")
        DTPickerAppointment.SetFocus
        DTPickerAppointment = Date
    End If
    frmPatientFacilityList.DTPickerAppointment.Value = DTPickerAppointment.Value
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    frmPatientFacilityList.FormatPatientFacilityList
    frmPatientFacilityList.FillPatientFacilityList
End Sub

Private Sub Form_Load()
    Call Setcolours
    
    PatientAdded = False
    Call PrepareToBook
'    Call PrepareToSearchPatient
 '   Call ClearSearchValues
    Call frmPatientFacilityList.FormatPatientFacilityList
    Call FormatPatientGrid
    Call FormatGrid
    DTPickerAppointment = Date
    frmPatientFacilityList.DTPickerAppointment.Value = DTPickerAppointment.Value
    frmPatientFacilityList.DataComboFacility.Text = Empty
    frmPatientFacilityList.DataComboDoctorStaff.Text = Empty
End Sub

Private Sub PrepareToSearchPatient()
    PatientAdded = False
    
    TemPatientID = Empty
    SSTab1.Visible = False
    frameSearchPatient.Visible = True
        
    FullPaid = False
    GrossTotal = Empty
    Discount = Empty
    NetTotal = Empty
    
    TemAppointmentTime = Empty
    TemBillID = Empty
    TemCatogery = Empty
    TemDailyMaximum = Empty
    TemHospitalFacilityID = Empty
    TemPatientFacilityID = Empty
    TemPatientID = Empty
    TemStaffFacilityID = Empty
    TemstaffID = Empty

    FormatPatientGrid
    FormatPaymentGrid
    SSTab1.Tab = 1
    
End Sub

Private Sub PrepareToBook()
    
    FullPaid = False
    GrossTotal = Empty
    Discount = Empty
    NetTotal = Empty
    
    TemAppointmentTime = Empty
    TemBillID = Empty
    TemCatogery = Empty
    TemDailyMaximum = Empty
    TemHospitalFacilityID = Empty
    TemPatientFacilityID = Empty
    TemPatientID = Empty
    TemStaffFacilityID = Empty
    TemstaffID = Empty

    FormatPatientGrid
    FormatPaymentGrid

    
    Call ClearPatientValues
    Call ClearSearchValues
    PatientAdded = False
    TemBillFinished = False
    SSTab1.Visible = True
    SSTab1.Tab = 0
    frameSearchPatient.Visible = False
    OptionCash.Value = True
    Call SetPaymentMethod
    
End Sub

Private Sub bttnClose_Click()
    Dim TemResponce  As Integer
    
    If GridPayment.Rows > 1 Then
        TemResponce = MsgBox("You can't exit when there is a bill to settle", vbCritical, "Settle Bill")
        Exit Sub
    End If
    
    Unload frmPatientFacilityList
    Unload Me
End Sub

Private Sub bttnSearch_Click()
    If Trim(txtSearchID.Text) = "" And Trim(txtSearchFirstName.Text) = "" And Trim(txtSearchSurname.Text) = "" Then
        ListAllPatients
    ElseIf Trim(txtSearchID.Text) <> "" Then
        SearchFromID
    ElseIf Trim(txtSearchID.Text) = "" And Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) = "" Then
        ListFirstNames
    ElseIf Trim(txtSearchID.Text) = "" And Trim(txtSearchFirstName.Text) = "" And Trim(txtSearchSurname.Text) <> "" Then
        ListSurname
    ElseIf Trim(txtSearchID.Text) = "" And Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) <> "" Then
        ListBothNames
    End If
    ClearSearchValues
End Sub

Private Sub ClearSearchValues()
    txtSearchFirstName.Text = Empty
    txtSearchID.Text = Empty
    txtSearchSurname.Text = Empty
End Sub

Private Sub FormatGrid()
    Dim BorderMargin As Long
    BorderMargin = 100
    With grid1
        .Clear
        .Rows = 1
        .Cols = 3
        .ColWidth(0) = 600
        .ColWidth(1) = ((.Width) - (.ColWidth(0)) - BorderMargin) * 2 / 5
        .ColWidth(2) = ((.Width) - (.ColWidth(0)) - BorderMargin) * 3 / 5
        .Col = 0
        .CellAlignment = 4
        .Text = "ID"
        .Col = 1
        .CellAlignment = 4
        .Text = "Firstname"
        .Col = 2
        .CellAlignment = 4
        .Text = "Surname"
    End With
    bttnSelect.Enabled = False
End Sub

Private Sub DisplayPhoto()
    ImagePatient.Picture = LoadPicture()
    ImagePatient.Stretch = True
    On Error Resume Next
    ImagePatient.Picture = LoadPicture(txtPhoto.Text)
End Sub


Private Sub FillGrid()
    Dim NowRow As Long
    With DataEnvironment1.rssqlPatientMain
        If .RecordCount = 0 Then
            bttnSelect.Enabled = False
            Exit Sub
        Else
            bttnSelect.Enabled = True
        End If
        .MoveFirst
        NowRow = 0
        While .EOF = False
            NowRow = NowRow + 1
            grid1.Rows = NowRow + 1
            grid1.Row = NowRow
            grid1.Col = 0
            grid1.CellAlignment = 7
            grid1.Text = !patient_ID
            grid1.Col = 1
            grid1.CellAlignment = 7
            If Not IsNull(!firstname) Then grid1.Text = !firstname
            grid1.Col = 2
            grid1.CellAlignment = 7
            If Not IsNull(!surname) Then grid1.Text = !surname
            .MoveNext
        Wend
    End With
End Sub

Private Sub ListAllPatients()
    Dim NowRow As Long
    Call FormatGrid
    With DataEnvironment1.rssqlPatientMain
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails order by Patient_ID"
        If .State = 0 Then .Open
        Call FillGrid
        .Close
    End With
End Sub

Private Sub SearchFromID()
    Dim NowRow As Long
    Dim TemResponce  As Integer
    Call FormatGrid
    With DataEnvironment1.rssqlPatientMain
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where Patient_ID = " & txtSearchID.Text & " order by Patient_ID"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then
            TemResponce = MsgBox("There is no record with the patient ID of " & txtSearchID.Text, vbCritical, "Wrong ID")
            txtSearchID.SetFocus
            SendKeys "{Home}+{end}"
        End If
        Call FillGrid
        .Close
    End With
End Sub

Private Sub ListFirstNames()
    Dim NowRow As Long
    Call FormatGrid
    With DataEnvironment1.rssqlPatientMain
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where firstname like '" & txtSearchFirstName.Text & "%' order by FIrstname"
        If .State = 0 Then .Open
        Call FillGrid
        .Close
    End With
End Sub

Private Sub ListSurname()
    Dim NowRow As Long
    Call FormatGrid
    With DataEnvironment1.rssqlPatientMain
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where surname like '" & txtSearchSurname.Text & "%' order by surname"
        If .State = 0 Then .Open
        Call FillGrid
        .Close
    End With
End Sub

Private Sub ListBothNames()
    Dim NowRow As Long
    Call FormatGrid
    With DataEnvironment1.rssqlPatientMain
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where (surname like '" & txtSearchSurname.Text & "%') and ( firstname like '" & txtSearchFirstName.Text & "%') order by FIrstname, surname"
        If .State = 0 Then .Open
        Call FillGrid
        .Close
    End With
End Sub

Private Sub Grid1_Click()
grid1.Col = 0
    If grid1.Row < 1 Or Not IsNumeric(grid1.Text) Then
        bttnSelect.Enabled = False
        Exit Sub
    Else
        grid1.Col = 0
        If Not IsNumeric(grid1.Text) Then Exit Sub
        TemPatientID = Val(grid1.Text)
        Call GetDetails
        'Call bttnSelect_Click
        grid1.Col = 0
        grid1.ColSel = grid1.Cols - 1
        bttnSelect.Enabled = True
    End If
End Sub



Private Sub grid1_DblClick()
grid1.Col = 0
    If grid1.Row < 1 Or Not IsNumeric(grid1.Text) Then
        bttnSelect.Enabled = False
        Exit Sub
    Else
        grid1.Col = 0
        If Not IsNumeric(grid1.Text) Then Exit Sub
        TemPatientID = Val(grid1.Text)
        Call GetDetails
        Call bttnSelect_Click
        grid1.Col = 0
        grid1.ColSel = grid1.Cols - 1
        bttnSelect.Enabled = True
    End If
End Sub

Private Sub OptionAgent_Click()
    Call SetPaymentMethod
End Sub

Private Sub OptionCash_Click()
    Call SetPaymentMethod
End Sub

Private Sub SetPaymentMethod()
    If OptionCash.Value = True Then
        FrameCash.Visible = True
        FrameCredit.Visible = False
        FrameCreditCard.Visible = False
        FrameCheque.Visible = False
        FrameAgent.Visible = False
    ElseIf OptionCredit.Value = True Then
        FrameCash.Visible = False
        FrameCredit.Visible = True
        FrameCreditCard.Visible = False
        FrameCheque.Visible = False
        FrameAgent.Visible = False
    ElseIf OptionAgent.Value = True Then
        FrameCash.Visible = False
        FrameCredit.Visible = False
        FrameCreditCard.Visible = False
        FrameCheque.Visible = False
        FrameAgent.Visible = True
    ElseIf OptionCreditCard.Value = True Then
        FrameCash.Visible = False
        FrameCredit.Visible = False
        FrameCreditCard.Visible = True
        FrameCheque.Visible = False
        FrameAgent.Visible = False
    ElseIf OptionCheque.Value = True Then
        FrameCash.Visible = False
        FrameCredit.Visible = False
        FrameCreditCard.Visible = False
        FrameCheque.Visible = True
        FrameAgent.Visible = False
    End If


End Sub

Private Sub OptionCheque_Click()
    Call SetPaymentMethod
End Sub

Private Sub OptionCredit_Click()
    Call SetPaymentMethod
End Sub

Private Sub OptionCreditCard_Click()
    Call SetPaymentMethod
End Sub

Private Sub OptionEvening_Click()
    If OptionEvening.Value = True Then
        frmPatientFacilityList.OptionEvening.Value = True
    Else
        frmPatientFacilityList.OptionEvening.Value = False
    End If
End Sub

Private Sub OptionMorning_Click()
    If OptionMorning.Value = True Then
        frmPatientFacilityList.OptionMorning.Value = True
    Else
        frmPatientFacilityList.OptionMorning.Value = False
    End If
End Sub

Private Sub OptionNoPreferance_Click()
    If OptionNoPreferance.Value = True Then
        frmPatientFacilityList.OptionNoPreferance.Value = True
    Else
        frmPatientFacilityList.OptionNoPreferance.Value = False
    End If
End Sub

Private Sub OptionNotRelevent_Click()
    If OptionNotRelevent.Value = True Then
        frmPatientFacilityList.OptionNotRelevent.Value = True
    Else
        frmPatientFacilityList.OptionNotRelevent.Value = False
    End If
End Sub


Private Sub PatientGrid_Click()
    If PatientGrid.Rows <= 1 Or PatientGrid.Row < 1 Then
        bttnRemove.Enabled = False
        Exit Sub
    Else
        bttnRemove.Enabled = True
        Exit Sub
    End If
    PatientGrid.ColSel = PatientGrid.Cols - 1
    PatientGrid.Col = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If PatientAdded = False And SSTab1.Tab <> 0 Then
        If AddToPatientDatabase = False Then Exit Sub
        Call TemperaryBill
    End If
End Sub

Private Function AddToPatientDatabase() As Boolean

AddToPatientDatabase = False
    
    Dim TemResponce  As Integer
    
    If Trim(txtFirstName.Text) = "" Then
        TemResponce = MsgBox("Please enter the Firstname", vbCritical, "First Name")
        SSTab1.Tab = 0
'        txtFirstName.SetFocus
        Exit Function
    End If

    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "Select tblpatientmaindetails.* from tblPatientMainDetails"
        If .State = 0 Then .Open
        .AddNew
        !firstname = txtFirstName.Text
        !surname = txtSurname.Text
        !othernames = txtOtherName.Text
        If IsNumeric(DataComboTitle.BoundText) Then !Title_ID = DataComboTitle.BoundText
        If IsNumeric(DataComboSex.BoundText) Then !sex_ID = DataComboSex.BoundText
        If IsNumeric(DataComboMarietal.BoundText) Then !Marital_ID = DataComboMarietal.BoundText
        If IsNumeric(DataComboRace.BoundText) Then !Race_ID = DataComboRace.BoundText
        !NICNo = txtNIC.Text
        !Address = txtAddress.Text
        !phone = txtTelephone.Text
        !fax = txtFax.Text
        !email = txtEmail.Text
        !DateOfBirth = DTPickerDOB.Value
        !notes = txtNotes.Text
        !registeredDate = Date
        .Update
        TemPatientID = !patient_ID
        .Close
    
AddToPatientDatabase = True
PatientAdded = True
Exit Function

ErrorHandler:
    TemResponce = MsgBox("An unknown error has occured", vbCritical, "Error")
    .CancelUpdate
    .Close
    Exit Function
    
    End With
End Function

Private Sub txtCashPaid_Change()
    lblCashBalance.Caption = Format((Val(txtCashPaid.Text) - Val(txtNetTotal.Text)), "#0.00")
End Sub

Private Sub txtCashPaid_GotFocus()
    txtCashPaid.Alignment = 0
End Sub

Private Sub txtCashPaid_LostFocus()
    txtCashPaid.Alignment = 1
    txtCashPaid.Text = Format(txtCashPaid.Text, "#0.00")
End Sub



Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDiscount_LostFocus
End Sub

Private Sub txtDiscount_LostFocus()
        txtNetTotal.Text = Format((Val(txtGrossTotal.Text) - Val(txtDiscount.Text)), "#0.00")
End Sub

Private Sub txtGrossTotal_Change()
    txtNetTotal.Text = Format((Val(txtGrossTotal.Text) - Val(txtDiscount.Text)), "#0.00")
    lblCashDue.Caption = Format((Val(txtGrossTotal.Text) - Val(txtDiscount.Text)), "#0.00")
    lblCashBalance.Caption = Format((Val(txtNetTotal.Text) - Val(txtCashPaid.Text)), "#0.00")
End Sub

Private Sub txtInstitutionFee_Change()
    CalculateFacilityFee
End Sub

Private Sub txtNetTotal_Change()
    
    lblCashDue.Caption = Format(Val(txtNetTotal.Text), "#0.00")
    lblCashBalance.Caption = Format(Val(txtCashPaid.Text) - Val(lblCashDue.Caption), "#0.00")
        
    lblThisTimeCredit.Caption = Format((Val(txtNetTotal.Text) - Val(txtPaidForCredit.Text)), "#0.00")
    
    lblChequeAmount.Caption = Format(Val(txtNetTotal.Text), "#0.00")
    
    lblAmount.Caption = Format(Val(txtNetTotal.Text), "#0.00")
    
    lblAgentAmount.Caption = Format(Val(txtNetTotal.Text), "#0.00")
    
End Sub

Private Sub txtOtherFee_Change()
    CalculateFacilityFee
End Sub

Private Sub txtPaidForCredit_Change()
    lblThisTimeCredit.Caption = Format((Val(txtNetTotal.Text) - Val(txtPaidForCredit.Text)), "#0.00")
End Sub

Private Sub txtPersonalFee_Change()
    CalculateFacilityFee
End Sub

Private Sub txtSearchFirstName_Change()
    ClearPatientValues
    If Trim(txtSearchFirstName.Text) = "" Then Exit Sub
    If Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) = "" Then ListFirstNames
    If Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) <> "" Then ListBothNames
End Sub

Private Sub txtSearchID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SearchFromID
    End If
    If Len(txtSearchID.Text) = 0 And KeyAscii = 45 Then
        KeyAscii = 0
    End If
    If KeyAscii >= 58 Or (KeyAscii <= 47 And KeyAscii <> 45 And KeyAscii <> 8 And KeyAscii <> 13) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSearchSurname_Change()
    ClearPatientValues
    If Trim(txtSearchSurname.Text) = "" Then Exit Sub
    If Trim(txtSearchFirstName.Text) = "" And Trim(txtSearchSurname.Text) = "" Then ListSurname
    If Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) <> "" Then ListBothNames
End Sub

Private Sub GetDetails()
    Call ClearPatientValues
    With DataEnvironment1.rssqlPatientMain
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientmaindetails.* FROM tblPatientmainDetails where (patient_ID =" & TemPatientID & ")"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!firstname) Then txtFirstName.Text = !firstname
        If Not IsNull(!surname) Then txtSurname.Text = !surname
        If Not IsNull(!othernames) Then txtOtherName.Text = !othernames
        If Not IsNull(!Title_ID) Then DataComboTitle.BoundText = !Title_ID
        If Not IsNull(!sex_ID) Then DataComboSex.BoundText = !sex_ID
        If Not IsNull(!Marital_ID) Then DataComboMarietal.BoundText = !Marital_ID
        If Not IsNull(!Race_ID) Then DataComboRace.BoundText = !Race_ID
        If Not IsNull(!NICNo) Then txtNIC.Text = !NICNo
        If Not IsNull(!Address) Then txtAddress.Text = !Address
        If Not IsNull(!phone) Then txtTelephone.Text = !phone
        If Not IsNull(!fax) Then txtFax.Text = !fax
        If Not IsNull(!email) Then txtEmail.Text = !email
        If Not IsNull(!DateOfBirth) Then DTPickerDOB.Value = !DateOfBirth
        If Not IsNull(!notes) Then txtNotes.Text = !notes
        If Not IsNull(!photo) Then txtPhoto.Text = !photo
        If Not IsNull(!Credit) Then
            TemPatientCredit = !Credit
        Else
            TemPatientCredit = 0
        End If
        If Not IsNull(!maxcredit) Then
            TemPatientMaxCredit = !maxcredit
        Else
            TemPatientMaxCredit = 0
        End If
        BlackListedPatient = !BlackListedPatient
        lblTotalCredit = Format(TemPatientCredit, "#0.00")
        Call DisplayPhoto
        .Close
    End With
End Sub


Private Sub ClearPatientValues()
    txtFirstName.Text = Empty
    txtOtherName.Text = Empty
    txtSurname.Text = Empty
    txtAddress.Text = Empty
    txtNIC.Text = Empty
    txtTelephone.Text = Empty
    txtFax.Text = Empty
    txtEmail.Text = Empty
    txtNotes.Text = Empty
    DataComboTitle.Text = Empty
    DataComboSex.Text = Empty
    DataComboMarietal.Text = Empty
    DataComboRace.Text = Empty
    txtPhoto.Text = Empty
    ImagePatient.Picture = LoadPicture()
    DTPickerDOB.Value = Date
    txtAge.Text = Empty
    
    
    txtPaidForCredit.Text = Empty
    txtAgentBalance.Caption = Empty
    txtAuthorizationCode.Text = Empty
    txtBranch.Text = Empty
    txtCardNumber.Text = Empty
    txtCashPaid.Text = Empty
    txtChequeNo.Text = Empty
    txtDiscount.Text = Empty
    txtGrossTotal.Text = Empty
    txtNetTotal.Text = Empty
    txtPaidForCredit.Text = Empty
    txtTotalFee.Text = Empty
    lblAgentAmount.Caption = Empty
    lblAmount.Caption = Empty
    lblCashBalance.Caption = Empty
    lblCashDue.Caption = Empty
    lblChequeAmount.Caption = Empty
    lblThisTimeCredit.Caption = Empty
    lblTotalCredit.Caption = Empty

    DTPickerTime.Value = TimeSerial(0, 0, 0)

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

bttnList.BackColor = BttnBackColour
bttnList.ForeColor = BttnForeColour

bttnPay.BackColor = BttnBackColour
bttnPay.ForeColor = BttnForeColour

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour

bttnSearch.BackColor = BttnBackColour
bttnSearch.ForeColor = BttnForeColour

OptionPrintSeperately.BackColor = TxtBackColour
OptionPrintSeperately.ForeColor = TxtForeColour

bttnRemove.BackColor = BttnBackColour
bttnRemove.ForeColor = BttnForeColour


frmBoooking.BackColor = FrmBackColour
frmBoooking.ForeColor = FrmForeColour

Frame1.BackColor = FrameBackColour
Frame1.ForeColor = FrameForeColour

FrameAgent.BackColor = FrameBackColour
FrameAgent.ForeColor = FrameForeColour

FrameBooking.BackColor = FrameBackColour
FrameBooking.ForeColor = FrameForeColour

FrameCash.BackColor = FrameBackColour
FrameCash.ForeColor = FrameForeColour

FrameCheque.BackColor = FrameBackColour
FrameCheque.ForeColor = FrameForeColour
FrameCredit.BackColor = FrameBackColour
FrameCredit.ForeColor = FrameForeColour
FrameCreditCard.BackColor = FrameBackColour
FrameCreditCard.ForeColor = FrameForeColour
FramePatient.BackColor = FrameBackColour
FramePatient.ForeColor = FrameForeColour
FramePayment.BackColor = FrameBackColour
FramePayment.ForeColor = FrameForeColour

FramePaymentMethod.BackColor = FrameBackColour
FramePaymentMethod.ForeColor = FrameForeColour
frameSearchPatient.BackColor = FrameBackColour
frameSearchPatient.ForeColor = FrameForeColour
'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour



'chk.BackColor = LblBackColour
'chkCurrentlyChanneling.ForeColor = LblForeColour

DataComboBank.BackColor = TxtBackColour
DataComboBank.ForeColor = TxtForeColour

DataComboAgent.BackColor = TxtBackColour
DataComboAgent.ForeColor = TxtForeColour

DataComboBank.BackColor = TxtBackColour
DataComboBank.ForeColor = TxtForeColour

DataComboDoctorStaff.BackColor = TxtBackColour
DataComboDoctorStaff.ForeColor = TxtForeColour

DataComboTitle.BackColor = TxtBackColour
DataComboTitle.ForeColor = TxtForeColour

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

Label11.BackColor = LblBackColour
Label11.ForeColor = LblForeColour
Label12.BackColor = LblBackColour
Label12.ForeColor = LblForeColour
Label13.BackColor = LblBackColour
Label13.ForeColor = LblForeColour
Label14.BackColor = LblBackColour
Label14.ForeColor = LblForeColour
Label15.BackColor = LblBackColour
Label15.ForeColor = LblForeColour
Label16.BackColor = LblBackColour
Label16.ForeColor = LblForeColour
Label2.BackColor = LblBackColour
Label2.ForeColor = LblForeColour
Label18.BackColor = LblBackColour
Label18.ForeColor = LblForeColour
Label3.BackColor = LblBackColour
Label3.ForeColor = LblForeColour
Label20.BackColor = LblBackColour
Label20.ForeColor = LblForeColour
Label21.BackColor = LblBackColour
Label21.ForeColor = LblForeColour
Label4.BackColor = LblBackColour
Label4.ForeColor = LblForeColour
Label23.BackColor = LblBackColour
Label23.ForeColor = LblForeColour
Label24.BackColor = LblBackColour
Label24.ForeColor = LblForeColour
Label25.BackColor = LblBackColour
Label25.ForeColor = LblForeColour
Label26.BackColor = LblBackColour
Label26.ForeColor = LblForeColour
Label27.BackColor = LblBackColour
Label27.ForeColor = LblForeColour
Label4.BackColor = LblBackColour
Label4.ForeColor = LblForeColour
Label5.BackColor = LblBackColour
Label5.ForeColor = LblForeColour
Label6.BackColor = LblBackColour
Label6.ForeColor = LblForeColour

lblAmount.BackColor = LblBackColour
lblAmount.ForeColor = LblForeColour

lblCashBalance.BackColor = LblBackColour
lblCashBalance.ForeColor = LblForeColour

'lblCashPaid.BackColor = LblBackColour
'lblCashPaid.ForeColor = LblForeColour

lblChequeAmount.BackColor = LblBackColour
lblChequeAmount.ForeColor = LblForeColour

lblThisTimeCredit.BackColor = LblBackColour
lblThisTimeCredit.ForeColor = LblForeColour

'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour



txtAddress.BackColor = TxtBackColour
txtAddress.ForeColor = TxtForeColour

txtAge.BackColor = TxtBackColour
txtAge.ForeColor = TxtForeColour

txtAgentBalance.BackColor = TxtBackColour
txtAgentBalance.ForeColor = TxtForeColour
txtAuthorizationCode.BackColor = TxtBackColour
txtAuthorizationCode.ForeColor = TxtForeColour
'txtCashDue.BackColor = TxtBackColour
'txtCashDue.ForeColor = TxtForeColour
txtChequeNo.BackColor = TxtBackColour
txtChequeNo.ForeColor = TxtForeColour
txtDiscount.BackColor = TxtBackColour
txtDiscount.ForeColor = TxtForeColour
txtEmail.BackColor = TxtBackColour
txtEmail.ForeColor = TxtForeColour
txtFax.BackColor = TxtBackColour
txtFax.ForeColor = TxtForeColour
txtFirstName.BackColor = TxtBackColour
txtFirstName.ForeColor = TxtForeColour
txtGrossTotal.BackColor = TxtBackColour
txtGrossTotal.ForeColor = TxtForeColour
txtNetTotal.BackColor = TxtBackColour
txtNetTotal.ForeColor = TxtForeColour

txtNIC.BackColor = TxtBackColour
txtNIC.ForeColor = TxtForeColour
txtNotes.BackColor = TxtBackColour
txtNotes.ForeColor = TxtForeColour
txtOtherName.BackColor = TxtBackColour
txtOtherName.ForeColor = TxtForeColour
txtPaidForCredit.BackColor = TxtBackColour
txtPaidForCredit.ForeColor = TxtForeColour
txtSearchFirstName.BackColor = TxtBackColour
txtSearchFirstName.ForeColor = TxtForeColour


OptionAgent.BackColor = TxtBackColour
OptionAgent.ForeColor = TxtForeColour

OptionCash.BackColor = TxtBackColour
OptionCash.ForeColor = TxtForeColour

OptionCheque.BackColor = TxtBackColour
OptionCheque.ForeColor = TxtForeColour
OptionDoNotPrint.BackColor = TxtBackColour
OptionDoNotPrint.ForeColor = TxtForeColour
OptionCredit.BackColor = TxtBackColour
OptionCredit.ForeColor = TxtForeColour

OptionCreditCard.BackColor = TxtBackColour
OptionCreditCard.ForeColor = TxtForeColour

OptionMaster.BackColor = TxtBackColour
OptionMaster.ForeColor = TxtForeColour

OptionPrintOne.BackColor = TxtBackColour
OptionPrintOne.ForeColor = TxtForeColour

'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour



txtSearchID.BackColor = TxtBackColour
txtSearchID.ForeColor = TxtForeColour
txtSearchSurname.BackColor = TxtBackColour
txtSearchSurname.ForeColor = TxtForeColour
txtSurname.BackColor = TxtBackColour
txtSurname.ForeColor = TxtForeColour
txtTelephone.BackColor = TxtBackColour
txtTelephone.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour


End Sub

