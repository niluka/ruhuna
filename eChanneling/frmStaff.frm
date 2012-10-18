VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Details"
   ClientHeight    =   9045
   ClientLeft      =   -165
   ClientTop       =   225
   ClientWidth     =   12045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStaff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   12045
   Begin VB.TextBox txtTemUserName 
      Height          =   495
      Left            =   6360
      TabIndex        =   71
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTabMain 
      Height          =   8055
      Left            =   5280
      TabIndex        =   9
      Top             =   240
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   14208
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Personal Details"
      TabPicture(0)   =   "frmStaff.frx":0582
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label20"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label21"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label29"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "SSTabSub"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "DataComboSpeciality"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DataComboSex"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DataComboTitle"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtListedName"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDesignation"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtRegistation"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtQualifications"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtName"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCode"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Payment Details"
      TabPicture(1)   =   "frmStaff.frx":059E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtCredit"
      Tab(1).Control(1)=   "txtBankBranch"
      Tab(1).Control(2)=   "txtAccount"
      Tab(1).Control(3)=   "txtComments"
      Tab(1).Control(4)=   "chkCurrentlyChanneling"
      Tab(1).Control(5)=   "DataComboPaymenyMethod"
      Tab(1).Control(6)=   "DataComboBank"
      Tab(1).Control(7)=   "MSFlexGrid1"
      Tab(1).Control(8)=   "Label9"
      Tab(1).Control(9)=   "Label14"
      Tab(1).Control(10)=   "Label8"
      Tab(1).Control(11)=   "Label7"
      Tab(1).Control(12)=   "Label16"
      Tab(1).Control(13)=   "Label10"
      Tab(1).Control(14)=   "Label6"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Program Usage"
      TabPicture(2)   =   "frmStaff.frx":05BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "OptionAnalyzer"
      Tab(2).Control(1)=   "OptionUser"
      Tab(2).Control(2)=   "OptionAccounting"
      Tab(2).Control(3)=   "OptionHR"
      Tab(2).Control(4)=   "OptionOwnerCovered"
      Tab(2).Control(5)=   "OptionAdministrator"
      Tab(2).Control(6)=   "OptionOwner"
      Tab(2).Control(7)=   "chkUser"
      Tab(2).Control(8)=   "txtReenterPassword"
      Tab(2).Control(9)=   "txtPassword"
      Tab(2).Control(10)=   "txtUserName"
      Tab(2).Control(11)=   "Label28"
      Tab(2).Control(12)=   "Label22"
      Tab(2).Control(13)=   "Label19"
      Tab(2).Control(14)=   "Label17"
      Tab(2).ControlCount=   15
      Begin VB.TextBox txtCode 
         Height          =   375
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   19
         Top             =   1920
         Width           =   3975
      End
      Begin VB.OptionButton OptionAnalyzer 
         Caption         =   "Analyzer"
         Height          =   255
         Left            =   -72120
         TabIndex        =   78
         Top             =   5400
         Width           =   3615
      End
      Begin VB.OptionButton OptionUser 
         Caption         =   "User"
         Height          =   255
         Left            =   -72120
         TabIndex        =   77
         Top             =   5040
         Width           =   3615
      End
      Begin VB.OptionButton OptionAccounting 
         Caption         =   "Accounting"
         Height          =   255
         Left            =   -72120
         TabIndex        =   76
         Top             =   4680
         Width           =   3615
      End
      Begin VB.OptionButton OptionHR 
         Caption         =   "Human Resources"
         Height          =   255
         Left            =   -72120
         TabIndex        =   75
         Top             =   4320
         Width           =   3615
      End
      Begin VB.OptionButton OptionOwnerCovered 
         Caption         =   "Owner (Covered)"
         Height          =   255
         Left            =   -72120
         TabIndex        =   74
         Top             =   3960
         Width           =   3615
      End
      Begin VB.OptionButton OptionAdministrator 
         Caption         =   "Administrator"
         Height          =   255
         Left            =   -72120
         TabIndex        =   73
         Top             =   3240
         Width           =   3615
      End
      Begin VB.OptionButton OptionOwner 
         Caption         =   "Owner"
         Height          =   255
         Left            =   -72120
         TabIndex        =   72
         Top             =   3600
         Width           =   3615
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "User of the program"
         Height          =   495
         Left            =   -72120
         TabIndex        =   70
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtReenterPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -72120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   69
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox txtPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -72120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   68
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtUserName 
         Height          =   375
         Left            =   -72120
         MaxLength       =   10
         TabIndex        =   67
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtCredit 
         Height          =   360
         Left            =   -72480
         MaxLength       =   100
         TabIndex        =   53
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txtBankBranch 
         Height          =   375
         Left            =   -72480
         TabIndex        =   52
         Top             =   2400
         Width           =   4095
      End
      Begin VB.TextBox txtAccount 
         Height          =   375
         Left            =   -72480
         MaxLength       =   100
         TabIndex        =   51
         Top             =   2880
         Width           =   4095
      End
      Begin VB.TextBox txtComments 
         Height          =   840
         Left            =   -72480
         MultiLine       =   -1  'True
         TabIndex        =   50
         Top             =   3360
         Width           =   4095
      End
      Begin VB.CheckBox chkCurrentlyChanneling 
         Caption         =   "Currently Channeling "
         Height          =   375
         Left            =   -74880
         TabIndex        =   49
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   15
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtQualifications 
         Height          =   375
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   21
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox txtRegistation 
         Height          =   375
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   23
         Top             =   2880
         Width           =   3975
      End
      Begin VB.TextBox txtDesignation 
         Height          =   375
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   27
         Top             =   3360
         Width           =   3975
      End
      Begin VB.TextBox txtListedName 
         Height          =   375
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   17
         Top             =   1440
         Width           =   3975
      End
      Begin MSDataListLib.DataCombo DataComboTitle 
         Bindings        =   "frmStaff.frx":05D6
         Height          =   360
         Left            =   2280
         TabIndex        =   11
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Title"
         BoundColumn     =   "Title_ID"
         Text            =   ""
         Object.DataMember      =   "sqlTitle"
      End
      Begin MSDataListLib.DataCombo DataComboSex 
         Bindings        =   "frmStaff.frx":05F5
         Height          =   360
         Left            =   4920
         TabIndex        =   13
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "Sex"
         BoundColumn     =   "Sex_ID"
         Text            =   ""
         Object.DataMember      =   "sqlSex"
      End
      Begin MSDataListLib.DataCombo DataComboSpeciality 
         Bindings        =   "frmStaff.frx":0614
         Height          =   360
         Left            =   2280
         TabIndex        =   30
         Top             =   3840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "StaffSpeciality"
         BoundColumn     =   "StaffSpeciality_ID"
         Text            =   ""
         Object.DataMember      =   "sqlStaffSpeciality"
      End
      Begin TabDlg.SSTab SSTabSub 
         Height          =   3615
         Left            =   240
         TabIndex        =   24
         Top             =   4320
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6376
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Private"
         TabPicture(0)   =   "frmStaff.frx":0633
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtPrivateMobile"
         Tab(0).Control(1)=   "txtPrivateEmail"
         Tab(0).Control(2)=   "txtPrivateFax"
         Tab(0).Control(3)=   "txtPrivateTel"
         Tab(0).Control(4)=   "txtPrivateAddress"
         Tab(0).Control(5)=   "Label27"
         Tab(0).Control(6)=   "Label5"
         Tab(0).Control(7)=   "Label4"
         Tab(0).Control(8)=   "Label3"
         Tab(0).Control(9)=   "Label2"
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Official"
         TabPicture(1)   =   "frmStaff.frx":064F
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lblOfficialEmail"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label23"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label24"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label25"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lblOfficialWebsite"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "txtOfficialEMail"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "txtOfficialFax"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "txtOfficialTel"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "txtOfficialAddress"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "txtOfficialWebsite"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).ControlCount=   10
         Begin VB.TextBox txtPrivateMobile 
            Height          =   375
            Left            =   -73800
            MaxLength       =   100
            TabIndex        =   38
            Top             =   1800
            Width           =   4695
         End
         Begin VB.TextBox txtPrivateEmail 
            Height          =   375
            Left            =   -73800
            MaxLength       =   100
            TabIndex        =   37
            Top             =   2760
            Width           =   4695
         End
         Begin VB.TextBox txtPrivateFax 
            Height          =   375
            Left            =   -73800
            MaxLength       =   100
            TabIndex        =   36
            Top             =   2280
            Width           =   4695
         End
         Begin VB.TextBox txtPrivateTel 
            Height          =   375
            Left            =   -73800
            MaxLength       =   100
            TabIndex        =   35
            Top             =   1320
            Width           =   4695
         End
         Begin VB.TextBox txtPrivateAddress 
            Height          =   735
            Left            =   -73800
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   480
            Width           =   4695
         End
         Begin VB.TextBox txtOfficialWebsite 
            Height          =   375
            Left            =   1200
            TabIndex        =   33
            Top             =   2760
            Width           =   4695
         End
         Begin VB.TextBox txtOfficialAddress 
            Height          =   735
            Left            =   1200
            MultiLine       =   -1  'True
            TabIndex        =   32
            Top             =   480
            Width           =   4695
         End
         Begin VB.TextBox txtOfficialTel 
            Height          =   375
            Left            =   1200
            TabIndex        =   31
            Top             =   1320
            Width           =   4695
         End
         Begin VB.TextBox txtOfficialFax 
            Height          =   375
            Left            =   1200
            TabIndex        =   29
            Top             =   1800
            Width           =   4695
         End
         Begin VB.TextBox txtOfficialEMail 
            Height          =   375
            Left            =   1200
            TabIndex        =   26
            Top             =   2280
            Width           =   4695
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile"
            Height          =   375
            Left            =   -74880
            TabIndex        =   48
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail"
            Height          =   375
            Left            =   -74880
            TabIndex        =   47
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            Height          =   375
            Left            =   -74880
            TabIndex        =   46
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone"
            Height          =   375
            Left            =   -74880
            TabIndex        =   45
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   375
            Left            =   -74880
            TabIndex        =   44
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblOfficialWebsite 
            BackStyle       =   0  'Transparent
            Caption         =   "Website"
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblOfficialEmail 
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   2280
            Width           =   2175
         End
      End
      Begin MSDataListLib.DataCombo DataComboPaymenyMethod 
         Bindings        =   "frmStaff.frx":066B
         Height          =   360
         Left            =   -72480
         TabIndex        =   54
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "PaymentMethod"
         BoundColumn     =   "PaymentMethod_ID"
         Text            =   ""
         Object.DataMember      =   "sqlPaymentMethod"
      End
      Begin MSDataListLib.DataCombo DataComboBank 
         Bindings        =   "frmStaff.frx":068A
         Height          =   360
         Left            =   -72480
         TabIndex        =   55
         Top             =   1920
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         ListField       =   "BankName"
         BoundColumn     =   "Bank_ID"
         Text            =   ""
         Object.DataMember      =   "sqlBank"
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   56
         Top             =   4800
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   4
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. No"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Authority Level"
         Height          =   255
         Left            =   -74640
         TabIndex        =   79
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-enter Password"
         Height          =   255
         Left            =   -74640
         TabIndex        =   66
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   255
         Left            =   -74640
         TabIndex        =   65
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   255
         Left            =   -74640
         TabIndex        =   64
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   375
         Left            =   -74880
         TabIndex        =   63
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch"
         Height          =   375
         Left            =   -74880
         TabIndex        =   62
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Account No"
         Height          =   375
         Left            =   -74880
         TabIndex        =   61
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         Height          =   375
         Left            =   -74880
         TabIndex        =   60
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Payments"
         Height          =   375
         Left            =   -74880
         TabIndex        =   59
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   375
         Left            =   -74880
         TabIndex        =   58
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Method"
         Height          =   375
         Left            =   -74880
         TabIndex        =   57
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label21 
         Caption         =   "Sex"
         Height          =   375
         Left            =   4320
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Qualifications"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Registation"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Listed Name"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   3840
         Width           =   2175
      End
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   10440
      TabIndex        =   6
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
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
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Sa&ve"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   7200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "E&dit"
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
      Left            =   720
      TabIndex        =   1
      Top             =   7200
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   600
      MaxLength       =   100
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   6495
      Left            =   600
      TabIndex        =   7
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11456
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnChange 
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label15 
      Caption         =   "Label15"
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim BorderMargin As Long
    Dim TemstaffID As Long
    Dim FromGrid As Boolean
    Dim TemUserName As String
    
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

frmStaff.BackColor = FrmBackColour
frmStaff.ForeColor = FrmForeColour


'chkCurrentlyChanneling.BackColor = LblBackColour
'chkCurrentlyChanneling.ForeColor = LblForeColour

DataComboBank.BackColor = TxtBackColour
DataComboBank.ForeColor = TxtForeColour

DataComboPaymenyMethod.BackColor = TxtBackColour
DataComboPaymenyMethod.ForeColor = TxtForeColour

DataComboSex.BackColor = TxtBackColour
DataComboSex.ForeColor = TxtForeColour

DataComboSpeciality.BackColor = TxtBackColour
DataComboSpeciality.ForeColor = TxtForeColour

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



'Label1.BackColor = LblBackColour
'Label1.ForeColor = LblForeColour
'
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
'Label16.BackColor = LblBackColour
'Label16.ForeColor = LblForeColour
'Label2.BackColor = LblBackColour
'Label2.ForeColor = LblForeColour
'Label18.BackColor = LblBackColour
'Label18.ForeColor = LblForeColour
'Label3.BackColor = LblBackColour
'Label3.ForeColor = LblForeColour
'Label20.BackColor = LblBackColour
'Label20.ForeColor = LblForeColour
''Label21.BackColor = LblBackColour
''Label21.ForeColor = LblForeColour
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
'Label4.BackColor = LblBackColour
'Label4.ForeColor = LblForeColour
'Label5.BackColor = LblBackColour
'Label5.ForeColor = LblForeColour
'Label6.BackColor = LblBackColour
'Label6.ForeColor = LblForeColour
'Label7.BackColor = LblBackColour
'Label7.ForeColor = LblForeColour
'
'Label8.BackColor = LblBackColour
'Label8.ForeColor = LblForeColour
'Label9.BackColor = LblBackColour
'Label9.ForeColor = LblForeColour
'
'lblOfficialEmail.BackColor = LblBackColour
'lblOfficialEmail.ForeColor = LblForeColour
'
'lblOfficialWebsite.BackColor = LblBackColour
'lblOfficialWebsite.ForeColor = LblForeColour


txtAccount.BackColor = TxtBackColour
txtAccount.ForeColor = TxtForeColour

txtBankBranch.BackColor = TxtBackColour
txtBankBranch.ForeColor = TxtForeColour

txtComments.BackColor = TxtBackColour
txtComments.ForeColor = TxtForeColour
txtCredit.BackColor = TxtBackColour
txtCredit.ForeColor = TxtForeColour
txtDesignation.BackColor = TxtBackColour
txtDesignation.ForeColor = TxtForeColour
txtListedName.BackColor = TxtBackColour
txtListedName.ForeColor = TxtForeColour
txtName.BackColor = TxtBackColour
txtName.ForeColor = TxtForeColour
txtCode.BackColor = TxtBackColour
txtCode.ForeColor = TxtForeColour

txtOfficialAddress.BackColor = TxtBackColour
txtOfficialAddress.ForeColor = TxtForeColour
txtOfficialEMail.BackColor = TxtBackColour
txtOfficialEMail.ForeColor = TxtForeColour
txtOfficialFax.BackColor = TxtBackColour
txtOfficialFax.ForeColor = TxtForeColour
txtOfficialTel.BackColor = TxtBackColour
txtOfficialTel.ForeColor = TxtForeColour
txtOfficialWebsite.BackColor = TxtBackColour
txtOfficialWebsite.ForeColor = TxtForeColour

txtPrivateAddress.BackColor = TxtBackColour
txtPrivateAddress.ForeColor = TxtForeColour
txtPrivateEmail.BackColor = TxtBackColour
txtPrivateEmail.ForeColor = TxtForeColour
txtPrivateFax.BackColor = TxtBackColour
txtPrivateFax.ForeColor = TxtForeColour
txtPrivateMobile.BackColor = TxtBackColour
txtPrivateMobile.ForeColor = TxtForeColour
txtPrivateTel.BackColor = TxtBackColour
txtPrivateTel.ForeColor = TxtForeColour

txtUserName.BackColor = TxtBackColour
txtUserName.ForeColor = TxtForeColour
txtPassword.BackColor = TxtBackColour
txtPassword.ForeColor = TxtForeColour
txtReenterPassword.BackColor = TxtBackColour
txtReenterPassword.ForeColor = TxtForeColour
'txtOfficialFax.BackColor = TxtBackColour
'txtOfficialFax.ForeColor = TxtForeColour
'txtOfficialTel.BackColor = TxtBackColour
'txtOfficialTel.ForeColor = TxtForeColour
'txtOfficialWebsite.BackColor = TxtBackColour
'txtOfficialWebsite.ForeColor = TxtForeColour





txtQualifications.BackColor = TxtBackColour
txtQualifications.ForeColor = TxtForeColour
txtRegistation.BackColor = TxtBackColour
txtRegistation.ForeColor = TxtForeColour
txtSearch.BackColor = TxtBackColour
txtSearch.ForeColor = TxtForeColour
End Sub
Private Sub bttnAdd_Click()
Dim TemResponce  As Integer
    
    If UserAuthority <> 1 And UserAuthority <> 2 And UserAuthority <> 3 And UserAuthority <> 4 Then
        If TemUserName <> UserName Then
            TemResponce = MsgBox("You can not add staff members. Only the owner and human resource manager can add staff members", vbCritical, "Not Allowed")
            Exit Sub
        End If
    End If
    
    Call AfterAdd
    Call ClearValues
    Call PrepairTabs
    txtName.SetFocus
End Sub

Private Sub PrepairTabs()
    SSTabMain.Tab = 0
    SSTabSub.Tab = 0
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub bttnChange_Click()
Dim TemResponce  As Integer
    
    If UserAuthority <> 1 And UserAuthority <> 2 And UserAuthority <> 3 And UserAuthority <> 4 Then
        If UCase(TemUserName) <> UCase(UserName) Then
            TemResponce = MsgBox("You can only edit your own details. Only the owner and human resource manager can change others details", vbCritical, "Not Allowed")
            Exit Sub
        End If
    End If
    
    If EditData = False Then Exit Sub
    Call ClearValues
    Call BeforeAddEdit
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub


Private Sub bttnEdit_Click()
    FromGrid = True
    grid1.Col = 2
    If Not IsNumeric(grid1.Text) Then Beep: Exit Sub
    TemstaffID = Val(grid1.Text)
    Call AfterEdit
    If SSTabMain.Tab = 0 Then txtName.SetFocus
End Sub

Private Sub bttnSave_Click()
    If SaveData = False Then Exit Sub
    Call ClearValues
    Call BeforeAddEdit
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub chkUser_Click()
    If chkUser.Value = 0 Then
        txtUserName.Enabled = False
        txtPassword.Enabled = False
        txtReenterPassword.Enabled = False
    Else
        txtUserName.Enabled = True
        txtPassword.Enabled = True
        txtReenterPassword.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Call BeforeAddEdit
    Call ClearValues
    Call FormatGrid
    Call FillGrid
    Call Setcolours
    Call SetAuthority
End Sub

Private Sub SetAuthority()
    Select Case UserAuthority
        Case AuthorityAdministrator
                OptionAdministrator.Visible = True
                OptionOwner.Visible = True
                OptionOwnerCovered.Visible = True
                OptionHR.Visible = True
                OptionAccounting.Visible = True
                OptionUser.Visible = True
                OptionAnalyzer.Visible = True
        Case AuthorityOwner
                OptionAdministrator.Visible = False
                OptionOwner.Visible = True
                OptionOwnerCovered.Visible = True
                OptionHR.Visible = True
                OptionAccounting.Visible = True
                OptionUser.Visible = True
                OptionAnalyzer.Visible = True
        Case AuthorityOwnerCOvered
                OptionAdministrator.Visible = False
                OptionOwner.Visible = True
                OptionOwnerCovered.Visible = False
                OptionHR.Visible = True
                OptionAccounting.Visible = True
                OptionUser.Visible = True
                OptionAnalyzer.Visible = True
        Case Else
                OptionAdministrator.Visible = False
                OptionOwner.Visible = False
                OptionOwnerCovered.Visible = False
                OptionHR.Visible = True
                OptionAccounting.Visible = True
                OptionUser.Visible = True
                OptionAnalyzer.Visible = True
        
    End Select
End Sub

Private Sub FormatGrid()
    BorderMargin = 100
    With grid1
        .Clear
        .Cols = 3
        .Rows = 1
        
        .ColWidth(0) = 600
        .ColWidth(2) = 1
        .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + BorderMargin)
        
        .Row = 0
        
        .Col = 0
        .Text = "No."
        .CellAlignment = 6
        
        .Col = 1
        .Text = "Name"
        
        .Col = 2
        .Text = "ID"
        .CellAlignment = 6
    End With
End Sub

Private Sub FillGrid()
    Dim NowROw As Long
    With DataEnvironment1.rssqlStaff
        If .State = 1 Then .Close
        .Source = "Select tblstaff.* from tblstaff order by staffListedName"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        NowROw = 0
        Do While .EOF = False
            If Not IsNull(!staffListedName) Then
                NowROw = NowROw + 1
                grid1.Rows = NowROw + 1
                grid1.Row = NowROw
                grid1.Col = 0
                grid1.CellAlignment = 7
                grid1.Text = NowROw
                grid1.Col = 1
                grid1.CellAlignment = 1
                grid1.Text = !staffListedName
                grid1.Col = 2
                grid1.Text = !Staff_ID
            End If
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub BeforeAddEdit()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = True
    grid1.Enabled = True
    
    bttnSave.Visible = False
    bttnChange.Visible = False
    bttnCancel.Visible = False
    
    SSTabMain.Enabled = False
    SSTabMain.Tab = 0
    SSTabSub.Tab = 0
   
    TemstaffID = Empty
    
        FromGrid = False

End Sub

Private Sub AfterAdd()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
    grid1.Enabled = False
    
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    
    SSTabMain.Enabled = True
    SSTabMain.Tab = 0
    SSTabSub.Tab = 0
    
    TemstaffID = Empty
End Sub
Private Sub AfterEdit()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
    grid1.Enabled = False
    
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    
    SSTabMain.Enabled = True
    SSTabMain.Tab = 0
    SSTabSub.Tab = 0
End Sub

Private Function SaveData() As Boolean
    SaveData = False
    Dim TemResponce  As Integer
    If Trim(txtName.Text) = "" Then
        TemResponce = MsgBox("Please enter a name of a staff to save", vbCritical + vbOKOnly, "No name")
        SSTabMain.Tab = 0
        txtName.SetFocus
        Exit Function
    End If
    
    If Trim(txtListedName.Text) = "" Then
        TemResponce = MsgBox("Please enter name to be listed before saving", vbCritical + vbOKOnly, "No name")
        txtName_LostFocus
        SSTabMain.Tab = 0
        txtListedName.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(DataComboTitle.BoundText) Then
        TemResponce = MsgBox("Please select the title", vbCritical, "Title?")
        SSTabMain.Tab = 0
        DataComboTitle.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(DataComboSex.BoundText) Then
        TemResponce = MsgBox("Please select the sex", vbCritical, "Sex?")
        SSTabMain.Tab = 0
        DataComboSex.SetFocus
        Exit Function
    End If

    If chkUser.Value = 1 Then
        If Trim(txtUserName.Text) = "" Then
            TemResponce = MsgBox("You have to enter a user name if you are going to use the program.", vbCritical, "No User Name")
            SSTabMain.Tab = 2
            txtUserName.SetFocus
            Exit Function
        End If
        If Trim(txtPassword.Text) = "" Then
            TemResponce = MsgBox("You have not entered the password", vbInformation, "No Password")
            SSTabMain.Tab = 2
            txtPassword.SetFocus
            Exit Function
        End If
        If Trim(txtPassword.Text) <> Trim(txtReenterPassword.Text) Then
            TemResponce = MsgBox("The passwords you entered are not identical", vbCritical, "Passwords?")
            SSTabMain.Tab = 2
            txtPassword.SetFocus
            Exit Function
        End If
            With DataEnvironment1.rssqlTem
                If .State = 1 Then .Close
                .Source = "SELECT tblStaff.* from tblStaff"
                If .State = 0 Then .Open
                If .RecordCount <> 0 Then
                    .MoveFirst
                    While .EOF = False
                        If Not IsNull(!StaffUserName) Then
                            txtTemUserName.Text = DecreptedWord(!StaffUserName)
                            If UCase(txtTemUserName.Text) = UCase(txtUserName.Text) Then
                                TemResponce = MsgBox("This username is already taken, Select another user name", vbCritical, "Username")
                                SSTabMain.Tab = 2
                                txtUserName.SetFocus
                                SendKeys "{home}+{end}"
                                Exit Function
                            End If
                        End If
                    .MoveNext
                    Wend
                End If
                .Close
            End With
    End If

    'On Error GoTo ErrorHandler

        With DataEnvironment1.rssqlStaff
            If .State = 0 Then .Open
            .AddNew
            !stafftitle_ID = DataComboTitle.BoundText
            !staffsex_ID = DataComboSex.BoundText
            !StaffName = txtName.Text
            !StaffCode = Val(txtCode.Text)
            !staffListedName = txtListedName.Text
            !staffqualifications = txtQualifications.Text
            !staffdesignation = txtDesignation.Text
            !staffregistation = txtRegistation.Text
            If IsNumeric(DataComboSpeciality.BoundText) Then !StaffSpeciality_Id = DataComboSpeciality.BoundText
            !staffprivateaddress = txtPrivateAddress.Text
            !staffprivatephone = txtPrivateTel.Text
            !staffprivatefax = txtPrivateFax.Text
            !staffprivateemail = txtPrivateEmail.Text
            !staffmobilephone = txtPrivateMobile.Text
            !staffofficialAddress = txtOfficialAddress.Text
            !staffofficialphone = txtOfficialTel.Text
            !staffofficialFax = txtOfficialFax.Text
            !staffofficialEmail = txtOfficialEMail.Text
            !staffWebsite = txtOfficialWebsite.Text
            !staffComments = txtComments.Text
            If IsNumeric(DataComboPaymenyMethod.BoundText) Then !staffPaymentMethod_id = DataComboPaymenyMethod.BoundText
            If IsNumeric(DataComboBank.BoundText) Then !staffBank_id = DataComboBank.BoundText
            !staffBankBranch = txtBankBranch.Text
            !staffAccount = txtAccount.Text
            If chkCurrentlyChanneling.Value = 1 Then
                !staffCurrentlyChanneling = 0
            Else
                !staffCurrentlyChanneling = 1
            End If
            If chkUser.Value = 1 Then
                !StaffUser = 0
            Else
                !StaffUser = 1
            End If
            !StaffUserName = EncreptedWord(txtUserName.Text)
            !staffpassword = EncreptedWord(txtPassword.Text)
            If OptionAdministrator.Value = True Then !StaffAuthority = 1
            If OptionOwner.Value = True Then !StaffAuthority = 2
            If OptionOwnerCovered.Value = True Then !StaffAuthority = 3
            If OptionHR.Value = True Then !StaffAuthority = 4
            If OptionAccounting.Value = True Then !StaffAuthority = 5
            If OptionUser.Value = True Then !StaffAuthority = 6
            If OptionAnalyzer.Value = True Then !StaffAuthority = 7
            .Update
            .Close
            SaveData = True
        Exit Function
ErrorHandler:
        TemResponce = MsgBox("An unknown error has occured, Please contact Lakmedipro (077 3177874) with the following details." & vbNewLine & Me.Name & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly, "Update Error")
        .CancelUpdate
        .Close
        Exit Function
    End With
End Function

Private Function EditData() As Boolean
    EditData = False
    Dim TemResponce  As Integer
    If Trim(txtName.Text) = "" Then
        TemResponce = MsgBox("Please enter a name of a staff to save", vbCritical + vbOKOnly, "No name")
        SSTabMain.Tab = 0
        txtName.SetFocus
        Exit Function
    End If
    
    If Trim(txtListedName.Text) = "" Then
        TemResponce = MsgBox("Please enter name to be listed before saving", vbCritical + vbOKOnly, "No name")
        txtName_LostFocus
        SSTabMain.Tab = 0
        txtListedName.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(DataComboTitle.BoundText) Then
        TemResponce = MsgBox("Please select the title", vbCritical, "Title?")
        SSTabMain.Tab = 0
        DataComboTitle.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(DataComboSex.BoundText) Then
        TemResponce = MsgBox("Please select the sex", vbCritical, "Sex?")
        SSTabMain.Tab = 0
        DataComboSex.SetFocus
        Exit Function
    End If
    
    If chkUser.Value = 1 Then
        If Trim(txtUserName.Text) = "" Then
            TemResponce = MsgBox("You have to enter a user name if you are going to use the program.", vbCritical, "No User Name")
            SSTabMain.Tab = 2
            txtUserName.SetFocus
            Exit Function
        End If
        If Trim(txtPassword.Text) = "" Then
            TemResponce = MsgBox("You have not entered the password", vbInformation, "No Password")
            SSTabMain.Tab = 2
            txtPassword.SetFocus
            Exit Function
        End If
        If Trim(txtPassword.Text) <> Trim(txtReenterPassword.Text) Then
            TemResponce = MsgBox("The passwords you entered are not identical", vbCritical, "Passwords?")
            SSTabMain.Tab = 2
            txtPassword.SetFocus
            Exit Function
        End If
        If TemUserName <> txtUserName.Text Then
            With DataEnvironment1.rssqlTem
                If .State = 1 Then .Close
                .Source = "SELECT tblStaff.* from tblStaff"
                If .State = 0 Then .Open
                If .RecordCount <> 0 Then
                    .MoveFirst
                    While .EOF = False
                        If IsNull(!StaffUserName) = False Then
                            txtTemUserName.Text = DecreptedWord(!StaffUserName)
                            If UCase(txtTemUserName.Text) = UCase(txtUserName.Text) Then
                                TemResponce = MsgBox("This username is already taken, Select another user name", vbCritical, "Username")
                                SSTabMain.Tab = 2
                                txtUserName.SetFocus
                                SendKeys "{home}+{end}"
                                Exit Function
                            End If
                        End If
                    .MoveNext
                    Wend
                End If
                .Close
            End With
        End If
    End If
    
    'On Error GoTo ErrorHandler

        With DataEnvironment1.rssqlStaff
            If .State = 1 Then .Close
            .Source = "SELECT tblstaff.* from tblstaff where staff_id = " & TemstaffID
            If .State = 0 Then .Open
            If .RecordCount = 0 Then Exit Function
            
            !stafftitle_ID = DataComboTitle.BoundText
            !staffsex_ID = DataComboSex.BoundText
            !StaffName = txtName.Text
            !StaffCode = Val(txtCode.Text)
            !staffListedName = txtListedName.Text
            !staffqualifications = txtQualifications.Text
            !staffdesignation = txtDesignation.Text
            !staffregistation = txtRegistation.Text
            If IsNumeric(DataComboSpeciality.BoundText) Then !StaffSpeciality_Id = DataComboSpeciality.BoundText
            !staffprivateaddress = txtPrivateAddress.Text
            !staffprivatephone = txtPrivateTel.Text
            !staffprivatefax = txtPrivateFax.Text
            !staffprivateemail = txtPrivateEmail.Text
            !staffmobilephone = txtPrivateMobile.Text
            !staffofficialAddress = txtOfficialAddress.Text
            !staffofficialphone = txtOfficialTel.Text
            !staffofficialFax = txtOfficialFax.Text
            !staffofficialEmail = txtOfficialEMail.Text
            !staffWebsite = txtOfficialWebsite.Text
            !staffComments = txtComments.Text
            If IsNumeric(DataComboPaymenyMethod.BoundText) Then !staffPaymentMethod_id = DataComboPaymenyMethod.BoundText
            If IsNumeric(DataComboBank.BoundText) Then !staffBank_id = DataComboBank.BoundText
            !staffBankBranch = txtBankBranch.Text
            !staffAccount = txtAccount.Text
'            !staffCredit = Val(txtCredit.Text)
            If chkCurrentlyChanneling.Value = 1 Then
                !staffCurrentlyChanneling = 0
            Else
                !staffCurrentlyChanneling = 1
            End If
            !StaffUserName = EncreptedWord(txtUserName.Text)
            !staffpassword = EncreptedWord(txtPassword.Text)
            If chkUser.Value = 1 Then
                !StaffUser = 1
            Else
                !StaffUser = 0
            End If
            If OptionAdministrator.Value = True Then !StaffAuthority = 1
            If OptionOwner.Value = True Then !StaffAuthority = 2
            If OptionOwnerCovered.Value = True Then !StaffAuthority = 3
            If OptionHR.Value = True Then !StaffAuthority = 4
            If OptionAccounting.Value = True Then !StaffAuthority = 5
            If OptionUser.Value = True Then !StaffAuthority = 6
            If OptionAnalyzer.Value = True Then !StaffAuthority = 7
            .Update
            .Close
        EditData = True
        
        Exit Function
ErrorHandler:
        TemResponce = MsgBox("An unknown error has occured, Please contact Lakmedipro (077 3177874) with the following details." & vbNewLine & Me.Name & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly, "Update Error")
        .CancelUpdate
        .Close
        Exit Function
    End With
End Function

Private Sub ClearValues()
    txtName.Text = Empty
    txtCode.Text = Empty
    txtListedName.Text = Empty
    DataComboSex.Text = Empty
    DataComboTitle.Text = Empty
    
    txtQualifications.Text = Empty
    txtDesignation.Text = Empty
    txtRegistation.Text = Empty
    DataComboSpeciality.Text = Empty
    txtOfficialAddress.Text = Empty
    txtOfficialTel.Text = Empty
    txtOfficialFax.Text = Empty
    txtOfficialEMail.Text = Empty
    txtPrivateAddress.Text = Empty
    txtPrivateEmail.Text = Empty
    txtPrivateFax.Text = Empty
    txtPrivateMobile.Text = Empty
    txtPrivateTel.Text = Empty
    txtOfficialWebsite.Text = Empty
    
    DataComboBank.Text = Empty
    DataComboPaymenyMethod.Text = Empty
    chkCurrentlyChanneling.Value = 0
    txtBankBranch.Text = Empty
    txtAccount.Text = Empty
    txtCredit.Text = Empty
    txtComments.Text = Empty
    
    txtSearch.Text = Empty
    
    txtPassword.Text = Empty
    txtUserName.Text = Empty
    txtReenterPassword.Text = Empty
    
    chkUser.Value = 0
    
    Call FormatGrid
    
End Sub


Private Sub GetData()
    grid1.Col = 2
    If IsNumeric(grid1.Text) = False Then Exit Sub
    TemstaffID = Val(grid1.Text)
    With DataEnvironment1.rssqlStaff
        If .State = 1 Then .Close
        .Source = "SELECT tblstaff.* from tblstaff where staff_ID = " & TemstaffID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!stafftitle_ID) Then DataComboTitle.BoundText = !stafftitle_ID
        If Not IsNull(!staffsex_ID) Then DataComboSex.BoundText = !staffsex_ID
        If Not IsNull(!StaffName) Then txtName.Text = !StaffName
        If Not IsNull(!StaffCode) Then txtCode.Text = !StaffCode
        If Not IsNull(!staffListedName) Then txtListedName.Text = !staffListedName
        If Not IsNull(!staffqualifications) Then txtQualifications.Text = !staffqualifications
        If Not IsNull(!staffdesignation) Then txtDesignation.Text = !staffdesignation
        If Not IsNull(!staffregistation) Then txtRegistation.Text = !staffregistation
        If Not IsNull(!StaffSpeciality_Id) Then DataComboSpeciality.BoundText = !StaffSpeciality_Id
        If Not IsNull(!staffprivateaddress) Then txtPrivateAddress.Text = !staffprivateaddress
        If Not IsNull(!staffprivatephone) Then txtPrivateTel.Text = !staffprivatephone
        If Not IsNull(!staffprivatefax) Then txtPrivateFax.Text = !staffprivatefax
        If Not IsNull(!staffprivateemail) Then txtPrivateEmail.Text = !staffprivateemail
        If Not IsNull(!staffmobilephone) Then txtPrivateMobile.Text = !staffmobilephone
        If Not IsNull(!staffofficialAddress) Then txtOfficialAddress.Text = !staffofficialAddress
        If Not IsNull(!staffofficialphone) Then txtOfficialTel.Text = !staffofficialphone
        If Not IsNull(!staffofficialFax) Then txtOfficialFax.Text = !staffofficialFax
        If Not IsNull(!staffofficialEmail) Then txtOfficialEMail.Text = !staffofficialEmail
        If Not IsNull(!staffWebsite) Then txtOfficialWebsite.Text = !staffWebsite
        If Not IsNull(!staffComments) Then txtComments.Text = !staffComments
        If Not IsNull(!staffPaymentMethod_id) Then DataComboPaymenyMethod.BoundText = !staffPaymentMethod_id
        If Not IsNull(!staffBank_id) Then DataComboBank.BoundText = !staffBank_id
        If Not IsNull(!staffBankBranch) Then txtBankBranch.Text = !staffBankBranch
        If Not IsNull(!staffAccount) Then txtAccount.Text = !staffAccount
        If Not IsNull(!staffCredit) Then txtCredit.Text = !staffCredit
        If Not IsNull(!staffpassword) Then txtPassword.Text = DecreptedWord(!staffpassword)
        If Not IsNull(!StaffUserName) Then txtUserName.Text = DecreptedWord(!StaffUserName)
        If Not IsNull(!staffpassword) Then txtReenterPassword.Text = DecreptedWord(!staffpassword)
        If Not IsNull(!staffCurrentlyChanneling) Then
            If !staffCurrentlyChanneling = True Then
                chkCurrentlyChanneling.Value = 1
            Else
                chkCurrentlyChanneling.Value = 0
            End If
        End If
        If !StaffUser = True Then chkUser.Value = 1
        If Not IsNull(!StaffAuthority) Then
            Select Case !StaffAuthority
                Case AuthorityAdministrator: OptionAdministrator.Value = True
                Case AuthorityOwner: OptionOwner.Value = True
                Case AuthorityOwnerCOvered: OptionOwnerCovered.Value = True
                Case AuthorityHumanResources: OptionHR.Value = True
                Case AuthorityAccount: OptionAccounting.Value = True
                Case AuthorityUser: OptionUser.Value = True
            End Select
        End If
        TemUserName = txtUserName.Text
    End With
End Sub

Private Sub Grid1_Click()
    FromGrid = True
    With grid1
        If .Row < 1 Then FromGrid = False: Exit Sub
        .Col = 2
        If Not IsNumeric(.Text) Then FromGrid = False: Exit Sub
        TemstaffID = Val(.Text)
        .Col = 1
        txtSearch.Text = .Text
        Call GetData
        
        .Col = 0
        .ColSel = .Cols - 1
        
        txtSearch.SetFocus
        SendKeys "{home}+{end}"
    FromGrid = False
    bttnAdd.Enabled = False
    bttnEdit.Enabled = True
End With
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then Grid1_Click
End Sub



Private Sub txtName_LostFocus()
    Dim TemFirstName As String
    Dim TemSurname As String
    Dim TemBreakingPoint  As Integer
    If Trim(txtName.Text) = "" Then Exit Sub
    txtName.Text = Trim(txtName.Text)
    TemBreakingPoint = InStr(1, txtName.Text, " ")
    If TemBreakingPoint > 1 Then
        TemFirstName = Left(txtName.Text, TemBreakingPoint - 1)
        TemSurname = Right(txtName.Text, Len(txtName.Text) - TemBreakingPoint)
        txtListedName.Text = TemSurname & ", " & TemFirstName
    Else
        txtListedName.Text = txtName.Text
    End If
End Sub

Private Sub txtSearch_Change()
    
' **************************************

    If FromGrid = True Then Exit Sub
    Dim TemFRows As Long
    Dim TemNowRow As Long
    Dim TemArray As Long
    Dim SearchSuccess As Boolean
    Dim TemLength As Single
    TemFRows = grid1.Rows
    grid1.Col = 1
    SearchSuccess = False
    If Len(txtSearch.Text) = 0 Then GoTo MeasureSuccess
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
        bttnAdd.Enabled = False
        grid1.Col = 2
        TemstaffID = grid1.Text
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
    End If
'**************************************
End Sub


