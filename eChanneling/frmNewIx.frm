VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNewIx 
   Caption         =   "New Investigation"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   15240
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   78
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   12938
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Select Patient"
      TabPicture(0)   =   "frmNewIx.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frameSearchPatient"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Patient Details"
      TabPicture(1)   =   "frmNewIx.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FramePatient"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ix Details"
      TabPicture(2)   =   "frmNewIx.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label7"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label8"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label11"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label12"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label13"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label14"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label15"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label16"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label26"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label27"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "DataComboPerformed"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "DTPickerPerformedTime"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "DTPickerPerformedDate"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "DTPickerCollectedTime"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "DTPickerCollectedDate"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "DataComboCollected"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "DataComboSpeciman"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "DataComboDepartment"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "DataComboInstitute"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "DataComboDoctor"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "DataComboIx"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "txtSpecimanNo"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "txtRack"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "txtTube"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "txtNotes"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).ControlCount=   27
      Begin VB.TextBox txtNotes 
         Height          =   600
         Left            =   840
         TabIndex        =   167
         Top             =   6600
         Width           =   3975
      End
      Begin VB.TextBox txtTube 
         Height          =   360
         Left            =   2880
         TabIndex        =   155
         Top             =   6120
         Width           =   1935
      End
      Begin VB.TextBox txtRack 
         Height          =   360
         Left            =   840
         TabIndex        =   154
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Frame FramePatient 
         Caption         =   "Patient Details"
         Height          =   6855
         Left            =   -74880
         TabIndex        =   119
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txtID 
            Height          =   345
            Left            =   1320
            TabIndex        =   133
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox txtFirstName 
            Height          =   345
            Left            =   1320
            TabIndex        =   123
            Top             =   1320
            Width           =   3375
         End
         Begin VB.TextBox txtOtherName 
            Height          =   345
            Left            =   1320
            TabIndex        =   122
            Top             =   1800
            Width           =   3375
         End
         Begin VB.TextBox txtSurname 
            Height          =   345
            Left            =   1320
            TabIndex        =   121
            Top             =   2280
            Width           =   3375
         End
         Begin VB.TextBox txtAge 
            Height          =   375
            Left            =   1320
            TabIndex        =   120
            Top             =   5400
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo DataComboTitle 
            Bindings        =   "frmNewIx.frx":0054
            Height          =   360
            Left            =   1320
            TabIndex        =   124
            Top             =   840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Title"
            BoundColumn     =   "Title_ID"
            Text            =   ""
            Object.DataMember      =   "sqlTitle"
         End
         Begin MSDataListLib.DataCombo DataComboSex 
            Bindings        =   "frmNewIx.frx":0073
            Height          =   315
            Left            =   1320
            TabIndex        =   125
            Top             =   3480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Sex"
            BoundColumn     =   "Sex_ID"
            Text            =   ""
            Object.DataMember      =   "sqlSex"
         End
         Begin MSDataListLib.DataCombo DataComboMarietal 
            Bindings        =   "frmNewIx.frx":0092
            Height          =   315
            Left            =   1320
            TabIndex        =   126
            Top             =   2880
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Marietal"
            BoundColumn     =   "Marietal_ID"
            Text            =   ""
            Object.DataMember      =   "sqlMarietal"
         End
         Begin MSDataListLib.DataCombo DataComboRace 
            Bindings        =   "frmNewIx.frx":00B1
            Height          =   315
            Left            =   1320
            TabIndex        =   127
            Top             =   4080
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Race"
            BoundColumn     =   "Race_ID"
            Text            =   ""
            Object.DataMember      =   "sqlRace"
         End
         Begin MSComCtl2.DTPicker DTPickerDOB 
            Height          =   375
            Left            =   1320
            TabIndex        =   128
            Top             =   4680
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39413
         End
         Begin btButtonEx.ButtonEx bttnPatientSave 
            Height          =   375
            Left            =   2400
            TabIndex        =   195
            Top             =   6360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            Appearance      =   3
            Caption         =   "Save Details"
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
         Begin VB.Label Label6 
            Caption         =   "Patient &ID"
            Height          =   255
            Left            =   240
            TabIndex        =   143
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label18 
            Caption         =   "Age"
            Height          =   255
            Left            =   240
            TabIndex        =   142
            Top             =   5400
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "D&ate of Birth"
            Height          =   255
            Left            =   240
            TabIndex        =   141
            Top             =   4800
            Width           =   3015
         End
         Begin VB.Label Label19 
            Caption         =   "&Race"
            Height          =   255
            Left            =   240
            TabIndex        =   140
            Top             =   4080
            Width           =   2655
         End
         Begin VB.Label Label20 
            Caption         =   "&Marietal"
            Height          =   255
            Left            =   240
            TabIndex        =   139
            Top             =   2880
            Width           =   2415
         End
         Begin VB.Label Label21 
            Caption         =   "S&ex"
            Height          =   255
            Left            =   240
            TabIndex        =   138
            Top             =   3480
            Width           =   2655
         End
         Begin VB.Label Label22 
            Caption         =   "&Title"
            Height          =   255
            Left            =   240
            TabIndex        =   137
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label Label23 
            Caption         =   "&Surname"
            Height          =   255
            Left            =   240
            TabIndex        =   136
            Top             =   2280
            Width           =   3135
         End
         Begin VB.Label Label24 
            Caption         =   "&Other Names"
            Height          =   255
            Left            =   240
            TabIndex        =   135
            Top             =   1800
            Width           =   3135
         End
         Begin VB.Label Label25 
            Caption         =   "&First Name"
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   1320
            Width           =   3015
         End
      End
      Begin VB.TextBox txtSpecimanNo 
         Height          =   360
         Left            =   600
         TabIndex        =   115
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Frame frameSearchPatient 
         Caption         =   "Search Patient"
         Height          =   6855
         Left            =   -74880
         TabIndex        =   79
         Top             =   360
         Width           =   4815
         Begin VB.TextBox txtSearchSurname 
            Height          =   345
            Left            =   840
            TabIndex        =   82
            Top             =   1200
            Width           =   3855
         End
         Begin VB.TextBox txtSearchFirstName 
            Height          =   345
            Left            =   840
            TabIndex        =   81
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox txtSearchID 
            Height          =   345
            Left            =   840
            TabIndex        =   80
            Top             =   240
            Width           =   3855
         End
         Begin MSFlexGridLib.MSFlexGrid Grid1 
            Height          =   4095
            Left            =   120
            TabIndex        =   83
            Top             =   2160
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   7223
            _Version        =   393216
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
         End
         Begin btButtonEx.ButtonEx bttnSearch 
            Height          =   375
            Left            =   120
            TabIndex        =   84
            Top             =   1680
            Width           =   4575
            _ExtentX        =   8070
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
            Left            =   120
            TabIndex        =   85
            Top             =   6360
            Width           =   2175
            _ExtentX        =   3836
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
         Begin btButtonEx.ButtonEx bttnPatientAdd 
            Height          =   375
            Left            =   2400
            TabIndex        =   194
            Top             =   6360
            Width           =   2175
            _ExtentX        =   3836
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
         Begin VB.Label Label5 
            Caption         =   "Surname"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Firstname"
            Height          =   375
            Left            =   120
            TabIndex        =   87
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "ID"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSDataListLib.DataCombo DataComboIx 
         Bindings        =   "frmNewIx.frx":00D0
         Height          =   315
         Left            =   600
         TabIndex        =   116
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "IX"
         BoundColumn     =   "IxID"
         Text            =   ""
         Object.DataMember      =   "sqlInvestigations"
      End
      Begin MSDataListLib.DataCombo DataComboDoctor 
         Bindings        =   "frmNewIx.frx":00EF
         Height          =   315
         Left            =   600
         TabIndex        =   129
         Top             =   2760
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "DoctorListedName"
         BoundColumn     =   "Doctor_ID"
         Text            =   ""
         Object.DataMember      =   "sqlDoctor"
      End
      Begin MSDataListLib.DataCombo DataComboInstitute 
         Bindings        =   "frmNewIx.frx":010E
         Height          =   315
         Left            =   600
         TabIndex        =   130
         Top             =   3480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "InstitutionName"
         BoundColumn     =   "Institution_ID"
         Text            =   ""
         Object.DataMember      =   "sqlInstitutions"
      End
      Begin MSDataListLib.DataCombo DataComboDepartment 
         Bindings        =   "frmNewIx.frx":012D
         Height          =   315
         Left            =   600
         TabIndex        =   148
         Top             =   3840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Department"
         BoundColumn     =   "Department_ID"
         Text            =   ""
         Object.DataMember      =   "sqlDepartment"
      End
      Begin MSDataListLib.DataCombo DataComboSpeciman 
         Bindings        =   "frmNewIx.frx":014C
         Height          =   315
         Left            =   600
         TabIndex        =   151
         Top             =   1440
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "Speciman"
         BoundColumn     =   "Speciman_ID"
         Text            =   ""
         Object.DataMember      =   "sqlSpeciman"
      End
      Begin MSDataListLib.DataCombo DataComboCollected 
         Bindings        =   "frmNewIx.frx":016B
         Height          =   315
         Left            =   1560
         TabIndex        =   152
         Top             =   4320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "StaffListedName"
         BoundColumn     =   "Staff_ID"
         Text            =   ""
         Object.DataMember      =   "sqlStaff"
      End
      Begin MSComCtl2.DTPicker DTPickerCollectedDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   153
         Top             =   4680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   20709379
         CurrentDate     =   39436
      End
      Begin MSComCtl2.DTPicker DTPickerCollectedTime 
         Height          =   375
         Left            =   3480
         TabIndex        =   157
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   20709378
         CurrentDate     =   39436
      End
      Begin MSComCtl2.DTPicker DTPickerPerformedDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   158
         Top             =   5520
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   20709379
         CurrentDate     =   39436
      End
      Begin MSComCtl2.DTPicker DTPickerPerformedTime 
         Height          =   375
         Left            =   3480
         TabIndex        =   159
         Top             =   5520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   20709378
         CurrentDate     =   39436
      End
      Begin MSDataListLib.DataCombo DataComboPerformed 
         Bindings        =   "frmNewIx.frx":018A
         Height          =   315
         Left            =   1560
         TabIndex        =   160
         Top             =   5160
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "StaffListedName"
         BoundColumn     =   "Staff_ID"
         Text            =   ""
         Object.DataMember      =   "sqlStaff"
      End
      Begin VB.Label Label27 
         Caption         =   "Notes"
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
         Left            =   120
         TabIndex        =   168
         Top             =   6600
         Width           =   2055
      End
      Begin VB.Label Label26 
         Caption         =   "Date / Time"
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
         Left            =   120
         TabIndex        =   161
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label Label16 
         Caption         =   "Performed By"
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
         Left            =   120
         TabIndex        =   156
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "Tube"
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
         Left            =   2400
         TabIndex        =   150
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "Rack"
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
         Left            =   120
         TabIndex        =   149
         Top             =   6120
         Width           =   2055
      End
      Begin VB.Label Label13 
         Caption         =   "Speciman"
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
         Left            =   120
         TabIndex        =   147
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "Date / Time"
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
         Left            =   120
         TabIndex        =   146
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Collected By"
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
         Left            =   120
         TabIndex        =   145
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Referring Institute"
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
         Left            =   120
         TabIndex        =   132
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Referring Doctor"
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
         Left            =   120
         TabIndex        =   131
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Investigation Name"
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
         Left            =   120
         TabIndex        =   118
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Speciman No."
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
         Left            =   120
         TabIndex        =   117
         Top             =   1800
         Width           =   2055
      End
   End
   Begin VB.Frame framePatientIX 
      Height          =   7335
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtValue25 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   89
         Top             =   6000
         Width           =   1695
      End
      Begin VB.TextBox txtValue24 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   90
         Top             =   5760
         Width           =   1695
      End
      Begin VB.TextBox txtValue23 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   91
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox txtValue22 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   92
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox txtValue21 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   93
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox txtValue20 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   94
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox txtValue19 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   95
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txtValue18 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   96
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox txtValue17 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   97
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtValue16 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   98
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox txtValue15 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   99
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox txtValue14 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   100
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txtValue13 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   101
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txtValue12 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   102
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtValue11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   103
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtValue10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   104
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtValue9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   105
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtValue8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   106
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtValue7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   107
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtValue6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   108
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtValue5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   109
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtValue4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   110
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtValue3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   111
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtValue2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   112
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtValue1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   113
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks25 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   193
         Top             =   6000
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks24 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   192
         Top             =   5760
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks23 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   191
         Top             =   5520
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks22 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   190
         Top             =   5280
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks21 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   189
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks20 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   188
         Top             =   4800
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks19 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   187
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks18 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   186
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks17 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   185
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks16 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   184
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks15 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   183
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks14 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   182
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks13 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   181
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks12 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   180
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   179
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   178
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   177
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   176
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   175
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   174
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   173
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   172
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   171
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   170
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtRemarks1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8040
         TabIndex        =   169
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtComments 
         Height          =   855
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   6360
         Width           =   7335
      End
      Begin VB.Label lblFeild2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblFeild3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label lblFeild4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblFeild5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblFeild6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label lblFeild7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label lblFeild8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   71
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label lblFeild9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label lblFeild10 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label lblFeild11 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label lblFeild12 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label lblFeild13 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label lblFeild14 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label lblFeild15 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label lblFeild1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblUnit1 
         Height          =   255
         Left            =   4800
         TabIndex        =   62
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblUnit2 
         Height          =   255
         Left            =   4800
         TabIndex        =   61
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblUnit3 
         Height          =   255
         Left            =   4800
         TabIndex        =   60
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblUnit4 
         Height          =   255
         Left            =   4800
         TabIndex        =   59
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblUnit5 
         Height          =   255
         Left            =   4800
         TabIndex        =   58
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblUnit6 
         Height          =   255
         Left            =   4800
         TabIndex        =   57
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblUnit7 
         Height          =   255
         Left            =   4800
         TabIndex        =   56
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblUnit8 
         Height          =   255
         Left            =   4800
         TabIndex        =   55
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblUnit9 
         Height          =   255
         Left            =   4800
         TabIndex        =   54
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblUnit10 
         Height          =   255
         Left            =   4800
         TabIndex        =   53
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblUnit11 
         Height          =   255
         Left            =   4800
         TabIndex        =   52
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblUnit12 
         Height          =   255
         Left            =   4800
         TabIndex        =   51
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblUnit13 
         Height          =   255
         Left            =   4800
         TabIndex        =   50
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lblUnit14 
         Height          =   255
         Left            =   4800
         TabIndex        =   49
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblUnit15 
         Height          =   255
         Left            =   4800
         TabIndex        =   48
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblRef1 
         Height          =   255
         Left            =   6240
         TabIndex        =   47
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblRef2 
         Height          =   255
         Left            =   6240
         TabIndex        =   46
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblRef3 
         Height          =   255
         Left            =   6240
         TabIndex        =   45
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblRef4 
         Height          =   255
         Left            =   6240
         TabIndex        =   44
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblRef5 
         Height          =   255
         Left            =   6240
         TabIndex        =   43
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblRef6 
         Height          =   255
         Left            =   6240
         TabIndex        =   42
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblRef7 
         Height          =   255
         Left            =   6240
         TabIndex        =   41
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblRef8 
         Height          =   255
         Left            =   6240
         TabIndex        =   40
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblRef9 
         Height          =   255
         Left            =   6240
         TabIndex        =   39
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblRef10 
         Height          =   255
         Left            =   6240
         TabIndex        =   38
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lblRef11 
         Height          =   255
         Left            =   6240
         TabIndex        =   37
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblRef12 
         Height          =   255
         Left            =   6240
         TabIndex        =   36
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblRef13 
         Height          =   255
         Left            =   6240
         TabIndex        =   35
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblRef14 
         Height          =   255
         Left            =   6240
         TabIndex        =   34
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lblRef15 
         Height          =   255
         Left            =   6240
         TabIndex        =   33
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Comments"
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
         TabIndex        =   32
         Top             =   6360
         Width           =   2895
      End
      Begin VB.Label lblRef24 
         Height          =   255
         Left            =   6240
         TabIndex        =   31
         Top             =   5760
         Width           =   1815
      End
      Begin VB.Label lblRef23 
         Height          =   255
         Left            =   6240
         TabIndex        =   30
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label lblRef22 
         Height          =   255
         Left            =   6240
         TabIndex        =   29
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label lblRef21 
         Height          =   255
         Left            =   6240
         TabIndex        =   28
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label lblRef20 
         Height          =   255
         Left            =   6240
         TabIndex        =   27
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label lblRef19 
         Height          =   255
         Left            =   6240
         TabIndex        =   26
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label lblRef18 
         Height          =   255
         Left            =   6240
         TabIndex        =   25
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lblRef17 
         Height          =   255
         Left            =   6240
         TabIndex        =   24
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label lblRef16 
         Height          =   255
         Left            =   6240
         TabIndex        =   23
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label lblUnit24 
         Height          =   255
         Left            =   4800
         TabIndex        =   22
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Label lblUnit23 
         Height          =   255
         Left            =   4800
         TabIndex        =   21
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label lblUnit22 
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label lblUnit21 
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label lblUnit20 
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label lblUnit19 
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label lblUnit18 
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label lblUnit17 
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label lblUnit16 
         Height          =   255
         Left            =   4800
         TabIndex        =   14
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label lblFeild24 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Label lblFeild23 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Label lblFeild22 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   5280
         Width           =   2655
      End
      Begin VB.Label lblFeild21 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Label lblFeild20 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Label lblFeild19 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Label lblFeild18 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label lblFeild17 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label lblFeild16 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label lblRef25 
         Height          =   255
         Left            =   6240
         TabIndex        =   4
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label lblUnit25 
         Height          =   255
         Left            =   4800
         TabIndex        =   3
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label lblFeild25 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   6000
         Width           =   2655
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   13440
      TabIndex        =   114
      Top             =   7560
      Width           =   1695
      _ExtentX        =   2990
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
   Begin btButtonEx.ButtonEx bttnNewPatient 
      Height          =   375
      Left            =   11640
      TabIndex        =   162
      Top             =   7560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "New Patient"
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
   Begin btButtonEx.ButtonEx bttnNewInvestigation 
      Height          =   375
      Left            =   9840
      TabIndex        =   163
      Top             =   7560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "New Investigation"
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
      Left            =   8040
      TabIndex        =   164
      Top             =   7560
      Width           =   1695
      _ExtentX        =   2990
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
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   375
      Left            =   6240
      TabIndex        =   165
      Top             =   7560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Print"
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
      Left            =   8040
      TabIndex        =   166
      Top             =   7920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save Changes"
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
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   495
      Left            =   7080
      TabIndex        =   144
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmNewIx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemPatientID As Long
    Dim IxSaved As Boolean
    Dim IxPrinted As Boolean
    Dim TemPatientIxID As Long


Private Sub bttnSave_Click()
Dim TemResponce As Byte
'On Error GoTo ErrorHandler


If Not IsNumeric(DataComboIx.BoundText) Then
    TemResponce = MsgBox("You have not selected an Investigation to save", vbCritical, "No Investigation")
    DataComboIx.SetFocus
    Exit Sub
End If

With DataEnvironment1.rssqlTem6
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientinvestigations"
    If .State = 0 Then .Open
    .AddNew
    !ixid = DataComboIx.BoundText
    !patient_ID = TemPatientID
    If IsNumeric(DataComboDoctor.BoundText) Then !doctor_id = DataComboDoctor.BoundText
    If IsNumeric(DataComboInstitute.BoundText) Then !institution_ID = DataComboInstitute.BoundText
    If IsNumeric(DataComboDepartment.BoundText) Then !department_ID = DataComboDepartment.BoundText
    If IsNumeric(DataComboCollected.BoundText) Then !collectedby = DataComboCollected.BoundText
    !collecteddate = DTPickerCollectedDate.Value
    !collectedtime = DTPickerCollectedTime.Value
    If IsNumeric(DataComboPerformed.BoundText) Then !Performedby = DataComboPerformed.BoundText
    !performeddate = DTPickerPerformedDate.Value
    !performedtime = DTPickerPerformedTime.Value
    If IsNumeric(DataComboSpeciman.BoundText) Then !speciman_ID = DataComboSpeciman.BoundText
    !specimanno = txtSpecimanNo.Text
    !rack = txtRack.Text
    !tube = txtTube.Text
    !notes = txtNotes.Text
    !fieldvalue1 = txtValue1.Text
    !fieldvalue2 = txtValue2.Text
    !fieldvalue3 = txtValue3.Text
    !fieldvalue4 = txtValue4.Text
    !fieldvalue5 = txtValue5.Text
    !fieldvalue6 = txtValue6.Text
    !fieldvalue7 = txtValue7.Text
    !fieldvalue8 = txtValue8.Text
    !fieldvalue9 = txtValue9.Text
    !fieldvalue10 = txtValue10.Text
    !fieldvalue11 = txtValue11.Text
    !fieldvalue12 = txtValue12.Text
    !fieldvalue13 = txtValue13.Text
    !fieldvalue14 = txtValue14.Text
    !fieldvalue15 = txtValue15.Text
    !fieldvalue16 = txtValue16.Text
    !fieldvalue17 = txtValue17.Text
    !fieldvalue18 = txtValue18.Text
    !fieldvalue19 = txtValue19.Text
    !fieldvalue20 = txtValue20.Text
    !fieldvalue21 = txtValue21.Text
    !fieldvalue22 = txtValue22.Text
    !fieldvalue23 = txtValue23.Text
    !fieldvalue24 = txtValue24.Text
    !fieldvalue25 = txtValue25.Text
    !remark1 = txtRemarks1.Text
    !remark2 = txtRemarks2.Text
    !remark3 = txtRemarks3.Text
    !remark4 = txtRemarks4.Text
    !remark5 = txtRemarks5.Text
    !remark6 = txtRemarks6.Text
    !remark7 = txtRemarks7.Text
    !remark8 = txtRemarks8.Text
    !remark9 = txtRemarks9.Text
    !remark10 = txtRemarks10.Text
    !remark11 = txtRemarks11.Text
    !remark12 = txtRemarks12.Text
    !remark13 = txtRemarks13.Text
    !remark14 = txtRemarks14.Text
    !remark15 = txtRemarks15.Text
    !remark16 = txtRemarks16.Text
    !remark17 = txtRemarks17.Text
    !remark18 = txtRemarks18.Text
    !remark19 = txtRemarks19.Text
    !remark20 = txtRemarks20.Text
    !remark21 = txtRemarks21.Text
    !remark22 = txtRemarks22.Text
    !remark23 = txtRemarks23.Text
    !remark24 = txtRemarks24.Text
    !remark25 = txtRemarks25.Text
    .Update
    TemPatientIxID = !PatientIxId
    .Close

    IxSaved = True
    
    bttnSave.Visible = False
    bttnChange.Visible = True
    
Exit Sub

ErrorHandler:
    TemResponce = MsgBox("An unknown error has occured. Please contact Lakmedipro(077 3177874) with following details." & vbNewLine & Me.Caption & vbNewLine & Err.Number & vbNewLine & Err.Description, vbCritical, "Error")
    .CancelUpdate
    Exit Sub






End With
End Sub

Private Sub DataComboInstitute_Click(Area As Integer)

On Error Resume Next

If Not IsNumeric(DataComboInstitute.BoundText) Then
    DataComboDepartment.RowMember = Empty
    DataComboDepartment.ListField = Empty
    DataComboDepartment.BoundColumn = Empty
Else
    With DataEnvironment1.rssqlDepartment
    If .State = 1 Then .Close
    .Source = "SELECT * from tbldepartments where institution_ID = " & DataComboInstitute.BoundText
    If .State = 0 Then .Open
    DataComboDepartment.RowMember = "sqlDepartment"
    DataComboDepartment.ListField = "Department"
    DataComboDepartment.BoundColumn = "Department_ID"
    End With
End If

End Sub

Private Sub DataComboIx_Change()
If Not IsNumeric(DataComboIx.BoundText) Then Exit Sub
Call ClearIxFields
Call PrepareForInvestigation


End Sub

Private Sub ClearIxFields()
    lblFeild1.Caption = Empty
    lblFeild2.Caption = Empty
    lblFeild3.Caption = Empty
    lblFeild4.Caption = Empty
    lblFeild5.Caption = Empty
    lblFeild6.Caption = Empty
    lblFeild7.Caption = Empty
    lblFeild8.Caption = Empty
    lblFeild9.Caption = Empty
    lblFeild10.Caption = Empty
    lblFeild11.Caption = Empty
    lblFeild12.Caption = Empty
    lblFeild13.Caption = Empty
    lblFeild14.Caption = Empty
    lblFeild15.Caption = Empty
    lblFeild16.Caption = Empty
    lblFeild17.Caption = Empty
    lblFeild18.Caption = Empty
    lblFeild19.Caption = Empty
    lblFeild20.Caption = Empty
    lblFeild21.Caption = Empty
    lblFeild22.Caption = Empty
    lblFeild23.Caption = Empty
    lblFeild24.Caption = Empty
    lblFeild25.Caption = Empty
    lblRef1.Caption = Empty
    lblRef2.Caption = Empty
    lblRef3.Caption = Empty
    lblRef4.Caption = Empty
    lblRef5.Caption = Empty
    lblRef6.Caption = Empty
    lblRef7.Caption = Empty
    lblRef8.Caption = Empty
    lblRef9.Caption = Empty
    lblRef10.Caption = Empty
    lblRef11.Caption = Empty
    lblRef12.Caption = Empty
    lblRef13.Caption = Empty
    lblRef14.Caption = Empty
    lblRef15.Caption = Empty
    lblRef16.Caption = Empty
    lblRef17.Caption = Empty
    lblRef18.Caption = Empty
    lblRef19.Caption = Empty
    lblRef20.Caption = Empty
    lblRef21.Caption = Empty
    lblRef22.Caption = Empty
    lblRef23.Caption = Empty
    lblRef24.Caption = Empty
    lblRef25.Caption = Empty
    lblUnit1.Caption = Empty
    lblUnit2.Caption = Empty
    lblUnit3.Caption = Empty
    lblUnit4.Caption = Empty
    lblUnit5.Caption = Empty
    lblUnit6.Caption = Empty
    lblUnit7.Caption = Empty
    lblUnit8.Caption = Empty
    lblUnit9.Caption = Empty
    lblUnit10.Caption = Empty
    lblUnit11.Caption = Empty
    lblUnit12.Caption = Empty
    lblUnit13.Caption = Empty
    lblUnit14.Caption = Empty
    lblUnit15.Caption = Empty
    lblUnit16.Caption = Empty
    lblUnit17.Caption = Empty
    lblUnit18.Caption = Empty
    lblUnit19.Caption = Empty
    lblUnit20.Caption = Empty
    lblUnit21.Caption = Empty
    lblUnit22.Caption = Empty
    lblUnit23.Caption = Empty
    lblUnit24.Caption = Empty
    lblUnit25.Caption = Empty
     txtComments.Text = Empty
    
    txtValue1.Text = Empty
    txtValue2.Text = Empty
    txtValue3.Text = Empty
    txtValue4.Text = Empty
    txtValue5.Text = Empty
    txtValue6.Text = Empty
    txtValue7.Text = Empty
    txtValue8.Text = Empty
    txtValue9.Text = Empty
    txtValue10.Text = Empty
    txtValue11.Text = Empty
    txtValue12.Text = Empty
    txtValue13.Text = Empty
    txtValue14.Text = Empty
    txtValue15.Text = Empty
    txtValue16.Text = Empty
    txtValue17.Text = Empty
    txtValue18.Text = Empty
    txtValue19.Text = Empty
    txtValue20.Text = Empty
    txtValue21.Text = Empty
    txtValue22.Text = Empty
    txtValue23.Text = Empty
    txtValue24.Text = Empty
    txtValue25.Text = Empty
    
    txtRemarks1.Text = Empty
    txtRemarks2.Text = Empty
    txtRemarks3.Text = Empty
    txtRemarks4.Text = Empty
    txtRemarks5.Text = Empty
    txtRemarks6.Text = Empty
    txtRemarks7.Text = Empty
    txtRemarks8.Text = Empty
    txtRemarks9.Text = Empty
    txtRemarks10.Text = Empty
    txtRemarks11.Text = Empty
    txtRemarks12.Text = Empty
    txtRemarks13.Text = Empty
    txtRemarks14.Text = Empty
    txtRemarks15.Text = Empty
    txtRemarks16.Text = Empty
    txtRemarks17.Text = Empty
    txtRemarks18.Text = Empty
    txtRemarks19.Text = Empty
    txtRemarks20.Text = Empty
    txtRemarks21.Text = Empty
    txtRemarks22.Text = Empty
    txtRemarks23.Text = Empty
    txtRemarks24.Text = Empty
    txtRemarks25.Text = Empty
    
    
    
End Sub

Private Sub PrepareForInvestigation()
With DataEnvironment1.rssqlTem7
    If .State = 1 Then .Close
    .Source = "SELECT * from tblinvestigationdetails where ixID = " & DataComboIx.BoundText
    .Open
    If .RecordCount = 0 Then Exit Sub
        If Not IsNull(!field1) Then lblFeild1.Caption = !field1
        If Not IsNull(!field2) Then lblFeild2.Caption = !field2
        If Not IsNull(!field3) Then lblFeild3.Caption = !field3
        If Not IsNull(!field4) Then lblFeild4.Caption = !field4
        If Not IsNull(!field5) Then lblFeild5.Caption = !field5
        If Not IsNull(!field6) Then lblFeild6.Caption = !field6
        If Not IsNull(!field7) Then lblFeild7.Caption = !field7
        If Not IsNull(!field8) Then lblFeild8.Caption = !field8
        If Not IsNull(!field9) Then lblFeild9.Caption = !field9
        If Not IsNull(!field10) Then lblFeild10.Caption = !field10
        If Not IsNull(!field11) Then lblFeild11.Caption = !field11
        If Not IsNull(!field12) Then lblFeild12.Caption = !field12
        If Not IsNull(!field13) Then lblFeild13.Caption = !field13
        If Not IsNull(!field14) Then lblFeild14.Caption = !field14
        If Not IsNull(!Field15) Then lblFeild15.Caption = !Field15
        If Not IsNull(!field16) Then lblFeild16.Caption = !field16
        If Not IsNull(!field17) Then lblFeild17.Caption = !field17
        If Not IsNull(!field18) Then lblFeild18.Caption = !field18
        If Not IsNull(!field19) Then lblFeild19.Caption = !field19
        If Not IsNull(!field20) Then lblFeild20.Caption = !field20
        If Not IsNull(!field21) Then lblFeild21.Caption = !field21
        If Not IsNull(!field22) Then lblFeild22.Caption = !field22
        If Not IsNull(!field23) Then lblFeild23.Caption = !field23
        If Not IsNull(!field24) Then lblFeild24.Caption = !field24
        If Not IsNull(!Field25) Then lblFeild25.Caption = !Field25
        If Not IsNull(!fieldref1) Then lblRef1.Caption = !fieldref1
        If Not IsNull(!fieldref2) Then lblRef2.Caption = !fieldref2
        If Not IsNull(!fieldref3) Then lblRef3.Caption = !fieldref3
        If Not IsNull(!fieldref4) Then lblRef4.Caption = !fieldref4
        If Not IsNull(!fieldref5) Then lblRef5.Caption = !fieldref5
        If Not IsNull(!fieldref6) Then lblRef6.Caption = !fieldref6
        If Not IsNull(!fieldref7) Then lblRef7.Caption = !fieldref7
        If Not IsNull(!fieldref8) Then lblRef8.Caption = !fieldref8
        If Not IsNull(!fieldref9) Then lblRef9.Caption = !fieldref9
        If Not IsNull(!fieldref10) Then lblRef10.Caption = !fieldref10
        If Not IsNull(!fieldref11) Then lblRef11.Caption = !fieldref11
        If Not IsNull(!fieldref12) Then lblRef12.Caption = !fieldref12
        If Not IsNull(!fieldref13) Then lblRef13.Caption = !fieldref13
        If Not IsNull(!fieldref14) Then lblRef14.Caption = !fieldref14
        If Not IsNull(!Fieldref15) Then lblRef15.Caption = !Fieldref15
        If Not IsNull(!fieldref16) Then lblRef16.Caption = !fieldref16
        If Not IsNull(!fieldref17) Then lblRef17.Caption = !fieldref17
        If Not IsNull(!fieldref18) Then lblRef18.Caption = !fieldref18
        If Not IsNull(!fieldref19) Then lblRef19.Caption = !fieldref19
        If Not IsNull(!fieldref20) Then lblRef20.Caption = !fieldref20
        If Not IsNull(!fieldref21) Then lblRef21.Caption = !fieldref21
        If Not IsNull(!fieldref22) Then lblRef22.Caption = !fieldref22
        If Not IsNull(!fieldref23) Then lblRef23.Caption = !fieldref23
        If Not IsNull(!fieldref24) Then lblRef24.Caption = !fieldref24
        If Not IsNull(!Fieldref25) Then lblRef25.Caption = !Fieldref25
        If Not IsNull(!fieldunit1) Then lblUnit1.Caption = !fieldunit1
        If Not IsNull(!fieldunit2) Then lblUnit2.Caption = !fieldunit2
        If Not IsNull(!fieldunit3) Then lblUnit3.Caption = !fieldunit3
        If Not IsNull(!fieldunit4) Then lblUnit4.Caption = !fieldunit4
        If Not IsNull(!fieldunit5) Then lblUnit5.Caption = !fieldunit5
        If Not IsNull(!fieldunit6) Then lblUnit6.Caption = !fieldunit6
        If Not IsNull(!fieldunit7) Then lblUnit7.Caption = !fieldunit7
        If Not IsNull(!fieldunit8) Then lblUnit8.Caption = !fieldunit8
        If Not IsNull(!fieldunit9) Then lblUnit9.Caption = !fieldunit9
        If Not IsNull(!fieldunit10) Then lblUnit10.Caption = !fieldunit10
        If Not IsNull(!fieldunit11) Then lblUnit11.Caption = !fieldunit11
        If Not IsNull(!fieldunit12) Then lblUnit12.Caption = !fieldunit12
        If Not IsNull(!fieldunit13) Then lblUnit13.Caption = !fieldunit13
        If Not IsNull(!fieldunit14) Then lblUnit14.Caption = !fieldunit14
        If Not IsNull(!Fieldunit15) Then lblUnit15.Caption = !Fieldunit15
        If Not IsNull(!fieldunit16) Then lblUnit16.Caption = !fieldunit16
        If Not IsNull(!fieldunit17) Then lblUnit17.Caption = !fieldunit17
        If Not IsNull(!fieldunit18) Then lblUnit18.Caption = !fieldunit18
        If Not IsNull(!fieldunit19) Then lblUnit19.Caption = !fieldunit19
        If Not IsNull(!fieldunit20) Then lblUnit20.Caption = !fieldunit20
        If Not IsNull(!fieldunit21) Then lblUnit21.Caption = !fieldunit21
        If Not IsNull(!fieldunit22) Then lblUnit22.Caption = !fieldunit22
        If Not IsNull(!fieldunit23) Then lblUnit23.Caption = !fieldunit23
        If Not IsNull(!fieldunit24) Then lblUnit24.Caption = !fieldunit24
        If Not IsNull(!Fieldunit25) Then lblUnit25.Caption = !Fieldunit25
        If Not IsNull(!Comments) Then txtComments.Text = !Comments

End With

End Sub

Private Sub txtSearchFirstName_Change()
    If Trim(txtSearchFirstName.Text) = "" Then Exit Sub
    If Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) = "" Then ListFirstNames
    If Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) <> "" Then ListBothNames
End Sub



Private Sub GetDetails()
    Call ClearPatientValues
    With DataEnvironment1.rssqlTem
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
        txtID.Text = !patient_ID
        .Close
    End With
End Sub


Private Sub ClearPatientValues()
    txtFirstName.Text = Empty
    txtSurname.Text = Empty
    txtOtherName.Text = Empty
    DataComboTitle.BoundText = Empty
    DataComboSex.BoundText = Empty
    DataComboMarietal.BoundText = Empty
    DataComboRace.BoundText = Empty
    txtID.Text = Empty
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
    If Trim(txtSearchSurname.Text) = "" Then Exit Sub
    If Trim(txtSearchFirstName.Text) = "" And Trim(txtSearchSurname.Text) = "" Then ListSurname
    If Trim(txtSearchFirstName.Text) <> "" And Trim(txtSearchSurname.Text) <> "" Then ListBothNames
End Sub



Private Sub grid1_DblClick()
Grid1.Col = 0
    If Grid1.Row < 1 Or Not IsNumeric(Grid1.Text) Then
        bttnSelect.Enabled = False
        Exit Sub
    Else
        Grid1.Col = 0
        If Not IsNumeric(Grid1.Text) Then Exit Sub
        TemPatientID = Val(Grid1.Text)
        Call GetDetails
        Grid1.Col = 0
        Grid1.ColSel = Grid1.Cols - 1
        bttnSelect.Enabled = True
        bttnSearch_Click
    End If
End Sub

Private Sub Grid1_Click()
Grid1.Col = 0
    If Grid1.Row < 1 Or Not IsNumeric(Grid1.Text) Then
        bttnSelect.Enabled = False
        Exit Sub
    Else
        Grid1.Col = 0
        If Not IsNumeric(Grid1.Text) Then Exit Sub
        TemPatientID = Val(Grid1.Text)
        Call GetDetails
        Grid1.Col = 0
        Grid1.ColSel = Grid1.Cols - 1
        bttnSelect.Enabled = True
    End If
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
            Grid1.Rows = NowRow + 1
            Grid1.Row = NowRow
            Grid1.Col = 0
            Grid1.CellAlignment = 7
            Grid1.Text = !patient_ID
            Grid1.Col = 1
            Grid1.CellAlignment = 7
            If Not IsNull(!firstname) Then Grid1.Text = !firstname
            Grid1.Col = 2
            Grid1.CellAlignment = 7
            If Not IsNull(!surname) Then Grid1.Text = !surname
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
    Dim TemResponce As Byte
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

Private Sub FormatGrid()
    Dim BorderMargin As Long
    BorderMargin = 100
    With Grid1
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

Private Sub Form_Load()
    Call Setcolours
    Call PrepareToSearchPatient
    Call ClearSearchValues
    Call FormatGrid
    SSTab1.Tab = 0
End Sub

Private Sub PrepareToSearchPatient()
    TemPatientID = Empty
    frameSearchPatient.Visible = True
    SSTab1.Tab = 1
    
End Sub

Private Sub PrepareToBook()
    SSTab1.Tab = 0
End Sub

Private Sub PrepareToAdd()
    SSTab1.Tab = 1
    Call FormatGrid
    Call ClearSearchValues
End Sub

Private Sub bttnSelect_Click()
    Call PrepareToAdd
End Sub

Private Sub Setcolours()

End Sub

Private Sub PrepareForNewPatient()

End Sub

Private Sub PrepareForNewIx()
    DataComboIx.Text = Empty
    DataComboSpeciman.Text = Empty
    txtSpecimanNo.Text = Empty
    DataComboDoctor.Text = Empty
    DataComboInstitute.Text = Empty
    DataComboDepartment.Text = Empty
    DataComboCollected.Text = Empty
    DTPickerCollectedDate.Value = Date
    DTPickerCollectedTime.Value = Time
    
    
    DataComboPerformed.Text = Empty
    

End Sub

Private Sub PrintIx()

End Sub

Private Sub SaveIx()

End Sub

Private Sub ChangeIx()

End Sub

