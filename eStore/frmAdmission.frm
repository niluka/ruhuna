VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdmission 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adimission"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   7545
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Patient Details"
      TabPicture(0)   =   "frmAdmission.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FramePatientMainDetails"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Admission Details"
      TabPicture(1)   =   "frmAdmission.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Admission Details"
         Height          =   8055
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   6975
         Begin VB.TextBox Text3 
            Height          =   345
            Left            =   2280
            TabIndex        =   38
            Top             =   1320
            Width           =   4455
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   2280
            TabIndex        =   39
            Top             =   360
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   73859075
            CurrentDate     =   39413
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   2280
            TabIndex        =   40
            Top             =   840
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   73859074
            CurrentDate     =   39413
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Checked By"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   2280
            Width           =   2655
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Admitted By"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "BHT No"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Admission Time"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   840
            Width           =   3495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Admission Date"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.Frame FramePatientMainDetails 
         Caption         =   "Patient Details"
         Height          =   8055
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   6975
         Begin VB.TextBox txtAge 
            Height          =   375
            Left            =   4560
            TabIndex        =   11
            Top             =   3240
            Width           =   1935
         End
         Begin VB.TextBox txtNotes 
            Height          =   825
            Left            =   2040
            TabIndex        =   10
            Top             =   6360
            Width           =   4455
         End
         Begin VB.TextBox txtEmail 
            Height          =   345
            Left            =   2040
            TabIndex        =   9
            Top             =   5880
            Width           =   4455
         End
         Begin VB.TextBox txtFax 
            Height          =   345
            Left            =   2040
            TabIndex        =   8
            Top             =   5400
            Width           =   4455
         End
         Begin VB.TextBox txtTelephone 
            Height          =   345
            Left            =   2040
            TabIndex        =   7
            Top             =   4920
            Width           =   4455
         End
         Begin VB.TextBox txtAddress 
            Height          =   1065
            Left            =   2040
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   3720
            Width           =   4455
         End
         Begin VB.TextBox txtNIC 
            Height          =   345
            Left            =   2040
            MaxLength       =   12
            TabIndex        =   5
            Top             =   2760
            Width           =   4455
         End
         Begin VB.TextBox txtSurname 
            Height          =   345
            Left            =   2040
            TabIndex        =   4
            Top             =   1320
            Width           =   4455
         End
         Begin VB.TextBox txtOtherName 
            Height          =   345
            Left            =   2040
            TabIndex        =   3
            Top             =   840
            Width           =   4455
         End
         Begin VB.TextBox txtFirstName 
            Height          =   345
            Left            =   2040
            TabIndex        =   2
            Top             =   360
            Width           =   4455
         End
         Begin MSDataListLib.DataCombo dtcTitle 
            Height          =   360
            Left            =   2040
            TabIndex        =   12
            Top             =   1800
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin MSDataListLib.DataCombo dtcSex 
            Height          =   360
            Left            =   2040
            TabIndex        =   13
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin MSDataListLib.DataCombo dtcMarietal 
            Height          =   360
            Left            =   4800
            TabIndex        =   14
            Top             =   1800
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin MSDataListLib.DataCombo dtcRace 
            Height          =   360
            Left            =   4800
            TabIndex        =   15
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin MSComCtl2.DTPicker DTPickerDOB 
            Height          =   375
            Left            =   2040
            TabIndex        =   16
            Top             =   3240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   73859073
            CurrentDate     =   39413
         End
         Begin VB.Label Label18 
            Caption         =   "(Age)"
            Height          =   255
            Left            =   3960
            TabIndex        =   31
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Date &of Birth"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   3240
            Width           =   2535
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "No&tes"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   6360
            Width           =   3615
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "E-&Mail"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   5880
            Width           =   2895
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "&Fax"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   5400
            Width           =   3615
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "&Telephone"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   4920
            Width           =   2895
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "&Address"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   3720
            Width           =   3615
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "NIC &No."
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   2760
            Width           =   2775
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "&Race"
            Height          =   255
            Left            =   3960
            TabIndex        =   23
            Top             =   2280
            Width           =   2655
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "&Marietal"
            Height          =   255
            Left            =   3960
            TabIndex        =   22
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Se&x"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   2280
            Width           =   2655
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "&Title"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "&Given name"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "&Family Names"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   840
            Width           =   3495
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "&Name"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   3615
         End
      End
   End
End
Attribute VB_Name = "frmAdmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
