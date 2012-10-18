VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmDoctor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctors Details"
   ClientHeight    =   8700
   ClientLeft      =   1710
   ClientTop       =   2085
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDoctor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   12375
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   5160
      TabIndex        =   36
      Top             =   120
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Doctors Details"
      TabPicture(0)   =   "frmDoctor.frx":0582
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameDoctor"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Payment Details"
      TabPicture(1)   =   "frmDoctor.frx":059E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FramePayment"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameDoctor 
         Caption         =   "Doctor Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7335
         Left            =   -74880
         TabIndex        =   45
         Top             =   360
         Width           =   6735
         Begin VB.TextBox txtName 
            Height          =   375
            Left            =   2280
            MaxLength       =   100
            TabIndex        =   6
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox txtQualifications 
            Height          =   375
            Left            =   2280
            MaxLength       =   100
            TabIndex        =   8
            Top             =   1680
            Width           =   3975
         End
         Begin VB.TextBox txtRegistation 
            Height          =   375
            Left            =   2280
            MaxLength       =   100
            TabIndex        =   9
            Top             =   2160
            Width           =   3975
         End
         Begin VB.TextBox txtDesignation 
            Height          =   375
            Left            =   2280
            MaxLength       =   100
            TabIndex        =   10
            Top             =   2640
            Width           =   3975
         End
         Begin VB.TextBox txtListedName 
            Height          =   375
            Left            =   2280
            MaxLength       =   100
            TabIndex        =   7
            Top             =   1200
            Width           =   3975
         End
         Begin MSDataListLib.DataCombo DataComboTitle 
            Bindings        =   "frmDoctor.frx":05BA
            Height          =   360
            Left            =   2280
            TabIndex        =   4
            Top             =   240
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
            Bindings        =   "frmDoctor.frx":05D9
            Height          =   360
            Left            =   4920
            TabIndex        =   5
            Top             =   240
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
            Bindings        =   "frmDoctor.frx":05F8
            Height          =   360
            Left            =   2280
            TabIndex        =   11
            Top             =   3120
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   635
            _Version        =   393216
            Style           =   2
            ListField       =   "Speciality"
            BoundColumn     =   "Speciality_ID"
            Text            =   ""
            Object.DataMember      =   "sqlSpeciality"
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   3615
            Left            =   120
            TabIndex        =   54
            Top             =   3600
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   6376
            _Version        =   393216
            Tabs            =   2
            Tab             =   1
            TabHeight       =   520
            TabCaption(0)   =   "Private"
            TabPicture(0)   =   "frmDoctor.frx":0617
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "FramePrivate"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Official"
            TabPicture(1)   =   "frmDoctor.frx":0633
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "FrameOfficial"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            Begin VB.Frame FrameOfficial 
               Caption         =   "Official"
               Height          =   3135
               Left            =   120
               TabIndex        =   61
               Top             =   360
               Width           =   6135
               Begin VB.TextBox txtOfficialEMail 
                  Height          =   375
                  Left            =   2040
                  TabIndex        =   20
                  Top             =   2040
                  Width           =   3975
               End
               Begin VB.TextBox txtOfficialFax 
                  Height          =   375
                  Left            =   2040
                  TabIndex        =   19
                  Top             =   1560
                  Width           =   3975
               End
               Begin VB.TextBox txtOfficialTel 
                  Height          =   375
                  Left            =   2040
                  TabIndex        =   18
                  Top             =   1080
                  Width           =   3975
               End
               Begin VB.TextBox txtOfficialAddress 
                  Height          =   735
                  Left            =   2040
                  MultiLine       =   -1  'True
                  TabIndex        =   17
                  Top             =   240
                  Width           =   3975
               End
               Begin VB.TextBox txtOfficialWebsite 
                  Height          =   375
                  Left            =   2040
                  TabIndex        =   21
                  Top             =   2520
                  Width           =   3975
               End
               Begin VB.Label lblOfficialEmail 
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-Mail"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   66
                  Top             =   2040
                  Width           =   2175
               End
               Begin VB.Label Label23 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Fax:"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   65
                  Top             =   1560
                  Width           =   2175
               End
               Begin VB.Label Label24 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Telephone"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   64
                  Top             =   1080
                  Width           =   2175
               End
               Begin VB.Label Label25 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   63
                  Top             =   360
                  Width           =   2175
               End
               Begin VB.Label lblOfficialWebsite 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Website"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   62
                  Top             =   2520
                  Width           =   2175
               End
            End
            Begin VB.Frame FramePrivate 
               Caption         =   "Private"
               Height          =   3135
               Left            =   -74880
               TabIndex        =   55
               Top             =   360
               Width           =   6135
               Begin VB.TextBox txtPrivateAddress 
                  Height          =   735
                  Left            =   2040
                  MultiLine       =   -1  'True
                  TabIndex        =   12
                  Top             =   240
                  Width           =   3975
               End
               Begin VB.TextBox txtPrivateTel 
                  Height          =   375
                  Left            =   2040
                  MaxLength       =   100
                  TabIndex        =   13
                  Top             =   1080
                  Width           =   3975
               End
               Begin VB.TextBox txtPrivateFax 
                  Height          =   375
                  Left            =   2040
                  MaxLength       =   100
                  TabIndex        =   15
                  Top             =   2040
                  Width           =   3975
               End
               Begin VB.TextBox txtPrivateEmail 
                  Height          =   375
                  Left            =   2040
                  MaxLength       =   100
                  TabIndex        =   16
                  Top             =   2520
                  Width           =   3975
               End
               Begin VB.TextBox txtPrivateMobile 
                  Height          =   375
                  Left            =   2040
                  MaxLength       =   100
                  TabIndex        =   14
                  Top             =   1560
                  Width           =   3975
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "&Telephone"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   60
                  Top             =   1080
                  Width           =   2175
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "&Fax:"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   59
                  Top             =   2040
                  Width           =   2175
               End
               Begin VB.Label Label5 
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-&Mail"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   58
                  Top             =   2520
                  Width           =   2175
               End
               Begin VB.Label Label27 
                  BackStyle       =   0  'Transparent
                  Caption         =   "&Mobile"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   57
                  Top             =   1560
                  Width           =   2175
               End
               Begin VB.Label Label2 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Add&ress"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   56
                  Top             =   240
                  Width           =   2175
               End
            End
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "&Title"
            Height          =   375
            Left            =   240
            TabIndex        =   53
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "&Qualifications"
            Height          =   375
            Left            =   240
            TabIndex        =   52
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "&Registation"
            Height          =   375
            Left            =   240
            TabIndex        =   51
            Top             =   2160
            Width           =   2415
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "&Designation"
            Height          =   375
            Left            =   240
            TabIndex        =   50
            Top             =   2640
            Width           =   2415
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "L&isted Name"
            Height          =   375
            Left            =   240
            TabIndex        =   49
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "&Sex"
            Height          =   375
            Left            =   4440
            TabIndex        =   48
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "S&peciality"
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "&Name"
            Height          =   375
            Left            =   240
            TabIndex        =   46
            Top             =   720
            Width           =   2415
         End
      End
      Begin VB.Frame FramePayment 
         Caption         =   "Payment"
         Height          =   7335
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   6735
         Begin VB.CheckBox chkCreditBooking 
            Caption         =   "Can Book For Credit"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txtComments 
            Height          =   840
            Left            =   2520
            MultiLine       =   -1  'True
            TabIndex        =   29
            Top             =   3720
            Width           =   4095
         End
         Begin VB.TextBox txtAccount 
            Height          =   375
            Left            =   2520
            MaxLength       =   100
            TabIndex        =   28
            Top             =   3240
            Width           =   4095
         End
         Begin VB.TextBox txtBankBranch 
            Height          =   375
            Left            =   2520
            TabIndex        =   27
            Top             =   2760
            Width           =   4095
         End
         Begin VB.TextBox txtCredit 
            Height          =   360
            Left            =   2520
            MaxLength       =   100
            TabIndex        =   25
            Top             =   1800
            Width           =   4095
         End
         Begin VB.CheckBox chkCurrentlyChanneling 
            Caption         =   "Currently Channeling "
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   3495
         End
         Begin MSDataListLib.DataCombo DataComboPaymenyMethod 
            Bindings        =   "frmDoctor.frx":064F
            Height          =   360
            Left            =   2520
            TabIndex        =   24
            Top             =   1320
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
            Bindings        =   "frmDoctor.frx":066E
            Height          =   360
            Left            =   2520
            TabIndex        =   26
            Top             =   2280
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
            Height          =   2055
            Left            =   120
            TabIndex        =   30
            Top             =   5160
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   3625
            _Version        =   393216
            Cols            =   4
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor's Comments"
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   3720
            Width           =   2175
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Payments"
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   4800
            Width           =   2175
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor's Bank"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor's Account"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   3240
            Width           =   2175
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor's Bank Branch"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Method"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor's Credit"
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   1800
            Width           =   2655
         End
      End
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   495
      Left            =   7680
      TabIndex        =   33
      Top             =   8040
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
      Left            =   10800
      TabIndex        =   34
      Top             =   8040
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
      TabIndex        =   31
      Top             =   8040
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11456
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnChange 
      Height          =   495
      Left            =   6000
      TabIndex        =   32
      Top             =   8040
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
      TabIndex        =   35
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "frmDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim BorderMargin As Long
    Dim TemDoctorID As Long
    Dim FromGrid As Boolean

Private Sub bttnAdd_Click()
    Call AfterAdd
    Call ClearValues
    Call PrepairTabs
    txtName.SetFocus
End Sub

Private Sub PrepairTabs()
    SSTab1.Tab = 0
    SSTab2.Tab = 0
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub bttnChange_Click()
    Call EditData
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
    Grid1.Col = 2
    If Not IsNumeric(Grid1.Text) Then Beep: Exit Sub
    TemDoctorID = Val(Grid1.Text)
    Call AfterEdit
    SSTab1.Tab = 0
    txtName.SetFocus
End Sub

Private Sub bttnSave_Click()
    Call SaveData
    Call ClearValues
    Call BeforeAddEdit
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call BeforeAddEdit
    Call ClearValues
    Call FormatGrid
    Call FillGrid
    Call Setcolours
    
End Sub

Private Sub FormatGrid()
    BorderMargin = 100
    With Grid1
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
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "Select tbldoctor.* from tbldoctor order by DoctorListedName"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        NowROw = 0
        Do While .EOF = False
            If Not IsNull(!doctorlistedname) Then
                NowROw = NowROw + 1
                Grid1.Rows = NowROw + 1
                Grid1.Row = NowROw
                Grid1.Col = 0
                Grid1.CellAlignment = 7
                Grid1.Text = NowROw
                Grid1.Col = 1
                Grid1.CellAlignment = 1
                Grid1.Text = !doctorlistedname
                Grid1.Col = 2
                Grid1.Text = !Doctor_ID
            End If
            .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub BeforeAddEdit()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = True
    
    bttnSave.Visible = False
    bttnChange.Visible = False
    bttnCancel.Visible = False
    
    FrameDoctor.Enabled = False
    FramePayment.Enabled = False
    
    SSTab1.Tab = 0
        
    
    TemDoctorID = Empty
    
        FromGrid = False

    
End Sub

Private Sub AfterAdd()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
    
    bttnSave.Visible = True
    bttnChange.Visible = False
    bttnCancel.Visible = True
    
    FrameDoctor.Enabled = True
    FramePayment.Enabled = True
    
    TemDoctorID = Empty
End Sub
Private Sub AfterEdit()
    bttnEdit.Enabled = False
    bttnAdd.Enabled = False
    
    bttnSave.Visible = False
    bttnChange.Visible = True
    bttnCancel.Visible = True
    
    FrameDoctor.Enabled = True
    FramePayment.Enabled = True
End Sub

Private Sub SaveData()
    Dim TemResponce  As Integer
    If Trim(txtName.Text) = "" Then
        TemResponce = MsgBox("Please enter a name of a doctor to save", vbCritical + vbOKOnly, "No name")
        SSTab1.Tab = 0
        txtName.SetFocus
        Exit Sub
    End If
    
    If Trim(txtListedName.Text) = "" Then
        TemResponce = MsgBox("Please enter name to be listed before saving", vbCritical + vbOKOnly, "No name")
        SSTab1.Tab = 0
        txtName_LostFocus
        txtListedName.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(DataComboTitle.BoundText) Then
        TemResponce = MsgBox("Please select the title", vbCritical, "Title?")
        SSTab1.Tab = 0
        DataComboTitle.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(DataComboSex.BoundText) Then
        TemResponce = MsgBox("Please select the sex", vbCritical, "Sex?")
        SSTab1.Tab = 0
        DataComboSex.SetFocus
        Exit Sub
    End If

    'On Error GoTo ErrorHandler

        With DataEnvironment1.rssqlDoctor
            If .State = 0 Then .Open
            .AddNew
            !DoctorTitle_ID = DataComboTitle.BoundText
            !doctorsex_ID = DataComboSex.BoundText
            !doctorname = txtName.Text
            !doctorlistedname = txtListedName.Text
            !doctorqualifications = txtQualifications.Text
            !doctordesignation = txtDesignation.Text
            !doctorregistation = txtRegistation.Text
            If IsNumeric(DataComboSpeciality.BoundText) Then !DoctorSpeciality_ID = DataComboSpeciality.BoundText
            !doctorprivateaddress = txtPrivateAddress.Text
            !doctorprivatephone = txtPrivateTel.Text
            !doctorprivatefax = txtPrivateFax.Text
            !doctorprivateemail = txtPrivateEmail.Text
            !doctormobilephone = txtPrivateMobile.Text
            !DoctorofficialAddress = txtOfficialAddress.Text
            !Doctorofficialphone = txtOfficialTel.Text
            !DoctorofficialFax = txtOfficialFax.Text
            !DoctorofficialEmail = txtOfficialEMail.Text
            !DoctorWebsite = txtOfficialWebsite.Text
            !DoctorComments = txtComments.Text
            If IsNumeric(DataComboPaymenyMethod.BoundText) Then !DoctorPaymentMethod_id = DataComboPaymenyMethod.BoundText
            If IsNumeric(DataComboBank.BoundText) Then !DoctorBank_id = DataComboBank.BoundText
            !DoctorBankBranch = txtBankBranch.Text
            !DoctorAccount = txtAccount.Text
'            !DoctorCredit = Val(txtCredit.Text)
            If chkCurrentlyChanneling.Value = 1 Then
                !DoctorCurrentlyChanneling = 1
            Else
                !DoctorCurrentlyChanneling = 0
            End If
            If chkCreditBooking.Value = 1 Then
                !CreditBookings = 1
            Else
                !CreditBookings = 0
            End If
            
            .Update
            .Close
        Exit Sub
ErrorHandler:
        TemResponce = MsgBox("An unknown error has occured, Please contact Lakmedipro (077 3177874) with the following details." & vbNewLine & Me.Name & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly, "Update Error")
        .CancelUpdate
        .Close
        Exit Sub
    End With
End Sub

Private Sub EditData()
    Dim TemResponce  As Integer
    If Trim(txtName.Text) = "" Then
        TemResponce = MsgBox("Please enter a name of a doctor to save", vbCritical + vbOKOnly, "No name")
        txtName.SetFocus
        SSTab1.Tab = 0
        Exit Sub
    End If
    
    If Trim(txtListedName.Text) = "" Then
        TemResponce = MsgBox("Please enter name to be listed before saving", vbCritical + vbOKOnly, "No name")
        SSTab1.Tab = 0
        txtName_LostFocus
        txtListedName.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(DataComboTitle.BoundText) Then
        TemResponce = MsgBox("Please select the title", vbCritical, "Title?")
        SSTab1.Tab = 0
        DataComboTitle.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(DataComboSex.BoundText) Then
        TemResponce = MsgBox("Please select the sex", vbCritical, "Sex?")
        SSTab1.Tab = 0
        DataComboSex.SetFocus
        Exit Sub
    End If

    'On Error GoTo ErrorHandler

        With DataEnvironment1.rssqlTem1
            If .State = 1 Then .Close
            .Source = "SELECT tblDoctor.* from tbldoctor where doctor_id = " & TemDoctorID
            If .State = 0 Then .Open
            If .RecordCount = 0 Then Exit Sub
            
            !DoctorTitle_ID = DataComboTitle.BoundText
            !doctorsex_ID = DataComboSex.BoundText
            !doctorname = txtName.Text
            !doctorlistedname = txtListedName.Text
            !doctorqualifications = txtQualifications.Text
            !doctordesignation = txtDesignation.Text
            !doctorregistation = txtRegistation.Text
            If IsNumeric(DataComboSpeciality.BoundText) Then !DoctorSpeciality_ID = DataComboSpeciality.BoundText
            !doctorprivateaddress = txtPrivateAddress.Text
            !doctorprivatephone = txtPrivateTel.Text
            !doctorprivatefax = txtPrivateFax.Text
            !doctorprivateemail = txtPrivateEmail.Text
            !doctormobilephone = txtPrivateMobile.Text
            !DoctorofficialAddress = txtOfficialAddress.Text
            !Doctorofficialphone = txtOfficialTel.Text
            !DoctorofficialFax = txtOfficialFax.Text
            !DoctorofficialEmail = txtOfficialEMail.Text
            !DoctorWebsite = txtOfficialWebsite.Text
            !DoctorComments = txtComments.Text
            If IsNumeric(DataComboPaymenyMethod.BoundText) Then !DoctorPaymentMethod_id = DataComboPaymenyMethod.BoundText
            If IsNumeric(DataComboBank.BoundText) Then !DoctorBank_id = DataComboBank.BoundText
            !DoctorBankBranch = txtBankBranch.Text
            !DoctorAccount = txtAccount.Text
'            !DoctorCredit = Val(txtCredit.Text)
            If chkCurrentlyChanneling.Value = 1 Then
                !DoctorCurrentlyChanneling = 1
            Else
                !DoctorCurrentlyChanneling = 0
            End If
            If chkCreditBooking.Value = 1 Then
                !CreditBookings = 1
            Else
                !CreditBookings = 0
            End If
            .Update
            .Close
        Exit Sub
ErrorHandler:
        TemResponce = MsgBox("An unknown error has occured, Please contact Lakmedipro (077 3177874) with the following details." & vbNewLine & Me.Name & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly, "Update Error")
        .CancelUpdate
        .Close
        Exit Sub
    End With
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
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
    chkCreditBooking.Value = 0
    txtBankBranch.Text = Empty
    txtAccount.Text = Empty
    txtCredit.Text = Empty
    txtComments.Text = Empty
    
    txtSearch.Text = Empty
    
    Call FormatGrid
    
End Sub


Private Sub GetData()
    Grid1.Col = 2
    If IsNumeric(Grid1.Text) = False Then Exit Sub
    TemDoctorID = Val(Grid1.Text)
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tbldoctor.* from tbldoctor where doctor_ID = " & TemDoctorID
        If .State = 0 Then .Open
        
        If .RecordCount = 0 Then Exit Sub
    
    
        If Not IsNull(!DoctorTitle_ID) Then DataComboTitle.BoundText = !DoctorTitle_ID
        If Not IsNull(!doctorsex_ID) Then DataComboSex.BoundText = !doctorsex_ID
        If Not IsNull(!doctorname) Then txtName.Text = !doctorname
        If Not IsNull(!doctorlistedname) Then txtListedName.Text = !doctorlistedname
        If Not IsNull(!doctorqualifications) Then txtQualifications.Text = !doctorqualifications
        If Not IsNull(!doctordesignation) Then txtDesignation.Text = !doctordesignation
        If Not IsNull(!doctorregistation) Then txtRegistation.Text = !doctorregistation
        If Not IsNull(!DoctorSpeciality_ID) Then DataComboSpeciality.BoundText = !DoctorSpeciality_ID
        If Not IsNull(!doctorprivateaddress) Then txtPrivateAddress.Text = !doctorprivateaddress
        If Not IsNull(!doctorprivatephone) Then txtPrivateTel.Text = !doctorprivatephone
        If Not IsNull(!doctorprivatefax) Then txtPrivateFax.Text = !doctorprivatefax
        If Not IsNull(!doctorprivateemail) Then txtPrivateEmail.Text = !doctorprivateemail
        If Not IsNull(!doctormobilephone) Then txtPrivateMobile.Text = !doctormobilephone
        If Not IsNull(!DoctorofficialAddress) Then txtOfficialAddress.Text = !DoctorofficialAddress
        If Not IsNull(!Doctorofficialphone) Then txtOfficialTel.Text = !Doctorofficialphone
        If Not IsNull(!DoctorofficialFax) Then txtOfficialFax.Text = !DoctorofficialFax
        If Not IsNull(!DoctorofficialEmail) Then txtOfficialEMail.Text = !DoctorofficialEmail
        If Not IsNull(!DoctorWebsite) Then txtOfficialWebsite.Text = !DoctorWebsite
        If Not IsNull(!DoctorComments) Then txtComments.Text = !DoctorComments
        If Not IsNull(!DoctorPaymentMethod_id) Then DataComboPaymenyMethod.BoundText = !DoctorPaymentMethod_id
        If Not IsNull(!DoctorBank_id) Then DataComboBank.BoundText = !DoctorBank_id
        If Not IsNull(!DoctorBankBranch) Then txtBankBranch.Text = !DoctorBankBranch
        If Not IsNull(!DoctorAccount) Then txtAccount.Text = !DoctorAccount
        If Not IsNull(!DoctorCredit) Then txtCredit.Text = !DoctorCredit
        If Not IsNull(!DoctorCurrentlyChanneling) Then
            If !DoctorCurrentlyChanneling = True Then
                chkCurrentlyChanneling.Value = 1
            Else
                chkCurrentlyChanneling.Value = 0
            End If
        End If
        If Not IsNull(!CreditBookings) Then
            If !CreditBookings = True Then
                chkCreditBooking.Value = 1
            Else
                chkCreditBooking.Value = 0
            End If
        Else
            chkCreditBooking.Value = 0
        End If
        
        
    End With
End Sub

Private Sub Grid1_Click()
    FromGrid = True
    With Grid1
        If .Row < 1 Then FromGrid = False: Exit Sub
        .Col = 2
        If Not IsNumeric(.Text) Then FromGrid = False: Exit Sub
        TemDoctorID = Val(.Text)
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
    TemFRows = Grid1.Rows
    Grid1.Col = 1
    SearchSuccess = False
    If Len(txtSearch.Text) = 0 Then GoTo MeasureSuccess
    For TemArray = 1 To (TemFRows - 1)
        Grid1.Row = TemArray
        If Len(txtSearch.Text) > Len(Grid1.Text) Then
            GoTo FinishLoop
        Else
            TemLength = Len(txtSearch.Text)
        End If
        If UCase(Left((Grid1.Text), TemLength)) = UCase(txtSearch.Text) Then
            SearchSuccess = True
            Exit For
        Else
            SearchSuccess = False
        End If
FinishLoop:
    Next
    
MeasureSuccess:
    
    If SearchSuccess = True Then
        Grid1.TopRow = TemArray
        Grid1.Row = TemArray
        Grid1.Col = 0
        Grid1.ColSel = (Grid1.Cols - 1)
        bttnEdit.Enabled = True
        bttnAdd.Enabled = False
        Grid1.Col = 2
        TemDoctorID = Grid1.Text
        Call GetData
        Grid1.Col = 0
        Grid1.ColSel = Grid1.Cols - 1
    Else
        Grid1.TopRow = 1
        Grid1.Row = 0
        Grid1.Col = 0
        Grid1.ColSel = 0
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
    End If
'**************************************
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

frmDoctor.BackColor = FrmBackColour
frmDoctor.ForeColor = FrmForeColour

FrameDoctor.BackColor = FrameBackColour
FrameDoctor.ForeColor = FrameForeColour




FramePrivate.BackColor = FrameBackColour
FramePrivate.ForeColor = FrameForeColour

FrameOfficial.BackColor = FrameBackColour
FrameOfficial.ForeColor = FrameForeColour

FramePayment.BackColor = FrameBackColour
FramePayment.ForeColor = FrameForeColour

chkCurrentlyChanneling.BackColor = LblBackColour
chkCurrentlyChanneling.ForeColor = LblForeColour

chkCreditBooking.BackColor = LblBackColour
chkCreditBooking.ForeColor = LblForeColour

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




Grid1.BackColor = GridBackColor
Grid1.ForeColor = GridForeColor

Grid1.BackColorBkg = GridBackColorBkg
Grid1.BackColorFixed = GridBackColorFixed
Grid1.BackColorSel = GridBackColorSel

Grid1.ForeColor = GridForeColor
Grid1.ForeColorFixed = GridForeColorFixed
Grid1.ForeColorSel = GridForeColorSel

'grid1.ForeColor = Grid



Label1.BackColor = LblBackColour
Label1.ForeColor = LblForeColour

Label10.BackColor = LblBackColour
Label10.ForeColor = LblForeColour
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
Label7.BackColor = LblBackColour
Label7.ForeColor = LblForeColour

Label8.BackColor = LblBackColour
Label8.ForeColor = LblForeColour
Label9.BackColor = LblBackColour
Label9.ForeColor = LblForeColour

lblOfficialEmail.BackColor = LblBackColour
lblOfficialEmail.ForeColor = LblForeColour

lblOfficialWebsite.BackColor = LblBackColour
lblOfficialWebsite.ForeColor = LblForeColour


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


txtQualifications.BackColor = TxtBackColour
txtQualifications.ForeColor = TxtForeColour
txtRegistation.BackColor = TxtBackColour
txtRegistation.ForeColor = TxtForeColour
txtSearch.BackColor = TxtBackColour
txtSearch.ForeColor = TxtForeColour
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
