VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDischarge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discharge"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
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
   ScaleHeight     =   7065
   ScaleWidth      =   11160
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   9631
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Patient Details"
      TabPicture(0)   =   "frmDischarge.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Payemnt Details"
      TabPicture(1)   =   "frmDischarge.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Other Charges"
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   6975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Charges for Medicines"
         Height          =   1575
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   6975
         Begin VB.OptionButton Option4 
            Caption         =   "By Drugs"
            Height          =   255
            Left            =   2040
            TabIndex        =   25
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option3 
            Caption         =   "By Drug Category"
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option2 
            Caption         =   "By Bill Totals"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Total"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   6975
         Begin VB.TextBox txtBHTBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   4440
            Width           =   2535
         End
         Begin VB.TextBox txtFName 
            Height          =   375
            Left            =   2280
            TabIndex        =   9
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox txtAddress 
            Height          =   1215
            Left            =   2280
            TabIndex        =   8
            Top             =   840
            Width           =   3975
         End
         Begin VB.TextBox txtPhone 
            Height          =   375
            Left            =   2280
            TabIndex        =   7
            Top             =   2160
            Width           =   3975
         End
         Begin VB.TextBox txtNIC 
            Height          =   375
            Left            =   2280
            TabIndex        =   6
            Top             =   2640
            Width           =   3975
         End
         Begin MSComCtl2.DTPicker dtpTOD 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "h:mm:ss AMPM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   11
            Top             =   3720
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "hour MIN sec"
            Format          =   65994754
            CurrentDate     =   39589
         End
         Begin MSComCtl2.DTPicker dtpDOD 
            Height          =   375
            Left            =   2280
            TabIndex        =   12
            Top             =   3120
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   65994755
            CurrentDate     =   39614
         End
         Begin VB.Label Label2 
            Caption         =   "&Date of Discharge"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   3240
            Width           =   2055
         End
         Begin VB.Label Label10 
            Caption         =   "&Time of Discharge"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   3840
            Width           =   2175
         End
         Begin VB.Label Label40 
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   4440
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "&Phone"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "&Address"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "First &Name"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "&NIC No"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2640
            Width           =   1575
         End
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9120
      TabIndex        =   3
      Top             =   6360
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
   Begin btButtonEx.ButtonEx bttnDischarge 
      Height          =   495
      Left            =   7800
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Discharge"
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
   Begin MSDataListLib.DataCombo dtcBHT 
      Height          =   5220
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9208
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin VB.Label Label38 
      Caption         =   "BH&T"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmDischarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsBHT As New ADODB.Recordset
    Dim rsTemBHT As New ADODB.Recordset
    Dim rsPatients As New ADODB.Recordset
    Dim temSQL As String


Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDischarge_Click()
    Dim TemBHTCredit As Double
    If IsNumeric(dtcBHT.BoundText) = False Then Exit Sub
    With rsTemBHT
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblBHT where BHTID = " & Val(dtcBHT.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !DIscharge = True
            !DOD = Format(dtpDOD.Value, "dd MMMM yyyy")
            !TOD = dtpTOD.Value
            .Update
            MsgBox "Discharged"
            If .State = 1 Then .Close
            Unload Me
        End If
        If .State = 1 Then .Close
    End With
    Call FillCombos
End Sub

Private Sub dtcBHT_Click(Area As Integer)
    Dim TemBHTCredit As Double
    Dim temPatientID As Long
    If IsNumeric(dtcBHT.BoundText) = False Then Exit Sub
    With rsTemBHT
        If .State = 1 Then .Close
        temSQL = "SELECT * from tblBHT where BHTID = " & Val(dtcBHT.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If Not IsNull(!Balance) Then
                TemBHTCredit = !Balance
            Else
                TemBHTCredit = 0
            End If
            If TemBHTCredit < 0 Then
                txtBHTBalance.Text = "(" & Format(Abs(TemBHTCredit), "#,##0.00") & ")"
            Else
                txtBHTBalance.Text = Format(TemBHTCredit, "#,##0.00")
            End If
        End If
        temPatientID = !PatientID
        If .State = 1 Then .Close
    End With
    
    
    With rsPatients
    If .State = 1 Then .Close
    temSQL = "SELECT * FROM tblPatientMainDetails WHERE PatientID = " & temPatientID
    .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        txtFName.Text = !FirstName
        If IsNull(!Address) = False Then txtAddress.Text = !Address
        If IsNull(!Phone) = False Then txtPhone.Text = !Phone
        If IsNull(!NICNo) = False Then txtNIC.Text = !NICNo
    End If
    .Close
    End With
    
End Sub

Private Sub Form_Load()
    Call FillCombos
End Sub

Private Sub FillCombos()
    With rsBHT
        If .State = 1 Then .Close
        temSQL = "SELECT tblBHT.* FROM tblBHT WHERE (((tblBHT.Discharge)=False)) ORDER BY tblBHT.BHT"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
End Sub

