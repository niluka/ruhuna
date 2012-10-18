VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditBHT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit BHT Details"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
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
   ScaleHeight     =   6210
   ScaleWidth      =   10395
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   8880
      TabIndex        =   27
      Top             =   5400
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
   Begin btButtonEx.ButtonEx bttnSave 
      Height          =   495
      Left            =   7560
      TabIndex        =   26
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
      Height          =   5055
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   8916
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Patient"
      TabPicture(0)   =   "frmEditBHT.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "BHT"
      TabPicture(1)   =   "frmEditBHT.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Room"
      TabPicture(2)   =   "frmEditBHT.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "dtcRoom"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtRoomID"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "GridRoomPatient"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin MSFlexGridLib.MSFlexGrid GridRoomPatient 
         Height          =   3015
         Left            =   240
         TabIndex        =   35
         Top             =   1440
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin VB.TextBox txtRoomID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Admission Details"
         Height          =   4575
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   5655
         Begin VB.TextBox txtBHTID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   3960
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkDischarged 
            Alignment       =   1  'Right Justify
            Caption         =   "Discharged"
            Height          =   495
            Left            =   240
            TabIndex        =   21
            Top             =   2280
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpDOA 
            Height          =   375
            Left            =   2280
            TabIndex        =   20
            Top             =   840
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   20578307
            CurrentDate     =   39614
         End
         Begin VB.TextBox txtBHT 
            Height          =   375
            Left            =   2280
            TabIndex        =   15
            Top             =   240
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtpTOA 
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
            TabIndex        =   16
            Top             =   1440
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "hour MIN sec"
            Format          =   20578306
            CurrentDate     =   39589
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
            TabIndex        =   22
            Top             =   3480
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "hour MIN sec"
            Format          =   20578306
            CurrentDate     =   39589
         End
         Begin MSComCtl2.DTPicker dtpDOD 
            Height          =   375
            Left            =   2280
            TabIndex        =   25
            Top             =   2880
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   20578307
            CurrentDate     =   39614
         End
         Begin VB.Label Label10 
            Caption         =   "&Time of Discharge"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   3480
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "&Date of Discharge"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "&Time of Admission"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label7 
            Caption         =   "&Date of Admission"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "BH&T"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   5655
         Begin VB.TextBox txtPatientID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   4080
            Width           =   1335
         End
         Begin VB.TextBox txtNIC 
            Height          =   375
            Left            =   1320
            TabIndex        =   8
            Top             =   3120
            Width           =   3975
         End
         Begin VB.TextBox txtPhone 
            Height          =   375
            Left            =   1320
            TabIndex        =   7
            Top             =   2640
            Width           =   3975
         End
         Begin VB.TextBox txtAddress 
            Height          =   1695
            Left            =   1320
            TabIndex        =   6
            Top             =   840
            Width           =   3975
         End
         Begin VB.TextBox txtFName 
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox txtBHTBalance 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   4080
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker dtpDOB 
            Height          =   375
            Left            =   1320
            TabIndex        =   33
            Top             =   3600
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd MMMM yyyy"
            Format          =   20578307
            CurrentDate     =   39614
         End
         Begin VB.Label Label12 
            Caption         =   "&Date of Birth"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "&NIC No"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "First &Name"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "&Address"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "&Phone"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label40 
            Caption         =   "Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   4080
            Width           =   1575
         End
      End
      Begin MSDataListLib.DataCombo dtcRoom 
         Height          =   360
         Left            =   1920
         TabIndex        =   34
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   "Select Room"
      End
      Begin VB.Label Label11 
         Caption         =   "Current Room"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   2175
      End
   End
   Begin MSDataListLib.DataCombo dtcBHT 
      Height          =   4740
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   8361
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin VB.Label Label38 
      Caption         =   "BH&T"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmEditBHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsBHT As New ADODB.Recordset
    Dim rsTemBHT As New ADODB.Recordset
    Dim rsPatients As New ADODB.Recordset
    Dim rsViewRoom As New ADODB.Recordset
    Dim temSql As String
    Dim temPatientID As Long

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnSave_Click()
    Dim TemBHTCredit As Double
    If CanSave = False Then Exit Sub
    With rsPatients
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPatientMainDetails where PatientID = " & txtPatientID.Text
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !FirstName = txtFName.Text
            !Address = txtAddress.Text
            !Phone = txtPhone.Text
            !NICNo = txtNIC.Text
            !DateOfBirth = dtpDOB.Value
            .Update
        End If
        .Close
    End With
    With rsTemBHT
        If .State = 1 Then .Close
        temSql = "SELECT * from tblBHT where BHTID = " & Val(dtcBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !TOA = dtpTOA.Value
            !DOA = dtpDOA.Value
            !BHT = txtBHT.Text
            
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
            If chkDischarged.Value = 1 Then
                !DIscharge = True
                !DOD = Format(dtpDOD.Value, "dd MMMM yyyy")
                !TOD = dtpTOD.Value
            Else
                !DIscharge = False
                !DOD = vbNull
                !TOD = vbNull
            End If
            .Update
        End If
        If .State = 1 Then .Close
    End With
    Call ClearValues
    Call FillCombos
    dtcBHT.Text = Empty
    dtcBHT.SetFocus
End Sub

Private Function CanSave() As Boolean
    CanSave = False
    Dim tr As Integer
    If Trim(txtFName.Text) = Empty Then
        tr = MsgBox("You have not entered the name of the patient", vbCritical, "Name?")
        txtFName.SetFocus
        Exit Function
    End If
    If Trim(txtBHT.Text) = Empty Then
        tr = MsgBox("You have not entered the BHT number", vbCritical, "BHT?")
        txtBHT.SetFocus
        Exit Function
    End If
    
    CanSave = True
End Function




Private Sub chkDischarged_Click()
    If chkDischarged.Value = 1 Then
        dtpDOD.Visible = True
        dtpTOD.Visible = True
    Else
        dtpDOD.Visible = False
        dtpTOD.Visible = False
    End If
End Sub

Private Sub dtcBHT_Change()
    Call ClearValues
    If IsNumeric(dtcBHT.BoundText) = False Then Exit Sub
    Call DisplayDetails
End Sub

Private Sub Form_Load()
    Call FillCombos
End Sub

Private Sub FillCombos()
    With rsBHT
        If .State = 1 Then .Close
        temSql = "SELECT tblBHT.* FROM tblBHT ORDER BY tblBHT.BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
    With rsViewRoom
        If .State = 1 Then .Close
        temSql = "SELECT * from tblRoom order by Room"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcRoom
        Set .RowSource = rsViewRoom
        .ListField = "Room"
        .BoundText = "RoomID"
    End With
    
End Sub

Private Sub ClearValues()
    txtAddress.Text = Empty
    txtBHT.Text = Empty
    txtBHTBalance.Text = Empty
    txtBHTID.Text = Empty
    txtFName.Text = Empty
    txtNIC.Text = Empty
    txtPatientID.Text = Empty
    txtPhone.Text = Empty
    txtRoomID.Text = Empty
    chkDischarged.Value = 0
    dtpDOA.Value = Date
    dtpDOD.Value = Date
    dtcRoom.Text = Empty

End Sub

Private Sub DisplayDetails()
    Dim TemBHTCredit As Double
    With rsTemBHT
        If .State = 1 Then .Close
        temSql = "SELECT * from tblBHT where BHTID = " & Val(dtcBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 1 Then .Close:   Exit Sub
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
            txtBHTID.Text = !BHTID
            txtBHT.Text = !BHT
            If Not IsNull(!RoomID) Then
                txtRoomID.Text = !RoomID
                dtcRoom.BoundText = !RoomID
            End If
        temPatientID = !PatientID
        txtPatientID.Text = temPatientID
        dtpTOA.Value = !TOA
        dtpDOA.Value = !DOA
        
        If !DIscharge = True Then
            chkDischarged.Value = 1
            dtpDOD.Visible = True
            dtpTOD.Visible = True
            If Not IsNull(!DOD) Then dtpDOD.Value = !DOD
            If Not IsNull(!TOD) Then dtpTOD.Value = !TOD
        Else
            chkDischarged.Value = 0
            dtpDOD.Visible = False
            dtpTOD.Visible = False
        End If
        If .State = 1 Then .Close
    End With
    
    
    With rsPatients
    If .State = 1 Then .Close
    temSql = "SELECT * FROM tblPatientMainDetails WHERE PatientID = " & temPatientID
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        txtFName.Text = !FirstName
        If IsNull(!Address) = False Then txtAddress.Text = !Address
        If IsNull(!Phone) = False Then txtPhone.Text = !Phone
        If IsNull(!NICNo) = False Then txtNIC.Text = !NICNo
    End If
    .Close
    End With

    If txtFName.Text = "Customer" Then
        txtFName.Enabled = False
    Else
        txtFName.Enabled = True
    End If

End Sub
