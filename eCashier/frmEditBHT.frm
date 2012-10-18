VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditBHT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admit Patients"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   15270
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "BHT"
      TabPicture(0)   =   "frmEditBHT.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmbBHT"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Patient"
      TabPicture(1)   =   "frmEditBHT.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmbPatient"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSDataListLib.DataCombo cmbBHT 
         Height          =   5940
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   10478
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPatient 
         Height          =   5940
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   10478
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   1
         Text            =   ""
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   11535
      Begin VB.TextBox txtComments 
         Alignment       =   1  'Right Justify
         Height          =   735
         Left            =   8280
         TabIndex        =   35
         Top             =   4800
         Width           =   3135
      End
      Begin VB.CheckBox chkDischarge 
         Caption         =   "Discharged"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   1200
         Width           =   9495
      End
      Begin VB.TextBox txtAddress 
         Height          =   360
         Left            =   1800
         TabIndex        =   13
         Top             =   1680
         Width           =   9495
      End
      Begin VB.TextBox txtGuardianPhone 
         Height          =   375
         Left            =   8280
         TabIndex        =   26
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox txtGunadianNIC 
         Height          =   375
         Left            =   1800
         TabIndex        =   24
         Top             =   3600
         Width           =   3975
      End
      Begin VB.TextBox txtBHT 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtAge 
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtGunardian 
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   2640
         Width           =   9495
      End
      Begin VB.TextBox txtGuardianAddress 
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   3120
         Width           =   9495
      End
      Begin VB.TextBox txtPtSurcharge 
         Height          =   360
         Left            =   240
         TabIndex        =   45
         Top             =   3840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtComSurcharge 
         Height          =   360
         Left            =   720
         TabIndex        =   44
         Top             =   3840
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dtpDOA 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22085635
         CurrentDate     =   39956
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   22085635
         CurrentDate     =   39589
      End
      Begin MSDataListLib.DataCombo cmbHealthSchemeSupplier 
         Height          =   360
         Left            =   8280
         TabIndex        =   30
         Top             =   4200
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
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
         Left            =   8280
         TabIndex        =   9
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "hour MIN sec"
         Format          =   22085634
         CurrentDate     =   39589
      End
      Begin MSDataListLib.DataCombo cmbSex 
         Height          =   360
         Left            =   8280
         TabIndex        =   18
         Top             =   2160
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPtCat 
         Height          =   360
         Left            =   1800
         TabIndex        =   28
         Top             =   4200
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpDOD 
         Height          =   375
         Left            =   3480
         TabIndex        =   38
         Top             =   5640
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   22085635
         CurrentDate     =   39956
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
         Left            =   8280
         TabIndex        =   40
         Top             =   5640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "hour MIN sec"
         Format          =   22085634
         CurrentDate     =   39589
      End
      Begin MSDataListLib.DataCombo cmbSpeciality 
         Height          =   360
         Left            =   1800
         TabIndex        =   32
         Top             =   4680
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbStaff 
         Height          =   360
         Left            =   1800
         TabIndex        =   33
         Top             =   5160
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label16 
         Caption         =   "Referring Doctor"
         Height          =   735
         Left            =   120
         TabIndex        =   31
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "C&omments"
         Height          =   255
         Left            =   6960
         TabIndex        =   34
         Top             =   4800
         Width           =   2895
      End
      Begin VB.Label Label15 
         Caption         =   "&Date of Discharge"
         Height          =   255
         Left            =   1800
         TabIndex        =   37
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "&Time"
         Height          =   255
         Left            =   6960
         TabIndex        =   39
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label lblHSS 
         Caption         =   "&Health Scheme Supplier"
         Height          =   495
         Left            =   6960
         TabIndex        =   29
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Age"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "&Phone"
         Height          =   255
         Left            =   6960
         TabIndex        =   25
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "&Address"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "First &Name"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "&NIC No"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "BH&T"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "&Date of Admission"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "&Time"
         Height          =   255
         Left            =   6960
         TabIndex        =   8
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "Sex"
         Height          =   255
         Left            =   6960
         TabIndex        =   17
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Name of &Guardian"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Address"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "&Patient Category"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   4200
         Width           =   1695
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   13800
      TabIndex        =   43
      Top             =   7080
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   495
      Left            =   9240
      TabIndex        =   42
      Top             =   7080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   7680
      TabIndex        =   41
      Top             =   7080
      Width           =   1335
      _ExtentX        =   2355
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
End
Attribute VB_Name = "frmEditBHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsPatients As New ADODB.Recordset
    Dim rsHSS As New ADODB.Recordset
    Dim rsBHT As New ADODB.Recordset
    Dim rsRoom As New ADODB.Recordset
    Dim rsTemRoom As New ADODB.Recordset
    Dim rsRoomPatient As New ADODB.Recordset
    Dim temSql As String
    Dim temPatientID As Long
    Dim temBHTID As Long
    Dim PCat As New clsPatientCategory
    Dim FirstActivation As Boolean
    Dim rsPt As New ADODB.Recordset
    Dim temPtID As Long
    Dim rsStaff As New ADODB.Recordset
    
Private Sub btnSave_Click()
    With rsPatients
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPatientMainDetails where PatientID = " & temPtID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !FirstName = UCase(txtName.Text)
            !Address = txtAddress.Text
            !DateOfBirth = dtpDOB.Value
            !SexID = Val(cmbSex.BoundText)
            .Update
        End If
        .Close
    End With
    With rsBHT
    If .State = 1 Then .Close
        temSql = "SELECT * FROM tblBHT where BHTID = " & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !BHT = txtBHT.Text
            !DOA = dtpDOA.Value
            !Discharge = False
            !TOA = Format(dtpDOA.Value, "hh:mm:ss") & " " & dtpTOA.Value
            !TemAge = CalculateAgeInWords(dtpDOB.Value)
            !HealthSchemeSupplierID = Val(cmbHealthSchemeSupplier.BoundText)
            !PtSurcharge = Val(txtPtSurcharge.Text)
            !ComSurcharge = Val(txtComSurcharge.Text)
            !PatientCategoryID = Val(cmbPtCat.BoundText)
            !GuardianName = txtGunardian.Text
            !GuardianPhone = txtGuardianPhone.Text
            !GuardianAddress = txtGuardianAddress.Text
            !GuardianNIC = txtGunadianNIC.Text
            !Comments = txtComments.Text
            !ReferringDoctorID = Val(cmbStaff.BoundText)
            .Update
        End If
        .Close
    End With
    Call ClearValues
    Call FillCombos
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub chkDischarge_Click()
    If chkDischarge.Value = 1 Then
        dtpDOD.Visible = True
        dtpTOD.Visible = True
    Else
        dtpDOD.Visible = False
        dtpTOD.Visible = False
    End If
End Sub

Private Sub chkDischarge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If dtpDOD.Visible = True Then
            dtpDOD.SetFocus
        Else
            btnSave.SetFocus
        End If
    End If
End Sub

Private Sub cmbBHT_Change()
    Call ClearValues
    Call DisplayDetails
End Sub

Private Sub cmbBHT_Click(Area As Integer)
    Call ClearValues
    Call DisplayDetails
End Sub

Private Sub ClearValues()
    txtAddress.Text = Empty
    txtAge.Text = Empty
    txtBHT.Text = Empty
    txtComments.Text = Empty
    txtComSurcharge.Text = Empty
    txtGuardianAddress.Text = Empty
    txtGuardianPhone.Text = Empty
    txtGunadianNIC.Text = Empty
    txtGunardian.Text = Empty
    txtName.Text = Empty
    txtPtSurcharge.Text = Empty
    cmbStaff.Text = Empty
    cmbHealthSchemeSupplier.Text = Empty
    cmbPtCat.Text = Empty
    cmbSex.Text = Empty
End Sub

Private Sub DisplayDetails(): On Error Resume Next
    With rsBHT
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblBHT where BHTID = " & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtBHT.Text = !BHT
            temPtID = !PatientID
            dtpDOA.Value = !DOA
            dtpTOA.Value = Format(!TOA, "hh:mm:ss")
            txtGuardianAddress.Text = Format(!GuardianAddress, "")
            txtGuardianPhone.Text = Format(!GuardianPhone, "")
            txtGunadianNIC.Text = Format(!GuardianNIC, "")
            txtGunardian.Text = Format(!GuardianName, "")
            If Not IsNull(!PatientCategoryID) Then cmbPtCat.BoundText = !PatientCategoryID
            cmbHealthSchemeSupplier.BoundText = !HealthSchemeSupplierID
            If !Discharge = True Then
                chkDischarge.Value = 1
                dtpDOD.Visible = True
                dtpTOD.Visible = True
                dtpDOD.Value = !DOD
                dtpTOD.Value = Format(!TOD, "hh:mm:ss")
            Else
                chkDischarge.Value = 0
                dtpDOD.Visible = False
                dtpTOD.Visible = False
            End If
            If Not IsNull(!Comments) Then txtComments.Text = !Comments
            cmbSpeciality.Text = Empty
            If IsNull(!ReferringDoctorID) = False Then cmbStaff.BoundText = !ReferringDoctorID
        End If
        .Close
    End With
    With rsPatients
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPatientMainDetails where PatientID = " & temPtID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            txtName.Text = !FirstName
            txtAddress.Text = !Address
            dtpDOB.Value = !DateOfBirth
            cmbSex.BoundText = !SexID
        End If
        .Close
    End With
End Sub

Private Sub cmbHealthSchemeSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbHealthSchemeSupplier.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbSpeciality.SetFocus
    End If
End Sub

Private Sub cmbPatient_Change()
    Call ClearValues
    On Error Resume Next
    cmbBHT.BoundText = Val(cmbPatient.BoundText)
End Sub

Private Sub cmbPatient_Click(Area As Integer)
    On Error Resume Next
    cmbBHT.BoundText = Val(cmbPatient.BoundText)
End Sub

Private Sub cmbPtCat_Change()
    If IsNumeric(cmbPtCat.BoundText) = False Then
        txtPtSurcharge.Text = 0
        Exit Sub
    End If
    PCat.ID = Val(cmbPtCat.BoundText)
    If LCase(PCat.PaymentMethod) = "credit" Then
        cmbHealthSchemeSupplier.Visible = True
        lblHSS.Visible = True
    Else
        cmbHealthSchemeSupplier.Visible = False
        lblHSS.Visible = False
    End If
    txtPtSurcharge.Text = PCat.Surcharge
    
End Sub

Private Sub cmbPtCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbPtCat.Text = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If cmbHealthSchemeSupplier.Visible = True Then
            cmbHealthSchemeSupplier.SetFocus
        Else
            cmbSpeciality.SetFocus
        End If
    End If
End Sub

Private Sub cmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtGunardian.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbSex.Text = Empty
    End If
End Sub

Private Sub cmbSpeciality_Change()
    With rsStaff
        If .State = 1 Then .Close
        If IsNumeric(cmbSpeciality.BoundText) = True Then
            temSql = "Select * from tblStaff where SpecialityID = " & Val(cmbSpeciality.BoundText) & " ORDER BY Name"
        Else
            temSql = "Select * from tblStaff Order BY Name"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbStaff
        Set .RowSource = rsStaff
        .ListField = "Name"
        .BoundColumn = "StaffID"
        .Text = Empty
    End With

End Sub

Private Sub cmbSpeciality_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbStaff.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbSpeciality.Text = Empty
    End If
End Sub

Private Sub cmbStaff_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtComments.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbStaff.Text = Empty
    End If
End Sub

Private Sub dtpDOB_Change()
    txtAge.Text = DateDiff("yyyy", dtpDOB.Value, Date)
End Sub

Private Sub dtpDOB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtpDOB.Value = Date
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbSex.SetFocus
    End If
End Sub

Private Sub dtpDOD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpTOD.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        dtpDOD.Value = Date
    End If
End Sub

Private Sub dtpTOA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtpTOA.Value = Date
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtName.SetFocus
    End If
End Sub

Private Sub dtpTOD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnSave_Click
    End If
End Sub

Private Sub Form_Activate()
    If FirstActivation = True Then
        txtBHT.SetFocus
        SendKeys "{home}+{end}"
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    FirstActivation = True
    txtBHT.Text = NewBHT
    dtpDOB.Value = Date
    dtpDOA.Value = Date
    dtpTOA.Value = Time
End Sub

Private Function NewBHT() As Long
    Dim rsTemBHT As New ADODB.Recordset
    With rsTemBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT order by BHTID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            NewBHT = Val(!BHT) + 1
        Else
            NewBHT = 1
        End If
        .Close
    End With
End Function

Private Sub FillCombos()
    Dim Sex As New clsFillCombos
    Sex.FillAnyCombo cmbSex, "Sex", False
    Dim PtCat As New clsFillCombos
    PtCat.FillAnyCombo cmbPtCat, "PatientCategory", True
    With rsHSS
        If .State = 1 Then .Close
        temSql = "SELECT * from tblHealthSchemeSuppliers order by HealthSchemeSupplierName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbHealthSchemeSupplier
        Set .RowSource = rsHSS
        .ListField = "HealthSchemeSupplierName"
        .BoundColumn = "HealthSchemeSupplierID"
    End With
    Dim BHT As New clsFillCombos
    BHT.FillAnyCombo cmbBHT, "BHT", False
    With rsPt
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientMainDetails.FirstName, tblBHT.BHTID " & _
                    "FROM tblBHT LEFT JOIN tblPatientMainDetails ON tblBHT.PatientID = tblPatientMainDetails.PatientID " & _
                    "ORDER BY tblPatientMainDetails.FirstName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbPatient
        Set .RowSource = rsPt
        .ListField = "FirstName"
        .BoundColumn = "BHTID"
    End With
    Dim Speciality As New clsFillCombos
    Speciality.FillAnyCombo cmbSpeciality, "Speciality", False
End Sub


Private Sub dtpDOA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtpDOA.Value = Date
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpTOA.SetFocus
    End If
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtAddress.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtAge.SetFocus
    End If
End Sub

Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpDOB.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtAge.Text = Empty
    End If
End Sub

Private Sub txtAge_LostFocus()
    Dim TemDOB As Date
    TemDOB = DateSerial(Year(Date) - Val(txtAge.Text), Month(Date), Day(Date))
    If TemDOB - dtpDOB.Value > 365 Or TemDOB - dtpDOB.Value < -365 Then
        dtpDOB.Value = TemDOB
    End If
End Sub

Private Sub txtBHT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtBHT.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpDOA.SetFocus
    End If
End Sub

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        chkDischarge.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtComments.Text = Empty
    End If
End Sub

Private Sub txtGuardianAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtGunadianNIC.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtGuardianAddress.Text = Empty
    End If
    
End Sub

Private Sub txtGunardian_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtGuardianAddress.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtGunardian.Text = Empty
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtName.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtAddress.SetFocus
    End If
End Sub

Private Sub txtGunadianNIC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtGunadianNIC.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtGuardianPhone.SetFocus
    End If
End Sub

Private Sub txtGuardianPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtGuardianPhone.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPtCat.SetFocus
    End If
End Sub

Private Sub cmbHealthSchemeSupplier_Change()
    If IsNumeric(cmbHealthSchemeSupplier.BoundText) = False Then
        txtComSurcharge.Text = 0
    Else
        Dim rsTem As New ADODB.Recordset
        With rsTem
            If .State = 1 Then .Close
            temSql = "Select * from tblHealthSchemeSuppliers where HealthSchemeSupplierID = " & Val(cmbHealthSchemeSupplier.BoundText)
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                If IsNull(!InwardAddition) = False Then
                    txtComSurcharge.Text = !InwardAddition
                Else
                    txtComSurcharge.Text = 0
                End If
            Else
                txtComSurcharge.Text = 0
            End If
            .Close
        End With
    End If
End Sub

