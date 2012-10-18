VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPatientFacilityList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Lists"
   ClientHeight    =   8310
   ClientLeft      =   375
   ClientTop       =   1755
   ClientWidth     =   7125
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
   Icon            =   "frmPatientFacilityList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FramePatientList 
      Caption         =   "Patients List"
      Height          =   7575
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6855
      Begin VB.Frame FrameSecession 
         Caption         =   "Secession"
         Height          =   1095
         Left            =   1680
         TabIndex        =   11
         Top             =   1680
         Width           =   4815
         Begin VB.OptionButton OptionMorning 
            Caption         =   "&Morning Secession"
            Height          =   255
            Left            =   840
            TabIndex        =   3
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton OptionEvening 
            Caption         =   "&Evening Secession"
            Height          =   255
            Left            =   840
            TabIndex        =   4
            Top             =   480
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton OptionNoPreferance 
            Caption         =   "No Preferance"
            Height          =   255
            Left            =   840
            TabIndex        =   12
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.OptionButton OptionNotRelevent 
            Caption         =   "&Not Relevent"
            Height          =   255
            Left            =   840
            TabIndex        =   5
            Top             =   720
            Width           =   2055
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GridList1 
         Height          =   4455
         Left            =   360
         TabIndex        =   6
         Top             =   3000
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   7858
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo DataComboFacility 
         Bindings        =   "frmPatientFacilityList.frx":0442
         Height          =   360
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "HospitalFacility"
         BoundColumn     =   "HospitalFacility_ID"
         Text            =   ""
         Object.DataMember      =   "sqlHospitalFacility"
      End
      Begin MSDataListLib.DataCombo DataComboDoctorStaff 
         Bindings        =   "frmPatientFacilityList.frx":0461
         Height          =   360
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "DoctorListedName"
         BoundColumn     =   "Doctor_ID"
         Text            =   ""
         Object.DataMember      =   "sqlDoctor"
      End
      Begin MSComCtl2.DTPicker DTPickerAppointment 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58392579
         CurrentDate     =   39421
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "D&ate :"
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
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblDoctorStaff 
         BackStyle       =   0  'Transparent
         Caption         =   "&Doctor :"
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
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblFacility 
         Caption         =   "Fa&cility :"
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
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
   End
   Begin btButtonEx.ButtonEx bttnCloseList 
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   7800
      Width           =   1575
      _ExtentX        =   2778
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
End
Attribute VB_Name = "frmPatientFacilityList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemPatientID As Long
    Dim TemCatogery  As Integer
    Dim TemDailyMaximum As Integer
    Dim TemStaffFacilityID As Long
    Dim TemHospitalFacilityID As Long
    Dim TemstaffID As Long
    Dim TemPatientFacilityID As Long
    Dim TemBillID As Long
    Dim TemSecession  As Integer
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

'bttnAdd.BackColor = BttnBackColour
'bttnAdd.ForeColor = BttnForeColour

'bttnCancel.BackColor = BttnBackColour
'bttnCancel.ForeColor = BttnForeColour

'bttnChange.BackColor = BttnBackColour
'bttnChange.ForeColor = BttnForeColour

bttnCloseList.BackColor = BttnBackColour
bttnCloseList.ForeColor = BttnForeColour

'bttnEdit.BackColor = BttnBackColour
'bttnEdit.ForeColor = BttnForeColour

'bttnSave.BackColor = BttnBackColour
'bttnSave.ForeColor = BttnForeColour

FramePatientList.BackColor = FrmBackColour
FramePatientList.ForeColor = FrmForeColour


'CheckAgent.BackColor = LblBackColour
'CheckAgent.ForeColor = LblForeColour

'chkBlackListed.BackColor = LblBackColour
'chkBlackListed.ForeColor = LblForeColour


DataComboDoctorStaff.BackColor = TxtBackColour
DataComboDoctorStaff.ForeColor = TxtForeColour

DataComboFacility.BackColor = TxtBackColour
DataComboFacility.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboSex.ForeColor = TxtForeColour

'DataComboSpeciality.BackColor = TxtBackColour
'DataComboSpeciality.ForeColor = TxtForeColour

'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour




'Grid1.BackColor = GridBackColor
'Grid1.ForeColor = GridForeColor

'Grid1.BackColorBkg = GridBackColorBkg
'Grid1.BackColorFixed = GridBackColorFixed
'Grid1.BackColorSel = GridBackColorSel
'
'Grid1.ForeColor = GridForeColor
'Grid1.ForeColorFixed = GridForeColorFixed
'Grid1.ForeColorSel = GridForeColorSel

'grid1.ForeColor = Grid

OptionMorning.BackColor = FrameBackColour
OptionMorning.ForeColor = FrameForeColour

OptionEvening.BackColor = FrameBackColour
OptionEvening.ForeColor = FrameForeColour

OptionNotRelevent.BackColor = FrameBackColour
OptionNotRelevent.ForeColor = FrameForeColour




lblFacility.BackColor = LblBackColour
lblFacility.ForeColor = LblForeColour
'
'lblDoctorStaff.BackColor = LblBackColour
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


'txtAccount.BackColor = TxtBackColour
'txtAccount.ForeColor = TxtForeColour

'txtBankBranch.BackColor = TxtBackColour
'txtBankBranch.ForeColor = TxtForeColour

'txtComment.BackColor = TxtBackColour
'txtComment.ForeColor = TxtForeColour
'txtCredit.BackColor = TxtBackColour
'txtCredit.ForeColor = TxtForeColour
'txtMaxCredit.BackColor = TxtBackColour
'txtMaxCredit.ForeColor = TxtForeColour
'txtListedName.BackColor = TxtBackColour
'txtListedName.ForeColor = TxtForeColour
'txtName.BackColor = TxtBackColour
'txtName.ForeColor = TxtForeColour
'txtAddress.BackColor = TxtBackColour
'txtAddress.ForeColor = TxtForeColour
'txtEmail.BackColor = TxtBackColour
'txtEmail.ForeColor = TxtForeColour
'txtFax.BackColor = TxtBackColour
'txtFax.ForeColor = TxtForeColour
'txtTel.BackColor = TxtBackColour
'txtTel.ForeColor = TxtForeColour
''txtOfficialWebsite.BackColor = TxtBackColour
''txtOfficialWebsite.ForeColor = TxtForeColour

'txtPrivateAddress.BackColor = TxtBackColour
'txtPrivateAddress.ForeColor = TxtForeColour
'txtPrivateEmail.BackColor = TxtBackColour
'txtPrivateEmail.ForeColor = TxtForeColour
'txtPrivateFax.BackColor = TxtBackColour
'txtPrivateFax.ForeColor = TxtForeColour
'txtPrivateMobile.BackColor = TxtBackColour
'txtPrivateMobile.ForeColor = TxtForeColour
'txtPrivateTel.BackColor = TxtBackColour
'txtPrivateTel.ForeColor = TxtForeColour

'txtUserName.BackColor = TxtBackColour
'txtUserName.ForeColor = TxtForeColour
'txtPassword.BackColor = TxtBackColour
'txtPassword.ForeColor = TxtForeColour
'txtReenterPassword.BackColor = TxtBackColour
'txtReenterPassword.ForeColor = TxtForeColour
'txtOfficialFax.BackColor = TxtBackColour
'txtOfficialFax.ForeColor = TxtForeColour
'txtOfficialTel.BackColor = TxtBackColour
'txtOfficialTel.ForeColor = TxtForeColour
'txtOfficialWebsite.BackColor = TxtBackColour
'txtOfficialWebsite.ForeColor = TxtForeColour

frmPatientFacilityList.BackColor = FrmBackColour
frmPatientFacilityList.ForeColor = FrmForeColour

FrameSecession.BackColor = FrmBackColour
FrameSecession.ForeColor = FrmForeColour


'txtQualifications.BackColor = TxtBackColour
'txtQualifications.ForeColor = TxtForeColour
'txtRegistation.BackColor = TxtBackColour
'txtRegistation.ForeColor = TxtForeColour
'txtSearch.BackColor = TxtBackColour
'txtSearch.ForeColor = TxtForeColour
End Sub
Public Sub FormatPatientFacilityList()
    Dim BorderMargin As Long
    BorderMargin = 150
    With GridList1
        .Clear
        .Rows = 1
        .Row = 0
        .Cols = 3
        .ColWidth(0) = 900
        .ColWidth(2) = 1
        .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + BorderMargin)
        .Col = 0
        .CellAlignment = 4
        .Text = "Serial"
        .Col = 1
        .CellAlignment = 4
        .Text = "Patient Name"
    End With
End Sub

Public Sub FillPatientFacilityList()
    Dim NowRow As Long
    Dim TemNum As Long
    If TemDailyMaximum > 1 Then
        GridList1.Rows = TemDailyMaximum + 1
    Else
        GridList1.Rows = 1
    End If
    GridList1.Col = 0
    For TemNum = 1 To TemDailyMaximum
        GridList1.Rows = TemNum + 1
        GridList1.Row = TemNum
        GridList1.Text = TemNum
    Next
    
    If OptionMorning.Value = True Then
        TemSecession = MorningSecession
    ElseIf OptionEvening.Value = True Then
        TemSecession = EveningSecession
    ElseIf OptionNotRelevent.Value = True Then
        TemSecession = NoReleventSecession
    End If
    
    With DataEnvironment1.rssqlTem3
        If .State = 1 Then .Close
        .Source = "select tblpatientfacility.* from tblpatientfacility where (FacilityStaff_ID = " & TemStaffFacilityID & ") and (appointmentdate = #" & DTPickerAppointment.Value & "#) and (secession = " & TemSecession & ") order by dayserial"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        While Not .EOF
            TemNum = !DaySerial
            If TemNum + 1 >= GridList1.Rows Then GridList1.Rows = TemNum + 1
            GridList1.Row = TemNum
            GridList1.Col = 1
            GridList1.Text = FindPatientByID(!patientid)
            GridList1.Col = 2
            GridList1.Text = !patientfacility_ID
            .MoveNext
        Wend
        .Close
    GridList1.Col = 0
    For TemNum = 1 To GridList1.Rows - 1
        GridList1.Row = TemNum
        GridList1.Text = TemNum
    Next
    End With
    GridList1.Row = 0
    GridList1.Col = 0
End Sub

Private Sub bttnCloseList_Click()
    frmBoooking.ZOrder 0
    Me.Hide
    frmBoooking.Top = 0
    frmBoooking.Left = 0
End Sub

Private Sub DataComboDoctorStaff_Change()

    On Error Resume Next

    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
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
        Case Staff:
            LblDoctorStaff.Caption = "Staff Member :"
        Case Investigation:
            LblDoctorStaff.Caption = "Investigation :"
        Case Other:
        End Select
        .Close
    
    End With
    Call FormatPatientFacilityList
    Call FillPatientFacilityList
End Sub

Private Sub DataComboFacility_Change()

    On Error Resume Next
    
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    TemHospitalFacilityID = DataComboFacility.BoundText
    DataComboDoctorStaff.Text = Empty
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
              On Error Resume Next
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
    FormatPatientFacilityList
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

Private Sub DTPickerAppointment_Change()
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    FormatPatientFacilityList
    FillPatientFacilityList
End Sub


Private Sub Form_Load()
Call Setcolours
End Sub

Private Sub OptionEvening_Click()
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    FormatPatientFacilityList
    FillPatientFacilityList
End Sub

Private Sub OptionMorning_Click()
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    FormatPatientFacilityList
    FillPatientFacilityList
End Sub

Private Sub OptionNotRelevent_Click()
    If Not IsNumeric(DataComboFacility.BoundText) Then Exit Sub
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    FormatPatientFacilityList
    FillPatientFacilityList
End Sub
