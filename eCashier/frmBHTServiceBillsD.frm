VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBHTServiceBillsD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BHT Service Bills"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13275
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
   ScaleHeight     =   8370
   ScaleWidth      =   13275
   Begin VB.TextBox txtDetails 
      Height          =   1815
      Left            =   10440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   64
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtSurcharge 
      Height          =   375
      Left            =   7680
      TabIndex        =   63
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkAddToMedicineCharge 
      Caption         =   "Add to Medicine Charge"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   2640
      Width           =   4575
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   6
      Left            =   7920
      TabIndex        =   62
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   5
      Left            =   7920
      TabIndex        =   61
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   4
      Left            =   7920
      TabIndex        =   60
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   3
      Left            =   7920
      TabIndex        =   59
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   2
      Left            =   7920
      TabIndex        =   58
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   57
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtSpecialityID 
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   56
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   6
      Left            =   8400
      TabIndex        =   55
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   54
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   4
      Left            =   8400
      TabIndex        =   53
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   3
      Left            =   8400
      TabIndex        =   52
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   2
      Left            =   8400
      TabIndex        =   51
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   1
      Left            =   8400
      TabIndex        =   50
      Top             =   4080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtServiceProfessionalChargesID 
      Height          =   375
      Index           =   0
      Left            =   8400
      TabIndex        =   49
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtCharge 
      Height          =   375
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtProfessionalCharge 
      Height          =   375
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   6
      Left            =   11520
      TabIndex        =   48
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   5
      Left            =   11520
      TabIndex        =   47
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   4
      Left            =   11520
      TabIndex        =   46
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   3
      Left            =   11520
      TabIndex        =   45
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   11520
      TabIndex        =   44
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   11520
      TabIndex        =   43
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFee1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   11520
      TabIndex        =   42
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   0
      Left            =   8880
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtEditID 
      Height          =   375
      Left            =   7200
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtDelID 
      Height          =   360
      Left            =   6720
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   8640
      TabIndex        =   29
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   67371011
      CurrentDate     =   39956
   End
   Begin VB.TextBox txtComments 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1680
      Width           =   4575
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11520
      TabIndex        =   37
      Top             =   7200
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   6720
      TabIndex        =   25
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin MSFlexGridLib.MSFlexGrid gridService 
      Height          =   3975
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7011
      _Version        =   393216
   End
   Begin VB.TextBox txtHospitalCharge 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSC 
      Height          =   360
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   6720
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Delete"
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   10200
      TabIndex        =   36
      Top             =   7200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
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
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   1
      Left            =   8880
      TabIndex        =   14
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   2
      Left            =   8880
      TabIndex        =   16
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   3
      Left            =   8880
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   4
      Left            =   8880
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   5
      Left            =   8880
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff1 
      Height          =   360
      Index           =   6
      Left            =   8880
      TabIndex        =   24
      Top             =   6480
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   375
      Left            =   8640
      TabIndex        =   31
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   67371010
      CurrentDate     =   39956
   End
   Begin VB.Label Label10 
      Caption         =   "Details"
      Height          =   255
      Left            =   10440
      TabIndex        =   65
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Time"
      Height          =   255
      Left            =   6720
      TabIndex        =   30
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Date"
      Height          =   255
      Left            =   6720
      TabIndex        =   28
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Total Charge"
      Height          =   255
      Left            =   6720
      TabIndex        =   34
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Professional Charge"
      Height          =   255
      Left            =   6720
      TabIndex        =   32
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblSpeciality1 
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   23
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   5
      Left            =   6720
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblSpeciality1 
      Caption         =   "lblSpeciality1"
      Height          =   255
      Index           =   0
      Left            =   6720
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   39
      Top             =   7080
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Hospital Charge"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblSC 
      Caption         =   "Service Subcategory"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Service Category"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "BHT"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBHTServiceBillsD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsSC As New ADODB.Recordset
    Dim temSql As String
    Dim rsSPC As New ADODB.Recordset
    Dim rsStaff() As New ADODB.Recordset
    Dim PSCCount As Long
    Dim rsBHT As New ADODB.Recordset
    Dim MyBHT As New clsBHT
    
Private Sub btnAdd_Click()
   
    Dim n As Integer
    If IsNumeric(cmbBHT.BoundText) = False Then
        MsgBox "BHT?"
        cmbBHT.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbCategory.BoundText) = False Then
        MsgBox "Service?"
        cmbCategory.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtEditID.Text) = True Then
            temSql = "Select * from tblPatientService where PatientServiceID = " & Val(txtEditID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount <= 0 Then
                .AddNew
            End If
        Else
            temSql = "Select * from tblPatientService  where PatientServiceID = 0 "
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
        End If
        !BHTID = Val(cmbBHT.BoundText)
        !ServiceCategoryID = Val(cmbCategory.BoundText)
        !ServicesubcategoryID = Val(cmbSC.BoundText)
        !Comments = txtComments.Text
        !ServiceDate = dtpDate.Value
        !ServiceTime = dtpTime.Value
        !Charge = Val(txtCharge.Text)
        !ProfessionalCharge = Val(txtProfessionalCharge.Text)
        !HospitalCharge = Val(txtHospitalCharge.Text)
        !UserID = UserID
        If chkAddToMedicineCharge.Value = 1 Then
            !AddToMedicineCharge = True
        Else
            !AddToMedicineCharge = False
        End If
        
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        txtEditID.Text = !NewID
        .Close
    End With
    For n = 0 To lblSpeciality1.UBound
        If lblSpeciality1(n).Visible = True Then
            With rsTem
                If .State = 1 Then .Close
                temSql = "Select * from tblProfessionalCharges where ServiceProfessionalChargesID = " & Val(txtServiceProfessionalChargesID(n).Text) & " AND PatientServiceID = " & Val(txtEditID.Text)
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    
                Else
                    .AddNew
                    !UserID = UserID
                    !ForBHTID = Val(cmbBHT.BoundText)
                    !PatientServiceID = Val(txtEditID.Text)
                    !ServiceProfessionalChargesID = Val(txtServiceProfessionalChargesID(n).Text)
                    !StaffID = Val(cmbStaff1(n).BoundText)
                End If
                !Date = dtpDate.Value
                !Time = dtpTime.Value
                !Fee = Val(txtFee1(n).Text)
                .Update
            End With
        End If
    Next n
    Call FillGrid
    Call ClearAddValues
    If Val(cmbCategory.BoundText) = 32 Then
        cmbSC.SetFocus
    Else
        cmbCategory.SetFocus
    End If
End Sub

Private Sub ClearAddValues()
    Dim n As Long
    If Val(cmbCategory.BoundText) <> 32 Then
        cmbCategory.Text = Empty
    End If
    cmbSC.Text = Empty
    txtComments.Text = Empty
    txtProfessionalCharge.Text = Empty
    txtHospitalCharge.Text = Empty
    txtCharge.Text = Empty
    txtEditID.Text = Empty
    txtDelID.Text = Empty
    For n = 0 To lblSpeciality1.UBound
        lblSpeciality1(n).Visible = False
        lblSpeciality1(n).Caption = Empty
        cmbStaff1(n).Visible = False
        cmbStaff1(n).Text = Empty
        txtServiceProfessionalChargesID(n).Text = Empty
        txtFee1(n).Visible = False
        txtFee1(n).Text = Empty
        txtSpecialityID(n).Text = Empty
    Next
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    If Val(txtDelID.Text) = 0 Then
        MsgBox "Please select one to delete"
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtDelID.Text) = True Then
            temSql = "Select * from tblPatientService where PatientServiceID = " & Val(txtDelID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Deleted = True
                !DeletedUserID = UserID
                !DeletedDate = Date
                !DeletedTime = Now
                .Update
            End If
            .Close
        End If
    End With
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblProfessionalCharges where PatientServiceID = " & Val(txtDelID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !Cancelled = True
            !CancelledDate = Date
            !CancelledTime = Now
            !CancelledDateTime = Now
            !CancelledUserID = UserID
            .Update
            .MoveNext
        Wend
        .Close
    End With
    
    Call FillGrid
    Call ClearAddValues
    cmbCategory.SetFocus
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    GridPrint gridService, ThisReportFormat
    Printer.EndDoc
End Sub

'Private Sub chkAddToMedicineCharge_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        btnAdd_Click
'    End If
'End Sub

Private Sub cmbBHT_Change()
    If IsNumeric(cmbBHT.BoundText) = False Then Exit Sub
    MyBHT.BHTID = Val(cmbBHT.BoundText)
    Call FillGrid
    Call DisplayDetails
End Sub

Private Sub DisplayDetails():  On Error Resume Next
    Dim temText As String
    Dim r As Long
    temText = "Patient Name : " & MyBHT.FirstName & vbNewLine
    temText = temText & "Guardian : " & MyBHT.GuardianName & vbNewLine
    temText = temText & "Address : " & MyBHT.PtAddress & vbNewLine
    temText = temText & "BHT : " & MyBHT.BHT & vbNewLine
    temText = temText & "Age : " & MyBHT.AgeInWords & vbNewLine
    temText = temText & "Admitted : " & Format(MyBHT.DOA, "dd MMMM yyyy") & " at " & Format(MyBHT.TOA, "HH:MM AMPM") & vbNewLine
    If MyBHT.Discharge = True Then
        temText = temText & "Discharged :" & Format(MyBHT.DOD, "dd MMMM yyyy") & " at " & Format(MyBHT.TOD, "HH:MM AMPM") & vbNewLine
    Else
        temText = temText & "Not yet discharged" & vbNewLine
    End If
    temText = temText & "Payment Method : " & MyBHT.PaymentMethod
    If MyBHT.HealthSchemeSupplier <> "" Then
        temText = temText & " (" & MyBHT.HealthSchemeSupplier & ")" & vbNewLine
    Else
        temText = temText & vbNewLine
    End If
    If MyBHT.Comments <> "" Then
        temText = temText & MyBHT.Comments & vbNewLine
    End If
    
    txtDetails.Text = temText
End Sub


Private Sub cmbBHT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbCategory.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbBHT.Text = Empty
    End If
End Sub

Private Sub cmbCategory_Change()

    txtSurcharge.Text = 0
    
    If IsNumeric(cmbCategory.BoundText) = False Then Exit Sub
    
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceCategory where ServiceCategoryID = " & Val(cmbCategory.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtSurcharge.Text = Format(!InwardSurcharge, "0")
            If MyBHT.Foreigner = True Then
                txtHospitalCharge.Text = Format(!Fee * 2 * ((Val(txtSurcharge.Text) + 100) / 100), "0.00")
            Else
                txtHospitalCharge.Text = Format(!Fee * ((Val(txtSurcharge.Text) + 100) / 100), "0.00")
            End If
            If !CanChange = True Then
                txtHospitalCharge.Locked = False
            Else
                txtHospitalCharge.Locked = True
            End If
            If !ToMedicineCharge = True Then
                chkAddToMedicineCharge.Value = 1
            Else
                chkAddToMedicineCharge.Value = 0
            End If
        End If
        .Close
    End With
    
    With rsSC
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSubCategory where   Deleted = 0 AND ForBHT = 1 AND ServiceCategoryID = " & Val(cmbCategory.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbSC.Visible = True
            lblSC.Visible = True
        Else
            cmbSC.Visible = False
            lblSC.Visible = False
        End If
    End With
    With cmbSC
        Set .RowSource = rsSC
        .ListField = "ServiceSubcategory"
        .BoundColumn = "ServiceSubcategoryID"
        .Text = Empty
    End With
    
    txtCharge.Text = Format(Val(txtHospitalCharge.Text) + Val(txtProfessionalCharge.Text), "0.00")
    
End Sub

Private Sub ClearServiceValues()
    Dim n As Integer
    txtComments.Text = Empty
    txtProfessionalCharge.Text = Empty
    txtHospitalCharge.Text = Empty
    txtCharge.Text = Empty
    txtEditID.Text = Empty
    txtDelID.Text = Empty
    For n = 0 To lblSpeciality1.UBound
        lblSpeciality1(n).Visible = False
        lblSpeciality1(n).Caption = Empty
        cmbStaff1(n).Visible = False
        cmbStaff1(n).Text = Empty
        txtServiceProfessionalChargesID(n).Text = Empty
        txtFee1(n).Visible = False
        txtFee1(n).Text = Empty
        txtSpecialityID(n).Text = Empty
    Next

End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If cmbSC.Visible = False Then
            txtComments.SetFocus
            SendKeys "{home}+{end}"
        Else
            cmbSC.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub cmbSC_Change()
    Call ClearServiceValues
    
    If IsNumeric(cmbCategory.BoundText) = False Then Exit Sub
    If IsNumeric(cmbSC.BoundText) = False Then Exit Sub
    If cmbSC.Visible = False Then Exit Sub
    
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSubCategory where ServiceSubCategoryID = " & Val(cmbSC.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If MyBHT.Foreigner = True Then
                txtHospitalCharge.Text = Format(!Fee * 2 * ((Val(txtSurcharge.Text) + 100) / 100), "0.00")
            Else
                txtHospitalCharge.Text = Format(!Fee * ((Val(txtSurcharge.Text) + 100) / 100), "0.00")
            End If

            If !CanChange = True Then
                txtHospitalCharge.Locked = False
            Else
                txtHospitalCharge.Locked = True
            End If
            If !ToMedicineCharge = True Then
                chkAddToMedicineCharge.Value = 1
            Else
                chkAddToMedicineCharge.Value = 0
            End If
            
        End If
        .Close
    End With
    
    Dim n As Integer
    
    With rsSPC
        If .State = 1 Then .Close
        temSql = "SELECT Top 7 tblSpeciality.Speciality, tblSpeciality.SpecialityID, tblServiceProfessionalCharges.Fee,  tblServiceProfessionalCharges.StaffID, tblServiceProfessionalCharges.ServiceProfessionalChargesID " & _
                    "FROM tblSpeciality RIGHT JOIN tblServiceProfessionalCharges ON tblSpeciality.SpecialityID = tblServiceProfessionalCharges.SpecialityID " & _
                    "Where (((tblServiceProfessionalCharges.ServiceSubcategoryID) = " & Val(cmbSC.BoundText) & ") AND ((tblServiceProfessionalCharges.Deleted)=0 ))" & _
                    "ORDER BY tblServiceProfessionalCharges.ServiceProfessionalChargesID DESC"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        PSCCount = .RecordCount
        ReDim rsStaff(.RecordCount)
        For n = 0 To PSCCount - 1
            lblSpeciality1(n).Visible = True
            lblSpeciality1(n).Caption = !Speciality
            txtServiceProfessionalChargesID(n).Text = !ServiceProfessionalChargesID
            txtSpecialityID(n).Text = !SpecialityID
            cmbStaff1(n).Visible = True
            If rsStaff(n).State = 1 Then rsStaff(n).Close
            temSql = "SELECT tblStaff.Name as TitleStaff, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID Where SpecialityID = " & !SpecialityID & " ORDER BY tblTitle.Title, tblStaff.Name"
            rsStaff(n).Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            Set cmbStaff1(n).RowSource = rsStaff(n)
            cmbStaff1(n).ListField = "TitleStaff"
            cmbStaff1(n).BoundColumn = "StaffID"
            cmbStaff1(n).BoundText = !StaffID
        
            txtFee1(n).Visible = True
            If MyBHT.Foreigner = True Then
                txtFee1(n).Text = Format(!Fee * 2, "0.00")
            Else
                txtFee1(n).Text = Format(!Fee, "0.00")
            End If
            .MoveNext
            
        Next
        If PSCCount = 0 Then
            For n = 0 To lblSpeciality1.UBound
                lblSpeciality1(n).Visible = False
                lblSpeciality1(n).Caption = Empty
                cmbStaff1(n).Visible = False
                cmbStaff1(n).Text = Empty
                txtServiceProfessionalChargesID(n).Text = Empty
                txtFee1(n).Visible = False
                txtFee1(n).Text = Empty
                txtSpecialityID(n).Text = Empty
            Next
        Else
            For n = PSCCount To lblSpeciality1.UBound
                lblSpeciality1(n).Visible = False
                lblSpeciality1(n).Caption = Empty
                cmbStaff1(n).Visible = False
                cmbStaff1(n).Text = Empty
                txtServiceProfessionalChargesID(n).Text = Empty
                txtFee1(n).Visible = False
                txtFee1(n).Text = Empty
                txtSpecialityID(n).Text = Empty
            Next
        End If
    End With
    
    txtCharge.Text = Format(Val(txtHospitalCharge.Text) + Val(txtProfessionalCharge.Text), "0.00")
    
    
End Sub

Private Sub cmbSC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtComments.SetFocus
        SendKeys "{home}+{end}"
    ElseIf KeyCode = vbKeyEscape Then
        cmbSC.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call GetSettings
    Call FillCombos
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    dtpTime.Value = Time
    GetCommonSettings Me
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    'Cat.FillBoolCombo cmbCategory, "ServiceCategory", "ServiceCategory", "ForOPD", True
    
    Cat.FillBoolCombo cmbCategory, "ServiceCategory", "ServiceCategory", "ForBHT", True
    
    With rsBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where IsBHT = 1 And Discharge = 1 order by BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
End Sub

Private Sub FillGrid()
    Call FormatGrid
    Dim rsTem As New ADODB.Recordset
    Dim TotalCharge As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientService.PatientServiceID, tblPatientService.ServiceDate, tblServiceCategory.ServiceCategory, tblServiceSubcategory.ServiceSubcategory, tblPatientService.Comments, tblPatientService.Charge, tblPatientService.UserID " & _
                    "FROM (tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID " & _
                    "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.BHTID)=" & Val(cmbBHT.BoundText) & ")) " & _
                    "ORDER BY tblPatientService.PatientServiceID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridService.Rows = gridService.Rows + 1
            gridService.Row = gridService.Rows - 1
            gridService.Col = 0
            gridService.Text = !PatientServiceID
            gridService.Col = 1
            gridService.Text = !ServiceDate
            gridService.Col = 2
            If IsNull(!ServiceSubcategory) = True Then
                gridService.Text = !ServiceCategory
            Else
                gridService.Text = !ServiceCategory & " - " & !ServiceSubcategory
            End If
            gridService.Col = 3
            gridService.Text = !Comments
            gridService.Col = 4
            gridService.Text = Format(!Charge, "0.00")
            gridService.Col = 5
            gridService.Text = FullStaffName(!UserID)
            
            TotalCharge = TotalCharge + !Charge
            .MoveNext
        Wend
    End With
    lblTotal.Caption = Format(TotalCharge, "0.00")
End Sub

Private Sub FormatGrid()
    '   0   ID
    '   1   Date
    '   2   Service
    '   3   Comments
    '   4   Charges
    With gridService
        .Cols = 6
        .Rows = 1
        .ColWidth(0) = 0
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        
        .Col = 1
        .Text = "Date"
        
        .Col = 2
        .Text = "Service"
        
        .Col = 3
        .Text = "Comments "
        
        .Col = 4
        .Text = "Charge"
        
        .Col = 5
        .Text = "User"
        
    End With
    lblTotal.Caption = "0.00"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    saveSettings
End Sub

Private Sub saveSettings()
    SaveCommonSettings Me
End Sub

Private Sub gridService_Click()
    With gridService
        txtDelID.Text = Val(.TextMatrix(.Row, 0))
        .Col = .Cols - 1
        .ColSel = 0
    End With
End Sub

Private Sub gridService_DblClick()
    Dim rsTem As New ADODB.Recordset
    With gridService
        txtEditID.Text = Val(.TextMatrix(.Row, 0))
        .Col = .Cols - 1
        .ColSel = 0
    End With
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtEditID.Text) = True Then
            temSql = "Select * from tblPatientService where PatientServiceID = " & Val(txtEditID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                cmbCategory.BoundText = !ServiceCategoryID
                cmbSC.BoundText = !ServicesubcategoryID
                txtComments.Text = !Comments
                dtpDate.Value = !ServiceDate
                dtpTime.Value = !ServiceTime
                txtCharge.Text = Format(!Charge, "0.00")
                txtHospitalCharge.Text = Format(!HospitalCharge, "0.00")
                txtProfessionalCharge.Text = Format(!ProfessionalCharge, "0.00")
                If !AddToMedicineCharge = True Then
                    chkAddToMedicineCharge.Value = 1
                Else
                    chkAddToMedicineCharge.Value = 0
                End If
                
            End If
            .Close
        End If
    End With
    Dim n As Integer
    For n = 0 To lblSpeciality1.UBound
        If lblSpeciality1(n).Visible = True Then
            With rsTem
                If .State = 1 Then .Close
                temSql = "Select * from tblProfessionalCharges where ServiceProfessionalChargesID = " & Val(txtServiceProfessionalChargesID(n).Text) & " AND PatientServiceID = " & Val(txtEditID.Text)
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    cmbStaff1(n).BoundText = !StaffID
                    txtFee1(n).Text = Format(!Fee, "0.00")
                End If
                .Close
            End With
        Else
            txtFee1(n).Text = 0
        End If
    Next n
End Sub

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtHospitalCharge.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        txtComments.Text = Empty
    End If
End Sub

Private Sub txtFee1_Change(Index As Integer)
    Dim n As Long
    Dim temTotal As Double
    For n = 0 To txtFee1.UBound
        temTotal = temTotal + Val(txtFee1(n).Text)
    Next
    txtProfessionalCharge.Text = Format(temTotal, "0.00")
End Sub

Private Sub txtHospitalCharge_Change()
    txtCharge.Text = Format(Val(txtHospitalCharge.Text) + Val(txtProfessionalCharge.Text), "0.00")
End Sub

Private Sub txtHospitalCharge_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtHospitalCharge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        chkAddToMedicineCharge.SetFocus
    End If
End Sub

Private Sub txtProfessionalCharge_Change()
    txtCharge.Text = Format(Val(txtHospitalCharge.Text) + Val(txtProfessionalCharge.Text), "0.00")
End Sub
