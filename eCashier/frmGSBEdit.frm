VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGSBEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit GSB Details"
   ClientHeight    =   8685
   ClientLeft      =   1890
   ClientTop       =   -2445
   ClientWidth     =   15000
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
   ScaleHeight     =   8685
   ScaleWidth      =   15000
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   495
      Left            =   11520
      TabIndex        =   48
      Tag             =   "Edit"
      Top             =   7200
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   12000
      TabIndex        =   47
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.OptionButton optCurrent 
      Caption         =   "&Current Patients"
      Height          =   240
      Left            =   120
      TabIndex        =   46
      Tag             =   "Select"
      Top             =   7920
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.OptionButton optAll 
      Caption         =   "&All Patients"
      Height          =   240
      Left            =   120
      TabIndex        =   45
      Tag             =   "Select"
      Top             =   8280
      Width           =   3015
   End
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   375
      Left            =   1680
      TabIndex        =   44
      Tag             =   "Select"
      Top             =   7440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Edit"
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
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   6900
      Left            =   120
      TabIndex        =   42
      Tag             =   "Select"
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   12171
      _Version        =   393216
      Style           =   1
      Text            =   ""
   End
   Begin VB.CheckBox chkForeigner 
      Caption         =   "&Foreigner"
      Height          =   240
      Left            =   10560
      TabIndex        =   37
      Tag             =   "Edit"
      Top             =   1320
      Width           =   1215
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   13440
      TabIndex        =   25
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
   Begin btButtonEx.ButtonEx bttnAdmit 
      Height          =   495
      Left            =   10080
      TabIndex        =   24
      Tag             =   "Edit"
      Top             =   7200
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
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   3720
      TabIndex        =   35
      Top             =   7080
      Width           =   6015
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   1920
         TabIndex        =   29
         Top             =   720
         Width           =   3975
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   1920
         TabIndex        =   27
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label18 
         Caption         =   "Paper"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8775
      Left            =   3600
      TabIndex        =   32
      Top             =   0
      Width           =   11655
      Begin VB.TextBox txtPtNIC 
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
         Left            =   2400
         TabIndex        =   39
         Tag             =   "Edit"
         Top             =   3000
         Width           =   3855
      End
      Begin VB.TextBox txtPtPhone 
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
         Left            =   7680
         TabIndex        =   38
         Tag             =   "Edit"
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox txtPaymentMethodID 
         Height          =   375
         Left            =   10800
         TabIndex        =   36
         Top             =   4440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtName 
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
         Left            =   2400
         TabIndex        =   5
         Tag             =   "Edit"
         Top             =   1560
         Width           =   8655
      End
      Begin VB.TextBox txtAddress 
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
         Left            =   2400
         TabIndex        =   7
         Tag             =   "Edit"
         Top             =   2040
         Width           =   8655
      End
      Begin VB.TextBox txtAge 
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
         Left            =   2400
         TabIndex        =   9
         Tag             =   "Edit"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtGunardian 
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
         Left            =   2400
         TabIndex        =   14
         Tag             =   "Edit"
         Top             =   3480
         Width           =   8655
      End
      Begin VB.TextBox txtInitialDeposit 
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
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Tag             =   "Edit"
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox txtComments 
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
         Left            =   7680
         TabIndex        =   20
         Tag             =   "Edit"
         Top             =   3960
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtpDOA 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   78053379
         CurrentDate     =   39956
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Tag             =   "Edit"
         Top             =   2520
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   78053379
         CurrentDate     =   39589
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
         Left            =   7560
         TabIndex        =   3
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hour MIN sec"
         Format          =   78053378
         CurrentDate     =   39589
      End
      Begin MSDataListLib.DataCombo cmbSex 
         Height          =   360
         Left            =   7680
         TabIndex        =   12
         Tag             =   "Edit"
         Top             =   2520
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPtCat 
         Height          =   360
         Left            =   2400
         TabIndex        =   16
         Tag             =   "Edit"
         Top             =   3960
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbSpeciality 
         Height          =   360
         Left            =   2400
         TabIndex        =   22
         Tag             =   "Edit"
         Top             =   4920
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbStaff 
         Height          =   360
         Left            =   2400
         TabIndex        =   23
         Tag             =   "Edit"
         Top             =   5400
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label22 
         Caption         =   "Patient Phone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   41
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "Patient NIC No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblSubtopic 
         Alignment       =   2  'Center
         Caption         =   "TOPIC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   11175
      End
      Begin VB.Label lblTopic 
         Alignment       =   2  'Center
         Caption         =   "TOPIC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   11175
      End
      Begin VB.Label Label2 
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Patient Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Date of Admission"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "Sex"
         Height          =   255
         Left            =   6360
         TabIndex        =   11
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Name of Guardian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Patient Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Hospital Comments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label16 
         Caption         =   "Referring Doctor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   19
         Top             =   3960
         Width           =   2895
      End
   End
   Begin VB.TextBox txtComSurcharge 
      Height          =   360
      Left            =   960
      TabIndex        =   31
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtPtSurcharge 
      Height          =   360
      Left            =   480
      TabIndex        =   30
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label21 
      Caption         =   "BHT"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmGSBEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsPatients As New ADODB.Recordset
    Dim rsHSS As New ADODB.Recordset
    Dim rsBHT As New ADODB.Recordset
    Dim rsViewBHT As New ADODB.Recordset
    Dim rsRoom As New ADODB.Recordset
    Dim rsTemRoom As New ADODB.Recordset
    Dim rsRoomPatient As New ADODB.Recordset
    Dim temSql As String
    Dim temPatientID As Long
    Dim temBHTID As Long
    Dim PCat As New clsPatientCategory
    Dim FirstActivation As Boolean
    Dim rsStaff As New ADODB.Recordset
    
    Dim CsetPrinter As New cSetDfltPrinter
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    Dim SuppliedWord As String
    
    Dim FSys As New Scripting.FileSystemObject


Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbBHT.SetFocus
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbBHT.BoundText) = False Then
        MsgBox "Please select a BHT to Edit"
        cmbBHT.SetFocus
        Exit Sub
    End If
    Call EditMode
End Sub

Private Sub EditMode()
    Dim MyCtrl As Control
    For Each MyCtrl In Controls
        If MyCtrl.Tag = "Edit" Then
            MyCtrl.Enabled = True
        ElseIf MyCtrl.Tag = "Select" Then
            MyCtrl.Enabled = False
        End If
    Next
End Sub

Private Sub SelectMode()
    Dim MyCtrl As Control
    For Each MyCtrl In Controls
        If MyCtrl.Tag = "Edit" Then
            MyCtrl.Enabled = False
        ElseIf MyCtrl.Tag = "Select" Then
            MyCtrl.Enabled = True
        End If
    Next
End Sub

Private Sub btnPrint_Click()
    Call PrintAdmissionBill
End Sub

Private Sub bttnAdmit_Click()
    If CanAdmit = False Then
        Exit Sub
    End If
    With rsBHT
    If .State = 1 Then .Close
        temSql = "SELECT * FROM tblBHT Where BHTID = " & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then Exit Sub
        !IsBHT = True
        temPatientID = !PatientID
        '!DOA = dtpDOA.Value
        '!Discharge = False
        '!TOA = Format(dtpDOA.Value, "hh:mm:ss") & " " & dtpTOA.Value
'        !RoomID = Val(cmbRoom.BoundText)
'        !HealthSchemeSupplierID = Val(cmbHealthSchemeSupplier.BoundText)
        !PtSurcharge = Val(txtPtSurcharge.Text)
        !ComSurcharge = Val(txtComSurcharge.Text)
        !PatientCategoryID = Val(cmbPtCat.BoundText)
        !Comments = txtComments.Text
        !ReferringDoctorID = Val(cmbStaff.BoundText)
        !GuardianName = txtGunardian.Text
        !TemAge = CalculateAgeInWords(dtpDOB.Value)
        
'        !GuardianAddress = txtGuardianAddress.Text
'        !GuardianNIC = txtGunadianNIC.Text
'        !GuardianPhone = txtGuardianPhone.Text
        !AdStaffID = UserID
        !PaymentMethodID = Val(txtPaymentMethodID.Text)
        If chkForeigner.Value = 1 Then
            !Foreigner = True
        Else
            !Foreigner = False
        End If
        !HospitalComments = txtInitialDeposit.Text
        .Update
        .Close
    End With
    With rsPatients
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPatientMainDetails Where PatientID = " & temPatientID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then Exit Sub
        !FirstName = UCase(txtName.Text)
        !Address = txtAddress.Text
        !DateOfBirth = dtpDOB.Value
        !SexID = Val(cmbSex.BoundText)
        !NICNo = txtPtNIC.Text
        If Trim(txtPtPhone.Text) <> "" Then
            !Phone = txtPtPhone.Text
        End If
        .Update
        .Close
    End With
    Call SelectMode
    cmbBHT.SetFocus
End Sub

Private Sub PrintAdmissionBill()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim PrintText As String
    On Error Resume Next
    Dim MyFontSize As Long
    MyFontSize = 12
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        Dim MyControl As Control
        For Each MyControl In Controls
            If MyControl.Container.Name = Frame1.Name And MyControl.Visible = True Then
                Printer.Font = DefaultFont.Name    'MyControl.Font.Name
                If MyFontSize <> 0 Then
                    Printer.Font.Size = MyControl.Font.Size
                End If
                Printer.Font.Bold = MyControl.Font.Bold
                Printer.Font.Italic = MyControl.Font.Italic
                Printer.Font.Underline = MyControl.Font.Underline
                PrintText = MyControl.Caption
                PrintText = MyControl.Text
                PrintText = MyControl.Value
                
                If IsDate(PrintText) = True Then
                    If Abs(Year(PrintText) - Year(Date)) < 4 Then
                        PrintText = Format(PrintText, "dd MMMM yyyy")
                    End If
                End If
                
                If MyControl.Alignment = 0 Then
                    Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
                ElseIf MyControl.Alignment = 1 Then
                    Printer.CurrentX = (MyControl.Left / Frame1.Width * Printer.Width) + (MyControl.Width / Frame1.Width * Printer.Width) - Printer.TextWidth(PrintText)
                ElseIf MyControl.Alignment = 2 Then
                    Printer.CurrentX = (MyControl.Left / Frame1.Width * Printer.Width) + ((MyControl.Width / Frame1.Width * Printer.Width) / 2) - (Printer.TextWidth(PrintText) / 2)
                Else
                    Printer.CurrentX = MyControl.Left / Frame1.Width * Printer.Width
                End If
                Printer.CurrentY = MyControl.Top / Frame1.Width * Printer.Height
                Printer.Print PrintText
        End If
        Next
        Printer.EndDoc
    End If
    
End Sub


Private Sub PrintInitialDepositBIll()

End Sub

Private Sub bttnClose_Click()
'    Call PrintAdmissionBill
    Unload Me
End Sub

Private Sub cmbBHT_Click(Area As Integer)
    Call ClearValues
    Call DisplayDetails
End Sub

Private Sub ClearValues()
    Dim MyCtrl As Control
    For Each MyCtrl In Controls
        If TypeOf MyCtrl Is TextBox Then
            MyCtrl.Text = Empty
        ElseIf TypeOf MyCtrl Is DataCombo Then
            If MyCtrl.Tag <> "Select" Then
                MyCtrl.Text = Empty
            End If
        ElseIf TypeOf MyCtrl Is CheckBox Then
            MyCtrl.Value = 0
        ElseIf TypeOf MyCtrl Is DTPicker Then
        
        End If
    Next
End Sub

Private Sub DisplayDetails(): On Error Resume Next
    With rsBHT
    If .State = 1 Then .Close
        temSql = "SELECT * FROM tblBHT where BHTID = " & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount <= 0 Then Exit Sub
'        txtBHT.Text = !BHT
        temPatientID = !PatientID
        dtpDOA.Value = !DOA
        dtpTOA.Value = Format(!TOA, "hh:mm:ss")
        'cmbRoom.BoundText = !RoomID
        cmbPtCat.BoundText = !PatientCategoryID
        txtComments.Text = Format(!Comments, "")
        cmbStaff.BoundText = !ReferringDoctorID
        txtGunardian.Text = !GuardianName
        
'        txtGuardianAddress.Text = Format(!GuardianAddress, "")
'        txtGunadianNIC.Text = !GuardianNIC
'        txtGuardianPhone.Text = !GuardianPhone
'        txtPaymentMethodID.Text = !PaymentMethodID
        If !Foreigner = True Then
            chkForeigner.Value = 1
        Else
            chkForeigner.Value = 0
        End If
        txtInitialDeposit.Text = Format(!HospitalComments, "")
'        cmbHealthSchemeSupplier.BoundText = !HealthSchemeSupplierID
        temBHTID = !BHTID
        .Close
    End With
    With rsPatients
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPatientMainDetails where PatientID = " & temPatientID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount <= 0 Then Exit Sub
        txtName.Text = !FirstName
        txtAddress.Text = !Address
        dtpDOB.Value = !DateOfBirth
        cmbSex.BoundText = !SexID
        txtPtNIC.Text = Format(!NICNo, "")
        txtPtPhone.Text = Format(!Phone, "")
        .Close
    End With
End Sub

'Private Sub cmbHealthSchemeSupplier_Change()
'    If IsNumeric(cmbHealthSchemeSupplier.BoundText) = False Then
'        txtComSurcharge.Text = 0
'    Else
'        Dim rsTem As New ADODB.Recordset
'        With rsTem
'            If .State = 1 Then .Close
'            temSql = "Select * from tblHealthSchemeSuppliers where HealthSchemeSupplierID = " & Val(cmbHealthSchemeSupplier.BoundText)
'            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'            If .RecordCount > 0 Then
'                If IsNull(!InwardAddition) = False Then
'                    txtComSurcharge.Text = !InwardAddition
'                Else
'                    txtComSurcharge.Text = 0
'                End If
'            Else
'                txtComSurcharge.Text = 0
'            End If
'            .Close
'        End With
'    End If
'End Sub
'


Private Sub cmbPtCat_Change()
    If IsNumeric(cmbPtCat.BoundText) = False Then
        txtPtSurcharge.Text = 0
        Exit Sub
    End If
    PCat.ID = Val(cmbPtCat.BoundText)
    txtPaymentMethodID.Text = PCat.PaymentMethodID
'    If LCase(PCat.PaymentMethod) = "credit" Then
'        cmbHealthSchemeSupplier.Visible = True
'        lblHSS.Visible = True
'    Else
'        cmbHealthSchemeSupplier.Visible = False
'        lblHSS.Visible = False
'    End If
    txtPtSurcharge.Text = PCat.Surcharge
'    cmbHealthSchemeSupplier.Text = Empty
    txtComments.Text = Empty
End Sub

Private Sub cmbPtCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbPtCat.Text = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
'        If cmbHealthSchemeSupplier.Visible = True Then
'            cmbHealthSchemeSupplier.SetFocus
'        Else
            txtInitialDeposit.SetFocus
'        End If
    End If
End Sub

Private Sub cmbRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpDOA.SetFocus
    End If
End Sub

Private Sub cmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPtNIC.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbSex.Text = Empty
    End If
End Sub

Private Sub cmbSpeciality_Change()
    With rsStaff
        If .State = 1 Then .Close
        If IsNumeric(cmbSpeciality.BoundText) = True Then
            temSql = "SELECT tblStaff.Name AS NameWithTitle, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID where SpecialityID = " & Val(cmbSpeciality.BoundText) & " ORDER BY Name"
        Else
            temSql = "SELECT tblStaff.Name AS NameWithTitle, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID Order BY Name"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbStaff
        Set .RowSource = rsStaff
        .ListField = "NameWithTitle"
        .BoundColumn = "StaffID"
        .Text = Empty
    End With

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

Private Sub dtpTOA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtpTOA.Value = Date
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtName.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
    Call GetSettings
    Call SelectMode
End Sub

Private Sub GetSettings(): On Error Resume Next
    cmbSpeciality.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbSpeciality.Name, "1"))
    cmbPtCat.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbPtCat.Name, "1"))
    
'    txtBHT.Text = NewBHT
    dtpDOB.Value = Date
    dtpDOA.Value = Date
    dtpTOA.Value = Time
    lblTopic.Caption = HospitalName
    lblSubtopic.Caption = HospitalDescreption
    On Error Resume Next
'    chkPrint.Value = GetSetting(App.EXEName, Me.Name, chkPrint.Name, "1")
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "")
End Sub

Private Function NewBHT() As Long
    Dim rsTemBHT As New ADODB.Recordset
    With rsTemBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT Where IsBHT = 1 order by BHTID DESC"
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
'    With rsRoom
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblRoom order by Room"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With cmbRoom
'        Set .RowSource = rsRoom
'        .ListField = "Room"
'        .BoundColumn = "RoomID"
'    End With
'    With rsHSS
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblHealthSchemeSuppliers order by HealthSchemeSupplierName"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'    End With
'    With cmbHealthSchemeSupplier
'        Set .RowSource = rsHSS
'        .ListField = "HealthSchemeSupplierName"
'        .BoundColumn = "HealthSchemeSupplierID"
'    End With
    Dim Speciality As New clsFillCombos
    Speciality.FillAnyCombo cmbSpeciality, "Speciality", False
    With rsViewBHT
        If .State = 1 Then .Close
        If optAll.Value = True Then
            temSql = "Select * from tblBHT where IsGSB = 1 order by BHT"
        Else
            temSql = "Select * from tblBHT where IsGSB = 1 And Discharge = 0 order by BHT"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbBHT
        Set .RowSource = rsViewBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
    
End Sub

Private Function CanAdmit() As Boolean
    CanAdmit = False
    Dim tr As Integer
    If Trim(txtName.Text) = Empty Then
        tr = MsgBox("You have not entered the name of the patient", vbCritical, "Name?")
        txtName.SetFocus
        Exit Function
    End If
'    If Trim(txtBHT.Text) = Empty Then
'        tr = MsgBox("You have not entered the BHT number", vbCritical, "BHT?")
'        txtBHT.SetFocus
'        Exit Function
'    End If
'    If IsNumeric(cmbRoom.BoundText) = False Then
'        tr = MsgBox("Please select a room", vbCritical, "Room?")
'        cmbRoom.SetFocus
'        Exit Function
'    End If
    If IsNumeric(cmbPtCat.BoundText) = False Then
        tr = MsgBox("Please select a patient category", vbCritical, "Patient category?")
        cmbPtCat.SetFocus
        Exit Function
    Else
        PCat.ID = Val(cmbPtCat.BoundText)
    End If
'    If LCase(PCat.PaymentMethod) = "cash" And Val(txtInitialDeposit.Text) <= 0 Then
'        tr = MsgBox("Please enter the initial payment for this cash patient", vbCritical, "Initial Payment")
'        txtInitialDeposit.SetFocus
'        Exit Function
'    End If
    If LCase(PCat.PaymentMethod) = "credit" Then ' And IsNumeric(cmbHealthSchemeSupplier.BoundText) = False Then
        tr = MsgBox("No Credit", vbCritical, "No Credit?")
        Exit Function
    End If
    CanAdmit = True
End Function

Private Sub dtpDOA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtpDOA.Value = Date
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpTOA.SetFocus
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbSpeciality.Name, cmbSpeciality.BoundText
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
End Sub

Private Sub optAll_Click()
    Call FillCombos
End Sub

Private Sub optCurrent_Click()
    Call FillCombos
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

Private Sub txtAddress_LostFocus()
    txtAddress.Text = UCase(txtAddress.Text)
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
'    If KeyCode = vbKeyEscape Then
'        txtBHT.Text = Empty
'        KeyCode = Empty
'    ElseIf KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        cmbRoom.SetFocus
'    End If
End Sub

Private Sub txtInitialDeposit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbSpeciality.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtInitialDeposit.Text = Empty
    End If
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

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        bttnAdmit.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtComments.Text = Empty
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


Private Sub txtName_LostFocus()
    txtName.Text = UCase(txtName.Text)
End Sub



Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub PopulatePapers(): On Error Resume Next
    cmbPaper.Clear
    SetPrinter = False
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
'        With FormSize
'            .cx = BillPaperHeight
'            .cy = BillPaperWidth
'        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                cmbPaper.AddItem PtrCtoVbString(.pName)
'                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle): DoEvents
    End If
End Sub

Private Sub cmbPrinter_Change()
    Call PopulatePapers
End Sub

Private Sub cmbPrinter_Click()
    Call PopulatePapers
End Sub

Private Sub txtPtNIC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPtPhone.SetFocus
    End If
End Sub

Private Sub txtPtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtGunardian.SetFocus
    End If
End Sub
