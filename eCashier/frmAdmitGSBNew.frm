VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdmitGSBNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admit GSB Patients"
   ClientHeight    =   9150
   ClientLeft      =   1890
   ClientTop       =   -2445
   ClientWidth     =   12000
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
   ScaleHeight     =   9150
   ScaleWidth      =   12000
   Begin VB.CheckBox chkForeigner 
      Caption         =   "&Foreigner"
      Height          =   240
      Left            =   7800
      TabIndex        =   40
      Top             =   5280
      Width           =   1215
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9720
      TabIndex        =   27
      Top             =   8160
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
      Left            =   8280
      TabIndex        =   26
      Top             =   8160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Admit"
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
   Begin VB.CheckBox chkPrint 
      Caption         =   "&Print"
      Height          =   255
      Left            =   8280
      TabIndex        =   28
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   360
      TabIndex        =   38
      Top             =   6840
      Width           =   6255
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   2160
         TabIndex        =   32
         Top             =   720
         Width           =   3975
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   2160
         TabIndex        =   30
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label18 
         Caption         =   "Paper"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Printer"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   10095
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   11655
      Begin VB.TextBox txtIncomeBillID 
         Height          =   375
         Left            =   4680
         TabIndex        =   45
         Top             =   5160
         Visible         =   0   'False
         Width           =   1095
      End
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
         TabIndex        =   42
         Top             =   3720
         Width           =   3255
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
         TabIndex        =   41
         Top             =   3720
         Width           =   3375
      End
      Begin VB.TextBox txtPaymentMethodID 
         Height          =   375
         Left            =   10800
         TabIndex        =   39
         Top             =   7200
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
         TabIndex        =   7
         Top             =   2280
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
         TabIndex        =   9
         Top             =   2760
         Width           =   8655
      End
      Begin VB.TextBox txtBHT 
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
         TabIndex        =   1
         Top             =   1320
         Width           =   3495
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
         TabIndex        =   11
         Top             =   3240
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
         TabIndex        =   16
         Top             =   4200
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
         TabIndex        =   20
         Top             =   5160
         Width           =   1455
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
         TabIndex        =   22
         Top             =   4680
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtpDOA 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   1800
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
         Format          =   75038723
         CurrentDate     =   39956
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   3240
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
         Format          =   75038723
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
         TabIndex        =   5
         Top             =   1800
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
         Format          =   75038722
         CurrentDate     =   39589
      End
      Begin MSDataListLib.DataCombo cmbSex 
         Height          =   360
         Left            =   7680
         TabIndex        =   14
         Top             =   3240
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
         TabIndex        =   18
         Top             =   4680
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
         TabIndex        =   24
         Top             =   5640
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
         TabIndex        =   25
         Top             =   6120
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
      Begin VB.Label lblBillNo 
         Caption         =   "Bill No."
         Height          =   255
         Left            =   3960
         TabIndex        =   46
         Top             =   5160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Phone"
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
         TabIndex        =   44
         Top             =   3720
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
         TabIndex        =   43
         Top             =   3720
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   10
         Top             =   3240
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
         TabIndex        =   8
         Top             =   2760
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
         TabIndex        =   6
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Green Sheet Bill No"
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
         Top             =   1320
         Width           =   1935
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
         TabIndex        =   2
         Top             =   1800
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
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label11 
         Caption         =   "Sex"
         Height          =   255
         Left            =   6360
         TabIndex        =   13
         Top             =   3240
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
         TabIndex        =   15
         Top             =   4200
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
         TabIndex        =   17
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Initial Deposit"
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
         TabIndex        =   19
         Top             =   5160
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
         TabIndex        =   23
         Top             =   5640
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
         TabIndex        =   21
         Top             =   4680
         Width           =   2895
      End
   End
   Begin VB.TextBox txtComSurcharge 
      Height          =   360
      Left            =   960
      TabIndex        =   34
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtPtSurcharge 
      Height          =   360
      Left            =   480
      TabIndex        =   33
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmAdmitGSBNew"
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
    Dim temSQL As String
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


Private Sub bttnAdmit_Click()
    On Error Resume Next
    If CanAdmit = False Then
        Exit Sub
    End If
    With rsBHT
    If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblBHT where IsGSB = 1 AND BHT = '" & Trim(txtBHT.Text) & "'"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            MsgBox "The GSB Number " & txtBHT.Text & " already exists" & vbNewLine & "Please enter another GSB number"
            .Close
            txtBHT.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        .Close
    End With
    With rsPatients
        If .State = 1 Then .Close
        temSQL = "Select * from tblPatientMainDetails where PatientID = 0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        
        !FirstName = UCase(txtName.Text)
        !Address = txtAddress.Text
        !DateOfBirth = dtpDOB.Value
        !SexID = Val(cmbSex.BoundText)
        !NICNo = txtPtNIC.Text
        !Phone = txtPtPhone.Text
        .Update
        temSQL = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        temPatientID = !NewID
        .Close
    End With
    With rsBHT
    If .State = 1 Then .Close
        temSQL = "SELECT * FROM tblBHT where BHTID = 0  "
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !IsGSB = True
        !BHT = txtBHT.Text
        !PatientID = temPatientID
        !DOA = dtpDOA.Value
        !TOA = Format(dtpDOA.Value, "dd MMMM yyyy") & " " & dtpTOA.Value
        !Discharge = False
        !TOA = Format(dtpDOA.Value, "hh:mm:ss") & " " & dtpTOA.Value
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
        temSQL = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        temBHTID = !NewID
        .Close
    End With
    Dim DisplayBillID As Long
    With rsBHT
    
    
'        If .State = 1 Then .Close
'        temSQL = "Select Count(IncomeBillID) as BillCount from tblIncomeBill where IsGSBill = 1 and Completed = 1"
'        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            If IsNull(!BillCount) = False Then
'                DisplayBillID = !BillCount + 1
'            Else
'                DisplayBillID = 1
'            End If
'        Else
'            DisplayBillID = 1
'        End If
    
    
        If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeBill where IncomeBillID = 0 "
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Time = Now
        !Date = Date
        !UserID = UserID
        !StoreID = UserStoreID
        !GrossTotal = Val(txtInitialDeposit.Text)
        !NetTotal = Val(txtInitialDeposit.Text)
        !PaymentMethodID = PCat.PaymentMethodID
        !BHTID = temBHTID
        !IsGSBill = True
        
        !Completed = True
        !CompletedUserID = UserID
        !CompletedDate = Date
        !CompletedTime = Now
        .Update
        temSQL = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        txtIncomeBillID.Text = !NewID
        .Close
        
        
        DisplayBillID = NewGSBDisplayBillID(Val(txtIncomeBillID.Text))

         If .State = 1 Then .Close
        temSQL = "Select * from tblIncomeBill where IncomeBillID = " & Val((txtIncomeBillID.Text))
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        !DisplayBillID = DisplayBillID
        .Update
       End If
       .Close
    End With
    
    If chkPrint.Value = 1 Then
       ' lblBillNo.Visible = True
       ' txtIncomeBillID.Visible = True
        
        PrintAdmissionBill
        
        lblBillNo.Visible = False
        txtIncomeBillID.Visible = False
    End If
    MsgBox "Patient Admitted"
    Unload Me
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
'
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
'Private Sub cmbHealthSchemeSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyEscape Then
'        cmbHealthSchemeSupplier.Text = Empty
'        KeyCode = Empty
'    ElseIf KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        txtInitialDeposit.SetFocus
'    End If
'End Sub



Private Sub cmbPtCat_Change()
    If IsNumeric(cmbPtCat.BoundText) = False Then
        txtPtSurcharge.Text = 0
        Exit Sub
    End If
    PCat.ID = Val(cmbPtCat.BoundText)
    txtPaymentMethodID.Text = PCat.PaymentMethodID
    If LCase(PCat.PaymentMethod) = "credit" Then
        MsgBox "Credit Not Allowed"
        cmbPtCat.Text = Empty
        cmbPtCat.SetFocus
'        cmbHealthSchemeSupplier.Visible = True
'        lblHSS.Visible = True
'    Else
'        cmbHealthSchemeSupplier.Visible = False
'        lblHSS.Visible = False
    End If
    txtPtSurcharge.Text = PCat.Surcharge
'    cmbHealthSchemeSupplier.Text = Empty
'    txtComments.Text = Empty
End Sub

Private Sub cmbPtCat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbPtCat.Text = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtInitialDeposit.SetFocus
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
            temSQL = "SELECT tblStaff.Name AS NameWithTitle, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID where SpecialityID = " & Val(cmbSpeciality.BoundText) & " ORDER BY Name"
        Else
            temSQL = "SELECT tblStaff.Name AS NameWithTitle, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID Order BY Name"
        End If
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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

Private Sub Form_Activate()
    If FirstActivation = True Then
        txtBHT.SetFocus
        SendKeys "{home}+{end}"
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
    Call GetSettings
    FirstActivation = True
End Sub

Private Sub GetSettings(): On Error Resume Next
    cmbSpeciality.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbSpeciality.Name, "1"))
    cmbPtCat.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbPtCat.Name, "1"))
    
    txtBHT.Text = NewBHT
    dtpDOB.Value = Date
    dtpDOA.Value = Date
    dtpTOA.Value = Time
    lblTopic.Caption = HospitalName
    lblSubtopic.Caption = HospitalDescreption
    On Error Resume Next
    chkPrint.Value = GetSetting(App.EXEName, Me.Name, chkPrint.Name, "1")
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "")
End Sub

Private Function NewBHT() As Long
    Dim rsTemBHT As New ADODB.Recordset
    With rsTemBHT
        If .State = 1 Then .Close
        temSQL = "Select * from tblBHT Where IsGSB = 1 order by BHTID DESC"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
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
End Sub

Private Function CanAdmit() As Boolean
    CanAdmit = False
    Dim tr As Integer
    If Trim(txtName.Text) = Empty Then
        tr = MsgBox("You have not entered the name of the patient", vbCritical, "Name?")
        txtName.SetFocus
        Exit Function
    End If
    If Trim(txtBHT.Text) = Empty Then
        tr = MsgBox("You have not entered the GSB number", vbCritical, "GSB?")
        txtBHT.SetFocus
        Exit Function
    End If
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
    If LCase(PCat.PaymentMethod) = "credit" Then 'And IsNumeric(cmbHealthSchemeSupplier.BoundText) = False Then
        tr = MsgBox("Only Cash allowed", vbCritical, "Credit Not Allowed")
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
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, cmbPtCat.Name, cmbPtCat.BoundText
    
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
    If KeyCode = vbKeyEscape Then
        txtBHT.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpDOA.SetFocus
    End If
End Sub

Private Sub txtGunardian_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPtCat.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = Empty
        txtGunardian.Text = Empty
    End If
End Sub

Private Sub txtGunardian_LostFocus()
    txtGunardian.Text = UCase(txtGunardian.Text)
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


Private Sub txtInitialDeposit_LostFocus()
    txtInitialDeposit.Text = Format(Format(txtInitialDeposit.Text), "0.00")
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
