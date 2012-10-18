VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmCDoctorPatientLists 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Patient Lists"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
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
   ScaleHeight     =   1230
   ScaleWidth      =   4545
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd   MMMM   yyyy"
      Format          =   56819715
      CurrentDate     =   39696
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCDoctorPatientLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsDoctorPatientList As New ADODB.Recordset
    Dim temSQL As String
    
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    If dtpDate.Value = Date Or UserAuthority = AuthorityOwner Then
        temSQL = "SELECT [Title] & ' ' & [DoctorName] AS DoctorName, tblPatientMainDetails.FirstName, tblPatientFacility.AppointmentTime " & _
        "FROM (tblPatientFacility LEFT JOIN (tblTitle RIGHT JOIN tblDoctor ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID " & _
        "WHERE (((tblPatientFacility.AppointmentDate) = #" & Format(dtpDate.Value, "dd MMMM yyyy") & "#))"
    Else
        temSQL = "SELECT [Title] & ' ' & [DoctorName] AS DoctorName, tblPatientMainDetails.FirstName, tblPatientFacility.AppointmentTime " & _
        "FROM (tblPatientFacility LEFT JOIN (tblTitle RIGHT JOIN tblDoctor ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID) ON tblPatientFacility.Staff_ID = tblDoctor.Doctor_ID) LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID " & _
        "WHERE (((tblPatientFacility.AppointmentDate) = #" & Format(dtpDate.Value, "dd MMMM yyyy") & "#) AND (([PatientFacility_ID] Mod " & IncomeDeflation & ")=0))"
    End If
    With rsDoctorPatientList
        If .State = 1 Then .Close
        .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
    End With
    With dtrDoctorPatientLists
        Set .DataSource = rsDoctorPatientList
        .Sections("Section4").Controls("lblHospital").Caption = InstitutionName
        .Sections("Section4").Controls("lblTopic").Caption = "Doctor Patient Lists"
        .Sections("Section4").Controls("lblSubTopic").Caption = Format(dtpDate.Value, DefaultLongDate)
        
        .Sections("Section1").Controls("txtDoctor").DataField = "DoctorName"
        .Sections("Section1").Controls("txtPatient").DataField = "FirstName"
        .Sections("Section1").Controls("txtAppointment").DataField = "AppointmentTime"
        
        .Show
    End With
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
End Sub
