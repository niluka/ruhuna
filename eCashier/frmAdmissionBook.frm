VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdmissionBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admission Book"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
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
   ScaleHeight     =   9345
   ScaleWidth      =   12270
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Process"
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
   Begin MSFlexGridLib.MSFlexGrid gridAdmission 
      Height          =   7455
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   13150
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   20709379
      CurrentDate     =   40201
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   8760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "To &Excel"
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
      Left            =   120
      TabIndex        =   5
      Top             =   8760
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   10920
      TabIndex        =   6
      Top             =   8760
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   20709379
      CurrentDate     =   40201
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmAdmissionBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    If dtpFrom.Value = dtpTo.Value Then
    GridToExcel gridAdmission, HospitalName & " - Admissions on ", Format(dtpFrom.Value, "dd MMMM yyyy")
    Else
    GridToExcel gridAdmission, HospitalName & " - Admissions From ", Format(dtpFrom.Value, "dd MMMM yyyy") & " To " & Format(dtpTo.Value, "dd MMMM yyyy")
    End If
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    ThisReportFormat.ReportPrintOrientation = Landscape
    
    
    
    GetPrintDefaults ThisReportFormat
    
    With ThisReportFormat
        
        .LeftMargin = 0
        .ColSpace = 70
        
        .ReportPrintOrientation = Landscape
        
        .TopicFontSize = 11
        .TopicFontName = "Tahoma"
        
        .SubTopicFontSize = 10
        .SubTopicFontName = "Tahoma"
        
        .HeaderFontName = "Tahoma"
        .HeaderFontSize = 8
        .HeaderFontBold = False
        .HeaderFontUnderline = False
        
        .ColTopicFontName = "Tahoma"
        .ColTopicFontSize = 8
        .ColTopicFontBold = False
        .ColTopicFontUnderline = False
        
        .ColFontSize = 7
        .ColFontName = "Tahoma"
        
    End With
    
    
    GridPrint gridAdmission, ThisReportFormat, HospitalName & " - Admissions", Format(dtpFrom.Value, "dd MMMM yyyy")
    Printer.EndDoc

End Sub

Private Sub btnProcess_Click()
    temSQL = "SELECT     dbo.tblBHT.BHT, dbo.tblRoom.Room, dbo.tblPatientMainDetails.FirstName as [Patient Name], dbo.tblStaff.Name as [Doctor Name], dbo.tblBHT.DOA, dbo.tblBHT.TOA,  " & _
                      "dbo.tblPatientMainDetails.Address, dbo.tblBHT.TemAge AS Age, dbo.tblSex.Sex, dbo.tblPatientMainDetails.NICNo, dbo.tblPatientMainDetails.Phone, " & _
                      "dbo.tblBHT.GuardianName, dbo.tblBHT.GuardianAddress, dbo.tblBHT.GuardianNIC, dbo.tblBHT.GuardianPhone, " & _
                      "dbo.tblPaymentMethod.PaymentMethod , dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierName " & _
"FROM         dbo.tblStaff RIGHT OUTER JOIN " & _
 "                     dbo.tblBHT ON dbo.tblStaff.StaffID = dbo.tblBHT.ReferringDoctorID LEFT OUTER JOIN " & _
  "                    dbo.tblHealthSchemeSuppliers ON dbo.tblBHT.HealthSchemeSupplierID = dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierID LEFT OUTER JOIN " & _
   "                   dbo.tblPaymentMethod ON dbo.tblBHT.PaymentMethodID = dbo.tblPaymentMethod.PaymentMethodID LEFT OUTER JOIN " & _
    "                  dbo.tblRoom ON dbo.tblBHT.RoomID = dbo.tblRoom.RoomID LEFT OUTER JOIN " & _
     "                 dbo.tblSex RIGHT OUTER JOIN " & _
      "                dbo.tblPatientMainDetails ON dbo.tblSex.SexID = dbo.tblPatientMainDetails.SexID ON " & _
       "               dbo.tblBHT.PatientID = dbo.tblPatientMainDetails.PatientID " & _
"WHERE     dbo.tblBHT.DOA between '" & dtpFrom.Value & "' AND  '" & dtpTo.Value & "' AND dbo.tblBHT.IsBHT = 1"

    FillAnyGrid temSQL, gridAdmission
End Sub

Private Sub Form_Load()
    Call GetCommonSettings(Me)
    dtpFrom.Value = Date
    dtpTo.Value = Date
    btnProcess_Click
    Call GetCommonSettings(Me)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub
