VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDoctorIncomeSecessionVice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctors Secessionvice Income"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDoctorIncomeSeceessionWise.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4920
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Close"
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
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   65732611
      CurrentDate     =   39597
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   65732611
      CurrentDate     =   39597
   End
   Begin VB.Label Label1 
      Caption         =   "To"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmDoctorIncomeSecessionVice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemTital As String
    Dim CSetPrinter As New cSetDfltPrinter
    Dim A

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    Const pershape = "SHAPE {"
    Const Sql = "SELECT tblTitle.Title , tblDoctor.DoctorListedName, tblPatientFacility.* FROM ((tblDoctor LEFT OUTER JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) RIGHT OUTER JOIN tblPatientFacility ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID)"
    Const PostSHape = "}  AS cmdDoctorIncomeC COMPUTE cmdDoctorIncomeC, SUM(cmdDoctorIncomeC.'PersonalFee') AS TotalPersonalFee, ANY(cmdDoctorIncomeC.'Title') AS TitalName, COUNT(cmdDoctorIncomeC.'PersonalFee') AS TotalPatients BY 'DoctorListedName','PersonalFee'"
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    
    With DataEnvironment1
        If .rscmdDoctorIncomeC_Grouping.State = 1 Then .rscmdDoctorIncomeC_Grouping.Close
        If PayToDoctor = True Then
            .Commands!cmdDoctorIncomeC_Grouping.CommandText = pershape & Sql & " Where (fullypaid = 1) and (cancelled = 0) and (refund = 0) AND (AppointmentDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  ) Order By DoctorListedName" & PostSHape
        Else
            .Commands!cmdDoctorIncomeC_Grouping.CommandText = pershape & Sql & " Where (fullypaid = 1) and (cancelled = 0) and (refund = 0) and (patientabsent = 0) AND (AppointmentDate between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  ) Order By DoctorName" & PostSHape
        End If
        .cmdDoctorIncomeC_Grouping
        
        Set dtrDoctorIncomeC.DataSource = DataEnvironment1
        dtrDoctorIncomeC.Sections("PageFooter").Controls("lblAdd").Caption = LongAd
        
        dtrDoctorIncomeC.Show
    End With
End Sub

Private Sub Form_Load()
    dtpTo.Value = Date
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    If UserAuthority <> AuthorityOwner Then
        dtpFrom.Enabled = False
        dtpTo.Enabled = False
    End If

End Sub

