VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAgentChannellingCount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agent Booking Counts"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
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
   ScaleHeight     =   2790
   ScaleWidth      =   4860
   Begin btButtonEx.ButtonEx bttnCalcel 
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   2040
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
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   64225283
      CurrentDate     =   39597
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   64225283
      CurrentDate     =   39597
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "To"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmAgentChannellingCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    
Private Sub bttnCalcel_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        temSQL = "SELECT (tblInstitutions.InstitutionName + ' ' + tblInstitutions.InstitutionCode) as InstitutionNameCode , Count(tblPatientFacility.PatientFacility_ID) AS CountOfPatientFacility_ID, sum(tblPatientFacility.RefundNull) AS SumOfRefundToAgent " & _
                    "FROM tblPatientFacility RIGHT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID " & _
                    "WHERE tblPatientFacility.AppointmentDate Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'  AND tblInstitutions.InstitutionName <> '' " & _
                    "GROUP BY tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode"
        .Open temSQL
    End With
    With dtrAgentCountt
        Set .DataSource = DataEnvironment1.rssqlTem
        .Sections("Section4").Controls("lblTopic").Caption = "From " & Format(dtpFrom.Value, "DD mmmm yyyy") & " To " & Format(dtpTo.Value, "DD mmmm yyyy")
        
        .Sections("Section1").Controls("txtName").DataField = "InstitutionNameCode"
        .Sections("Section1").Controls("txtCount").DataField = "CountOfPatientFacility_ID"
        .Sections("Section1").Controls("txtRefund").DataField = "SumOfRefundToAgent"
        .Sections("Section5").Controls("funCount").DataField = "CountOfPatientFacility_ID"
        .Sections("Section5").Controls("funRefund").DataField = "SumOfRefundToAgent"
        .Show
    End With
End Sub

Private Sub Form_Load()
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
End Sub
