VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAllAppointments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Appointments"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11115
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
   ScaleHeight     =   8235
   ScaleWidth      =   11115
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   6975
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   12303
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Excel"
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   62390275
      CurrentDate     =   40464
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9600
      TabIndex        =   3
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
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
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmAllAppointments"
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
    GridToExcel grid, "All APPOINTMENTS", Format(dtpDate.Value, "dd MMMM yyyy")
End Sub

Private Sub btnProcess_Click()

    Dim D(3) As Integer
    Dim P(0) As Integer
    
    D(0) = 3
    D(1) = 4
    D(2) = 5

    temSQL = "SELECT     TOP 100 PERCENT dbo.tblTitle.Title + ' ' + dbo.tblDoctor.DoctorName AS Consultant, dbo.tblFacilitySecession.SecessionName AS Secession, " & _
                      "dbo.tblPatientFacility.DaySerial AS Number, dbo.tblPatientFacility.PersonalFee AS [Doctor Fee], dbo.tblPatientFacility.InstitutionFee AS [Hospital Fee],  " & _
                      "dbo.tblPatientFacility.PersonalFee + dbo.tblPatientFacility.InstitutionFee AS [Total Fee], dbo.tblPatientFacility.PaymentMode,  " & _
                      "dbo.tblPatientFacility.Cancelled AS Cancellations, dbo.tblPatientFacility.Refund AS Refunds, dbo.tblPatientFacility.PatientAbsent AS Absents,  " & _
                      "dbo.tblInstitutions.InstitutionName AS Agent, dbo.tblStaff.StaffName AS Staff, dbo.tblPatientFacility.PatientFacility_ID AS Serial  " & _
"FROM         dbo.tblStaff RIGHT OUTER JOIN  " & _
 "                     dbo.tblPatientFacility ON dbo.tblStaff.Staff_ID = dbo.tblPatientFacility.CreditStaff_ID LEFT OUTER JOIN  " & _
  "                    dbo.tblInstitutions ON dbo.tblPatientFacility.Agent_ID = dbo.tblInstitutions.Institution_ID LEFT OUTER JOIN  " & _
   "                   dbo.tblFacilitySecession ON dbo.tblPatientFacility.Secession = dbo.tblFacilitySecession.FacilitySecession_ID LEFT OUTER JOIN  " & _
    "                  dbo.tblDoctor LEFT OUTER JOIN  " & _
     "                 dbo.tblTitle ON dbo.tblDoctor.DoctorTitle_ID = dbo.tblTitle.Title_ID ON dbo.tblPatientFacility.Staff_ID = dbo.tblDoctor.Doctor_ID  " & _
"WHERE     (dbo.tblPatientFacility.AppointmentDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102))  " & _
"ORDER BY dbo.tblDoctor.DoctorName, dbo.tblFacilitySecession.SecessionName, dbo.tblPatientFacility.DaySerial"

    FillTotalGrid temSQL, grid, 0, D, P

    ReplaceGridCOlText grid, 7, "True", "Cancellation"
    ReplaceGridCOlText grid, 8, "True", "Refund"
    ReplaceGridCOlText grid, 9, "True", "Absent"
    ReplaceGridCOlText grid, 7, "False", ""
    ReplaceGridCOlText grid, 8, "False", ""
    ReplaceGridCOlText grid, 9, "False", ""

    Dim i As Integer
    
    With grid
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 6) = "Credit" And Trim(.TextMatrix(i, 11)) = "" Then
                .TextMatrix(i, 6) = "Telephone"
            ElseIf .TextMatrix(i, 6) = "Credit" And Trim(.TextMatrix(i, 11)) <> "" Then
                .TextMatrix(i, 6) = "Staff"
            End If
        Next
    End With


End Sub

Private Sub Form_Load()
    getSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    saveSettings
End Sub


Private Sub getSettings()
    GetCommonSettings Me
    dtpDate.Value = Date
End Sub

Private Sub saveSettings()
    SaveCommonSettings Me
End Sub
