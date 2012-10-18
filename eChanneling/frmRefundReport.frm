VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmRefundReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refunds Report"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRefundReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12885
   Begin MSFlexGridLib.MSFlexGrid gridPt 
      Height          =   4215
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   7435
      _Version        =   393216
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2820
      Left            =   6600
      TabIndex        =   7
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   62390273
      CurrentDate     =   39476
   End
   Begin VB.ListBox ListConsultants 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      ItemData        =   "frmRefundReport.frx":0442
      Left            =   3000
      List            =   "frmRefundReport.frx":0444
      TabIndex        =   1
      ToolTipText     =   "List of Consultants of selected speciality"
      Top             =   360
      Width           =   3495
   End
   Begin VB.ListBox ListSpecialities 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      ItemData        =   "frmRefundReport.frx":0446
      Left            =   120
      List            =   "frmRefundReport.frx":0448
      TabIndex        =   0
      ToolTipText     =   "List of Specialities"
      Top             =   360
      Width           =   2775
   End
   Begin VB.ListBox ListConsultantIDs 
      Height          =   2700
      Left            =   5760
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox ListSpecialityIDs 
      Height          =   2700
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   10560
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2820
      Left            =   9720
      TabIndex        =   10
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   62390273
      CurrentDate     =   39476
   End
   Begin btButtonEx.ButtonEx btnToExcel 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Excel"
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
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2055
      _ExtentX        =   3625
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
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "P&rocess"
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
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   255
      Left            =   9720
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Consultant"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Speciality"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmRefundReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
    Dim temSql As String
    
    Dim TemRoomNo As String
    Dim TemDoctorFee As Double
    Dim TemFDoctorFee As Double
    Dim TemADoctorFee As Double
    Dim TemInstitutionFee As Double
    Dim TemFInstitutionFee As Double
    Dim TemAInstitutionFee As Double
    Dim TemOtherFee As Double
    Dim TemFOtherFee As Double
    Dim TemAOtherFee As Double
    Dim TemSecession As Long
    Dim CSetPrinter As New cSetDfltPrinter
    Dim SecessionMax As Long
    Dim TemCanByPassOrder As Boolean
    Dim TemCalculateAppointment As Boolean
    Dim TemAgentRefNo As String
'    Dim TemSecession  As Integer
    Dim TemAgentCredit As Double
    Dim TemPatientID As Long
    Dim TemAgentMaxCredit As Double
    Dim TemPatientFacilityID As Long
    Dim TemAppointmentDate As Date
    Dim TemAppointmentTime As Date
    Dim TemDaySerial As Long
    Dim TemAgentBookingID As Long
    Dim TemSecessionStartingTime As Date
    Dim TemUsualDuration As Long
    Dim TemPatient As String
    Dim TemConsultant As String
    Dim TemNonCancelledVisits As Long
    Dim TemBillId As Long
    Dim TemPreviousDate As Date
    Dim TemTextForList As String


Private Sub btnPrint_Click()
    Dim temTopic As String
    temTopic = ListConsultants.Text
    Dim myPR As PrintReport
        GetPrintDefaults myPR

    GridPrint gridPt, myPR, temTopic, "From " & Format(MonthView1.Value, "dd MMMM yyyy") & " to " & Format(MonthView2.Value, "dd MMMM yyyy")
End Sub

Private Sub btnProcess_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblFacilitySecession.SecessionName, dbo.tblPatientFacility.DaySerial, dbo.tblPatientMainDetails.FirstName, dbo.tblPatientFacility.RepayComments " & _
                    "FROM dbo.tblFacilitySecession RIGHT OUTER JOIN " & _
                      "dbo.tblPatientFacility ON dbo.tblFacilitySecession.FacilitySecession_ID = dbo.tblPatientFacility.Secession LEFT OUTER JOIN " & _
                      "dbo.tblPatientMainDetails ON dbo.tblPatientFacility.PatientID = dbo.tblPatientMainDetails.Patient_ID " & _
                        "WHERE (dbo.tblPatientFacility.Refund = 1) AND (dbo.tblPatientFacility.Staff_ID = " & Val(ListConsultantIDs.Text) & ") AND (dbo.tblPatientFacility.AppointmentDate BETWEEN CONVERT(DATETIME, '" & MonthView1.Value & "', 102) AND CONVERT(DATETIME, '" & MonthView2.Value & "', 102)) " & _
                        "ORDER BY dbo.tblFacilitySecession.SecessionName, dbo.tblPatientFacility.DaySerial "
        .Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridPt.Rows = gridPt.Rows + 1
            gridPt.Row = gridPt.Rows - 1
            
            gridPt.Col = 0
            gridPt.Text = !SecessionName
            
            gridPt.Col = 1
            gridPt.Text = !DaySerial
            
            gridPt.Col = 2
            gridPt.Text = !FirstName
            
            gridPt.Col = 3
            gridPt.Text = Format(!RepayComments, "")
            
            
            .MoveNext
        Wend
        .Close
    End With

End Sub

Private Sub FormatGrid()
    With gridPt
        .Clear
        
        .Rows = 1
        
        .Cols = 5
        
        .Row = 0
        
        .Col = 0
        .Text = "Secession"
        
        .Col = 1
        .Text = "Serial"
        
        .Col = 2
        .Text = "Patient"
        
        .Col = 3
        .Text = "Remarks"
    
    End With
End Sub

Private Sub btnToExcel_Click()
    GridToExcel gridPt, ListConsultants.Text, "From " & Format(MonthView1.Value, "dd MMMM yyyy") & " to " & Format(MonthView2.Value, "dd MMMM yyyy")
End Sub

Private Sub Form_Load()
    Call FormatGridSpeciality
    Call FormatGridConsultants
    Call FillSpeciality
    Dim ingRet As Long
    Dim TabDates(1) As Long
    Dim TabPatientFacilities(6) As Long
    TabDates(0) = 48
    TabDates(1) = 166
    TabPatientFacilities(0) = 3 * 4
    TabPatientFacilities(1) = 15 * 4
    TabPatientFacilities(2) = 20 * 4
    TabPatientFacilities(3) = 28 * 4
    TabPatientFacilities(4) = 29 * 4
    MonthView1.Value = Date
    MonthView2.Value = Date
    FormatGrid
    GetCommonSettings Me
End Sub

Private Sub FormatGridSpeciality()
    ListSpecialities.Clear
    ListSpecialityIDs.Clear
End Sub

Private Sub FormatGridConsultants()
    ListConsultants.Clear
    ListConsultantIDs.Clear
End Sub


Private Sub FillSpeciality()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblspeciality order by speciality "
    .Open
    If NoAllNames = False Then
        ListSpecialities.AddItem "All"
        ListSpecialityIDs.AddItem "All"
    End If
    If .RecordCount <> 0 Then
        While Not .EOF
            ListSpecialities.AddItem !Speciality
            ListSpecialityIDs.AddItem !speciality_ID
            .MoveNext
        Wend
    End If
    .Close
End With
End Sub


Private Sub ListAllConsultants()
Call FormatGridConsultants
With DataEnvironment1.rssqlTem1
    If .State = 1 Then .Close
    If SurnameFirst = True Then
        .Source = "SELECT  tbldoctor.*  FROM  tbldoctor  order by DoctorListedName"
    Else
        .Source = "SELECT  tbldoctor.*  FROM  tbldoctor  order by DoctorName"
    End If
    .Open
    If .RecordCount = 0 Then Exit Sub
    While Not .EOF
            If SurnameFirst = True Then
                ListConsultants.AddItem !doctorlistedname
            Else
                ListConsultants.AddItem !doctorname
            End If
        ListConsultantIDs.AddItem !Doctor_ID
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub ListSelectedConsultants()
    Call FormatGridConsultants
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        If SurnameFirst = True Then
            .Source = "SELECT tbldoctor.* FROM tbldoctor where  doctorspeciality_ID = " & Val(ListSpecialityIDs.Text) & " order by doctorlistedname"
        Else
            .Source = "SELECT tbldoctor.* FROM tbldoctor where  doctorspeciality_ID = " & Val(ListSpecialityIDs.Text) & " order by DoctorName"
        End If
        .Open
        If .RecordCount = 0 Then Exit Sub
        While Not .EOF
            
            If SurnameFirst = True Then
                ListConsultants.AddItem !doctorlistedname
            Else
                ListConsultants.AddItem !doctorname
            End If
            
            ListConsultantIDs.AddItem !Doctor_ID
            .MoveNext
        Wend
        .Close
    End With
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub

Private Sub ListSpecialities_Click()
    ListSpecialityIDs.ListIndex = ListSpecialities.ListIndex
    ListConsultantIDs.Clear
    ListConsultants.Clear
    If ListSpecialities.Text = "All" Then
        ListAllConsultants
    ElseIf ListSpecialities.Text <> "All" And IsNumeric(ListSpecialityIDs.Text) = True Then
        ListSelectedConsultants
    Else
        FormatGridConsultants
    End If
End Sub

Private Sub ListSpecialities_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Then
    ListConsultants.SetFocus
    KeyCode = Empty
Else

End If
End Sub

Private Sub ListConsultants_Click()
    ListConsultantIDs.ListIndex = ListConsultants.ListIndex
    TemPatientFacilityID = 0
    TemDoctorFee = 0
    TemFDoctorFee = 0
    TemInstitutionFee = 0
    TemFInstitutionFee = 0
    TemOtherFee = 0
    TemAppointmentDate = Empty
    TemAppointmentTime = Empty
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
End Sub


