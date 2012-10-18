VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDoctorLeave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Leave"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
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
   ScaleHeight     =   7560
   ScaleWidth      =   6870
   Begin VB.ListBox ListLeaveDayIDs 
      Height          =   2460
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Add"
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
   Begin VB.ListBox ListLeaveDays 
      Height          =   2820
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   3375
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2820
      Left            =   3600
      TabIndex        =   4
      Top             =   3480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   62128129
      CurrentDate     =   39490
   End
   Begin VB.ListBox ListConsultantIDs 
      Height          =   2700
      Left            =   3600
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox ListConsultants 
      Height          =   2700
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.ListBox ListSpecialityIDs 
      Height          =   2700
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox ListSpecialities 
      Height          =   2700
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Delete"
      Enabled         =   0   'False
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
   Begin btButtonEx.ButtonEx bttnClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   6960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Close"
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
   Begin VB.Label Label4 
      Caption         =   "Date to add leave"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Consultant"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Speciality"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Leave Days"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   3375
   End
End
Attribute VB_Name = "frmDoctorLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bttnAdd_Click()
    Dim tr As Integer
    On Error GoTo EH
    If ListConsultantIDs.ListIndex < 0 Or ListConsultants.ListIndex < 0 Or Not IsNumeric(ListConsultantIDs.Text) Then
        tr = MsgBox("You have not selected a doctor", vbCritical, "Doctor?")
        ListConsultants.SetFocus
        Exit Sub
    End If
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT * from tblfacilitysecession"
        .Open
        .AddNew
        !fulldayleave = 1
        !altereddate = MonthView1.Value
        !HospitalFacility_ID = 10
        !Staff_ID = Val(ListConsultantIDs.Text)
        .Update
        .Close
        Call FillLeaveDays
        Exit Sub
EH:
    .CancelUpdate
    MsgBox "Unknown Error" & vbNewLine & Err.Description
    Exit Sub
    End With
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDelete_Click()
    ListLeaveDayIDs.ListIndex = ListLeaveDays.ListIndex
    If ListConsultants.ListIndex < 0 Or ListConsultantIDs.ListIndex < 0 Or Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    If Not IsNumeric(ListLeaveDayIDs.Text) Then Exit Sub
On Error GoTo EH
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "Select * from tblFacilitySecession where facilitysecession_ID = " & Val(ListLeaveDayIDs.Text)
        .Open
        If .RecordCount > 0 Then
            .Delete adAffectCurrent
        End If
        If .State = 1 Then .Close

    Call FillLeaveDays
    bttnDelete.Enabled = False
    Exit Sub
EH:
    .Close
    Exit Sub
    End With
End Sub

Private Sub Form_Load()
    Call FormatGridSpeciality
    Call FormatGridConsultants
    Call FillSpeciality
    MonthView1.Value = Date
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
        .Source = "SELECT  tbldoctor.*  FROM  tbldoctor  order by doctorlistedname"
    Else
        .Source = "SELECT  tbldoctor.*  FROM  tbldoctor  order by doctorname"
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
            .Source = "SELECT tbldoctor.* FROM tbldoctor where  doctorspeciality_ID = " & Val(ListSpecialityIDs.Text) & " order by doctorname"
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



Private Sub ListLeaveDays_Click()
    bttnDelete.Enabled = False
    ListLeaveDayIDs.ListIndex = ListLeaveDays.ListIndex
    If ListConsultants.ListIndex < 0 Or ListConsultantIDs.ListIndex < 0 Or Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    If Not IsNumeric(ListLeaveDayIDs.Text) Then Exit Sub
    ListLeaveDayIDs.ListIndex = ListLeaveDays.ListIndex
    bttnDelete.Enabled = True
End Sub

Private Sub ListSpecialities_Click()
    ListSpecialityIDs.ListIndex = ListSpecialities.ListIndex
    ListConsultantIDs.Clear
    ListConsultants.Clear
    ListLeaveDayIDs.Clear
    ListLeaveDays.Clear
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
    Call FillLeaveDays
End Sub

Private Sub FillLeaveDays()
    If ListConsultants.ListIndex < 0 Then Exit Sub
    If ListConsultantIDs.ListIndex < 0 Then Exit Sub
    If Not IsNumeric(ListConsultantIDs.Text) Then Exit Sub
    ListLeaveDays.Clear
    ListLeaveDayIDs.Clear
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT * from tblfacilitysecession where hospitalfacility_ID = 10 and staff_ID = " & ListConsultantIDs.Text & " and altereddate >= '" & Date & "' and fulldayleave = 1 order by altereddate"
        If .State = 0 Then .Open
        If .RecordCount < 1 Then
            ListLeaveDayIDs.Clear
            ListLeaveDays.Clear
            If .State = 1 Then .Close
            Exit Sub
        Else
            While .EOF = False
                ListLeaveDays.AddItem Format(!altereddate, DefaultLongDate)
                ListLeaveDayIDs.AddItem !facilitysecession_ID
                .MoveNext
            Wend
        End If
    End With
End Sub

