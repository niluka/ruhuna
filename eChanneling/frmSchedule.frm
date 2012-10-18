VERSION 5.00
Begin VB.Form frmSchedule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prepairing Schedule"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   360
   End
   Begin VB.Label Label1 
      Caption         =   "Please Wait ..."
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temSQL As String
Dim MyTime As Long

Private Sub WriteToDatabase()

Dim Speciality As String
Dim Consultant As String
Dim MyWeekday As Long
Dim TemSpeciality As String
Dim TemConsultant As String
Dim TemWeekday As Long

    With DataEnvironment1
        If .rssqlTem.State = 1 Then .rssqlTem.Close
        temSQL = "Truncate table tbltemtext"
        .rssqlTem.Open temSQL
        If .rssqlTem.State = 1 Then .rssqlTem.Close
        temSQL = "Select tbltemtext.* from tbltemtext"
        .rssqlTem.Source = temSQL
        .rssqlTem.Open
        temSQL = "SELECT * FROM tblSpeciality ORDER BY tblSpeciality.Speciality"
        If .rssqlTem1.State = 1 Then .rssqlTem1.Close
        .rssqlTem1.Source = temSQL
        .rssqlTem1.Open
        If .rssqlTem1.RecordCount <> 0 Then
            While .rssqlTem1.EOF = False
                Speciality = .rssqlTem1!Speciality
            
                temSQL = "SELECT tblDoctor.Doctor_ID, tblDoctor.DoctorSpeciality_ID, tblDoctor.DoctorTitle_ID, tblDoctor.DoctorName FROM tblDoctor where  tblDoctor.DoctorSpeciality_ID = " & .rssqlTem1!speciality_ID & " ORDER BY tblDoctor.DoctorName"
                If .rssqlTem2.State = 1 Then .rssqlTem2.Close
                .rssqlTem2.Source = temSQL
                .rssqlTem2.Open
                If .rssqlTem2.RecordCount <> 0 Then
                    While .rssqlTem2.EOF = False
                        Consultant = .rssqlTem2!doctorname
                
                        For MyWeekday = 2 To 8
                            If MyWeekday = 8 Then
                                temSQL = "SELECT tblFacilitySecession.FacilitySecession_ID, tblFacilitySecession.SecessionName, tblFacilitySecession.StartingTime, tblFacilitySecession.AgentDoctorFee, tblFacilitySecession.AgentHospitalFee, tblFacilitySecession.SecessionWeekday, tblFacilitySecession.Staff_ID FROM tblFacilitySecession WHERE (((tblFacilitySecession.SecessionWeekday) = " & 1 & ") AND ((tblFacilitySecession.Staff_ID)=" & .rssqlTem2!Doctor_ID & ")) order by tblFacilitySecession.StartingTime"
                            Else
                                temSQL = "SELECT tblFacilitySecession.FacilitySecession_ID, tblFacilitySecession.SecessionName, tblFacilitySecession.StartingTime, tblFacilitySecession.AgentDoctorFee, tblFacilitySecession.AgentHospitalFee, tblFacilitySecession.SecessionWeekday, tblFacilitySecession.Staff_ID FROM tblFacilitySecession WHERE (((tblFacilitySecession.SecessionWeekday) = " & MyWeekday & ") AND ((tblFacilitySecession.Staff_ID)=" & .rssqlTem2!Doctor_ID & ")) order by tblFacilitySecession.StartingTime"
                            End If
                            If .rssqlTem3.State = 1 Then .rssqlTem3.Close
                            .rssqlTem3.Source = temSQL
                            .rssqlTem3.Open
                            If .rssqlTem3.RecordCount <> 0 Then
                                While .rssqlTem3.EOF = False
                                    If .rssqlTem.State = 0 Then .rssqlTem.Open
                                    .rssqlTem.AddNew
                                    If Consultant <> TemConsultant Then
                                        .rssqlTem!txt2 = FindTitleFromID(.rssqlTem2!DoctorTitle_ID) & " " & Consultant
                                        TemConsultant = Consultant
                                    End If
                                    If Speciality <> TemSpeciality Then
                                        .rssqlTem!txt1 = Speciality
                                        TemSpeciality = Speciality
                                    End If
                                    If MyWeekday <> TemWeekday Then
                                        Select Case MyWeekday
                                            Case 1: .rssqlTem!txt3 = "Sunday"
                                            Case 2: .rssqlTem!txt3 = "Monday"
                                            Case 3: .rssqlTem!txt3 = "Tuesday"
                                            Case 4: .rssqlTem!txt3 = "Wednesday"
                                            Case 5: .rssqlTem!txt3 = "Thursday"
                                            Case 6: .rssqlTem!txt3 = "Friday"
                                            Case 7: .rssqlTem!txt3 = "Saturday"
                                        End Select
                                    TemWeekday = MyWeekday
                                    End If
                                    .rssqlTem!txt4 = .rssqlTem3!SecessionName
                                    .rssqlTem!txt5 = .rssqlTem3!startingtime
                                    .rssqlTem!txt6 = Format(.rssqlTem3!agentDoctorFee, "0.00")
                                    .rssqlTem!txt7 = Format(.rssqlTem3!AgentHospitalFee, "0.00")
                                    .rssqlTem.Update
                                    .rssqlTem3.MoveNext
                                Wend
                            End If
                
                        Next
                
                
                        .rssqlTem2.MoveNext
                    Wend
                End If
            
                .rssqlTem1.MoveNext
            Wend
        End If
    End With
   
    If HospitalDetails = True Then
        dtrSecessions.Sections.Item("Section4").Controls.Item("lblInstitutionName").Caption = InstitutionName
        dtrSecessions.Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
    Else
        dtrSecessions.Sections.Item("Section4").Controls.Item("lblInstitutionName").Caption = Empty
        dtrSecessions.Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = Empty
    End If
    dtrSecessions.Sections.Item("Section4").Controls.Item("lblreport").Caption = "Shedule for All Consultants for Agents"
    dtrSecessions.Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
    dtrSecessions.Show

End Sub

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub Timer1_Timer()
MyTime = MyTime + 1
If MyTime = 2 Then Call WriteToDatabase
If MyTime > 10 Then Unload Me: MyTime = 0
End Sub
