Attribute VB_Name = "ModuleFind"
Public Function FindPatientByID(ByVal patientid As Long) As String
With DataEnvironment1.rssqlFunction1
    If .State = 1 Then .Close
    .Source = "SELECT tblpatientmaindetails.* from tblpatientmaindetails where (patient_ID = " & patientid & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindPatientByID = Empty: .Close: Exit Function
    FindPatientByID = FindTitleFromID(!Title_ID) & " " & !FirstName & " " & !surname
    .Close
End With
End Function

Public Function FindPhoneByID(ByVal patientid As Long) As String
With DataEnvironment1.rssqlFunction1
    If .State = 1 Then .Close
    .Source = "SELECT tblpatientmaindetails.* from tblpatientmaindetails where (patient_ID = " & patientid & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindPhoneByID = Empty: .Close: Exit Function
    FindPhoneByID = Format(!Phone, "")
    .Close
End With
End Function

Public Function FindPatientContactNoByID(ByVal patientid As Long) As String
With DataEnvironment1.rssqlFunction1
    If .State = 1 Then .Close
    .Source = "SELECT tblpatientmaindetails.phone from tblpatientmaindetails where (patient_ID = " & patientid & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindPatientContactNoByID = Empty: .Close: Exit Function
    FindPatientContactNoByID = Format(!Phone, "")
    .Close
End With
End Function

Public Function FindTitleFromID(ByVal TitleID As Long) As String
With DataEnvironment1.rssqlFunction2
    If .State = 1 Then .Close
    .Source = ("SELECT * from tbltitle where title_ID = " & TitleID)
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindTitleFromID = Empty:   .Close: Exit Function
    FindTitleFromID = !Title
    .Close
End With
End Function

Public Function FindHospitalFacilityFromID(ByVal sendingID As Long) As String
With DataEnvironment1.rssqlFunction3
    If .State = 1 Then .Close
    .Source = "SELECT tblhospitalfacility.* from tblhospitalfacility where (hospitalfacility_ID = " & sendingID & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindHospitalFacilityFromID = Empty: .Close: Exit Function
    FindHospitalFacilityFromID = !hospitalfacility
    .Close
End With
End Function

Public Function FindLDoctorFromID(ByVal sendingID As Long) As String
With DataEnvironment1.rssqlFunction4
    If .State = 1 Then .Close
    .Source = "SELECT tbldoctor.* from tbldoctor where (doctor_ID = " & sendingID & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindLDoctorFromID = Empty:  .Close: Exit Function
    'FindDoctorFromID = FindTitleFromID(!DoctorTitle_ID) & " " & !doctorname
    FindLDoctorFromID = FindTitleFromID(!DoctorTitle_ID) & " " & !doctorlistedname
    .Close
End With
End Function

Public Function FindDoctorFromID(ByVal sendingID As Long) As String
With DataEnvironment1.rssqlFunction4
    If .State = 1 Then .Close
    .Source = "SELECT tbldoctor.* from tbldoctor where (doctor_ID = " & sendingID & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindDoctorFromID = Empty:  .Close: Exit Function
    FindDoctorFromID = FindTitleFromID(!DoctorTitle_ID) & " " & !doctorname
    .Close
End With
End Function


Public Function CanChannellByPhone(ByVal sendingID As Long) As Boolean
With DataEnvironment1.rssqlFunction4
    If .State = 1 Then .Close
    .Source = "SELECT tbldoctor.* from tbldoctor where (doctor_ID = " & sendingID & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then CanChannellByPhone = False:   .Close: Exit Function
    If !CreditBookings = True Then
        CanChannellByPhone = True
    Else
        CanChannellByPhone = False
    End If
    .Close
End With
End Function

Public Function FindInvestigationFromID(ByVal sendingID As Long) As String
With DataEnvironment1.rssqlFunction5
    If .State = 1 Then .Close
    .Source = "SELECT tblinvestigations.* from tblinvestigations where (investigation_ID = " & sendingID & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindInvestigationFromID = Empty: .Close: Exit Function
    FindInvestigationFromID = !Investigation
    .Close
End With
End Function

Public Function FindStaffFromID(ByVal sendingID As Long) As String
With DataEnvironment1.rssqlFunction6
    If .State = 1 Then .Close
    .Source = "SELECT tblstaff.* from tblstaff where (staff_ID = " & sendingID & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindStaffFromID = Empty: .Close: Exit Function
    FindStaffFromID = !StaffName
    .Close
End With
End Function

Public Function FindSecessionFromID(ByVal sendingID As Long)
With DataEnvironment1.rssqlFunction1
    If .State = 1 Then .Close
    .Source = "select * from tblfacilitysecession where facilitysecession_ID = " & sendingID
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindSecessionFromID = Empty: .Close: Exit Function
    FindSecessionFromID = !SecessionName
    .Close
End With
End Function

Public Function FindAgentFromID(ByVal sendingID As Long)
With DataEnvironment1.rssqlFunction1
    If .State = 1 Then .Close
    .Source = "select * from tblinstitutions where institution_ID = " & sendingID
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindAgentFromID = Empty: .Close: Exit Function
    FindAgentFromID = !InstitutionName & " (" & !InstitutionCode & ")"
    .Close
End With
End Function

Public Function FindAgentCodeFromID(ByVal sendingID As Long)
With DataEnvironment1.rssqlFunction1
    If .State = 1 Then .Close
    .Source = "select * from tblinstitutions where institution_ID = " & sendingID
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindAgentCodeFromID = Empty: .Close: Exit Function
    FindAgentCodeFromID = !InstitutionCode
    .Close
End With
End Function


Public Function FindStaffCodeFromID(ByVal sendingID As Long)
With DataEnvironment1.rssqlFunction1
    If .State = 1 Then .Close
    .Source = "select * from tblStaff where Staff_ID = " & sendingID
    If .State = 0 Then .Open
    If .RecordCount = 0 Then FindStaffCodeFromID = Empty: .Close: Exit Function
    FindStaffCodeFromID = !StaffCode
    .Close
End With
End Function

Public Function GetControlType(MyControl As Control) As ControlType
    GetControlType = Unknown
    If TypeOf MyControl Is TextBox Then
        GetControlType = TextBox
    ElseIf TypeOf MyControl Is ComboBox Then
        GetControlType = ComboBox
    ElseIf TypeOf MyControl Is Button Then
        GetControlType = Button
    ElseIf TypeOf MyControl Is CheckBox Then
        GetControlType = CheckBox
    ElseIf TypeOf MyControl Is DataCombo Then
        GetControlType = DataCombo
    ElseIf TypeOf MyControl Is DataList Then
        GetControlType = DataList
    ElseIf TypeOf MyControl Is DTPicker Then
        GetControlType = DateTimePicker
    ElseIf TypeOf MyControl Is MSFlexGrid Then
        GetControlType = grid
    ElseIf TypeOf MyControl Is Label Then
        GetControlType = Label
    ElseIf TypeOf MyControl Is ListBox Then
        GetControlType = ListBox
    ElseIf TypeOf MyControl Is Menu Then
        GetControlType = MenuItem
    ElseIf TypeOf MyControl Is OptionButton Then
        GetControlType = OptionButton
    ElseIf TypeOf MyControl Is SSTab Then
        GetControlType = SSTab
    End If

End Function

Public Function GetControlText(MyControl As Control) As String
    GetControlText = Empty
    If TypeOf MyControl Is TextBox Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is ComboBox Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is Button Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is CheckBox Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is DataCombo Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is DataList Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is DTPicker Then
        GetControlText = Right(MyControl.Name, Len(MyControl.Name) - 3)
    ElseIf TypeOf MyControl Is MSFlexGrid Then
        GetControlText = Right(MyControl.Name, Len(MyControl.Name) - 4)
    ElseIf TypeOf MyControl Is Label Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is ListBox Then
        GetControlText = MyControl.Text
    ElseIf TypeOf MyControl Is Menu Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is OptionButton Then
        GetControlText = MyControl.Caption
    ElseIf TypeOf MyControl Is SSTab Then
        GetControlText = MyControl.Caption
    End If
End Function

Public Function GetFormID(FormName As String, FormText As String) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblForm where Form = '" & FormName & "'"
        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !FormText = FormText
        Else
            .AddNew
            !FormText = FormText
            !Form = FormName
        End If
        .Update
        GetFormID = !FormID
        .Close
    End With
End Function


Public Sub EnableControls(MyForm As Form)
    Dim MyControl As Control
    Dim TemText As String
    Dim rsTem As New ADODB.Recordset
    On Error Resume Next
    For Each MyControl In MyForm.Controls
        With rsTem
            If .State = 1 Then .Close
            temSQL = "Select * from tblUserAuthorityControl where AuthorityID = " & UserAuthority & " AND ControlID = " & GetControlID(GetFormID(MyForm.Name, MyForm.Caption), MyControl)
            .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                MyControl.Enabled = !Enabled
            Else
                MyControl.Enabled = True
            End If
            .Close
        End With
    Next
End Sub

Public Sub VisibleControls(MyForm As Form)
    Dim MyControl As Control
    Dim TemText As String
    Dim rsTem As New ADODB.Recordset
    On Error Resume Next
    For Each MyControl In MyForm.Controls
        With rsTem
            If .State = 1 Then .Close
            temSQL = "Select * from tblUserAuthorityControl where AuthorityID = " & UserAuthority & " AND ControlID = " & GetControlID(GetFormID(MyForm.Name, MyForm.Caption), MyControl)
            .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                MyControl.Visible = !Visible
            Else
                MyControl.Visible = True
            End If
            .Close
        End With
    Next
End Sub


Private Function GetControlID(FormID As Long, MyControl As Control) As Long
    GetControlID = 0
    Dim rsForm As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
            With rsForm
                If TypeOf MyControl Is SSTab Then
                    For i = 0 To MyControl.Tabs - 1
                        If .State = 1 Then .Close
                        temSQL = "Select * from tblCOntrol where FormID = " & FormID & " AND COntrol = '" & MyControl.Name & "' AND ControlIndex = " & i
                        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
                        MyControl.Tab = i
                        If .RecordCount > 0 Then
                            !ControlText = GetControlText(MyControl)
                        Else
                            .AddNew
                            !FormID = FormID
                            !Control = MyControl.Name
                            !ControlType = GetControlType(MyControl)
                            !ControlText = GetControlText(MyControl)
                            !ControlIndex = i
                        End If
                        .Update
                        GetControlID = !ControlID
                        .Close
                    Next i
                Else
                    If .State = 1 Then .Close
                    temSQL = "Select * from tblCOntrol where FormID = " & FormID & " AND COntrol = '" & MyControl.Name & "'"
                    .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        !ControlText = GetControlText(MyControl)
                    Else
                        .AddNew
                        !FormID = FormID
                        !Control = MyControl.Name
                        !ControlType = GetControlType(MyControl)
                        !ControlText = GetControlText(MyControl)
                    End If
                    .Update
                    GetControlID = !ControlID
                    .Close
                End If
            End With
End Function
