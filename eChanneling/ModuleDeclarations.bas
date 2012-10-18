Attribute VB_Name = "ModuleDeclarations"
Option Explicit

Public Enum ControlType
    TextBox = 1
    ComboBox = 2
    ListBox = 3
    DataCombo = 4
    DataList = 5
    Grid = 6
    Button = 7
    MenuItem = 8
    CheckBox = 9
    OptionButton = 10
    DateTimePicker = 11
    Label = 12
    SSTab = 13
    Unknown = 100
End Enum

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Type IncomeByDates
    PersonalIncome As Double
    InstitutionIncome As Double
    OtherIncome As Double
End Type

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Const LB_SETTABSTOPS = &H192


Public Function CalculateAgeInMonths(ByVal DateOfBirth As Date) As Long
    CalculateAgeInMonths = DateDiff("m", DateOfBirth, Now)
End Function

Public Function CalculateAgeInWords(ByVal DateOfBirth As Date) As String
    Dim Age As Long
    Age = DateDiff("yyyy", DateOfBirth, Now)
    If Age >= 5 Then
        CalculateAgeInWords = Age & " Years"
        Exit Function
    Else
        Age = DateDiff("m", DateOfBirth, Now)
        If Age > 48 Then CalculateAgeInWords = "4" & " Years and " & Age - 48 & " months": Exit Function
        If Age = 48 Then CalculateAgeInWords = "4" & " Years": Exit Function
        If Age > 36 Then CalculateAgeInWords = "3" & " Years and " & Age - 36 & " months": Exit Function
        If Age = 36 Then CalculateAgeInWords = "3" & " Years": Exit Function
        If Age > 24 Then CalculateAgeInWords = "2" & " Years and " & Age - 24 & " months": Exit Function
        If Age = 24 Then CalculateAgeInWords = "2" & " Years": Exit Function
        If Age > 12 Then CalculateAgeInWords = "1" & " Years and " & Age - 12 & " months": Exit Function
        If Age = 12 Then CalculateAgeInWords = "1" & " Year": Exit Function
        If Age >= 1 Then CalculateAgeInWords = Age & " Months": Exit Function
        Age = DateDiff("d", DateOfBirth, Now)
        CalculateAgeInWords = Age & " Days"
        Exit Function
    End If
End Function


Public Function GetIncomeByDates(ByVal StartDate As Date, ByVal EndDate As Date, ByVal DaySecession As Integer, ByVal HospitalFacilityID As Long, ByVal StaffID As Long) As IncomeByDates

Dim TemIncome As IncomeByDates

With DataEnvironment1.rssqlTem3
    Dim TempTotalIncome         As Double
    Dim TempPersonalIncome As Double
    Dim TempInstitutionIncome   As Double
    Dim TempOtherIncome   As Double
    
    If .State = 1 Then .Close
    
    If HospitalFacilityID <> 0 And StaffID <> 0 Then
        .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & HospitalFacilityID & ") and (staff_ID = " & StaffID & ") and (bookingdate between '" & StartDate & "' and '" & EndDate & "')"
    ElseIf HospitalFacilityID <> 0 Then
        .Source = "SELECT * from tblpatientfacility where (hospitalfacility_ID = " & HospitalFacilityID & ") and (bookingdate between '" & StartDate & "' and '" & EndDate & "')"
    ElseIf StaffID <> 0 Then
        .Source = "SELECT * from tblpatientfacility where  (staff_ID = " & StaffID & ") and (bookingdate between '" & StartDate & "' and '" & EndDate & "')"
    Else
        TemIncome.PersonalIncome = TempPersonalIncome
        TemIncome.InstitutionIncome = TempInstitutionIncome
        TemIncome.PersonalIncome = TempOtherIncome
        Exit Function
    End If
    
    If .State = 0 Then .Open
    
    
    TempTotalIncome = 0
    TempPersonalIncome = 0
    TempInstitutionIncome = 0
    TempOtherIncome = 0
        
        If .RecordCount <> 0 Then
        .MoveFirst
            While Not .EOF
                If DaySecession = MorningSecession Then
                    If !Secession = MorningSecession Then
                        TempPersonalIncome = TempPersonalIncome + (!personalfee)
                        TempInstitutionIncome = TempInstitutionIncome + (!institutionfee)
                        TempOtherIncome = TempOtherIncome + (!otherfee)
                    End If
                ElseIf DaySecession = EveningSecession Then
                    If !Secession = EveningSecession Then
                        TempPersonalIncome = TempPersonalIncome + (!personalfee)
                        TempInstitutionIncome = TempInstitutionIncome + (!institutionfee)
                        TempOtherIncome = TempOtherIncome + (!otherfee)
                    End If
                Else
                    TempPersonalIncome = TempPersonalIncome + (!personalfee)
                    TempInstitutionIncome = TempInstitutionIncome + (!institutionfee)
                    TempOtherIncome = TempOtherIncome + (!otherfee)
                End If
                .MoveNext
            Wend
        End If
        TempTotalIncome = TempPersonalIncome + TempInstitutionIncome + TempOtherIncome
    End With
    TemIncome.PersonalIncome = TempPersonalIncome
    TemIncome.InstitutionIncome = TempInstitutionIncome
    TemIncome.PersonalIncome = TempOtherIncome
    GetIncomeByDates.PersonalIncome = TempPersonalIncome
    GetIncomeByDates.InstitutionIncome = TempInstitutionIncome
    GetIncomeByDates.OtherIncome = TempOtherIncome
End Function

