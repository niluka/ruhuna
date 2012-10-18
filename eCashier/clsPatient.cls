VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim TitleV As String
    Dim SexV As String
    Dim AgeInWordsV As String
    Dim FirstNameV As String
    
Public Property Let ID(GiveID As Long)
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblTitle.Title, tblPatientMainDetails.FirstName, tblSex.Sex, tblPatientMainDetails.DateOfBirth " & _
                    "FROM (tblPatientMainDetails LEFT JOIN tblTitle ON tblPatientMainDetails.TitleID = tblTitle.TitleID) LEFT JOIN tblSex ON tblPatientMainDetails.SexID = tblSex.SexID " & _
                    "WHERE (((tblPatientMainDetails.PatientID)=" & GiveID & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            TitleV = Format(!Title, "")
            SexV = Format(!Sex, "")
            If IsNull(!DateOfBirth) = False Then
                AgeInWordsV = CalculateAgeInWords(![DateOfBirth])
            End If
            FirstNameV = !FirstName
            .Close
        End If
    
    End With

End Property

Public Property Get AgeInWords() As String
    AgeInWords = AgeInWordsV
End Property

Public Property Get Title() As String
    Title = TitleV
End Property

Public Property Get Sex() As String
    Sex = SexV
End Property


Public Property Get FirstName() As String
    FirstName = FirstNameV
End Property

