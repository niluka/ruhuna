VERSION 5.00
Begin VB.Form frmWritingBalance 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4320
      Top             =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Wrintg to Database. Please Wait ..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "frmWritingBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub WriteToAgentBalance()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT tblInstitutionBalance.* From tblInstitutionBalance ORDER BY tblInstitutionBalance.Date"
    .Open
    If .RecordCount = 0 Then
        .Close
        WriteBalance
    Else
        If .State = 1 Then .Close
        .Source = "SELECT tblInstitutionBalance.* From tblInstitutionBalance ORDER BY tblInstitutionBalance.Date DESC"
        .Open
        If Format(!Date, "dd MMMM yyyy") < Format(Date, "dd MMMM yyyy") Then .Close: WriteBalance
    End If
End With
End Sub

'Private Sub WriteBalance()
'    DoEvents
'    Call TodaysBalance
'End Sub

Private Sub WriteBalance()
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "Select * from tblinstitutions order by institution_ID"
        .Open
        If .RecordCount > 0 Then
            While .EOF = False
                If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
                DataEnvironment1.rssqlTem1.Source = "SELECT tblInstitutionBalance.InstitutionBalance_Id, tblInstitutionBalance.Institution_Id, tblInstitutionBalance.Date, tblInstitutionBalance.SBalance, tblInstitutionBalance.EBalance From tblInstitutionBalance where tblInstitutionBalance.Institution_Id = " & !institution_ID & " AND tblInstitutionBalance.Date = #" & Format(Date, "dd MMMM yyyy") & "#"
                DataEnvironment1.rssqlTem1.Open
                If DataEnvironment1.rssqlTem1.RecordCount < 1 Then
                    DataEnvironment1.rssqlTem1.AddNew
                    DataEnvironment1.rssqlTem1!institution_ID = !institution_ID
                    DataEnvironment1.rssqlTem1!Date = Format(Date, "dd MMMM yyyy")
                    DataEnvironment1.rssqlTem1!SBalance = !InstitutionCredit
                    DataEnvironment1.rssqlTem1.Update
                Else
                    DataEnvironment1.rssqlTem1!EBalance = !InstitutionCredit
                End If
                .MoveNext
            Wend
        End If
        .Close
        DataEnvironment1.rssqlTem1.Close
    End With
End Sub



Private Sub TodaysStartingBalance()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "Select * from tblinstitutions order by institution_ID"
    .Open
    If .RecordCount < 1 Then Exit Sub
    If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
    DataEnvironment1.rssqlTem1.Source = "SELECT tblInstitutionBalance.InstitutionBalance_Id, tblInstitutionBalance.Institution_Id, tblInstitutionBalance.Date, tblInstitutionBalance.SBalance, tblInstitutionBalance.EBalance From tblInstitutionBalance"
    DataEnvironment1.rssqlTem1.Open
    While .EOF = False
        DataEnvironment1.rssqlTem1.AddNew
        DataEnvironment1.rssqlTem1!institution_ID = !institution_ID
        DataEnvironment1.rssqlTem1!Date = Format(Date, "dd MMMM yyyy")
        DataEnvironment1.rssqlTem1!SBalance = !InstitutionCredit
        DataEnvironment1.rssqlTem1.Update
        .MoveNext
    Wend
    .Close
    DataEnvironment1.rssqlTem1.Close
End With
End Sub

Private Sub YesterdaysEndingBalance()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "Select * from tblinstitutions order by institution_ID"
    .Open
    If .RecordCount < 1 Then Exit Sub
    While .EOF = False
        If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
        DataEnvironment1.rssqlTem1.Source = "SELECT tblInstitutionBalance.InstitutionBalance_Id, tblInstitutionBalance.Institution_Id, tblInstitutionBalance.Date, tblInstitutionBalance.SBalance, tblInstitutionBalance.EBalance From tblInstitutionBalance where tblInstitutionBalance.Institution_Id = " & !institution_ID & " AND tblInstitutionBalance.Date = #" & Format(Date - 1, "dd MMMM yyyy") & "#"
        DataEnvironment1.rssqlTem1.Open
        If DataEnvironment1.rssqlTem1.RecordCount = 0 Then
            DataEnvironment1.rssqlTem1.AddNew
            DataEnvironment1.rssqlTem1!SBalance = 0
            DataEnvironment1.rssqlTem1!institution_ID = !institution_ID
            DataEnvironment1.rssqlTem1!Date = Format(Date - 1, "dd MMMM yyyy")
        End If
        DataEnvironment1.rssqlTem1!EBalance = !InstitutionCredit
        DataEnvironment1.rssqlTem1.Update
        .MoveNext
    Wend
    .Close
    DataEnvironment1.rssqlTem1.Close
End With
End Sub

Private Sub Form_Load()
Me.MousePointer = vbHourglass
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.MousePointer = vbDefault
End Sub

Private Sub Timer1_Timer()
    DoEvents
    Call WriteBalance
    Timer1.Interval = 0
    Unload Me
End Sub
