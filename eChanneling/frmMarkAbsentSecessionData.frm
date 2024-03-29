VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmMarkAbsentSecessionData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make Absent to Secession Data"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
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
   ScaleHeight     =   2520
   ScaleWidth      =   7500
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   1800
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Mark Absent"
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
   Begin MSComctlLib.Slider sldPercent 
      Height          =   675
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1191
      _Version        =   393216
      Max             =   100
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy MMMM dd"
      Format          =   62193667
      CurrentDate     =   39974
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy MMMM dd"
      Format          =   62193667
      CurrentDate     =   39974
   End
   Begin VB.Label Label3 
      Caption         =   "Percentage to Mark Absent"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMarkAbsentSecessionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTem As New ADODB.Recordset
    Dim temSql As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    Dim PtID As Long
    Dim TotalCount As Long
    Dim DeleteCount As Long
    Dim DeletedCount As Long
    Dim rsDelete As New ADODB.Recordset
    Dim rsDoc As New ADODB.Recordset
    
    i = MsgBox("Are you sure you want to Mark Absent the data in all the secessions?", vbYesNo)
    If i = vbNo Then Exit Sub
    If rsDoc.State = 1 Then rsDoc.Close
    temSql = "Select * from tblDoctor"
    rsDoc.Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
    While rsDoc.EOF = False
        With rsTem
            If .State = 1 Then .Close
            temSql = "SELECT Count(tblPatientFacility.PatientFacility_ID) AS CountOfPatientFacility_ID, tblPatientFacility.AppointmentDate, tblPatientFacility.Secession " & _
                        "From tblPatientFacility " & _
                        "GROUP BY  tblPatientFacility.Staff_ID, tblPatientFacility.AppointmentDate, tblPatientFacility.Secession " & _
                        "HAVING (((tblPatientFacility.AppointmentDate) Between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "') AND  ((tblPatientFacility.Staff_ID) = " & rsDoc!Doctor_ID & " )  ) " & _
                        "ORDER BY tblPatientFacility.Staff_ID, tblPatientFacility.AppointmentDate, tblPatientFacility.Secession"
            .Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
            While .EOF = False
                TotalCount = !CountOfPatientFacility_ID
                DeleteCount = Round(TotalCount * sldPercent.Value / 100)
                DeletedCount = 0
                If rsDelete.State = 1 Then rsDelete.Close
                temSql = "Select * from tblPatientFacility where  tblPatientFacility.AppointmentDate = '" & !AppointmentDate & "' AND tblPatientFacility.Secession = " & !Secession & " ORDER BY tblPatientFacility.Staff_ID, tblPatientFacility.AppointmentDate, tblPatientFacility.Secession"
                rsDelete.Open temSql, cnnChannelling, adOpenStatic, adLockOptimistic
                While rsDelete.EOF = False
                    If DeletedCount < DeleteCount Then
                        If DateDiff("n", rsDelete!bookingtime, Now) Mod 3 = 0 Then
                            rsDelete!PatientAbsent = True
                            rsDelete!PatientAbsentNull = 1
                            rsDelete.Update
                            DeletedCount = DeletedCount + 1
                        End If
                    End If
                    rsDelete.MoveNext
                Wend
                rsDelete.MoveLast
                While rsDelete.BOF = False
                    If DeletedCount < DeleteCount Then
                        If DateDiff("n", rsDelete!bookingtime, Now) Mod 3 = 1 Then
                            rsDelete!PatientAbsent = True
                            rsDelete!PatientAbsentNull = 1
                            rsDelete.Update
                            DeletedCount = DeletedCount + 1
                        End If
                    End If
                    rsDelete.MovePrevious
                Wend
                rsDelete.MoveFirst
                While rsDelete.EOF = False
                    If DeletedCount < DeleteCount Then
                        If DateDiff("n", rsDelete!bookingtime, Now) Mod 3 = 2 Then
                            rsDelete!PatientAbsent = True
                            rsDelete!PatientAbsentNull = 1
                            rsDelete.Update
                            DeletedCount = DeletedCount + 1
                        End If
                    End If
                    rsDelete.MoveNext
                Wend
                
                rsDelete.Close
                .MoveNext
            Wend
            .Close
        End With
        rsDoc.MoveNext
    Wend
    MsgBox "Records Updated"
End Sub

Private Sub Form_Load()
    Call GetSettings
End Sub

Private Sub GetSettings()
    dtpFrom.Value = GetSetting(App.EXEName, Me.Name, dtpFrom.Name, Date)
    dtpTo.Value = GetSetting(App.EXEName, Me.Name, dtpTo.Name, Date)
    sldPercent.Value = Val(GetSetting(App.EXEName, Me.Name, sldPercent.Name, 60))
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, sldPercent.Name, sldPercent.Value
    SaveSetting App.EXEName, Me.Name, dtpFrom.Name, dtpFrom.Value
    SaveSetting App.EXEName, Me.Name, dtpTo.Name, dtpTo.Value
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub
