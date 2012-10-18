VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDeleteAllData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete All Data"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
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
   ScaleHeight     =   2010
   ScaleWidth      =   4965
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   1440
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
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Delete"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy MMMM dd"
      Format          =   56492035
      CurrentDate     =   39974
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy MMMM dd"
      Format          =   56492035
      CurrentDate     =   39974
   End
   Begin VB.Label Label1 
      Caption         =   "&From"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "&To"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmDeleteAllData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTem As New ADODB.Recordset
    Dim temSQL As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    Dim PtID As Long
    Dim rsDelete As New ADODB.Recordset
    
    i = MsgBox("Are you sure you want to delete the excess data in all the secessions?", vbYesNo)
    If i = vbNo Then Exit Sub
    
    If dtpFrom.Value > dtpTo.Value Then
        MsgBox "Please select a valid period"
        Exit Sub
    End If
    
    If dtpTo.Value > Date - 60 Then
        MsgBox "Data is less than 2 months old. You can not delete them"
        Exit Sub
    End If
    
   ' Exit Sub
    
    With rsTem
        
        
'        ALTER TABLE Persons
'ADD DateOfBirth date
'
        If .State = 1 Then .Close
        temSQL = "DELETE tblPatientFacility " & _
                    "From tblPatientFacility WHERE (((tblPatientFacility.AppointmentDate) Between '" & Format(dtpFrom.Value) & "' AND '" & Format(dtpTo.Value) & "'))"
        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
        
        If .State = 1 Then .Close
        temSQL = "Select * from tblPatientFacility order by PatientFacility_ID desc"
        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            PtID = !patientid
        End If
        
'        If .State = 1 Then .Close
'        temSQL = "Delete tblPatientMainDetails from tblPatientMainDetails where Patient_ID > " & PtID
'        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
        
        If .State = 1 Then .Close
        temSQL = "Delete tblPatientBill from tblPatientBill  WHERE (((tblPatientBill.Date) Between '" & Format(dtpFrom.Value) & "' AND '" & Format(dtpTo.Value) & "'))"
        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
        
        If .State = 1 Then .Close
        temSQL = "Delete tblPatientRepay from tblPatientRepay  WHERE (((tblPatientRepay.RepayDate) Between '" & Format(dtpFrom.Value) & "' AND '" & Format(dtpTo.Value) & "'))"
        .Open temSQL, cnnChannelling, adOpenStatic, adLockOptimistic
        

    End With
    MsgBox "Records Deleted"
End Sub

Private Sub Form_Load()
    Call GetSettings
End Sub

Private Sub GetSettings()
    dtpFrom.Value = GetSetting(App.EXEName, Me.Name, dtpFrom.Name, Date)
    dtpTo.Value = GetSetting(App.EXEName, Me.Name, dtpTo.Name, Date)
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, dtpFrom.Name, dtpFrom.Value
    SaveSetting App.EXEName, Me.Name, dtpTo.Name, dtpTo.Value
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub


