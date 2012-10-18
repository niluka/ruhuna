VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmCancelRepayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancellations of Repayments"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
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
   ScaleHeight     =   765
   ScaleWidth      =   6180
   Begin btButtonEx.ButtonEx bttnCancel 
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cancel Repayment"
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
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Receipt ID"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmCancelRepayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bttnCancel_Click()
Dim tr As Integer
On Error GoTo EH
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientfacility where PatientFacility_ID = " & Val(txtID.Text)
    .Open
    If .RecordCount = 0 Then
        tr = MsgBox("There is no such Receipt ID. Cancellation of repayment NOT done", vbCritical, "Error")
        txtID.SetFocus
        SendKeys "{Home}+{end}"
        Exit Sub
    End If
    If !cancelled = False And !Refund = False Then
        tr = MsgBox("This visit is not nither Cancelled nor Refunded, therefore Cancellation of repayment NOT done", vbCritical, "Can't Cancell")
        txtID.SetFocus
        SendKeys "{Home}+{end}"
        Exit Sub
    End If
    If !repaydate <> Date Then
        tr = MsgBox("This cancellation is not done today. If you cancel this repayment, there can be erronous results of records. Are you sure you want to cancel?", vbInformation + vbYesNo, "Can't Cancell")
            If tr = vbNo Then
                txtID.SetFocus
                SendKeys "{Home}+{end}"
                Exit Sub
            End If
    End If
    If !FullyPaid = 0 Then
        tr = MsgBox("This is not validated. Therefore you can't cancel?", vbInformation + vbYesNo, "Can't Cancell")
            If tr = vbNo Then
                txtID.SetFocus
                SendKeys "{Home}+{end}"
                Exit Sub
            End If
    End If
    tr = MsgBox("Are you sure you want to cancel this repayment", vbQuestion + vbYesNo, "Cancel Repayment?")
    If tr = vbNo Then
                txtID.SetFocus
                SendKeys "{Home}+{end}"
                Exit Sub
    End If
    If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
    DataEnvironment1.rssqlTem1.Source = "SELECT * from tblcancelrepayment"
    DataEnvironment1.rssqlTem1.Open
    DataEnvironment1.rssqlTem1.AddNew
    If !RefundToPatient = 1 Then
        DataEnvironment1.rssqlTem1!repaidtopatient = True
    ElseIf !RefundToAgent = 1 Then
        DataEnvironment1.rssqlTem1!repaidtopatient = True
        If DataEnvironment1.rssqlTem2.State = 1 Then DataEnvironment1.rssqlTem2.Close
        DataEnvironment1.rssqlTem2.Source = "SELECT * from tblinstitutions where institution_ID = " & !Agent_ID
        DataEnvironment1.rssqlTem2.Open
        If DataEnvironment1.rssqlTem2.RecordCount = 0 Then Exit Sub
        DataEnvironment1.rssqlTem2!InstitutionCredit = DataEnvironment1.rssqlTem2!InstitutionCredit - !totalrefund
        DataEnvironment1.rssqlTem2.Update
        DataEnvironment1.rssqlTem2.Close
    Else
    End If
    !repayUser_ID = Null
    !Personalrefund = Null
    !institutionrefund = Null
    !otherrefund = Null
    !totalrefund = Null
    !repaycomments = ""
    !repaydate = Null
    !repaytime = Null
    !personaldue = !personalfee
    !institutiondue = !institutionfee
    !otherdue = !otherfee
    !cancellednull = 0
    !refundnull = 0
    DataEnvironment1.rssqlTem1!patientfacility_ID = !patientfacility_ID
    DataEnvironment1.rssqlTem1!Date = Date
    DataEnvironment1.rssqlTem1!Time = Time
    DataEnvironment1.rssqlTem1!user_ID = UserID
    If Not IsNull(!Agent_ID) Then DataEnvironment1.rssqlTem1!Agent_ID = !Agent_ID
    If !cancelled = True Then
        DataEnvironment1.rssqlTem1!cancellation = True
        !cancelled = False
    ElseIf !Refund = True Then
        DataEnvironment1.rssqlTem1!Refund = True
        !Refund = False
    Else
    
    End If
    
    .Update
    DataEnvironment1.rssqlTem1.Update
    tr = MsgBox("Repayment was sucessfully cancelled", vbCritical, "OK")
    Exit Sub
EH:
    If .State = 1 Then .CancelUpdate: .Close
    If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.CancelUpdate: DataEnvironment1.rssqlTem1.Close
    If DataEnvironment1.rssqlTem2.State = 1 Then DataEnvironment1.rssqlTem2.CancelUpdate: DataEnvironment1.rssqlTem2.Close
    tr = MsgBox("An error occured. Cancellation was not done" & vbNewLine & Err.Description, vbCritical, "Error")
    Exit Sub
End With
End Sub

