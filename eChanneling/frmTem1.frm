VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmTem1 
   Caption         =   "Form1"
   ClientHeight    =   10785
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
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
   ScaleHeight     =   10785
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSAgentOtherFee 
      Height          =   360
      Left            =   4080
      MaxLength       =   250
      TabIndex        =   18
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtSForeignOtherFee 
      Height          =   360
      Left            =   4080
      MaxLength       =   250
      TabIndex        =   17
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtSLocalOtherFee 
      Height          =   360
      Left            =   4080
      MaxLength       =   250
      TabIndex        =   16
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtSAgentHospitalFee 
      Height          =   360
      Left            =   2400
      MaxLength       =   250
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtSAgentDoctorFee 
      Height          =   360
      Left            =   600
      MaxLength       =   250
      TabIndex        =   12
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtSFogrignerHospitalFee 
      Height          =   360
      Left            =   2400
      MaxLength       =   250
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtSForeginerDoctorFee 
      Height          =   360
      Left            =   600
      MaxLength       =   250
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtSLocalHospitalFee 
      Height          =   360
      Left            =   2400
      MaxLength       =   250
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtSLocalDoctorFee 
      Height          =   360
      Left            =   600
      MaxLength       =   250
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   6960
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   5340
      Left            =   7440
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   5340
      Left            =   7680
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin btButtonEx.ButtonEx bttnCloseList 
      Height          =   375
      Left            =   13440
      TabIndex        =   0
      Top             =   10200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Close List"
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Other Fee"
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Fee"
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Fee"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lblDocAgentFeeAgentRepaymentsO 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblDocAgentFeeCashRepaymentsO 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label lblDocFeeBookingsAgentO 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblDocFeeAgentO 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmTem1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblstaff"
    .Open
    While .EOF = False
        List1.AddItem DecreptedWord(!StaffUserName)
        List2.AddItem DecreptedWord(!staffpassword)
        .MoveNext
    Wend
    .Close
End With
End Sub




Private Sub bttnPrint_Click()
    Dim temSQL As String
    Dim TemWhere As String
    Dim TemCash As Double
    
' *******************************************************
    temSQL = "SELECT tblPatientFacility.PatientFacility_ID, tblDoctor.DoctorName, tblPatientMainDetails.FirstName, tblPatientFacility.PersonalFee, tblPatientFacility.InstitutionFee, tblPatientFacility.TotalFee, tblPatientFacility.BookingDate, tblPatientFacility.BookingTime "
    temSQL = temSQL & " FROM tblDoctor RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID "
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) "
    With DataEnvironment1.rssqlTem20
        If .State = 1 Then .Close
        .Source = "Delete from tbltemreport1" ' where user_ID = " & UserID
        .Open '("delete from tbltemreport1 where user_ID = " & UserID)
        If .State = 1 Then .Close
        .Source = " select * from tbltemreport1"
        .Open
    End With
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            While .EOF = False
                DataEnvironment1.rssqlTem20.AddNew
                DataEnvironment1.rssqlTem20!Catogery = "Cash Patients"
                DataEnvironment1.rssqlTem20!patientfacility_ID = !patientfacility_ID
                DataEnvironment1.rssqlTem20!doctorname = !doctorname
                DataEnvironment1.rssqlTem20!FirstName = !FirstName
                DataEnvironment1.rssqlTem20!PersonalFee = !PersonalFee
                DataEnvironment1.rssqlTem20!InstitutionFee = !InstitutionFee
                DataEnvironment1.rssqlTem20!totalfee = !totalfee
                DataEnvironment1.rssqlTem20!BookingTime = !BookingTime
                DataEnvironment1.rssqlTem20!user_ID = UserID
                DataEnvironment1.rssqlTem20!TotalIncome = !totalfee
                DataEnvironment1.rssqlTem20.Update
                .MoveNext
            Wend
        End If
    End With
' *******************************************************
    temSQL = "SELECT tblPatientFacility.PatientFacility_ID, tblDoctor.DoctorName, tblPatientMainDetails.FirstName, tblPatientFacility.PersonalFee, tblPatientFacility.InstitutionFee, tblPatientFacility.TotalFee, tblPatientFacility.SettleCashDate, tblPatientFacility.SettleCashTime "
    temSQL = temSQL & " FROM tblDoctor RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID "
    TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.PaymentMode)='Credit')) "
    With DataEnvironment1.rssqlTem20
        If .State = 1 Then .Close
        .Source = " select * from tbltemreport1"
        .Open
    End With
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            While .EOF = False
                DataEnvironment1.rssqlTem20.AddNew
                DataEnvironment1.rssqlTem20!Catogery = "Cash Settling For Credit Patients"
                DataEnvironment1.rssqlTem20!patientfacility_ID = !patientfacility_ID
                DataEnvironment1.rssqlTem20!doctorname = !doctorname
                DataEnvironment1.rssqlTem20!FirstName = !FirstName
                DataEnvironment1.rssqlTem20!PersonalFee = !PersonalFee
                DataEnvironment1.rssqlTem20!InstitutionFee = !InstitutionFee
                DataEnvironment1.rssqlTem20!totalfee = !totalfee
                DataEnvironment1.rssqlTem20!BookingTime = !SettleCashTime
                DataEnvironment1.rssqlTem20!user_ID = UserID
                DataEnvironment1.rssqlTem20!TotalIncome = !totalfee
                DataEnvironment1.rssqlTem20.Update
                .MoveNext
            Wend
        End If
    End With
' *******************************************************
    temSQL = "SELECT tblPatientFacility.PatientFacility_ID, tblDoctor.DoctorName, tblPatientMainDetails.FirstName, tblPatientFacility.PersonalRefund, tblPatientFacility.InstitutionRefund, tblPatientFacility.TotalRefund, tblPatientFacility.RepayDate, tblPatientFacility.RepayTime, tblPatientFacility.Cancelled, tblPatientFacility.Refund "
    temSQL = temSQL & " FROM tblDoctor RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
    With DataEnvironment1.rssqlTem20
        If .State = 1 Then .Close
        .Source = " select * from tbltemreport1"
        .Open
    End With
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            While .EOF = False
                DataEnvironment1.rssqlTem20.AddNew
                DataEnvironment1.rssqlTem20!Catogery = "Cash Repayments"
                DataEnvironment1.rssqlTem20!patientfacility_ID = !patientfacility_ID
                DataEnvironment1.rssqlTem20!doctorname = !doctorname
                DataEnvironment1.rssqlTem20!FirstName = !FirstName
                DataEnvironment1.rssqlTem20!Personalrefund = !Personalrefund
                DataEnvironment1.rssqlTem20!institutionrefund = !institutionrefund
                DataEnvironment1.rssqlTem20!totalrefund = !totalrefund
                DataEnvironment1.rssqlTem20!BookingTime = !repaytime
                DataEnvironment1.rssqlTem20!user_ID = UserID
                DataEnvironment1.rssqlTem20!TotalIncome = -!totalrefund
                If !Cancelled = True Then
                    DataEnvironment1.rssqlTem20!remarks = "Cancellation"
                ElseIf !Refund = True Then
                    DataEnvironment1.rssqlTem20!remarks = "Refund"
                End If
                DataEnvironment1.rssqlTem20.Update
                .MoveNext
            Wend
        End If
    End With
' *******************************************************
    temSQL = "SELECT tblPatientFacility.PatientFacility_ID, tblDoctor.DoctorName, tblPatientMainDetails.FirstName, tblPatientFacility.PersonalFee, tblPatientFacility.InstitutionFee, tblPatientFacility.TotalFee, tblPatientFacility.BookingDate, tblPatientFacility.BookingTime, tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode"
    temSQL = temSQL & " FROM tblInstitutions RIGHT JOIN (tblDoctor RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID "
    TemWhere = "WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent'))"
    With DataEnvironment1.rssqlTem20
        If .State = 1 Then .Close
        .Source = " select * from tbltemreport1"
        .Open
    End With
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            While .EOF = False
                DataEnvironment1.rssqlTem20.AddNew
                DataEnvironment1.rssqlTem20!Catogery = "Agent Bookings"
                DataEnvironment1.rssqlTem20!patientfacility_ID = !patientfacility_ID
                DataEnvironment1.rssqlTem20!doctorname = !doctorname
                DataEnvironment1.rssqlTem20!FirstName = !FirstName
                DataEnvironment1.rssqlTem20!PersonalFee = !PersonalFee
                DataEnvironment1.rssqlTem20!InstitutionFee = !InstitutionFee
                DataEnvironment1.rssqlTem20!totalfee = !totalfee
                DataEnvironment1.rssqlTem20!agent = !InstitutionName
                DataEnvironment1.rssqlTem20!agentcode = !InstitutionCode
                DataEnvironment1.rssqlTem20!BookingTime = !BookingTime
                DataEnvironment1.rssqlTem20!user_ID = UserID
                DataEnvironment1.rssqlTem20.Update
                .MoveNext
            Wend
        End If
    End With
' *******************************************************
    temSQL = "SELECT tblPatientFacility.PatientFacility_ID, tblDoctor.DoctorName, tblPatientMainDetails.FirstName, tblPatientFacility.PersonalRefund, tblPatientFacility.InstitutionRefund, tblPatientFacility.TotalRefund, tblPatientFacility.RepayDate, tblPatientFacility.RepayTime, tblPatientFacility.Cancelled, tblPatientFacility.Refund,  tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode "
    temSQL = temSQL & " FROM tblInstitutions RIGHT JOIN (tblDoctor RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToAgent)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
    
    With DataEnvironment1.rssqlTem20
        If .State = 1 Then .Close
        .Source = " select * from tbltemreport1"
        .Open
    End With
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            While .EOF = False
                DataEnvironment1.rssqlTem20.AddNew
                DataEnvironment1.rssqlTem20!Catogery = "Agent Repayments"
                DataEnvironment1.rssqlTem20!patientfacility_ID = !patientfacility_ID
                DataEnvironment1.rssqlTem20!doctorname = !doctorname
                DataEnvironment1.rssqlTem20!FirstName = !FirstName
                DataEnvironment1.rssqlTem20!Personalrefund = !Personalrefund
                DataEnvironment1.rssqlTem20!institutionrefund = !institutionrefund
                DataEnvironment1.rssqlTem20!totalrefund = !totalrefund
                DataEnvironment1.rssqlTem20!BookingTime = !repaytime
                DataEnvironment1.rssqlTem20!user_ID = UserID
                DataEnvironment1.rssqlTem20!agent = !InstitutionName
                DataEnvironment1.rssqlTem20!agentcode = !InstitutionCode
                If !Cancelled = True Then
                    DataEnvironment1.rssqlTem20!remarks = "Cancellation"
                ElseIf !Refund = True Then
                    DataEnvironment1.rssqlTem20!remarks = "Refund"
                End If
                DataEnvironment1.rssqlTem20.Update
                .MoveNext
            Wend
        End If
    End With
    lblAgentRepayments.Caption = Format(TemCash, "00.00")
' ***********************************************
    temSQL = "SELECT tblStaffPayment.* From tblStaffPayment "
    TemWhere = ""
    
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        TemCash = 0
        If .RecordCount <> 0 Then
            While .EOF = False
                TemCash = TemCash + !PaidAmount
                .MoveNext
            Wend
        End If
    End With
    
    Const PreSHape As String = "SHAPE { "
    Dim Sql As String
    Const PostSHape As String = "}  AS cmmdTemReport1 COMPUTE cmmdTemReport1, ANY(cmmdTemReport1.'Catogery') AS NameOfCatogery, SUM(cmmdTemReport1.'PersonalFee') AS GrandPersonalIncome, SUM(cmmdTemReport1.'InstitutionFee') AS GrandInstitutionIncome, SUM(cmmdTemReport1.'TotalFee') AS GrandTotalIncome, SUM(cmmdTemReport1.'TotalIncome') AS GrandCashCollection, SUM(cmmdTemReport1.'PersonalRefund') AS GrandPersonalExpence, SUM(cmmdTemReport1.'InstitutionRefund') AS GrandInstitutionExpence, SUM(cmmdTemReport1.'TotalRefund') AS GrandTotalExpence BY 'Catogery'"
    Sql = "SELECT tblTemReport1.* FROM tblTemReport1" ' where user_ID = " & UserID
    With DataEnvironment1
        If .rscmmdTemReport1_Grouping.State = 1 Then .rscmmdTemReport1_Grouping.Close
        .Commands!cmmdTemReport1_Grouping.CommandText = PreSHape & Sql & PostSHape
        .cmmdTemReport1_Grouping
        
        
    dtrNewUserSummery.Sections("PageFooter").Controls.Item("lblAds").Caption = LongAd
    dtrNewUserSummery.Sections("PageHeader").Controls.Item("lblDate").Caption = "Date    " & Date & "     " & "Time   " & Time
    dtrNewUserSummery.Sections("PageHeader").Controls.Item("RptCashierName").Caption = "Cashier Name   :  " & UserName
        
        
        Set dtrNewUserSummery.DataSource = DataEnvironment1
        dtrNewUserSummery.Show
        
    End With
End Sub


Private Sub CalculateIncome()
    
    
    
    Exit Sub
    
    Dim TemCash As Double
    Dim TemCash1 As Double
    Dim temSQL As String
    Dim TemWhere As String
    
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum (TotalFee) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblCashBookings.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum (TotalFee) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.PaymentMode)='Credit')) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblSettlingCredit.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum (Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblCashRepayments.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum (TotalFee) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = "WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Agent'))"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblAgentBoolings.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToAgent)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1))"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblAgentRepayments.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(cash) as TotalGrand "
    temSQL = temSQL & " FROM tblagentcashsettle "
    TemWhere = " where SettledDate = '" & Date & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblAgentCashPayments.Caption = Format(TemCash, "0.00")
' ***********************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(paidamount) as TotalGrand "
    temSQL = temSQL & " FROM tblstaffpayment "
    TemWhere = " where PaidDate = '" & Date & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblDoctorPayments.Caption = Format(TemCash, "0.00")
' ***********************************************
    TemCash = 0
    TemCash1 = 0
    lblDoctorPayments.Caption = Format(TemCash, "0.00")
    TemCash = lblCashBookings.Caption
    TemCash = TemCash + Val(lblSettlingCredit.Caption)
    TemCash = TemCash + Val(lblAgentCashPayments.Caption)
    TemCash = TemCash - Val(lblCashRepayments.Caption)
    TemCash = TemCash - Val(lblDoctorPayments.Caption)
    lblNetCash.Caption = Format(TemCash, "0.00")
    
   
' Doctor Cash ******************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum (personalFee) as TotalGrand ,  sum(institutionfee) as TotalInstitutionFee "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!TotalInstitutionFee) Then TemCash1 = !TotalInstitutionFee
        End If
    End With
    lblDoctorCashDC.Caption = Format(TemCash, "0.00")
    lblHospitalCashDC.Caption = Format(TemCash1, "0.00")
    temSQL = "SELECT sum (personalFee) as TotalGrand,  sum(institutionfee) as TotalInstitutionFee "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Credit')) "
    TemCash = 0
    TemCash1 = 0
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
            If Not IsNull(!TotalInstitutionFee) Then TemCash1 = !TotalInstitutionFee
        End If
    End With
    lblDoctorCashSC.Caption = Format(TemCash, "0.00")
    lblHospitalCashSC.Caption = Format(TemCash1, "0.00")
    lblDoctorCash.Caption = Format(Val(lblDoctorCashDC.Caption) + Val(lblDoctorCashSC.Caption), "0.00")
    lblHospitalCash.Caption = Format(Val(lblHospitalCashDC.Caption) + Val(lblHospitalCashSC.Caption), "0.00")
' Doctor Cash By AppointmentDate
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum (personalFee) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) and appointmentdate = '" & Date & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblTOdaysDoctorCashDC.Caption = Format(TemCash, "0.00")
    
    temSQL = "SELECT sum (personalFee) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Credit')) and appointmentdate = '" & Date & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash1 = !TotalGrand
        End If
    End With
    lblTodaysDoctorCashSC.Caption = Format(TemCash1, "0.00")
    
    lblTodaysDoctorCash.Caption = Format((Val(lblTOdaysDoctorCashDC.Caption) + Val(lblTodaysDoctorCashSC.Caption)), "0.00")
    ListDOctorCash.Clear
    ListDOctorCash.AddItem "Date     " & vbTab & vbTab & "Direct Cash" & vbTab & "Credit settling" & vbTab & "Total"
    Dim TemNum As Long
    Dim temText As String
    Dim TemCashDC As Double
    Dim TemCashSC As Double
    Dim TemCash2 As Double
    Dim TemCash3 As Double
    Dim TemMaxDate As Date
    Dim TemDate As Date
    
    With DataEnvironment1.rssqlTem
        temSQL = "Select max(appointmentDate) as MaxBookingDate from tblpatientfacility "
        TemWhere = " where (paymentmode = 'Cash' and bookingdate ='" & Date & "') or (paymentmode = 'Credit' and settlecashdate = '" & Date & "')"
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If Not IsNull(!MaxBookingDate) Then
            TemMaxDate = !MaxBookingDate
        Else
            TemMaxDate = Date
        End If
        If .RecordCount > 0 Then
            TemNum = 1
            TemCash = 0
            TemCashDC = 0
            TemCashSC = 0
            While Date + TemNum <= TemMaxDate
                TemDate = Date + TemNum
                TemCash1 = 0
                TemCash2 = 0
                TemCash3 = 0
                
                temSQL = "SELECT sum (personalFee) as TotalGrand "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) and appointmentdate = '" & TemDate & "'"
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!TotalGrand) Then TemCash1 = !TotalGrand: TemCashDC = TemCashDC + TemCash1
                    End If
                temSQL = "SELECT sum (personalFee) as TotalGrand "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Credit')) and appointmentdate = '" & TemDate & "'"
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!TotalGrand) Then TemCash2 = !TotalGrand: TemCashSC = TemCashSC + TemCash2
                    End If
                TemCash3 = TemCash1 + TemCash2
                If TemCash3 > 0 Then
                   temText = Format(TemDate, DefaultShortDate)
                   temText = temText & vbTab & vbTab & Right(Format(TemCash1, "0.00"), 10)
                   temText = temText & vbTab & vbTab & Right(Format(TemCash2, "0.00"), 10)
                   temText = temText & vbTab & vbTab & Right(Format(TemCash3, "0.00"), 10)
                   ListDOctorCash.AddItem temText
                   TemCash = TemCash + TemCash3
                End If
                TemNum = TemNum + 1
            Wend
        End If
        If .State = 1 Then .Close
    End With
    lblOtherdaysDoctorCashDC.Caption = Format(TemCashDC, "0.00")
    lblOtherDaysDoctorCashSC.Caption = Format(TemCashSC, "0.00")
    lblOtherDaysDoctorCash.Caption = Format(TemCash, "0.00")

' ************* TotalRepayments ****************
            
        
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum (Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblRepaidToPatient.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToAgent)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1))"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblRepaidToAgent.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE tblPatientFacility.RepayDate='" & Date & "' AND tblPatientFacility.HospitalFacility_ID=10  AND (tblPatientFacility.Cancelled=1 or tblPatientFacility.Refund=1)"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblTotalRepayments.Caption = Format(TemCash, "0.00")



End Sub

Private Sub tem1()
' Doctor Cash By AppointmentDate
    TemCash = 0
    TemCash1 = 0
    TemCash2 = 0
    temSQL = "SELECT sum (personalFee) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) and appointmentdate = '" & Date & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblTOdaysDoctorCashDC.Caption = Format(TemCash, "0.00")
    temSQL = "SELECT sum (personalFee) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Credit')) and appointmentdate = '" & Date & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash1 = !TotalGrand
        End If
    End With
    lblTodaysDoctorCashSC.Caption = Format(TemCash1, "0.00")
    temSQL = "SELECT sum(personalrefund) as TotalDoctorFee "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and appointmentdate = '" & Date & "'"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!totaldoctorfee) Then TemCash2 = !totaldoctorfee
        End If
    End With
    lblTodaysDoctorCashRepay.Caption = Format(TemCash2, "0.00")
    lblTodaysDoctorCash.Caption = Format((Val(lblTOdaysDoctorCashDC.Caption) + Val(lblTodaysDoctorCashSC.Caption) - Val(lblTodaysDoctorCashRepay.Caption)), "0.00")
    ListDOctorCash.Clear
    ListDOctorCash.AddItem "Date     " & vbTab & "Direct Cash" & vbTab & "Credit settling" & vbTab & "Cash Repayments" & vbTab & vbTab & "Total"
    With DataEnvironment1.rssqlTem
        temSQL = "Select max(appointmentDate) as MaxBookingDate , min(appointmentdate) as MinBookingDate from tblpatientfacility "
        TemWhere = " where (paymentmode = 'Cash' and bookingdate ='" & Date & "') or (paymentmode = 'Credit' and settlecashdate = '" & Date & "') or ( repaydate = '" & Date & "'  ) "
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If Not IsNull(!MaxBookingDate) Then
            TemMaxDate = !MaxBookingDate
        Else
            TemMaxDate = Date
        End If
        If Not IsNull(!minbookingdate) Then
            TemMinDate = !minbookingdate
        Else
            TemMinDate = Date
        End If
        If .RecordCount > 0 Then
            TemNum = 0
            TemCash = 0
            TemCashDC = 0
            TemCashSC = 0
            TemCashRepay = 0
            TemDate = TemMinDate
            While TemMinDate + TemNum <= TemMaxDate
                TemCash1 = 0
                TemCash2 = 0
                TemCash3 = 0
                TemCash4 = 0
                temSQL = "SELECT sum (personalFee) as TotalGrand "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Cash')) and appointmentdate = '" & TemDate & "'"
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!TotalGrand) Then TemCash1 = !TotalGrand: TemCashDC = TemCashDC + TemCash1
                    End If
                temSQL = "SELECT sum (personalFee) as TotalGrand "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.PaymentMode)='Credit')) and appointmentdate = '" & TemDate & "'"
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!TotalGrand) Then TemCash2 = !TotalGrand: TemCashSC = TemCashSC + TemCash2
                    End If
                temSQL = "SELECT sum(personalrefund) as TotalDoctorFee "
                temSQL = temSQL & " FROM tblPatientFacility "
                TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) and appointmentdate = '" & TemDate & "'"
                With DataEnvironment1.rssqlTem1
                    If .State = 1 Then .Close
                    .Source = temSQL & TemWhere
                    .Open
                    If .RecordCount <> 0 Then
                        If Not IsNull(!totaldoctorfee) Then TemCash4 = !totaldoctorfee: TemCashRepay = TemCashRepay + !totaldoctorfee
                    End If
                End With
                
                TemCash3 = TemCash1 + TemCash2 - TemCash4
                
                If TemCash1 + TemCash2 + TemCash4 > 0 Then
                   temText = Format(TemDate, DefaultShortDate)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash1, "0.00"), 10)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash2, "0.00"), 10)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash4, "0.00"), 10)
                   temText = temText & vbTab & Right(Space(10) & Format(TemCash3, "0.00"), 10)
                   ListDOctorCash.AddItem temText
                   TemCash = TemCash + TemCash3
                End If
                TemNum = TemNum + 1
                TemDate = TemMinDate + TemNum
            Wend
        End If
        If .State = 1 Then .Close
    End With
    lblOtherdaysDoctorCashDC.Caption = Format(TemCashDC - Val(lblTOdaysDoctorCashDC.Caption), "0.00")
    lblOtherDaysDoctorCashSC.Caption = Format(TemCashSC - Val(lblTodaysDoctorCashSC.Caption), "0.00")
    lblOtherDaysDoctorCashRepay.Caption = Format(TemCashRepay - Val(lblTodaysDoctorCashRepay.Caption), "0.00")
    lblOtherDaysDoctorCash.Caption = Format(TemCash - Val(lblTodaysDoctorCash.Caption), "0.00")
' ******************************************











































' ************* TotalRepayments ****************
            
        
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum (Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblRepaidToPatient.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10)  AND ((tblPatientFacility.RefundToAgent)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1))"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblRepaidToAgent.Caption = Format(TemCash, "0.00")
' *******************************************************
    TemCash = 0
    TemCash1 = 0
    temSQL = "SELECT sum(Totalrefund) as TotalGrand "
    temSQL = temSQL & " FROM tblPatientFacility "
    TemWhere = " WHERE tblPatientFacility.RepayDate='" & Date & "' AND tblPatientFacility.HospitalFacility_ID=10  AND (tblPatientFacility.Cancelled=1 or tblPatientFacility.Refund=1)"
    With DataEnvironment1.rssqlTem1
        If .State = 1 Then .Close
        .Source = temSQL & TemWhere
        .Open
        If .RecordCount <> 0 Then
            If Not IsNull(!TotalGrand) Then TemCash = !TotalGrand
        End If
    End With
    lblTotalRepayments.Caption = Format(TemCash, "0.00")



End Sub
