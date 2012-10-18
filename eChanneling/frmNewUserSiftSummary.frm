VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmNewUserShiftSummary 
   Caption         =   "User Sift Summary"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
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
   Icon            =   "frmNewUserSiftSummary.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   12240
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Print"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Bindings        =   "frmNewUserSiftSummary.frx":038A
      Height          =   6735
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   23
      DataMember      =   "cmmdTemReport1_Grouping"
      _NumberOfBands  =   2
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   9
      _Band(0)._MapCol(0)._Name=   "Catogery"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(0)._Hidden=   -1  'True
      _Band(0)._MapCol(1)._Name=   "NameOfCatogery"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "GrandPersonalIncome"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "GrandInstitutionIncome"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "GrandTotalIncome"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "GrandPersonalExpence"
      _Band(0)._MapCol(5)._RSIndex=   6
      _Band(0)._MapCol(6)._Name=   "GrandInstitutionExpence"
      _Band(0)._MapCol(6)._RSIndex=   7
      _Band(0)._MapCol(7)._Name=   "GrandTotalExpence"
      _Band(0)._MapCol(7)._RSIndex=   8
      _Band(0)._MapCol(8)._Name=   "GrandCashCollection"
      _Band(0)._MapCol(8)._RSIndex=   5
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   14
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   0
      _Band(1)._ParentBand=   0
      _Band(1)._NumMapCols=   18
      _Band(1)._MapCol(0)._Name=   "ID"
      _Band(1)._MapCol(0)._RSIndex=   0
      _Band(1)._MapCol(0)._Alignment=   7
      _Band(1)._MapCol(0)._Hidden=   -1  'True
      _Band(1)._MapCol(1)._Name=   "Catogery"
      _Band(1)._MapCol(1)._RSIndex=   1
      _Band(1)._MapCol(1)._Hidden=   -1  'True
      _Band(1)._MapCol(2)._Name=   "PaymentMethod"
      _Band(1)._MapCol(2)._RSIndex=   2
      _Band(1)._MapCol(2)._Hidden=   -1  'True
      _Band(1)._MapCol(3)._Name=   "PatientFacility_ID"
      _Band(1)._MapCol(3)._RSIndex=   3
      _Band(1)._MapCol(3)._Alignment=   7
      _Band(1)._MapCol(4)._Name=   "DoctorName"
      _Band(1)._MapCol(4)._RSIndex=   4
      _Band(1)._MapCol(5)._Name=   "FirstName"
      _Band(1)._MapCol(5)._RSIndex=   5
      _Band(1)._MapCol(6)._Name=   "Agent"
      _Band(1)._MapCol(6)._RSIndex=   6
      _Band(1)._MapCol(7)._Name=   "AgentCode"
      _Band(1)._MapCol(7)._RSIndex=   7
      _Band(1)._MapCol(8)._Name=   "PersonalFee"
      _Band(1)._MapCol(8)._RSIndex=   8
      _Band(1)._MapCol(8)._Alignment=   7
      _Band(1)._MapCol(9)._Name=   "InstitutionFee"
      _Band(1)._MapCol(9)._RSIndex=   9
      _Band(1)._MapCol(9)._Alignment=   7
      _Band(1)._MapCol(10)._Name=   "TotalFee"
      _Band(1)._MapCol(10)._RSIndex=   10
      _Band(1)._MapCol(10)._Alignment=   7
      _Band(1)._MapCol(11)._Name=   "PersonalRefund"
      _Band(1)._MapCol(11)._RSIndex=   11
      _Band(1)._MapCol(11)._Alignment=   7
      _Band(1)._MapCol(12)._Name=   "InstitutionRefund"
      _Band(1)._MapCol(12)._RSIndex=   12
      _Band(1)._MapCol(12)._Alignment=   7
      _Band(1)._MapCol(13)._Name=   "TotalRefund"
      _Band(1)._MapCol(13)._RSIndex=   13
      _Band(1)._MapCol(13)._Alignment=   7
      _Band(1)._MapCol(14)._Name=   "BookingTime"
      _Band(1)._MapCol(14)._RSIndex=   14
      _Band(1)._MapCol(15)._Name=   "Remarks"
      _Band(1)._MapCol(15)._RSIndex=   15
      _Band(1)._MapCol(16)._Name=   "User_ID"
      _Band(1)._MapCol(16)._RSIndex=   16
      _Band(1)._MapCol(16)._Alignment=   7
      _Band(1)._MapCol(16)._Hidden=   -1  'True
      _Band(1)._MapCol(17)._Name=   "TotalIncome"
      _Band(1)._MapCol(17)._RSIndex=   17
      _Band(1)._MapCol(17)._Alignment=   7
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   13680
      TabIndex        =   0
      Top             =   7080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmNewUserShiftSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CSetPrinter As New cSetDfltPrinter

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnPrint_Click()
Const PreSHape As String = "SHAPE { "
Dim Sql As String
Const PostSHape As String = "}  AS cmmdTemReport1 COMPUTE cmmdTemReport1, ANY(cmmdTemReport1.'Catogery') AS NameOfCatogery, SUM(cmmdTemReport1.'PersonalFee') AS GrandPersonalIncome, SUM(cmmdTemReport1.'InstitutionFee') AS GrandInstitutionIncome, SUM(cmmdTemReport1.'TotalFee') AS GrandTotalIncome, SUM(cmmdTemReport1.'TotalIncome') AS GrandCashCollection, SUM(cmmdTemReport1.'PersonalRefund') AS GrandPersonalExpence, SUM(cmmdTemReport1.'InstitutionRefund') AS GrandInstitutionExpence, SUM(cmmdTemReport1.'TotalRefund') AS GrandTotalExpence BY 'Catogery'"
Sql = "SELECT tblTemReport1.* FROM tblTemReport1 where user_ID = " & UserID
CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    
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

Private Sub Form_Load()
Call CalculateIncome
End Sub

Private Sub CalculateIncome()
    Dim temSQL As String
    Dim TemWhere As String
    
' *******************************************************
    temSQL = "SELECT tblPatientFacility.PatientFacility_ID, tblDoctor.DoctorName, tblPatientMainDetails.FirstName, tblPatientFacility.PersonalFee, tblPatientFacility.InstitutionFee, tblPatientFacility.TotalFee, tblPatientFacility.BookingDate, tblPatientFacility.BookingTime "
    temSQL = temSQL & " FROM tblDoctor RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID "
    TemWhere = " WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.User_ID)=" & UserID & ") AND ((tblPatientFacility.PaymentMode)='Cash')) "
    
    With DataEnvironment1.rssqlTem20
        If .State = 1 Then .Close
        .Source = "Delete from tbltemreport1 where user_ID = " & UserID
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
    TemWhere = " WHERE (((tblPatientFacility.SettleCashDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.CreditSettleUser_ID)=" & UserID & ") AND ((tblPatientFacility.PaymentMode)='Credit')) "
    
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
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ")  AND ((tblPatientFacility.RefundToPatient)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
    
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
    TemWhere = "WHERE (((tblPatientFacility.BookingDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.User_ID)=" & UserID & ") AND ((tblPatientFacility.PaymentMode)='Agent'))"
    
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
'                DataEnvironment1.rssqlTem20!totalincome = !totalfee
                DataEnvironment1.rssqlTem20.Update
                .MoveNext
            Wend
        End If
    End With
    


' *******************************************************

    temSQL = "SELECT tblPatientFacility.PatientFacility_ID, tblDoctor.DoctorName, tblPatientMainDetails.FirstName, tblPatientFacility.PersonalRefund, tblPatientFacility.InstitutionRefund, tblPatientFacility.TotalRefund, tblPatientFacility.RepayDate, tblPatientFacility.RepayTime, tblPatientFacility.Cancelled, tblPatientFacility.Refund,  tblInstitutions.InstitutionName, tblInstitutions.InstitutionCode "
    temSQL = temSQL & " FROM tblInstitutions RIGHT JOIN (tblDoctor RIGHT JOIN (tblPatientMainDetails RIGHT JOIN tblPatientFacility ON tblPatientMainDetails.Patient_ID = tblPatientFacility.PatientID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID "
    TemWhere = " WHERE (((tblPatientFacility.RepayDate)='" & Date & "') AND ((tblPatientFacility.HospitalFacility_ID)=10) AND ((tblPatientFacility.RepayUser_ID)=" & UserID & ")  AND ((tblPatientFacility.RefundToAgent)=1))  AND (((tblPatientFacility.Cancelled)=1) or ((tblPatientFacility.Refund)=1)) "
    
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
'                DataEnvironment1.rssqlTem20!totalincome = -!totalrefund
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

' ***********************************************
    
    Const PreSHape As String = "SHAPE { "
    Dim Sql As String
    Const PostSHape As String = "}  AS cmmdTemReport1 COMPUTE cmmdTemReport1, ANY(cmmdTemReport1.'Catogery') AS NameOfCatogery, SUM(cmmdTemReport1.'PersonalFee') AS GrandPersonalIncome, SUM(cmmdTemReport1.'InstitutionFee') AS GrandInstitutionIncome, SUM(cmmdTemReport1.'TotalFee') AS GrandTotalIncome, SUM(cmmdTemReport1.'TotalIncome') AS GrandCashCollection, SUM(cmmdTemReport1.'PersonalRefund') AS GrandPersonalExpence, SUM(cmmdTemReport1.'InstitutionRefund') AS GrandInstitutionExpence, SUM(cmmdTemReport1.'TotalRefund') AS GrandTotalExpence BY 'Catogery'"
    'SHAPE {SELECT tblTemReport1.* FROM tblTemReport1}  AS cmmdTemReport1 COMPUTE cmmdTemReport1, ANY(cmmdTemReport1.'Catogery') AS NameOfCatogery, SUM(cmmdTemReport1.'PersonalFee') AS GrandPersonalIncome, SUM(cmmdTemReport1.'InstitutionFee') AS GrandInstitutionIncome, SUM(cmmdTemReport1.'TotalFee') AS GrandTotalIncome, SUM(cmmdTemReport1.'TotalIncome') AS GrandCashCollection, SUM(cmmdTemReport1.'PersonalRefund') AS GrandPersonalExpence, SUM(cmmdTemReport1.'InstitutionRefund') AS GrandInstitutionExpence, SUM(cmmdTemReport1.'TotalRefund') AS GrandTotalExpence BY 'Catogery'
    Sql = "SELECT tblTemReport1.* FROM tblTemReport1 where user_ID = " & UserID
    With DataEnvironment1
        If .rscmmdTemReport1_Grouping.State = 1 Then .rscmmdTemReport1_Grouping.Close
        .Commands!cmmdTemReport1_Grouping.CommandText = PreSHape & Sql & PostSHape
        .cmmdTemReport1_Grouping
        Set Grid1.DataSource = DataEnvironment1
    End With
    With Grid1
        .CollapseAll
        .FormatString = "     |         Catogery                    | Doctor Income | Hospital Income | Total Income | Repayments(Doctor Fee) | Repayments(Hospital Fee) | Total Repayments | Net Cash Collection | Receipt | Patient | Doctor | Doctor Fee | Hospital Fee | Total Fee |  Repayments(Doctor Fee) | Repayments(Hospital Fee) | Total Repayments | Net Cash Collection "
        
        .ColAlignmentHeader(0, 0) = 4
        .ColAlignmentHeader(0, 1) = 4
        .ColAlignmentHeader(0, 2) = 4
        .ColAlignmentHeader(0, 3) = 4
        .ColAlignmentHeader(0, 4) = 4
        
        .ColAlignmentBand(0, 1) = 1
        .ColAlignmentBand(0, 2) = 7
        .ColAlignmentBand(0, 3) = 7
        .ColAlignmentBand(0, 4) = 7
        .ColAlignmentBand(0, 5) = 7
        
    End With
End Sub
