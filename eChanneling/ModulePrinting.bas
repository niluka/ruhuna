Attribute VB_Name = "ModulePrinting"
Public Sub PrintMyShiftEndSummery()

DataReportCashIncome.Visible = False
dtrAgentCashReceive.Visible = False
dtrCredittBookingsPayment.Visible = False
DataReportCashRefunds.Visible = False
DataReportDoctorPayment.Visible = False
DataReportAgentBookings.Visible = False


With DataEnvironment1.rssqlCashireRepost
    If .State = 1 Then .Close
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.User_ID = " & UserID & ") and ((tblPatientFacility.PaymentMode = 'Cash') Or (tblPatientFacility.PaymentMode = 'Cheque'))and (tblPatientFacility.BookingDate = '" & Date & "')  order by tblPatientFacility.patientfacility_ID ")
    Set DataReportCashIncome.DataSource = DataEnvironment1.rssqlCashireRepost
    
    If HospitalDetails = True Then
        DataReportCashIncome.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        DataReportCashIncome.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        DataReportCashIncome.Sections("Section2").Controls.Item("RptCashierName").Caption = UserName
        DataReportCashIncome.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        DataReportCashIncome.Sections("Section2").Controls.Item("RptToDate").Caption = ""
        
       If .RecordCount <> 0 Then DataReportCashIncome.PrintReport False
    Else
        DataReportCashIncome.Sections("Section4").Controls.Item("RptName").Caption = Empty
        DataReportCashIncome.Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        DataReportCashIncome.Sections("Section2").Controls.Item("RptCashierName").Caption = UserName
        DataReportCashIncome.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        DataReportCashIncome.Sections("Section2").Controls.Item("RptToDate").Caption = ""
        
    End If

End With


With DataEnvironment1.rssqlAgentPayment1
    If .State = 1 Then .Close
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  Where (tblAgentCashSettle.User_ID = " & UserID & ") and (tblAgentCashSettle.SettledDate = '" & Date & "')   ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    Set dtrAgentCashReceive.DataSource = DataEnvironment1.rssqlAgentPayment1
    
    If HospitalDetails = True Then
        dtrAgentCashReceive.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        dtrAgentCashReceive.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        dtrAgentCashReceive.Sections("Section2").Controls.Item("RptCashierName").Caption = UserName
        dtrAgentCashReceive.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        dtrAgentCashReceive.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    Else
        dtrAgentCashReceive.Sections("Section4").Controls.Item("RptName").Caption = Empty
        dtrAgentCashReceive.Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        dtrAgentCashReceive.Sections("Section2").Controls.Item("RptCashierName").Caption = UserName
        dtrAgentCashReceive.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        dtrAgentCashReceive.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    End If
    
    
    
    If .RecordCount <> 0 Then dtrAgentCashReceive.PrintReport False
End With




With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.CreditSettleUser_ID = " & UserID & ") and (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.SettleCashDate = '" & Date & "')  order by tblPatientFacility.patientfacility_ID ")
    End With
    
    With dtrCredittBookingsPayment
        Set .DataSource = DataEnvironment1.rssqlTem10
        .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        .Sections("Section2").Controls.Item("RptCashierName").Caption = UserName
        .Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        .Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
        If DataEnvironment1.rssqlTem10.RecordCount <> 0 Then .PrintReport False
    End With
    
With DataEnvironment1.rssqlTem4
    If .State = 1 Then .Close
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.repayUser_ID = " & UserID & ") and ( (tblPatientFacility.cancelled = 1)or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate = '" & Date & "') and (RefundToPatient = 1) order by patientfacility_ID")
    Set DataReportCashRefunds.DataSource = DataEnvironment1.rssqlTem4
    DataReportCashRefunds.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportCashRefunds.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptCashierName").Caption = UserName
    DataReportCashRefunds.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    If .RecordCount <> 0 Then DataReportCashRefunds.PrintReport False
End With

With DataEnvironment1.rssqlTem9
    If .State = 1 Then .Close
        .Open ("Select tblstaffpayment.*,tbldoctor.* From tblstaffpayment Left Join tbldoctor On tblstaffpayment.Staff_ID = tbldoctor.Doctor_ID Where  (tblstaffpayment.User_ID = " & UserID & " and (tblstaffpayment.PaidDate = '" & Date & "' ))")
    Set DataReportDoctorPayment.DataSource = DataEnvironment1.rssqlTem9
    DataReportDoctorPayment.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportDoctorPayment.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptCashierName").Caption = UserName
    DataReportDoctorPayment.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    If .RecordCount <> 0 Then DataReportDoctorPayment.PrintReport False
End With

With DataEnvironment1.rssqlAgentReport
    If .State = 1 Then .Close
    .Open ("SELECT tblPatientFacility.*, tblInstitutions.* FROM tblPatientFacility LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (User_ID = " & UserID & ") and (PaymentMode = 'Agent') and (BookingDate = '" & Date & "') ")
    Set DataReportAgentBookings.DataSource = DataEnvironment1.rssqlAgentReport
End With
    If HospitalDetails = True Then
        DataReportAgentBookings.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        DataReportAgentBookings.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        DataReportAgentBookings.Sections("Section2").Controls.Item("RptCashierName").Caption = UserName
        DataReportAgentBookings.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        DataReportAgentBookings.Sections("Section2").Controls.Item("RptToDate").Caption = ""
        DataReportAgentBookings.Sections("Section2").Controls.Item("label6").Caption = ""
        DataReportAgentBookings.Sections("Section2").Controls.Item("label18").Caption = ""
        DataReportAgentBookings.Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
    Else
        DataReportAgentBookings.Sections("Section4").Controls.Item("RptName").Caption = Empty
        DataReportAgentBookings.Sections("Section4").Controls.Item("RptAddress").Caption = Empty
        DataReportAgentBookings.Sections("Section2").Controls.Item("RptCashierName").Caption = UserName
        DataReportAgentBookings.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        DataReportAgentBookings.Sections("Section2").Controls.Item("RptToDate").Caption = ""
        DataReportAgentBookings.Sections("Section2").Controls.Item("label6").Caption = ""
        DataReportAgentBookings.Sections("Section2").Controls.Item("label18").Caption = ""
        DataReportAgentBookings.Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
 End If
        
If DataEnvironment1.rssqlAgentReport.RecordCount <> 0 Then DataReportAgentBookings.PrintReport False
    
With DataEnvironment1.rssqlAgentReport
    If .State = 1 Then .Close
    .Open ("SELECT tblPatientFacility.*, tblInstitutions.* FROM tblPatientFacility LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (tblPatientFacility.repayUser_ID = " & UserID & ") and ( (tblPatientFacility.cancelled = 1)or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate = '" & Date & "') and (RefundToAgent = 1) order by patientfacility_ID")
    Set DataReportAgentBookings.DataSource = DataEnvironment1.rssqlAgentReport
End With
    DataReportAgentBookings.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportAgentBookings.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportAgentBookings.Sections("Section2").Controls.Item("RptCashierName").Caption = UserName
    DataReportAgentBookings.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    DataReportAgentBookings.Sections("Section2").Controls.Item("RptToDate").Caption = ""
    DataReportAgentBookings.Sections("Section2").Controls.Item("label6").Caption = ""
    DataReportAgentBookings.Sections("Section2").Controls.Item("label18").Caption = ""
If DataEnvironment1.rssqlAgentReport.RecordCount <> 0 Then DataReportAgentBookings.PrintReport False
    
    
    
    
Unload DataReportCashIncome
Unload dtrAgentCashReceive
Unload dtrCredittBookingsPayment
Unload DataReportCashRefunds
Unload DataReportDoctorPayment
Unload DataReportAgentBookings
    
End Sub

Public Sub PrintDayEndSummery()

DataReportCashIncome.Visible = False
dtrAgentCashReceive.Visible = False
dtrCredittBookingsPayment.Visible = False
DataReportCashRefunds.Visible = False
DataReportDoctorPayment.Visible = False
DataReportAgentBookings.Visible = False

With DataEnvironment1.rssqlCashireRepost

    If .State = 1 Then .Close
    
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where  (tblPatientFacility.PaymentMode = 'Cash')and (tblPatientFacility.BookingDate = '" & Date & "')  order by tblPatientFacility.patientfacility_ID ")
      
    Set DataReportCashIncome.DataSource = DataEnvironment1.rssqlCashireRepost
    
    DataReportCashIncome.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportCashIncome.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportCashIncome.Sections("Section2").Controls.Item("RptCashierName").Caption = "Day End Summery"
    DataReportCashIncome.Sections("Section2").Controls.Item("rptlHeding1").Caption = "Report :"
  
    DataReportCashIncome.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    DataReportCashIncome.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    
    If .RecordCount <> 0 Then DataReportCashIncome.PrintReport False

End With

With DataEnvironment1.rssqlAgentPayment1

    If .State = 1 Then .Close
    
    .Open ("SELECT tblAgentCashSettle.*, tblInstitutions.* FROM tblAgentCashSettle LEFT OUTER JOIN tblInstitutions ON tblAgentCashSettle.Institution_ID = tblInstitutions.Institution_ID  Where  (tblAgentCashSettle.SettledDate = '" & Date & "')   ORDER BY tblAgentCashSettle.AgentCashSettle_ID ")
    
    Set dtrAgentCashReceive.DataSource = DataEnvironment1.rssqlAgentPayment1
    
    
    dtrAgentCashReceive.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    dtrAgentCashReceive.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    dtrAgentCashReceive.Sections("Section2").Controls.Item("RptCashierName").Caption = "Day End Summery"
    dtrAgentCashReceive.Sections("Section2").Controls.Item("rptLHeding1").Caption = "Report :"

    dtrAgentCashReceive.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    dtrAgentCashReceive.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    
    If .RecordCount <> 0 Then dtrAgentCashReceive.PrintReport False

End With

With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where   (tblPatientFacility.PaymentMode = 'Credit')and (tblPatientFacility.BookingDate = '" & Date & "')  order by tblPatientFacility.patientfacility_ID ")
    
    With dtrCredittBookingsPayment
        Set .DataSource = DataEnvironment1.rssqlTem10
        
        .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
        .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
        .Sections("Section2").Controls.Item("RptCashierName").Caption = "Day End Summery"
        .Sections("Section2").Controls.Item("rptlHeding1").Caption = "Report :"
        .Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
        .Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
        
        If DataEnvironment1.rssqlTem10.RecordCount <> 0 Then .PrintReport False
    End With
    
End With
    
With DataEnvironment1.rssqlTem4

    If .State = 1 Then .Close
    
    .Open ("Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where ( (tblPatientFacility.cancelled = 1)or (tblPatientFacility.refund = 1 ))and (tblPatientFacility.RepayDate = '" & Date & "')")

    Set DataReportCashRefunds.DataSource = DataEnvironment1.rssqlTem4
    
    DataReportCashRefunds.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportCashRefunds.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptCashierName").Caption = "Day End Summery"
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptlHeding1").Caption = "Report :"

    DataReportCashRefunds.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    DataReportCashRefunds.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    
    If .RecordCount <> 0 Then DataReportCashRefunds.PrintReport False
    
End With
 
With DataEnvironment1.rssqlTem9
    If .State = 1 Then .Close
   
    .Open ("Select tblstaffpayment.*,tbldoctor.* From tblstaffpayment Left Join tbldoctor On tblstaffpayment.Staff_ID = tbldoctor.Doctor_ID Where  ((tblstaffpayment.PaidDate = '" & Date & "' ))")
    
'    If .RecordCount = 0 Then A = MsgBox("No Transaction to view", vbInformation + vbOKOnly, "No Transactions"): Exit Sub
    
    Set DataReportDoctorPayment.DataSource = DataEnvironment1.rssqlTem9
    
    DataReportDoctorPayment.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportDoctorPayment.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptCashierName").Caption = "Day End Summery"
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptlHeading1").Caption = "Report :"

    DataReportDoctorPayment.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    DataReportDoctorPayment.Sections("Section2").Controls.Item("RptToDate").Caption = Format(Date, DefaultLongDate)
    
    If .RecordCount <> 0 Then DataReportDoctorPayment.PrintReport False

End With
 
 
With DataEnvironment1.rssqlAgentReport
    If .State = 1 Then .Close
    .Open ("SELECT tblPatientFacility.*, tblInstitutions.* FROM tblPatientFacility LEFT JOIN tblInstitutions ON tblPatientFacility.Agent_ID = tblInstitutions.Institution_ID Where (PaymentMode = 'Agent') and (BookingDate = '" & Date & "') ")
    Set DataReportAgentBookings.DataSource = DataEnvironment1.rssqlAgentReport
End With
    DataReportAgentBookings.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportAgentBookings.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportAgentBookings.Sections("Section2").Controls.Item("RptCashierName").Caption = "All Cashiers"
    DataReportAgentBookings.Sections("Section2").Controls.Item("rptFromdate").Caption = Format(Date, DefaultLongDate)
    DataReportAgentBookings.Sections("Section2").Controls.Item("RptToDate").Caption = ""
    DataReportAgentBookings.Sections("Section2").Controls.Item("label6").Caption = ""
    DataReportAgentBookings.Sections("Section2").Controls.Item("label18").Caption = ""
If DataEnvironment1.rssqlAgentReport.RecordCount <> 0 Then DataReportAgentBookings.PrintReport False
 
Unload DataReportCashIncome
Unload dtrAgentCashReceive
Unload dtrCredittBookingsPayment
Unload DataReportCashRefunds
Unload DataReportDoctorPayment
Unload DataReportAgentBookings

End Sub


Public Function PrintingPlainText(ByVal StartX As Double, ByVal StartY As Double, ByVal TextToPrint As String) As Boolean
On Error Resume Next
With Printer
    .CurrentX = .ScaleWidth * StartX
    .CurrentY = .ScaleHeight * StartY
    Printer.Print TextToPrint
End With
End Function
