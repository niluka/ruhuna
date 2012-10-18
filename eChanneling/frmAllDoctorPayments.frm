VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDoctorPaymentsAll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Payments"
   ClientHeight    =   11145
   ClientLeft      =   375
   ClientTop       =   1755
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAllDoctorPayments.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FramePatientList 
      Height          =   8895
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   15015
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51052545
         CurrentDate     =   39462
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   13455
         Begin VB.OptionButton OptionPayable 
            Caption         =   "All Payable"
            Height          =   255
            Left            =   2400
            TabIndex        =   10
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton OptionPaid 
            Caption         =   "Paid"
            Height          =   255
            Left            =   8520
            TabIndex        =   7
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton OptionToPay 
            Caption         =   "To Pay"
            Height          =   255
            Left            =   5760
            TabIndex        =   6
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton OptionAll 
            Caption         =   "All"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5520
         TabIndex        =   0
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51052547
         CurrentDate     =   39421
      End
      Begin MSFlexGridLib.MSFlexGrid GridList1 
         Height          =   5415
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   9551
         _Version        =   393216
      End
      Begin btButtonEx.ButtonEx bttnCloseList 
         Height          =   375
         Left            =   13200
         TabIndex        =   19
         Top             =   8280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin btButtonEx.ButtonEx bttnView 
         Height          =   375
         Left            =   11520
         TabIndex        =   20
         Top             =   8280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "View"
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
      Begin btButtonEx.ButtonEx bttnPayDoctor 
         Height          =   375
         Left            =   9840
         TabIndex        =   21
         Top             =   8280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Pay Doctor"
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
      Begin VB.Label lblPaidToDoctor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   18
         Top             =   7800
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Payments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   7800
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancellations / Refunds Rs."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   7440
         Width           =   2655
      End
      Begin VB.Label lblTotalDoctorFee 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   7080
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Doctor Fee : Rs."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   7080
         Width           =   1935
      End
      Begin VB.Label lblDuePayments 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   8400
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Due : Rs."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   8400
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "List Criteria"
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
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "D&ate To :"
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
         Left            =   4560
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblDoctorStaff 
         BackStyle       =   0  'Transparent
         Caption         =   "&Date From"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblTotalRepayment 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   7440
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmDoctorPaymentsAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemPatientID As Long
    Dim TemHospitalFacilityID As Long
    Dim TemstaffID As Long
    Dim TemPatientFacilityID As Long
    Dim TemBillId As Long
    Dim IsACancellation As Boolean
    Dim IsARefund As Boolean
    Dim ChoosenOption As OptionButton
    
    Dim TemPreviousDate As Date
    Dim TemPreviousSecession As Long
    Dim TemPreviousDoctorID As Long
    Dim TemPreviousOptionChanged As Boolean
    Dim i

    

Private Sub bttnView_Click()
Dim TemResponce As Integer

If Val(lblDuePayments.Caption) <= 0 Then TemResponce = MsgBox("You have already Paid to doctor", vbCritical, "No Due Payments"): Exit Sub

With DataEnvironment1.rssqlTem14
    If .State = 1 Then .Close
    .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.secession = " & ListSecessionIDs.Text & " and tblPatientFacility.paidtostaff = false ORDER BY tblPatientFacility.DaySerial")
    Set DataReportDoctorPayments.DataSource = DataEnvironment1.rssqlTem14
    
    DataReportDoctorPayments.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportDoctorPayments.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportDoctorPayments.Sections("Section4").Controls.Item("Rptlbltopic").Caption = "Doctor Payment"
    DataReportDoctorPayments.Sections("Section2").Controls.Item("rptdoctorname").Caption = FindDoctorFromID(DataComboDoctorStaff.BoundText)
    DataReportDoctorPayments.Sections("Section2").Controls.Item("RptDate").Caption = Format(DTPickerAppointment.Value, "dd/MM/YYYY")
    DataReportDoctorPayments.Sections("Section2").Controls.Item("rptSecession").Caption = ListSecessions.Text
    DataReportDoctorPayments.Sections("Section5").Controls.Item("lblpaidamount").Caption = Format(lblDuePayments.Caption, "#0.00")
    DataReportDoctorPayments.Sections("Section5").Controls.Item("lblusername").Caption = FindStaffFromID(UserID)
    DataReportDoctorPayments.Sections("Section5").Controls.Item("lbldoctorname").Caption = FindDoctorFromID(DataComboDoctorStaff.BoundText)
DataReportDoctorPayments.Show
End With
End Sub

Private Sub Form_Load()
    Call Setcolours
    
    DTPicker2.Value = Date
    DTPicker1.Value = Date
    FormatPatientFacilityList
End Sub


Public Sub FormatPatientFacilityList()
    Dim BorderMargin As Long
    BorderMargin = 350
    With GridList1
        .Clear
        .Rows = 1
        .Row = 0
        .Cols = 12
        
        .ColWidth(0) = 600
        .Col = 0
        .CellAlignment = 4
        .Text = "Serial"
        
        .ColWidth(1) = 3000
        .Col = 1
        .CellAlignment = 4
        .Text = "Doctor Name"
    
        .ColWidth(2) = 1100
        .Col = 2
        .CellAlignment = 4
        .Text = "Fully Paid"
    
        .ColWidth(3) = 1100
        .Col = 3
        .CellAlignment = 4
        .Text = "Refunds"
        
        .ColWidth(4) = 1100
        .Col = 4
        .CellAlignment = 4
        .Text = "Doc. Fee"
        
        .ColWidth(5) = 1000
        .Col = 5
        .CellAlignment = 4
        .Text = "Hos. Fee"

        .ColWidth(6) = 1000
        .Col = 6
        .CellAlignment = 4
        .Text = "Other Fee"

        .ColWidth(7) = 1100
        .Col = 7
        .CellAlignment = 4
        .Text = "Doc. Refund"
        
        .ColWidth(8) = 1000
        .Col = 8
        .CellAlignment = 4
        .Text = "Hos. Refund"

        .ColWidth(9) = 1000
        .Col = 9
        .CellAlignment = 4
        .Text = "Other Refund"

        .ColWidth(10) = 1100
        .Col = 10
        .CellAlignment = 4
        .Text = "Paid to Doctor"
        
        .ColWidth(11) = 500
        .Col = 11
        .Text = "PatientBill_ID"
    End With
End Sub

Public Sub FillPatientFacilityList()
    
    Me.MousePointer = vbHourglass
    
    Dim NowRow As Long
    Dim TemNum As Long
    With DataEnvironment1.rssqlTem3
        If .State = 1 Then .Close
        
        If OptionAll.Value = True Then
            .Source = "select tblpatientfacility.*,tblDoctor.* from tblpatientfacility Left Join tblDoctor On tblpatientfacility.PatientFacility_ID = tblDoctor.Doctor_Id where (hospitalFacility_ID = 10) and (tblpatientfacility.appointmentdate between #" & DTPicker1.Value & "# And #" & DTPicker2.Value & "#) and .tblpatientfacility.fullypaid = true and personalfee > 0   order by tblpatientfacility.dayserial"
        ElseIf OptionPayable.Value = True Then
            .Source = "select tblpatientfacility.* from tblpatientfacility where (hospitalFacility_ID = 10) and (appointmentdate between #" & DTPicker1.Value & "# And #" & DTPicker2.Value & "#)  and fullypaid = true and personalfee > 0 order by dayserial"
        ElseIf OptionPaid.Value = True Then
            .Source = "select tblpatientfacility.* from tblpatientfacility where (hospitalFacility_ID = 10) and (appointmentdate between #" & DTPicker1.Value & "# And #" & DTPicker2.Value & "#)   and fullypaid = true and personalfee > 0  and paidtostaff = true order by dayserial"
        ElseIf OptionToPay.Value = True Then
            .Source = "select tblpatientfacility.* from tblpatientfacility where (hospitalFacility_ID = 10) and (appointmentdate between #" & DTPicker1.Value & "# And #" & DTPicker2.Value & "#) and  fullypaid = true and personalfee > 0  and paidtostaff = false order by dayserial"
        End If
        
        
        
        
        If .State = 0 Then .Open
        
        If .RecordCount = 0 Then
            GridList1.Visible = True
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        .MoveFirst
        
        NowRow = 1
        i = 1
        
        Do While .EOF = False
            NowRow = NowRow + 1
            
            GridList1.Rows = NowRow
            GridList1.Row = i
            
            GridList1.Col = 0
            GridList1.CellAlignment = 7
            GridList1.Text = i
            
            GridList1.Col = 1
            GridList1.CellAlignment = 1
            GridList1.Text = !DoctorListedName ' FindPatientByID(Val(!patientid))
        
            GridList1.Col = 2
            GridList1.CellAlignment = 7
            If !fullypaid = True Then
                GridList1.Text = "Yes"
            Else
                GridList1.Text = "No"
            End If
        
            GridList1.Col = 3
            GridList1.CellAlignment = 7
            If !cancelled = True Then
                GridList1.Text = "Cancelled"
            ElseIf !refund = True Then
                GridList1.Text = "Repaied"
            End If
            
            
            GridList1.Col = 4
            GridList1.CellAlignment = 7
            GridList1.Text = Format(!PersonalFee, "#0.00")
            
            GridList1.Col = 5
            GridList1.CellAlignment = 7
            GridList1.Text = Format(!InstitutionFee, "#0.00")
    
            GridList1.Col = 6
            GridList1.CellAlignment = 7
            GridList1.Text = Format(!otherfee, "#0.00")
    
            GridList1.Col = 7
            GridList1.CellAlignment = 7
            If Not IsNull(!personalrefund) Then
                GridList1.Text = Format(!personalrefund, "#0.00")
            Else
                GridList1.Text = Empty
            End If
            
            GridList1.Col = 8
            GridList1.CellAlignment = 7
            If Not IsNull(!InstitutionRefund) Then
                GridList1.Text = Format(!InstitutionRefund, "#0.00")
            Else
                GridList1.Text = Empty
            End If
    
            GridList1.Col = 9
            GridList1.CellAlignment = 7
             If Not IsNull(!OtherRefund) Then
                GridList1.Text = Format(!OtherRefund, "#0.00")
            Else
                GridList1.Text = Empty
            End If
            
            GridList1.Col = 10
            GridList1.CellAlignment = 7
            If !paidtostaff = True Then
                GridList1.Text = "Yes"
            Else
                GridList1.Text = "No"
            End If
                 
            GridList1.Col = 11
            GridList1.Text = !PatientBill_ID
            
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    With GridList1
    
    Dim TemDoctorTotalFee As Double
    Dim TemHospitalTotalFee As Double
    Dim TemOtherTotalFee As Double
    Dim TemDoctorTotalRepayment As Double
    Dim TemHospitalTotalRepayment As Double
    Dim TemOtherTotalRepayment As Double
    Dim TemPaidToDoctor As Double

    TemDoctorTotalFee = 0
    TemHospitalTotalFee = 0
    TemOtherTotalFee = 0
    TemDoctorTotalRepayment = 0
    TemHospitalTotalRepayment = 0
    TemOtherTotalRepayment = 0
    TemPaidToDoctor = 0

    For TemNum = 1 To GridList1.Rows - 1
        .Col = 0
        .Row = TemNum
        .Text = TemNum

        .Col = 4
        TemDoctorTotalFee = TemDoctorTotalFee + Val(.Text)

        If .TextMatrix(TemNum, 11) <> Empty Then
            TemPaidToDoctor = TemPaidToDoctor + Val(.Text) - Val(.TextMatrix(TemNum, 7))
        Else
            TemPaidToDoctor = TemPaidToDoctor + Val(.Text)
        End If

        .Col = 5
        TemHospitalTotalFee = TemHospitalTotalFee + Val(.Text)

        .Col = 6
        TemOtherTotalFee = TemOtherTotalFee + Val(.Text)

        .Col = 7
        TemDoctorTotalRepayment = TemDoctorTotalRepayment + Val(.Text)

        .Col = 8
        TemHospitalTotalRepayment = TemHospitalTotalRepayment + Val(.Text)

        .Col = 9
        TemOtherTotalRepayment = TemOtherTotalRepayment + Val(.Text)

    Next

    .Rows = .Rows + 1
    .Row = .Rows - 1

        .Col = 4
        .CellAlignment = 7
        .Text = Format(TemDoctorTotalFee, "#0.00")

        .Col = 5
        .CellAlignment = 7
        .Text = Format(TemHospitalTotalFee, "#0.00")

        .Col = 6
        .CellAlignment = 7
        .Text = Format(TemOtherTotalFee, "#0.00")

        .Col = 7
        .CellAlignment = 7
        .Text = Format(TemDoctorTotalRepayment, "#0.00")

        .Col = 8
        .CellAlignment = 7
        .Text = Format(TemHospitalTotalRepayment, "#0.00")

        .Col = 9
        .CellAlignment = 7
        .Text = Format(TemOtherTotalRepayment, "#0.00")

    lblPaidToDoctor.Caption = Format(TemPaidToDoctor, "#0.00")
    lblTotalDoctorFee.Caption = Format(TemDoctorTotalFee, "#0.00")
    lblTotalRepayment.Caption = Format(TemDoctorTotalRepayment, "#0.00")

    lblDuePayments.Caption = Format((TemDoctorTotalFee - TemDoctorTotalRepayment - TemPaidToDoctor), "#0.00")



    .Row = 0
    .Col = 0


    End With

    TemPreviousOptionChanged = False
    
    GridList1.Visible = True
    Me.MousePointer = vbDefault
    
End Sub

Private Sub bttnCloseList_Click()
    Unload Me
End Sub

Private Sub bttnPayDoctor_Click()
Dim TemResponce As Integer

'If Val(lblDuePayments.Caption) <= 0 Then TemResponce = MsgBox("You have already Paid to doctor", vbCritical, "No Due Payments"): Exit Sub


With DataEnvironment1.rssqlTem14
    If .State = 1 Then .Close
    .Open ("SELECT tblPatientFacility.*,tblPatientMainDetails.* FROM tblPatientFacility LEFT JOIN tblPatientMainDetails ON tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID WHERE (tblPatientFacility.HospitalFacility_ID = 10) AND (tblPatientFacility.FullyPaid = true) AND (tblPatientFacility.PersonalFee > 0) and tblPatientFacility.paidtostaff = false ORDER BY tblPatientFacility.DaySerial")
    Set DataReportDoctorPayments.DataSource = DataEnvironment1.rssqlTem14
    
    DataReportDoctorPayments.Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    DataReportDoctorPayments.Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    DataReportDoctorPayments.Sections("Section4").Controls.Item("Rptlbltopic").Caption = "Doctor Payment"
    DataReportDoctorPayments.Sections("Section2").Controls.Item("rptdoctorname").Caption = FindDoctorFromID(DataComboDoctorStaff.BoundText)
    DataReportDoctorPayments.Sections("Section2").Controls.Item("RptDate").Caption = Format(DTPickerAppointment.Value, "dd/MM/YYYY")
    DataReportDoctorPayments.Sections("Section2").Controls.Item("rptSecession").Caption = ListSecessions.Text
    DataReportDoctorPayments.Sections("Section5").Controls.Item("lblpaidamount").Caption = Format(lblDuePayments.Caption, "#0.00")
    DataReportDoctorPayments.Sections("Section5").Controls.Item("lblusername").Caption = FindStaffFromID(UserID)
    DataReportDoctorPayments.Sections("Section5").Controls.Item("lbldoctorname").Caption = FindDoctorFromID(DataComboDoctorStaff.BoundText)
DataReportDoctorPayments.PrintReport False

End With

End Sub

Private Sub DataComboDoctorStaff_Change()
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    Call FormatSecessionsList
    Call FillSecessionsList
    Call FillPatientFacilityList
End Sub

Private Sub DTPickerAppointment_Change()
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    FormatSecessionsList
    FillSecessionsList
    FillPatientFacilityList
End Sub

Private Sub ListSecessions_Click()
    If ListSecessions.ListIndex < 0 Then Exit Sub
    ListSecessionIDs.ListIndex = ListSecessions.ListIndex
    If Not IsNumeric(DataComboDoctorStaff.BoundText) Then Exit Sub
    FillPatientFacilityList
End Sub

Private Sub OptionAgentBookings_Click()
    If OptionAgentBookings.Value = True Then
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionAll_Click()
    If OptionAll.Value = True Then
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionCancellations_Click()
    If OptionCancellations.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionCancellationsAndRefunds_Click()
    If OptionCancellationsAndRefunds.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionCashBookings_Click()
    If OptionCashBookings.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionFullyPaid_Click()
    If OptionFullyPaid.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub


Private Sub Setcolours()
    bttnCloseList.BackColor = BttnBackColour
    bttnCloseList.ForeColor = BttnForeColour
    bttnView.BackColor = BttnBackColour
    bttnView.ForeColor = BttnForeColour
    FramePatientList.BackColor = FrmBackColour
    FramePatientList.ForeColor = FrmForeColour
    Me.BackColor = FrameBackColour
    Me.ForeColor = FrameForeColour
'    DataComboDoctorStaff.BackColor = FrmBackColour
'    DataComboDoctorStaff.ForeColor = FrmForeColour
    OptionPayable.BackColor = FrmBackColour
    OptionPayable.ForeColor = FrmForeColour
    OptionAll.BackColor = FrmBackColour
    OptionAll.ForeColor = FrmForeColour
    OptionToPay.BackColor = FrmBackColour
    OptionToPay.ForeColor = FrmForeColour
    OptionPaid.BackColor = FrmBackColour
    OptionPaid.ForeColor = FrmForeColour
    Frame1.BackColor = FrmBackColour
    Frame1.ForeColor = FrmForeColour
End Sub

Private Sub OptionRefunds_Click()
    If OptionRefunds.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionWithoutCancellation_Click()
    If OptionWithoutCancellation.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionPaid_Click()
    If OptionPaid.Value = True Then
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionPayable_Click()
    If OptionPayable.Value = True Then
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionToPay_Click()
    If OptionToPay.Value = True Then
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionWithoutCancellationsAndRefunds_Click()
    If OptionWithoutCancellationsAndRefunds.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub

Private Sub OptionWithoutRefunds_Click()
    If OptionWithoutRefunds.Value = True Then
        
        FillPatientFacilityList
    End If
    TemPreviousOptionChanged = True
End Sub
