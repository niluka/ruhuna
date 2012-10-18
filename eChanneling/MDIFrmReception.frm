VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIFrmReception 
   BackColor       =   &H8000000C&
   Caption         =   "Lakmedipro eHospital - Reception"
   ClientHeight    =   10665
   ClientLeft      =   60
   ClientTop       =   870
   ClientWidth     =   11400
   Icon            =   "MDIFrmReception.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   12360
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   12240
      Top             =   1320
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmReception.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmReception.frx":9A69E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmReception.frx":135CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmReception.frx":1CFFF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmReception.frx":26B63A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmReception.frx":306C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmReception.frx":37B82A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New Bookings"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Past Bookings"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Doctor Payments"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Shift End Summery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Day End Summery"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Channeling Scheduling"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   12360
         ScaleHeight     =   315
         ScaleWidth      =   2595
         TabIndex        =   1
         Top             =   120
         Width           =   2655
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   2535
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuChannelling 
      Caption         =   "&Channelling"
      Begin VB.Menu mnuNewChanneling 
         Caption         =   "New Bookings"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuDoctorPayment 
         Caption         =   "Pay Doctors"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTraceBookings 
         Caption         =   "Past Bookings"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuOldBookings 
         Caption         =   "Old Bookings"
      End
      Begin VB.Menu mnuChannallingLists 
         Caption         =   "Channelling Lists"
      End
      Begin VB.Menu mnuDoctorLeave 
         Caption         =   "Doctor Leave"
      End
      Begin VB.Menu mnuChangeHospitalFee 
         Caption         =   "Change Hospital Fee"
      End
      Begin VB.Menu mnuChannellingSheduling 
         Caption         =   "Channeling Sheduling"
      End
      Begin VB.Menu mnuDoctorSecessionView 
         Caption         =   "Doctor Secession View"
      End
      Begin VB.Menu mnuAnnouncements 
         Caption         =   "Announcements"
      End
   End
   Begin VB.Menu mnuIncome 
      Caption         =   "Back &Office"
      Begin VB.Menu mnuShiftSummery 
         Caption         =   "Shift Summery"
         Shortcut        =   ^S
      End
      Begin VB.Menu munuDayendSummery 
         Caption         =   "Day End Summery"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuNewDayEndSummery 
         Caption         =   "Day-End Summeries"
      End
      Begin VB.Menu mnuNewShiftEndSummery 
         Caption         =   "Shift-End Summeries"
      End
      Begin VB.Menu mnuBackOfficeAllBookings 
         Caption         =   "All Bookings"
      End
      Begin VB.Menu mnuBackOfficeAllAppointments 
         Caption         =   "All Appointments"
      End
      Begin VB.Menu mnuBackOfficeCredit 
         Caption         =   "Credit"
         Begin VB.Menu mnuBackOfficeCreditBookings 
            Caption         =   "Credit Bookings"
         End
         Begin VB.Menu mnuBackOfficeCreditAppointments 
            Caption         =   "Credit Appointments"
         End
      End
      Begin VB.Menu mnuAgent 
         Caption         =   "Agent "
         Begin VB.Menu mnuPatientCounts 
            Caption         =   "Patient Counts"
         End
         Begin VB.Menu mnuAgentSummery 
            Caption         =   "All Agent Summery"
         End
         Begin VB.Menu mnuAgentDetails 
            Caption         =   "Selected Agent Summery"
         End
         Begin VB.Menu mnuAgentBookings 
            Caption         =   "Agent Bookings"
         End
         Begin VB.Menu mnuAllAgentTransactions 
            Caption         =   "All Agent Transactions"
         End
         Begin VB.Menu mnuSingleAgentTransactions 
            Caption         =   "Single Agent Transctions"
         End
         Begin VB.Menu mnuAgentReferranceNumbers 
            Caption         =   "Agent Referance Numbers"
         End
         Begin VB.Menu mnuCalculateDailyAgentBalance 
            Caption         =   "Calculate Daily Agent Balance"
         End
         Begin VB.Menu mnuAgentBookingChange 
            Caption         =   "Agent Booking Change"
         End
         Begin VB.Menu mnuAgentAgentPaymentCancellation 
            Caption         =   "All Agent Payment Cancellation"
         End
      End
      Begin VB.Menu mnuchannelingReport 
         Caption         =   "Doctor Reports"
         Begin VB.Menu mnuSecessionViceIncome 
            Caption         =   "Secession-vice Income"
         End
         Begin VB.Menu mnuChanelingPatients 
            Caption         =   "Channelling Patients"
         End
         Begin VB.Menu mnuDoctorsDetails 
            Caption         =   "Doctors Details"
         End
      End
      Begin VB.Menu mnuPayments 
         Caption         =   "Payments"
         Visible         =   0   'False
         Begin VB.Menu mnuTodaysPayments 
            Caption         =   "Doctor Payments - Today"
         End
         Begin VB.Menu mnuSelectedDaysDoctorPayments 
            Caption         =   "Doctor Payments - Selected days"
         End
         Begin VB.Menu mnuDoctorsPayment 
            Caption         =   "Doctor Payments (Summery)"
         End
         Begin VB.Menu mnuDoctorPaymentDetails 
            Caption         =   "Doctor Payments (Details)"
         End
      End
      Begin VB.Menu mnuBookings 
         Caption         =   "Bookings"
      End
      Begin VB.Menu mnuAppointments 
         Caption         =   "Appointments"
      End
      Begin VB.Menu mnuAbsenties 
         Caption         =   "Absenties"
         Begin VB.Menu mnuAbsentPatientList 
            Caption         =   "Absent PatientList"
         End
      End
      Begin VB.Menu mnuCancelRepayments 
         Caption         =   "Cancel Repayments"
         Begin VB.Menu mnuCancelRepaymentReports 
            Caption         =   "Cancel Repayment Reports"
         End
         Begin VB.Menu mnuCancelARepayment 
            Caption         =   "Cancel A Repayment"
         End
      End
      Begin VB.Menu mnuDoctorPatientLists 
         Caption         =   "Doctor Patient Lists"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDoctorPatientCount 
         Caption         =   "Doctor Patient Count"
      End
      Begin VB.Menu mnuBackOfficeManageRecords 
         Caption         =   "Manage Records"
         Begin VB.Menu mnuManageRecordsRefundReport 
            Caption         =   "Refund Report"
         End
         Begin VB.Menu mnuBackOfficeWHT 
            Caption         =   "WHT"
         End
         Begin VB.Menu mnuBulkRefund 
            Caption         =   "B. Refund"
         End
      End
      Begin VB.Menu mnuScanIncome 
         Caption         =   "Scan Income - Doctorvice"
      End
      Begin VB.Menu mnuScanIncomeAll 
         Caption         =   "Scan Income - All"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User"
      Begin VB.Menu mnuMyShiftSummary 
         Caption         =   "My Shift Details"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuShiftEndSummery 
         Caption         =   "New My Shift End Summery"
      End
      Begin VB.Menu mnuMySummary 
         Caption         =   "New Day End Summery"
      End
      Begin VB.Menu mnuTodaysMyBookings 
         Caption         =   "Today's Bookings"
      End
      Begin VB.Menu mnuTodaysMyCreditBookings 
         Caption         =   "Today's Credit Bookings"
      End
      Begin VB.Menu mnuTodaysCreditAppointments 
         Caption         =   "Today's Credit Appointments"
      End
      Begin VB.Menu mnuTodaysAllBookings 
         Caption         =   "Todays' All Bookings"
      End
      Begin VB.Menu mnuTodayEndSummery 
         Caption         =   "Day End Summery"
      End
   End
   Begin VB.Menu mnuDataEntering 
      Caption         =   "&Data Entering"
      Begin VB.Menu mnuPatients 
         Caption         =   "P&atients"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDoctors 
         Caption         =   "&Doctors"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuStaff 
         Caption         =   "Sta&ff"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuAuthorityPrevilages 
         Caption         =   "Authority Previlages"
         Begin VB.Menu mnuAuthorityPrevilagesMenuVisible 
            Caption         =   "Visibility of Menus"
         End
      End
      Begin VB.Menu mnuInstitutions 
         Caption         =   "Instit&utions"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuSpeciality 
         Caption         =   "Consultant Specialities"
         Shortcut        =   +^{F2}
      End
      Begin VB.Menu mnuSpecialityStaff 
         Caption         =   "Staff Specialities"
         Shortcut        =   +^{F1}
      End
   End
   Begin VB.Menu mnuPayment 
      Caption         =   "Pa&yment"
      Begin VB.Menu mnuInstitutionPayment 
         Caption         =   "Agent &Payment"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuAgentCancellation 
         Caption         =   "Agent Payment Cancellation"
      End
   End
   Begin VB.Menu mnuPreferances 
      Caption         =   "Pr&eferances"
      Begin VB.Menu mnuProgramPreferances 
         Caption         =   "Program Preferances"
      End
      Begin VB.Menu mnuPrintingPreferances 
         Caption         =   "Printing Preferances"
      End
      Begin VB.Menu mnuOwnerPreferances 
         Caption         =   "Owner Preferances"
      End
      Begin VB.Menu mnuLoggin 
         Caption         =   "Loggin"
      End
      Begin VB.Menu mnuInstitutionPreferances 
         Caption         =   "Institution Preferances"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTOC 
         Caption         =   "Table of Contexts"
      End
      Begin VB.Menu mnuTipOfDay 
         Caption         =   "Tip Of Day"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "Details"
      End
      Begin VB.Menu mnuAdministrator 
         Caption         =   "Administrator"
      End
   End
End
Attribute VB_Name = "MDIFrmReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim SuppliedWord As String
    
    Private Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
        (ByVal hWndCaller As Long, _
         ByVal pszFile As String, _
         ByVal uCommand As Long, _
         dwData As Any) As Long
    
    Const HH_DISPLAY_TOPIC As Long = 0
    Const HH_HELP_CONTEXT As Long = &HF

    Dim CSetPrinter As New cSetDfltPrinter


Private Sub SetColour()
    Select Case ColourScheme
        Case 1:
            BttnBackColour = 5341695
            BttnForeColour = 1314458
            FrmBackColour = 11066623
            FrmForeColour = 1314458
            FrameBackColour = 11066623
            FrameForeColour = 1314458
            TxtBackColour = 9881851
            TxtForeColour = 1314458
            LblBackColour = 11066623
            LblForeColour = 1314458
            GridBackColor = 9881855
            GridBackColorBkg = 10474239
            GridBackColorFixed = 8566015
            GridBackColorSel = 5341695
            GridForeColor = 1314458
            GridForeColorFixed = 11944
            GridForeColorSel = 3014824
    Case 2:
            BttnBackColour = 14803300
            BttnForeColour = 5539362
            FrmBackColour = 16766120
            FrmForeColour = 5539362
            FrameBackColour = 16766120
            FrameForeColour = 5539362
            TxtBackColour = 16760450
            TxtForeColour = 5539362
            LblBackColour = 16766120
            LblForeColour = 5539362
            GridBackColor = 16760450
            GridBackColorBkg = 16771260
            GridBackColorFixed = 16105620
            GridBackColorSel = 16737380
            GridForeColor = 5539362
            GridForeColorFixed = 5539362
            GridForeColorSel = 16765588
    Case 3:
            BttnBackColour = 51455
            BttnForeColour = 942490
            FrmBackColour = 11070719
            FrmForeColour = 942490
            FrameBackColour = 11070719
            FrameForeColour = 942490
            TxtBackColour = 11528439
            TxtForeColour = 1314458
            LblBackColour = 11070719
            LblForeColour = 942490
            GridBackColor = 16760450
            GridBackColorBkg = 16771260
            GridBackColorFixed = 16105620
            GridBackColorSel = 16737380
            GridForeColor = 5539362
            GridForeColorFixed = 5539362
            GridForeColorSel = 16765588
    End Select

    MDIFrmReception.BackColor = FrameBackColour

End Sub


Private Sub MDIForm_Load()
    App.HelpFile = App.Path & "\help.chm"
    Call SetColour
    If GetSetting(App.EXEName, "Options", "Show Tips at Startup", True) = True Then
        frmTip.Show
    End If
'    Call EnableControls(Me)
 '   Call VisibleControls(Me)
    Call WriteBalance
End Sub




Private Sub OldFormLoad()
            
        
' ************************************
    mnuAnnouncements.Enabled = False
' *************************************
 If UserAuthority = AuthorityOwnerCOvered Then
 mnuDoctorPatientCount.Visible = True
 Else
 mnuDoctorPatientCount.Visible = False
 End If
    
 Select Case UserAuthority
 
 Case AuthorityAdministrator
 
 Case AuthorityOwner
    'mnuInstitutionPreferances.Enabled = False
    
 Case AuthorityAccount
 
 Case AuthorityHumanResources
    mnuOwnerPreferances.Visible = False
    mnuProgramPreferances.Visible = False
    mnuChangeHospitalFee.Visible = False
    mnuCancelARepayment.Visible = False
    mnuBackOfficeManageRecords.Visible = False
    
 Case AuthorityOwnerCOvered
    mnuOwnerPreferances.Visible = False
    mnuInstitutionPayment.Visible = False
    mnuAgentSummery.Visible = False
    mnuAgentDetails.Visible = False
    mnuAgentBookings.Visible = False
    mnuchannelingReport.Visible = False
    mnuPayment.Visible = False
    mnuPayments.Visible = False
    mnuPreferances.Visible = False
    mnuStaff.Visible = False
    mnuAgent.Visible = False
    mnuDoctorPayment.Visible = False
    mnuNewDayEndSummery.Visible = False
    mnuCancelARepayment.Visible = False
    mnuAbsenties.Visible = False
    mnuAllAgentTransactions.Visible = False
    mnuChannallingLists.Visible = False
    mnuTraceBookings.Visible = False
    mnuBackOfficeManageRecords.Visible = False

 Case AuthorityUser
    mnuChangeHospitalFee.Visible = False
    mnuShiftSummery.Visible = False
    mnuOwnerPreferances.Visible = False
    mnuProgramPreferances.Visible = False
    mnuInstitutionPreferances.Visible = False
    mnuAgentCancellation.Visible = False
    mnuCancelARepayment.Visible = False
    mnuTraceBookings.Visible = False
    mnuChannallingLists.Visible = False
    mnuAgent.Visible = False
    mnuDoctorPayment.Visible = False
    mnuFile.Visible = False
    mnuShiftSummery.Visible = False
    munuDayendSummery.Visible = False
    mnuNewDayEndSummery.Visible = False
    mnuNewShiftEndSummery.Visible = False
    mnuAgent.Visible = False
    mnuBookings.Visible = False
    mnuAppointments.Visible = False
    mnuCancelRepayments.Visible = False
    mnuDoctorPatientLists.Visible = False
    mnuDoctorPatientCount.Visible = False
    mnuMyShiftSummary.Visible = False
    mnuTodaysMyBookings.Visible = False
    mnuTodaysAllBookings.Visible = False
    mnuStaff.Visible = False
    mnuAnnouncements.Visible = False
    mnuBackOfficeManageRecords.Visible = False
    
    mnuAgent.Visible = True
    mnuSingleAgentTransactions.Visible = True
    mnuAllAgentTransactions.Visible = True
    
    mnuPatientCounts.Visible = False
    mnuAgentBookings.Visible = False
    'mnuAllAgentTransactions.Visible = False
    'mnuSingleAgentTransactions.Visible = False
    mnuAgentReferranceNumbers.Visible = False
    mnuCalculateDailyAgentBalance.Visible = False
    mnuAgentBookingChange.Visible = False
    mnuAgentAgentPaymentCancellation.Visible = False
    
 Case AuthorityAnalyzer
    mnuInstitutionPayment.Enabled = False
    mnuNewChanneling.Enabled = False
    mnuDoctorPayment.Enabled = False
    mnuChangeHospitalFee.Enabled = False
    mnuChannallingLists.Enabled = False
    mnuDoctorLeave.Enabled = False
    mnuChangeHospitalFee.Enabled = False
    mnuChannellingSheduling.Enabled = False
    mnuInstitutionPreferances.Enabled = False
    mnuOwnerPreferances.Enabled = False
    mnuProgramPreferances.Enabled = False
    mnuInstitutionPreferances.Enabled = False
    mnuPayment.Visible = False
    mnuCancelARepayment.Enabled = False
    mnuBackOfficeManageRecords.Visible = False
 End Select

End Sub


Private Sub WriteBalance()
    Dim rsTem As New ADODB.Recordset
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "Select * from tblinstitutions order by institution_ID"
        .Open
        If .RecordCount > 0 Then
            While .EOF = False
                If DataEnvironment1.rssqlTem1.State = 1 Then DataEnvironment1.rssqlTem1.Close
                DataEnvironment1.rssqlTem1.Source = "SELECT tblInstitutionBalance.InstitutionBalance_Id, tblInstitutionBalance.Institution_Id, tblInstitutionBalance.Date, tblInstitutionBalance.SBalance, tblInstitutionBalance.EBalance From tblInstitutionBalance where tblInstitutionBalance.Institution_Id = " & !Institution_Id & " AND tblInstitutionBalance.Date = '" & Format(Date, "dd MMMM yyyy") & "'"
                DataEnvironment1.rssqlTem1.Open
                If DataEnvironment1.rssqlTem1.RecordCount < 1 Then
                    DataEnvironment1.rssqlTem1.AddNew
                    DataEnvironment1.rssqlTem1!Institution_Id = !Institution_Id
                    DataEnvironment1.rssqlTem1!Date = Format(Date, "dd MMMM yyyy")
                    DataEnvironment1.rssqlTem1!SBalance = !InstitutionCredit
                    DataEnvironment1.rssqlTem1!EBalance = !InstitutionCredit
                    DataEnvironment1.rssqlTem1.Update
                Else
                    DataEnvironment1.rssqlTem1!EBalance = !InstitutionCredit
                    DataEnvironment1.rssqlTem1.Update
                End If
                .MoveNext
            Wend
        End If
        .Close
        DataEnvironment1.rssqlTem1.Close
    End With
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim TemResponce  As Integer
    Dim TemForm As Form
    Dim AllFOrms As Form
    Call WriteBalance
'    TemResponce = MsgBox("Back up the database?", vbQuestion + vbYesNo, "Backup")
'    If TemResponce = vbYes Then
'        frmBackUpOnExit.Show 1
'    End If
    
'    TemResponce = MsgBox("Do you want to print the shift end summery before log off ?", vbQuestion + vbYesNo, "Print shift end summery?")
'    If TemResponce = vbYes Then
'        MDIFrmReception.mnuPrintMyShiftEndSummery_Click
'    End If
'
'    TemResponce = MsgBox("Print the day end summery report?", vbQuestion + vbYesNo, "Day end summery")
'    If TemResponce = vbYes Then
'        PrintDayEndSummery
'    End If
    
    TemResponce = MsgBox("Are you sure you want to exit Lakmedipro eChanneling ?", vbInformation + vbYesNo, "EXIT?")
    If TemResponce = vbNo Then Cancel = True: Exit Sub
    For Each TemForm In Forms
        Unload TemForm
    Next
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
With DataEnvironment1.rssqlTem
    If CheckLogin = True Then
        If .State = 1 Then .Close
        .Source = "select * from tblstaff where staff_ID = " & UserID
        .Open
            If .RecordCount <> 0 Then
                !logged = 0
                .Update
            End If
        End If
    If .State = 1 Then .Close
End With
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub


Private Sub mnuAbsentPatientList_Click()
On Error GoTo ErrorHandler
    frmAbsantPatientList.Show
    frmAbsantPatientList.ZOrder 0
    frmAbsantPatientList.Top = 0
    frmAbsantPatientList.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuAdministrator_Click()
    frmAdministrator.Show
    frmAdministrator.ZOrder 0
End Sub

Private Sub mnuAgentAgentPaymentCancellation_Click()
    frmAllAgentPaymentsCancellation.Show
    frmAllAgentPaymentsCancellation.ZOrder 0
    frmAllAgentPaymentsCancellation.Top = 0
    frmAllAgentPaymentsCancellation.Left = 0
End Sub

Private Sub mnuAgentBookingChange_Click()
    frmAgentBookingChange.Show
    frmAgentBookingChange.ZOrder 0
End Sub

Private Sub mnuAgentBookings_Click()
On Error GoTo ErrorHandler
'If UserAuthority <> AuthorityOwner And UserAuthority <> AuthorityAdministrator Then Exit Sub
    frmAgentBookings.ZOrder 0
    frmAgentBookings.Top = 0
    frmAgentBookings.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuAgentCancellation_Click()

On Error GoTo ErrorHandler
    frmCancellationAgentPayments.Show
    frmCancellationAgentPayments.ZOrder 0
    frmCancellationAgentPayments.Top = 0
    frmCancellationAgentPayments.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub


End Sub

Private Sub mnuAgentDetails_Click()
On Error GoTo ErrorHandler
'If UserAuthority <> AuthorityOwner And UserAuthority <> AuthorityAdministrator Then Exit Sub

    frmAgentSummery.ZOrder 0
    frmAgentSummery.Top = 0
    frmAgentSummery.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuAgentSummary_Click()
On Error GoTo ErrorHandler

    frmNewAgentSummary.Show
    frmNewAgentSummary.ZOrder 0
    frmNewAgentSummary.Top = 0
    frmNewAgentSummary.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub
'frmNewAgentSummary
End Sub

Private Sub mnuAgentReferranceNumbers_Click()
On Error GoTo ErrorHandler
    frmAgentRefranceNo.ZOrder 0
    frmAgentRefranceNo.Top = 0
    frmAgentRefranceNo.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuAgentSummery_Click()
'If UserAuthority <> AuthorityOwner And UserAuthority <> AuthorityAdministrator Then Exit Sub
On Error GoTo ErrorHandler

    frmAllAgentSummery.ZOrder 0
    frmAllAgentSummery.Top = 0
    frmAllAgentSummery.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuAllAgentTransactions_Click()
On Error GoTo ErrorHandler
    frmAllAgentTransactions.Show
    frmAllAgentTransactions.ZOrder 0
    frmAllAgentTransactions.Top = 0
    frmAllAgentTransactions.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuAnnouncements_Click()
On Error GoTo ErrorHandler
    frmAnnouncements.ZOrder 0
    frmAnnouncements.Top = 0
    frmAnnouncements.Left = 0
    frmAnnouncements.Show
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuAppointments_Click()
On Error GoTo ErrorHandler
    If UserAuthority = AuthorityOwnerCOvered Then
        frmAppointmentsC.ZOrder 0
        frmAppointmentsC.Top = 0
        frmAppointmentsC.Left = 0
        frmAppointmentsC.Show
    Else
        frmAppointments.ZOrder 0
        frmAppointments.Top = 0
        frmAppointments.Left = 0
        frmAppointments.Show
    End If
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuAuthorityPrevilagesControlEnable_Click()
    frmAuthorityPrevilagesControlEnable.Show
    frmAuthorityPrevilagesControlEnable.ZOrder 0
End Sub

Private Sub mnuAuthorityPrevilagesControlLocked_Click()
    frmAuthorityPrevilagesControlLocked.Show
    frmAuthorityPrevilagesControlLocked.ZOrder 0
End Sub

Private Sub mnuAuthorityPrevilagesControlVisibility_Click()
    frmAuthorityPrevilagesControlVisible.Show
    frmAuthorityPrevilagesControlVisible.ZOrder 0
End Sub

Private Sub mnuAuthorityPrevilagesMenuEnable_Click()
    frmAuthorityPrevilagesMenuEnable.Show
    frmAuthorityPrevilagesMenuEnable.ZOrder 0
End Sub

Private Sub mnuAuthorityPrevilagesMenuVisible_Click()
    frmAuthorityPrevilagesMenuVisible.Show
    frmAuthorityPrevilagesMenuVisible.ZOrder 0
End Sub

Private Sub mnuBackup_Click()
On Error GoTo ErrorHandler

    frmBackUp.ZOrder 0
    frmBackUp.Top = 0
    frmBackUp.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub

End Sub


Private Sub mnuBackOfficeAllAppointments_Click()
    frmAllAppointments.Show
    frmAllAppointments.ZOrder 0
End Sub

Private Sub mnuBackOfficeAllBookings_Click()
    frmAllBookings.Show
    frmAllBookings.ZOrder 0
End Sub

Private Sub mnuBackOfficeCreditAppointments_Click()
    frmCreditAppointments.Show
    frmCreditAppointments.ZOrder 0
End Sub

Private Sub mnuBackOfficeCreditBookings_Click()
    frmCreditBookings.Show
    frmCreditBookings.ZOrder 0
End Sub

Private Sub mnuBackOfficeWHT_Click()
    frmWHT.Show
    frmWHT.ZOrder 0
End Sub

Private Sub mnuBookings_Click()
On Error GoTo ErrorHandler
    If UserAuthority = AuthorityOwnerCOvered Then
        frmBookingsC.Show
        frmBookingsC.ZOrder 0
        frmBookingsC.WindowState = 0
        frmBookingsC.Top = 0
        frmBookingsC.Left = 0
    Else
        frmBookings.Show
        frmBookings.ZOrder 0
        frmBookings.WindowState = 0
        frmBookings.Top = 0
        frmBookings.Left = 0
    End If
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuBulkRefund_Click()
    frmBulkRefund.Show
    frmBulkRefund.ZOrder 0
End Sub

Private Sub mnuCalculateDailyAgentBalance_Click()
    frmCalculateAgentBalance.Show
    frmCalculateAgentBalance.ZOrder 0
End Sub

Private Sub mnuCancelARepayment_Click()
On Error GoTo ErrorHandler
    frmCancelRepayments.Top = 0
    frmCancelRepayments.Left = 0
    frmCancelRepayments.Show
    frmCancelRepayments.ZOrder 0
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuCancellation_Click()
On Error GoTo ErrorHandler
    frmChannelingCancellation.Show
    frmChannelingCancellation.ZOrder 0
    frmChannelingCancellation.WindowState = 0
    frmChannelingCancellation.Top = 0
    frmChannelingCancellation.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuCancelRepaymentReports_Click()
On Error GoTo ErrorHandler
    frmCancelRepamantsDetails.Show
    frmCancelRepamantsDetails.ZOrder 0
    frmCancelRepamantsDetails.Top = 0
    frmCancelRepamantsDetails.Left = 0
    
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuChanelingPatients_Click()
On Error GoTo ErrorHandler
    frmPeriodReport.Show
    frmPeriodReport.ZOrder 0
    frmPeriodReport.Top = 0
    frmPeriodReport.Left = 0
    
    Exit Sub
ErrorHandler:
    Exit Sub


End Sub

Private Sub mnuChangeHospitalFee_Click()
On Error GoTo ErrorHandler
    frmAllHospitalfeechange.Show
    frmAllHospitalfeechange.ZOrder 0
    frmAllHospitalfeechange.Top = 0
    frmAllHospitalfeechange.Left = 0
    
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuChannallingLists_Click()
On Error GoTo ErrorHandler
    frmChannelingLists.Show
    frmChannelingLists.ZOrder 0
    frmChannelingLists.WindowState = 2
Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuChannellingSheduling_Click()
On Error GoTo ErrorHandler
    frmChannellingEditing.Show
    frmChannellingEditing.ZOrder 0
    frmChannellingEditing.Top = 0
    frmChannellingEditing.Left = 0
    
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuCustomerPayment_Click()
'On Error GoTo ErrorHandler
'
'    frmPayment.Show
'    frmPayment.ZOrder 0
'    frmPayment.Top = 0
'    frmPayment.Left = 0
'    Exit Sub
'
'ErrorHandler:
'    Exit Sub
'
End Sub

Private Sub mnuDetails_Click()
    frmSplash.Show
End Sub


Private Sub mnuDoctorLeave_Click()
On Error GoTo ErrorHandler
    frmDoctorLeave.Show
    frmDoctorLeave.ZOrder 0
    frmDoctorLeave.Top = 0
    frmDoctorLeave.Left = 0

    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuDoctorPatientCount_Click()
    frmPeriodReport.Show
    frmPeriodReport.ZOrder 0
End Sub

Private Sub mnuDoctorPatientLists_Click()
    frmDoctorPatientLists.Show
    frmDoctorPatientLists.ZOrder 0
End Sub

Private Sub mnuDoctorPayment_Click()
On Error GoTo ErrorHandler

    frmDoctorPayments.Show
    frmDoctorPayments.ZOrder 0
    frmDoctorPayments.WindowState = 2
    Exit Sub

ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuDoctorPaymentDetails_Click()
On Error GoTo ErrorHandler
'If UserAuthority = AuthorityAdministrator Or UserAuthority = AuthorityOwner Then
    frmDoctorPaymentDetails.Show
    frmDoctorPaymentDetails.ZOrder 0
    frmDoctorPaymentDetails.WindowState = 2
    Exit Sub
'End If
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuDoctors_Click()
On Error GoTo ErrorHandler

    frmDoctor.Show
    frmDoctor.ZOrder 0
    frmDoctor.Top = 0
    frmDoctor.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub

End Sub


Private Sub mnuDoctorsDetails_Click()
On Error GoTo ErrorHandler
    frmPeriodDoctorReport.Show
    frmPeriodDoctorReport.ZOrder 0
    frmPeriodDoctorReport.Top = 0
    frmPeriodDoctorReport.Left = 0
    
    Exit Sub
ErrorHandler:
    Exit Sub


End Sub

Private Sub mnuDoctorSecessionView_Click()
On Error GoTo ErrorHandler
    frmSchedule.ZOrder 0
    frmSchedule.Top = 0
    frmSchedule.Left = 0
    frmSchedule.Show

Exit Sub

ErrorHandler:
Exit Sub

'Const PreSHape = "SHAPE {"
'Const Sql = "SELECT tblDoctor.*, tblTitle.Title AS Expr1 FROM tblDoctor LEFT OUTER JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID ORDER BY tblDoctor.DoctorName}  AS cmdDoc APPEND (( SHAPE {SELECT tblWeekDay.*,(AgentHospitalFee + AgentDoctorFee) As Total, tblFacilitySecession.* FROM tblWeekDay RIGHT OUTER JOIN tblFacilitySecession ON tblWeekDay.WeekDay_Id = tblFacilitySecession.SecessionWeekday Order By tblWeekDay.WeekDay_ID"
'Const PostSHape = "}  AS cmdWeekday COMPUTE cmdWeekday, ANY(cmdWeekday.'WeekDay') AS WeekdayName BY 'Staff_ID','WeekDay') AS cmdWeekday_Grouping RELATE 'Doctor_ID' TO 'Staff_ID') AS cmdWeekday_Grouping "
'csetPrinter.SetPrinterAsDefault (ReportPrinterName)
'
'' SHAPE {SELECT tblDoctor.*, tblTitle.Title AS Expr1 FROM tblDoctor LEFT OUTER JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID ORDER BY tblDoctor.DoctorName}  AS cmdDoc APPEND (( SHAPE {SELECT tblWeekDay.*, tblFacilitySecession.* FROM tblWeekDay RIGHT OUTER JOIN tblFacilitySecession ON tblWeekDay.WeekDay_Id = tblFacilitySecession.SecessionWeekday }  AS cmdWeekday COMPUTE cmdWeekday, ANY(cmdWeekday.'WeekDay') AS WeekdayName BY 'Staff_ID','WeekDay') AS cmdWeekday_Grouping RELATE 'Doctor_ID' TO 'Staff_ID') AS cmdWeekday_Grouping
''On Error GoTo ErrorHandler
'With DataEnvironment1
'
'    If .rscmdDoc.State = 1 Then .rscmdDoc.Close
'
'    .Commands!cmdDoc.CommandText = PreSHape & Sql & PostSHape
'    .cmdDoc
'
'    Set dtrDoctorSecession.DataSource = DataEnvironment1
'
'End With
'    dtrDoctorSecession.Sections("PageFooter").Controls("lblAdd").Caption = ShortAd
'    dtrDoctorSecession.Show
'    Exit Sub
'

'ErrorHandler:
'    Exit Sub

End Sub

Private Sub mnuDoctorsPayment_Click()
On Error GoTo ErrorHandler
'If UserAuthority = AuthorityAdministrator Or UserAuthority = AuthorityOwner Then
    frmDoctorsIncome.Show
    frmDoctorsIncome.ZOrder 0
    frmDoctorsIncome.Top = 0
    frmDoctorsIncome.Left = 0
'End If
Exit Sub

ErrorHandler:
Exit Sub
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub





Private Sub mnuInstitutionPayment_Click()
On Error GoTo ErrorHandler
    frmInstitutionPayment.Show
    frmInstitutionPayment.ZOrder 0
    frmInstitutionPayment.Top = 0
    frmInstitutionPayment.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuInstitutionPreferances_Click()
On Error GoTo ErrorHandler
    frmInstitutionPreferances.Show
    frmInstitutionPreferances.ZOrder 0
    frmInstitutionPreferances.Top = 0
    frmInstitutionPreferances.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuInstitutions_Click()
On Error GoTo ErrorHandler
    FrmInstitutions1.Show
    FrmInstitutions1.ZOrder 0
    FrmInstitutions1.Top = 0
    FrmInstitutions1.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub

End Sub



Private Sub mnuLoggin_Click()
On Error GoTo ErrorHandler
    frmLogged.Show
    frmLogged.ZOrder 0
    frmLogged.Top = 0
    frmLogged.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuManageRecordsDeleteAllRecords_Click()
    frmDeleteAllData.Show
    frmDeleteAllData.ZOrder 0
End Sub

Private Sub mnuManageRecordsDeleteSecessions_Click()
    frmDeletedSecessionData.Show
    frmDeletedSecessionData.ZOrder
End Sub

Private Sub mnuManageRecordsMakeAbsent_Click()
    frmMarkAbsentSecessionData.Show
    frmMarkAbsentSecessionData.ZOrder 0
End Sub

Private Sub mnuManageRecordsRefundReport_Click()
    frmRefundReport.Show
    frmRefundReport.ZOrder 0
End Sub

Private Sub mnuMySummary_Click()
On Error GoTo ErrorHandler
    frmNewDayEndSummary.Show
    frmNewDayEndSummary.ZOrder 0
    frmNewDayEndSummary.Left = 0
    frmNewDayEndSummary.Top = 0
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuMyShiftSummary_Click()
On Error GoTo ErrorHandler
    FrmMyShiftSummery.Show
    FrmMyShiftSummery.ZOrder 0
    FrmMyShiftSummery.Top = 0
    FrmMyShiftSummery.Left = 0

'    FrmUserShiftSummery.Show
'    FrmUserShiftSummery.ZOrder 0
'    FrmUserShiftSummery.Top = 0
'    FrmUserShiftSummery.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuNewChanneling_Click()
On Error GoTo ErrorHandler
    frmChannelingMS.Show
    frmChannelingMS.ZOrder 0
    frmChannelingMS.WindowState = 2
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub


Private Sub mnuNewDayEndSummery_Click()
On Error GoTo ErrorHandler
    frmNewAnyDayEndSummary.Show
    frmNewAnyDayEndSummary.ZOrder 0
    frmNewAnyDayEndSummary.Top = 0
    frmNewAnyDayEndSummary.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuNewMyShiftSummery_Click()
On Error GoTo ErrorHandler
    frmNewUserShiftSummary.Show
    frmNewUserShiftSummary.ZOrder 0
    frmNewUserShiftSummary.Top = 0
    frmNewUserShiftSummary.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuNewShiftEndSummery_Click()
    On Error GoTo ErrorHandler
    
    If UserAuthority = AuthorityOwnerCOvered Then
        FrmCoveredUserShiftSummery.Show
        FrmCoveredUserShiftSummery.ZOrder 0
        FrmCoveredUserShiftSummery.Top = 0
        FrmCoveredUserShiftSummery.Left = 0
    ElseIf UserAuthority = AuthorityOwner Or UserAuthority = AuthorityAdministrator Then
        frmNewAnyShiftEndSummary.Show
        frmNewAnyShiftEndSummary.ZOrder 0
        frmNewAnyShiftEndSummary.Top = 0
        frmNewAnyShiftEndSummary.Left = 0
    Else
'        Dim TemResponce As Integer
'        TemResponce = MsgBox("You are not allowed to view the reports", vbCritical, "No authority")
'        Exit Sub
    End If
        
    Exit Sub
    
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuOldBookings_Click()
On Error GoTo ErrorHandler
    frmOldChannelingMS.Show
    frmOldChannelingMS.ZOrder 0
    frmOldChannelingMS.WindowState = 2
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuOwnerPreferances_Click()
On Error GoTo ErrorHandler
If UserAuthority = AuthorityAdministrator Or UserAuthority = AuthorityOwner Then
    frmOwnersPreferances.Show
    frmOwnersPreferances.ZOrder 0
    frmOwnersPreferances.Top = 0
    frmOwnersPreferances.Left = 0
    Exit Sub
End If

ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuPatientCounts_Click()
    frmAgentChannellingCount.Show
End Sub

Private Sub mnuPatients_Click()
On Error GoTo ErrorHandler
    frmPatientMain.Show
    frmPatientMain.ZOrder 0
    frmPatientMain.Top = 0
    frmPatientMain.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub

End Sub



Private Sub mnuPreferances_Click()
On Error GoTo ErrorHandler
  '  frmPreferances.Show
  '  frmPreferances.ZOrder 0
'    frmPreferances.Left = 0
  '  frmPreferances.Top = 0
    Exit Sub

ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuPrintingPreferances_Click()
On Error GoTo ErrorHandler
    frmPrintingPreferances.Show
    frmPrintingPreferances.ZOrder 0
    frmPrintingPreferances.Top = 0
    frmPrintingPreferances.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

'Public Sub mnuPrintMyShiftEndSummery_Click()
'    PrintMyShiftEndSummery
'End Sub

Private Sub mnuPrintToDayEndSummery_Click()
Call PrintDayEndSummery
End Sub

Private Sub mnuProgramPreferances_Click()
On Error GoTo ErrorHandler
    frmProgramPreferances.Show
    frmProgramPreferances.ZOrder 0
    frmProgramPreferances.Top = 0
    frmProgramPreferances.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuRefund_Click()
On Error GoTo ErrorHandler
    frmChannelingRefund.Show
    frmChannelingRefund.ZOrder 0
    frmChannelingRefund.WindowState = 0
    frmChannelingRefund.Top = 0
    frmChannelingRefund.Left = 0
    frmChannelingRefund.Width = 7080
    frmChannelingRefund.Height = 9390
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuRestore_Click()
On Error GoTo ErrorHandler
    frmRestore.Show
    frmRestore.ZOrder 0
    frmRestore.Top = 0
    frmRestore.Left = 0
    Exit Sub

ErrorHandler:
    Exit Sub

    
End Sub


Private Sub mnuScanIncome_Click()
On Error GoTo ErrorHandler
    frmPeriodScanReport.Show
    frmPeriodScanReport.ZOrder 0
    frmPeriodScanReport.Top = 0
    frmPeriodScanReport.Left = 0
    
    Exit Sub
ErrorHandler:
    Exit Sub



End Sub

Private Sub mnuScanIncomeAll_Click()
    frmAllPeriodScanReport.Show
    frmAllPeriodScanReport.ZOrder 0
End Sub

Private Sub mnuSecessionViceIncome_Click()
    frmDoctorIncomeSecessionVice.Show
    frmDoctorIncomeSecessionVice.ZOrder 0
End Sub

Private Sub mnuSelectedDaysDoctorPayments_Click()
On Error GoTo ErrorHandler
'If UserAuthority = AuthorityAdministrator Or UserAuthority = AuthorityOwner Then
    frmSelectedDoctorPayments.Show
    frmSelectedDoctorPayments.ZOrder 0
    frmSelectedDoctorPayments.Top = 0
    frmSelectedDoctorPayments.Left = 0
    Exit Sub
'End If

ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuShiftEndSummery_Click()
On Error GoTo ErrorHandler
    frmNewShiftEndSummary.Show
    frmNewShiftEndSummary.ZOrder 0
    frmNewShiftEndSummary.Left = 0
    frmNewShiftEndSummary.Top = 0
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuShiftSummery_Click()
    On Error GoTo ErrorHandler
    
    If UserAuthority = AuthorityOwnerCOvered Then
        FrmCoveredUserShiftSummery.Show
        FrmCoveredUserShiftSummery.ZOrder 0
        FrmCoveredUserShiftSummery.Top = 0
        FrmCoveredUserShiftSummery.Left = 0
    ElseIf UserAuthority = AuthorityOwner Or UserAuthority = AuthorityAdministrator Then
        FrmShiftSummery.Show
        FrmShiftSummery.ZOrder 0
        FrmShiftSummery.Top = 0
        FrmShiftSummery.Left = 0
    Else
'        Dim TemResponce As Integer
'        TemResponce = MsgBox("You are not allowed to view the reports", vbCritical, "No authority")
'        Exit Sub
    End If
        
    Exit Sub
    
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuShiftUser_Click()
'    frmShiftUser.Show 1
End Sub

Private Sub mnuSingleAgentTransactions_Click()
On Error GoTo ErrorHandler
    frmSingleAgentTransactions.Show
    frmSingleAgentTransactions.ZOrder 0
    frmSingleAgentTransactions.Top = 0
    frmSingleAgentTransactions.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuSpeciality_Click()

On Error GoTo ErrorHandler

    
    frmSpecialtyDoctor.Show
    frmSpecialtyDoctor.ZOrder 0
    frmSpecialtyDoctor.Top = 0
    frmSpecialtyDoctor.Left = 0
    frmSpecialtyDoctor.Show
    Exit Sub

ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuSpecialityStaff_Click()

On Error GoTo ErrorHandler

    
    frmSpecialtyStaff.Show
    frmSpecialtyStaff.ZOrder 0
    frmSpecialtyStaff.Top = 0
    frmSpecialtyStaff.Left = 0
    frmSpecialtyStaff.Show
    Exit Sub

ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuStaff_Click()
On Error GoTo ErrorHandler
    
    frmStaff.Show
    frmStaff.ZOrder 0
    frmStaff.Top = 0
    frmStaff.Left = 0
Exit Sub

ErrorHandler:
    Exit Sub

End Sub

Private Sub mnutem_Click()
frmTem1.Show
End Sub

'Private Sub mnuTem_Click()
''    frmDoctorIncome.Show
''    dtrPrint.ReportWidth = 3 * 1440
''    dtrPrint.PrintReport = 6 * 1440
''    dtrPrint.Sections("Section4").Controls.Item("label3").Caption = "Lakmedipro Pvt Ltd.," & vbNewLine & "PO BOX 1," & vbNewLine & "Kamburupitiya"
''    dtrPrint.Show
''    dtrPrint.PrintReport True
''DataReportPatientViceDoctorPayments.Show
'With DataEnvironment1
'    If .rscAgents.State = 1 Then .rscAgents.Close
'    .Commands!cAgents.CommandText = " SHAPE {select * from tblinstitutions order by institutionname}  AS cAgents APPEND ({SELECT tblPatientFacility.* FROM tblPatientFacility where bookingdate = #02/02/2008# }  AS ccFacilities RELATE 'Institution_ID' TO 'Agent_ID') AS ccFacilities"
'    .cAgents
'
'
'
'End With
'
'DataReport1.Show
'End Sub

Private Sub mnuTipOfDay_Click()
    frmTip.Show
End Sub

Private Sub mnuTOC_Click()
    HtmlHelp hwnd, App.Path & "\help.chm", HH_DISPLAY_TOPIC, ByVal "refund.htm"
End Sub

Private Sub mnuTodayEndSummery_Click()
On Error GoTo ErrorHandler
    
    FrmtodayEndShiftSummery.Show
    FrmtodayEndShiftSummery.ZOrder 0
    FrmtodayEndShiftSummery.Left = 0
    FrmtodayEndShiftSummery.Top = 0
    Exit Sub
ErrorHandler:
    Exit Sub
    
End Sub

Private Sub mnuTodaysAllBookings_Click()
On Error GoTo ErrorHandler
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    Dim TemResponce As Long
    Dim RetVal As Integer
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    
    dtrDayEndSummery.DataMember = Empty
    
    With DataEnvironment1.rsDayEndSummery
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientFacility.*, tblStaff_Booked.StaffName as BStaffName, tblTitle.Title, tblDoctor.DoctorName, tblInstitutions.InstitutionName, tblStaff_CreditSettle.StaffName as CStaffName, tblStaff_Repay.StaffName as RStaffName FROM tblStaff AS tblStaff_CreditSettle RIGHT JOIN (tblStaff AS tblStaff_Repay RIGHT JOIN (tblInstitutions RIGHT JOIN ((tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) RIGHT JOIN (tblStaff AS tblStaff_Booked RIGHT JOIN tblPatientFacility ON tblStaff_Booked.Staff_ID = tblPatientFacility.User_ID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID) ON tblStaff_Repay.Staff_ID = tblPatientFacility.RepayUser_ID) ON tblStaff_CreditSettle.Staff_ID = tblPatientFacility.CreditSettleUser_ID where bookingdate = '" & Date & "' order by patientfacility_ID "
        .Open
    End With
        
    dtrDayEndSummery.DataMember = "dayendsummery"
    
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Report Not Printed")
            Exit Sub
        Case FORM_SELECTED   ' 1
            If HospitalDetails = True Then
                dtrDayEndSummery.Sections.Item("Section4").Controls.Item("lblInstitutionName").Caption = InstitutionName
                dtrDayEndSummery.Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
            End If
            dtrDayEndSummery.Sections.Item("Section4").Controls.Item("lblreport").Caption = "Current State of All Bookings on " & Format(Date, DefaultLongDate)
            dtrDayEndSummery.Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
            dtrDayEndSummery.Show
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added. Report NOT Printed", vbExclamation, "New Paper size")
            Exit Sub
    End Select
    Exit Sub
ErrorHandler:
    Exit Sub

End Sub

Private Sub mnuTodaysCreditAppointments_Click()
    frmCreditAppointments.Show
    frmCreditAppointments.ZOrder 0
End Sub

Private Sub mnuTodaysMyBookings_Click()
On Error GoTo ErrorHandler
    CSetPrinter.SetPrinterAsDefault (ReportPrinterName)
    Dim TemResponce As Long
    Dim RetVal As Integer
    RetVal = SelectForm(ReportPaperName, Me.hwnd)
    dtrShiftEndSummery.DataMember = Empty
    With DataEnvironment1.rsDayEndSummery
        If .State = 1 Then .Close
        .Source = "SELECT tblPatientFacility.*, tblStaff_Booked.StaffName as BStaffName, tblTitle.Title, tblDoctor.DoctorName, tblInstitutions.InstitutionName, tblStaff_CreditSettle.StaffName as CStaffName, tblStaff_Repay.StaffName as RStaffName FROM tblStaff AS tblStaff_CreditSettle RIGHT JOIN (tblStaff AS tblStaff_Repay RIGHT JOIN (tblInstitutions RIGHT JOIN ((tblDoctor LEFT JOIN tblTitle ON tblDoctor.DoctorTitle_ID = tblTitle.Title_ID) RIGHT JOIN (tblStaff AS tblStaff_Booked RIGHT JOIN tblPatientFacility ON tblStaff_Booked.Staff_ID = tblPatientFacility.User_ID) ON tblDoctor.Doctor_ID = tblPatientFacility.Staff_ID) ON tblInstitutions.Institution_ID = tblPatientFacility.Agent_ID) ON tblStaff_Repay.Staff_ID = tblPatientFacility.RepayUser_ID) ON tblStaff_CreditSettle.Staff_ID = tblPatientFacility.CreditSettleUser_ID " & _
        " where (bookingdate = '" & Date & "' and tblPatientFacility.user_ID = " & UserID & ") or " & _
        " (SettleCashDate = '" & Date & "' and tblPatientFacility.CreditSettleUser_ID = " & UserID & ") order by patientfacility_ID "
        .Open
    End With
    dtrShiftEndSummery.DataMember = "dayendsummery"
    Select Case RetVal
        Case FORM_NOT_SELECTED   ' 0
            TemResponce = MsgBox("You have not selected a printer form to print, Please goto Preferances and Printing preferances to set a valid printer form.", vbExclamation, "Report Not Printed")
            Exit Sub
        Case FORM_SELECTED   ' 1
            If HospitalDetails = True Then
                dtrShiftEndSummery.Sections.Item("Section4").Controls.Item("lblInstitutionName").Caption = InstitutionName
                dtrShiftEndSummery.Sections.Item("Section4").Controls.Item("lblInstitutionaddress").Caption = InstitutionAddress
            End If
            dtrShiftEndSummery.Sections.Item("Section4").Controls.Item("lblreport").Caption = "Current State of All Bookings on " & Format(Date, DefaultLongDate) & " by " & UserName
            dtrShiftEndSummery.Sections.Item("Section3").Controls.Item("lblad").Caption = LongAd
            dtrShiftEndSummery.Show
        Case FORM_ADDED   ' 2
            TemResponce = MsgBox("New paper size added. Report NOT Printed", vbExclamation, "New Paper size")
            Exit Sub
    End Select
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuTodaysMyCreditBookings_Click()
    frmCreditBookings.Show
    frmCreditBookings.ZOrder 0
End Sub

Private Sub mnuTodaysPayments_Click()
On Error GoTo ErrorHandler
'If UserAuthority <> AuthorityOwner And UserAuthority <> AuthorityAdministrator Then Exit Sub
    frmTodaysDoctorPayments.Show
    frmTodaysDoctorPayments.ZOrder 0
    frmTodaysDoctorPayments.Top = 0
    frmTodaysDoctorPayments.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub

Private Sub mnuTraceBookings_Click()
On Error GoTo ErrorHandler
    frmLocatePatients.Show
    frmLocatePatients.ZOrder 0
    frmLocatePatients.Top = 0
    frmLocatePatients.Left = 0
    Exit Sub
ErrorHandler:
    Exit Sub


End Sub

Private Sub munuDayendSummery_Click()
    On Error GoTo ErrorHandler
    If UserAuthority = AuthorityOwnerCOvered Then
        FrmCoveredDayEndShiftSummery.Show
        FrmCoveredDayEndShiftSummery.ZOrder 0
        FrmCoveredDayEndShiftSummery.Top = 0
        FrmCoveredDayEndShiftSummery.Left = 0
    ElseIf UserAuthority = AuthorityOwner Or UserAuthority = AuthorityAdministrator Then
        FrmtodayEndShiftSummery.Show
        FrmtodayEndShiftSummery.ZOrder 0
        FrmtodayEndShiftSummery.Top = 0
        FrmtodayEndShiftSummery.Left = 0
    Else
'        Dim TemResponce As Integer
'        TemResponce = MsgBox("You are not allowed to view the reports", vbCritical, "No authority")
'        Exit Sub
    End If
    Exit Sub

ErrorHandler:
    Exit Sub
    
End Sub

Private Sub Timer1_Timer()
lblTime.Caption = Time
End Sub

Private Sub Timer2_Timer()
'    frmWritingBalance.Show 1
'    Timer2.Interval = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1: mnuNewChanneling_Click
    Case 2: mnuTraceBookings_Click
    Case 3: mnuDoctorPayment_Click
    Case 4: mnuNewShiftEndSummery_Click
    Case 5: mnuNewDayEndSummery_Click
    Case 6: mnuChannellingSheduling_Click
    Case 7: mnuExit_Click
End Select
End Sub
