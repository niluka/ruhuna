VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "eCashier - Ruhunu Hospitals (Pvt) Ltd"
   ClientHeight    =   8820
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15240
   Icon            =   "MDIMainUser.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   7080
      Top             =   4200
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6600
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":29C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":302C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":38C75
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":409BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":48BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":4E3AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":55B28
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
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   8640
         TabIndex        =   1
         Top             =   0
         Width           =   6615
         Begin VB.Label lblDateTime 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   6375
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":5C800
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":62EB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":6B863
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":735AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":7B7B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":80F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMainUser.frx":88716
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuCategories 
         Caption         =   "Catogeries"
         Begin VB.Menu mnuEditPatientCategory 
            Caption         =   "Patient Catogery"
         End
         Begin VB.Menu mnuEditServiceCategory 
            Caption         =   "Service Catogeries"
         End
         Begin VB.Menu mnuEditServiceSubcategory 
            Caption         =   "Service Subcatogery"
         End
         Begin VB.Menu mnuEditCategoriesServiceSecessions 
            Caption         =   "Service Secessions"
         End
         Begin VB.Menu mnuEditProfessionalChargesForServiceSubcategories 
            Caption         =   "Professional Charges for Service Subcategories"
         End
         Begin VB.Menu mnuEditSurgeries 
            Caption         =   "Surgeries"
         End
      End
      Begin VB.Menu mnuHospital 
         Caption         =   "Hospital"
         Begin VB.Menu mnuEditSpeciality 
            Caption         =   "Speciality"
         End
         Begin VB.Menu mnuStaff 
            Caption         =   "Staff"
         End
         Begin VB.Menu mnuDepartments 
            Caption         =   "Departments"
         End
         Begin VB.Menu mnuEditRoomCategory 
            Caption         =   "Room Categories"
         End
         Begin VB.Menu mnuRooms 
            Caption         =   "Rooms"
         End
         Begin VB.Menu mnuEditPrevilages 
            Caption         =   "Previlages"
            Begin VB.Menu mnuEditPrevilagesMenuEnable 
               Caption         =   "Enable Menus"
            End
            Begin VB.Menu mnuPrevilagesMenuVisible 
               Caption         =   "Visible Menu"
            End
            Begin VB.Menu mnuPrevilagesEnableControls 
               Caption         =   "Enable Controls"
            End
            Begin VB.Menu mnuPrevilagesVisibleControls 
               Caption         =   "Visible Controls"
            End
         End
      End
      Begin VB.Menu mnuHealthSchemeSuppliers 
         Caption         =   "Health Scheme Suppliers"
      End
      Begin VB.Menu mnuEditAgents 
         Caption         =   "Agents"
      End
      Begin VB.Menu mnuEditHospitalCharges 
         Caption         =   "Charges"
      End
   End
   Begin VB.Menu mnuBHT 
      Caption         =   "BHT"
      Begin VB.Menu mnuAdmit 
         Caption         =   "Admit"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mmnuBHTRoomOccupancy 
         Caption         =   "Room Occupancy"
      End
      Begin VB.Menu mnuBHTCashierSummery 
         Caption         =   "Cashier Summery"
      End
      Begin VB.Menu mnuBHTCashierDischarge 
         Caption         =   "Cashier Discharge"
      End
      Begin VB.Menu mnuBHTServices 
         Caption         =   "Services"
      End
      Begin VB.Menu mnuBHTProfessionalCHarges 
         Caption         =   "Professional Charges"
      End
      Begin VB.Menu mnuChangeRoom 
         Caption         =   "Change Room"
      End
      Begin VB.Menu mnuBHTEditRoomDetails 
         Caption         =   "Edit Room Details"
      End
      Begin VB.Menu mnuBHTPayments 
         Caption         =   "Payments"
         Begin VB.Menu mnuBHTPaymentsPay 
            Caption         =   "Pay"
         End
         Begin VB.Menu mnuBHTPaymentsReprints 
            Caption         =   "Reprients"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuBHTPaymentCancellations 
            Caption         =   "Payment Cancellations"
         End
         Begin VB.Menu mnuBHTRefunds 
            Caption         =   "Refunds"
         End
         Begin VB.Menu mnuBHTPaymentBillSearch 
            Caption         =   "Search Payment Bills"
         End
      End
      Begin VB.Menu mnuBHTBHTSummery 
         Caption         =   "BHT Summery"
      End
      Begin VB.Menu mnuBHTSummeryAllF 
         Caption         =   "All BHT Summeries"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditBHT 
         Caption         =   "Edit BHT Details"
      End
      Begin VB.Menu mnuBHTServiceValues 
         Caption         =   "BHT Service Values (By Service Date)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBHTServiceValuesByDOD 
         Caption         =   "BHT Service Values (By Date of Discharge)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBHTAddDetailsOfSx 
         Caption         =   "Add Details of Surgery"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBHTGetDetailsOfSx 
         Caption         =   "Get Details of Surgeries"
      End
      Begin VB.Menu mnuBHTAdmissionBook 
         Caption         =   "Admission Book"
      End
      Begin VB.Menu mnuBHTAdvancBHTFunctions 
         Caption         =   "Advance BHT Functions"
         Begin VB.Menu mnuBHTSummeryA 
            Caption         =   "BHT Summery A"
         End
         Begin VB.Menu mnuBHTAdvanceBulkDischarge 
            Caption         =   "Bulk Discharge"
         End
         Begin VB.Menu mnuBHTServicesD 
            Caption         =   "Services(Discharged Patients)"
         End
         Begin VB.Menu mnuBHTProfessionalCHargesD 
            Caption         =   "Professional Charges (Discharged Patients)"
         End
         Begin VB.Menu mnuBHTAdditionalCharges 
            Caption         =   "Additional Charges"
         End
         Begin VB.Menu mnuBHTReverseDischarge 
            Caption         =   "Reverse Discharge"
         End
      End
   End
   Begin VB.Menu mnuGSB 
      Caption         =   "Green Sheet Bills"
      Begin VB.Menu mnuGSBAdmit 
         Caption         =   "Admit"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuGSBServices 
         Caption         =   "Services"
      End
      Begin VB.Menu mnuGSBProfessionalCharges 
         Caption         =   "Professional Charges"
      End
      Begin VB.Menu mnuGSBPayments 
         Caption         =   "Payments"
         Begin VB.Menu mnuGSBPay 
            Caption         =   "Pay"
         End
         Begin VB.Menu mnuGSBPaymentReprint 
            Caption         =   "Reprint"
         End
         Begin VB.Menu mnuGSBPaymentCancellations 
            Caption         =   "Payment Cancellations"
         End
         Begin VB.Menu mnuGSBRefunds 
            Caption         =   "Refunds"
         End
         Begin VB.Menu mnuGSBPaymentSearch 
            Caption         =   "GSB Search"
         End
      End
      Begin VB.Menu mnuGSBSummery 
         Caption         =   "Summery"
      End
      Begin VB.Menu mnuGSBEdit 
         Caption         =   "Edit GSB"
      End
      Begin VB.Menu mnuGSBReprint 
         Caption         =   "Reprint GSB"
      End
      Begin VB.Menu mnuGSBServiceValuesByDOD 
         Caption         =   "Service Values by DOD"
      End
   End
   Begin VB.Menu mnuOPD 
      Caption         =   "OPD"
      Begin VB.Menu mnuOPDBIlls 
         Caption         =   "OPD Bills"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuOPDBillCancellation 
         Caption         =   "OPD Bill Cancellation"
      End
      Begin VB.Menu mnuOPDRefunds 
         Caption         =   "OPD Bill Refunds"
      End
      Begin VB.Menu mnuOPDBillReprints 
         Caption         =   "OPD Bill Reprints"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOPDSearchBills 
         Caption         =   "Search OPD Bills"
      End
      Begin VB.Menu mnuOPDServiceBills 
         Caption         =   "Service Bills"
      End
      Begin VB.Menu mnuOPDSErviceCounts 
         Caption         =   "ServiceCounts"
      End
      Begin VB.Menu mnuOPDServiceValues 
         Caption         =   "Service Values"
      End
   End
   Begin VB.Menu mnuR 
      Caption         =   "Roentgents"
      Begin VB.Menu mnuRBills 
         Caption         =   "Roentgents Bills"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuRBillCancellation 
         Caption         =   "Roentgents Bill Cancellation"
      End
      Begin VB.Menu mnuRBillRefunds 
         Caption         =   "Roentgents Bill Refunds"
      End
      Begin VB.Menu mnuRSarchBills 
         Caption         =   "Search Roentgents Bills"
      End
      Begin VB.Menu mnuRBillReprints 
         Caption         =   "Roentgents Bill Reprints"
      End
      Begin VB.Menu mnuRBillValues 
         Caption         =   "Roentgents Bill Values"
      End
      Begin VB.Menu mnuRBillsByCategories 
         Caption         =   "Roentgents Bills By Categories"
      End
   End
   Begin VB.Menu mnuLab 
      Caption         =   "Lab"
      Begin VB.Menu mnuLabBills 
         Caption         =   "Lab Bills"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuLabBillCancellation 
         Caption         =   "Lab Bill Cancellations"
      End
      Begin VB.Menu mnuLabBillReprint 
         Caption         =   "Lab Bill Reprint"
      End
      Begin VB.Menu mnuLabSearchBills 
         Caption         =   "Search Lab Bills"
      End
      Begin VB.Menu mnuLabBillsList 
         Caption         =   "Lab Bills List"
      End
      Begin VB.Menu mnuLabPatientList 
         Caption         =   "Lab Service Patient List"
      End
      Begin VB.Menu mnuLabServiceBills 
         Caption         =   "Lab Service Bills"
      End
      Begin VB.Menu mnuLabLabServiceValues 
         Caption         =   "Lab Service Values"
      End
   End
   Begin VB.Menu mnuMedicalTest 
      Caption         =   "Medical Test Bills"
      Begin VB.Menu mnuMedicalTestBills 
         Caption         =   "Medical Test Bills"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuMedicalTestBillCancellation 
         Caption         =   "Medical Test Bill Cancellation"
      End
      Begin VB.Menu mnuMedicalTestBillReprints 
         Caption         =   "Medical Test Bill Reprints"
      End
      Begin VB.Menu mnuMedicalTestBillRefunds 
         Caption         =   "Medical Test Bill Refunds"
      End
      Begin VB.Menu mnuMTBillSearch 
         Caption         =   "Search Medical Test Bills"
      End
      Begin VB.Menu mnuMTServiceValues 
         Caption         =   "Medical Test Service Values"
      End
   End
   Begin VB.Menu mnuPharmacy 
      Caption         =   "Pharmacy"
      Begin VB.Menu mnuPharmacyBills 
         Caption         =   "Pharmacy Bills"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuPharmacyBillCancellation 
         Caption         =   "Pharmacy Bill Cancellation"
      End
      Begin VB.Menu mnuPharmacyBillReturn 
         Caption         =   "Pharmacy Bill Return"
      End
   End
   Begin VB.Menu mnuPayments 
      Caption         =   "Payments"
      Begin VB.Menu mnuPaymentAgent 
         Caption         =   "Agents"
         Begin VB.Menu mnuAgentPayments 
            Caption         =   "Agent Payments"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnuAgentCancellations 
            Caption         =   "Agent Cancellations"
         End
         Begin VB.Menu mnuAgentReprints 
            Caption         =   "Agent Payment Reprints"
         End
      End
      Begin VB.Menu mnuPaymentsHSS 
         Caption         =   "Health Scheme Suppliers"
         Begin VB.Menu mnuPaymentsHSSBHT 
            Caption         =   "Settling BHT"
         End
         Begin VB.Menu mnuPaymentsHSSMT 
            Caption         =   "Medical Test"
         End
      End
      Begin VB.Menu mnuProfessionalFee 
         Caption         =   "Professional Fee"
         Begin VB.Menu mnuPaymentsProfessionalFeePaymentsForInwardPatients 
            Caption         =   "Professional Fee Payments for Inward Patients"
         End
         Begin VB.Menu mnuPaymentsProfessionalFeePaymentsForGSB 
            Caption         =   "Professional Fee Payments for Greensheets"
         End
         Begin VB.Menu mnuPaymentsProfessionalFeePaymentsForOPDPatients 
            Caption         =   "Professional Fee Payments for OPD Patients"
         End
         Begin VB.Menu mnuPaymentsProfessionalFeePaymentsForLabPatients 
            Caption         =   "Professional Fee Payments for Lab Patients"
         End
         Begin VB.Menu mnuPaymentsProfessionalFeePaymentsForRPatients 
            Caption         =   "Professional Fee Payments for Roentgents Patients"
         End
      End
      Begin VB.Menu mnuPaymentsExpence 
         Caption         =   "Expences"
         Begin VB.Menu mnuPaymentsAddExpences 
            Caption         =   "Add Expences"
         End
         Begin VB.Menu mnuPaymentsCancelExpences 
            Caption         =   "Cancel Expences"
         End
      End
   End
   Begin VB.Menu mnuBackOffice 
      Caption         =   "Back Office"
      Begin VB.Menu mnuShiftEndSummery 
         Caption         =   "Shift End Summery"
      End
      Begin VB.Menu mnuDayEndSummery 
         Caption         =   "Day End Summery"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuBackOfficeAllShiftEndSummeries 
         Caption         =   "All Shift End Sumeries"
      End
      Begin VB.Menu mnuBackOfficeShiftEndCounts 
         Caption         =   "Shift End Counts"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBackOfficeDayEndCounts 
         Caption         =   "Day End Counts"
      End
      Begin VB.Menu mnuBackOfficeDetailedSummery 
         Caption         =   "Detailed Summery"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuBackOfficePeriodSummeryReport 
         Caption         =   "Period Detailed Summery"
      End
      Begin VB.Menu mnuBackOfficePeriodSummeryReportByDate 
         Caption         =   "Period Detailed Summery - By Date"
      End
      Begin VB.Menu mnuBackOfficeDailySummeryReport2 
         Caption         =   "Daily Summery Report - Detailed"
      End
      Begin VB.Menu mnuBackOfficeDailySummeryReport 
         Caption         =   "Daily Summery Report"
      End
      Begin VB.Menu mnuBackOfficeDailySummeryReportHospital 
         Caption         =   "Daily Summery Report - Hospital"
      End
      Begin VB.Menu mnuBackOfficeExpenceDetails 
         Caption         =   "Expence Details"
      End
      Begin VB.Menu mnuBackOfficePayments 
         Caption         =   "Payments"
         Begin VB.Menu mnuBackOfficePaymentSummery 
            Caption         =   "Payment Summery"
         End
         Begin VB.Menu frmBackOfficePeriodPaymentSummery 
            Caption         =   "Period Payment"
         End
         Begin VB.Menu mnuBackOfficePaymentsAgeAnalysis 
            Caption         =   "Age Analysis"
         End
      End
      Begin VB.Menu mnuBackOfficeBHT 
         Caption         =   "BHT"
         Begin VB.Menu mnuBackOfficeBHTSummery 
            Caption         =   "BHT Summery"
         End
         Begin VB.Menu mnuBackOfficeBHTBookKeepingSummery 
            Caption         =   "BHT Book-keeping Summery"
         End
         Begin VB.Menu mnuBackOfficeBHTBookKeepingSummeryExtended 
            Caption         =   "BHT Book-keeping Summery (Extended)"
         End
         Begin VB.Menu mnuBackOfficeBHTPaymentBookKeepingSummery 
            Caption         =   "BHT Payment Book-keeping Summery"
         End
         Begin VB.Menu mnuBackOfficeBHTServicesActualSummery 
            Caption         =   "BHT Services Actual Summery"
         End
         Begin VB.Menu mnuBackOfficeBHTServicesActualSummeryD 
            Caption         =   "BHT Services Actual Summery (With Discreptions)"
         End
         Begin VB.Menu mnuBackOfficeBHTServicesOPDValue 
            Caption         =   "BHT Services OPD Value"
         End
         Begin VB.Menu mnuBackOfficeBHTServicesOPDValueD 
            Caption         =   "BHT Services OPD Value (With Discreption)"
         End
      End
      Begin VB.Menu mnuBackOfficeGreenSheet 
         Caption         =   "Green Sheet"
         Begin VB.Menu mnnuBackOfficeGSBSummery 
            Caption         =   "Green Sheet Bill Summery"
         End
         Begin VB.Menu mnuBackOfficeBookKeepingSummery 
            Caption         =   "GSB Book-keeping Sumery"
         End
         Begin VB.Menu mnuBackOfficeBookKeepingSummeryE 
            Caption         =   "GSB Book-keeping Sumery (Extended)"
         End
         Begin VB.Menu mnuBackOfficeGSBPaymentBookKeepingSummery 
            Caption         =   "GSB Payment Book-keeping Summery"
         End
      End
      Begin VB.Menu mnuBackOfficeHealthSchemeSuppliers 
         Caption         =   "Health Scheme Suppliers"
         Begin VB.Menu mnuBackOfficeHSSBHT 
            Caption         =   "BHT"
            Begin VB.Menu mnuBackOfficeBHTPayments 
               Caption         =   "Payments"
            End
            Begin VB.Menu mnuBackOfficeHealthSchemeSuppliersAgeAnalysis 
               Caption         =   "Age Analysis"
            End
         End
         Begin VB.Menu mnuBackOfficeHSSMT 
            Caption         =   "Medical Tests"
            Begin VB.Menu mnuBackOfficeHSSMTPayments 
               Caption         =   "Payments"
            End
            Begin VB.Menu mnuBackOfficeHSSMTAgeAnalysis 
               Caption         =   "Age Analysis"
            End
            Begin VB.Menu mnuBackOfficeMTCreditLetters 
               Caption         =   "Credit Letters"
            End
         End
         Begin VB.Menu mnuBackOfficeHSSCurrentBalance 
            Caption         =   "Current Balance"
         End
      End
      Begin VB.Menu mnuBackOfficePatients 
         Caption         =   "Patients"
         Begin VB.Menu mnuBackOfficePatientsAgeAnalysis 
            Caption         =   "BHT Age Analysis"
         End
         Begin VB.Menu mnuBackOfficeGSBAgeAnalysis 
            Caption         =   "GSB Age Analysis"
         End
         Begin VB.Menu mnuBackOfficePatientsMTAgeAnalysis 
            Caption         =   "Medical Test Age Analysis"
         End
         Begin VB.Menu mnuBackOfficeBHTBalance 
            Caption         =   "BHT Balance"
         End
         Begin VB.Menu mnuBackOfficeGSBBalance 
            Caption         =   "GSB Blance"
         End
      End
      Begin VB.Menu mnuBackOfficeMedicalTests 
         Caption         =   "Medical Tests"
         Begin VB.Menu mnuBackOfficeMedicalTestBillsWithLabServices 
            Caption         =   "Medical Test Bills with Lab Services"
         End
      End
      Begin VB.Menu mnuBackOfficeAgentPayments 
         Caption         =   "Agent Payments"
      End
      Begin VB.Menu mnuBackOfficeViewServices 
         Caption         =   "View Services"
      End
   End
   Begin VB.Menu mnuPreferances 
      Caption         =   "Preferances"
      Begin VB.Menu mnuProgramPreferances 
         Caption         =   "Program Preferances"
      End
      Begin VB.Menu mnuPrintingPreferances 
         Caption         =   "Printing Preferances"
      End
      Begin VB.Menu mnuPreferanceHospitalCharges 
         Caption         =   "Inward Charges"
      End
      Begin VB.Menu mnuHospitalDetails 
         Caption         =   "Hospital Details"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuTipOfTheDay 
         Caption         =   "Tip of the day"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuTableOfContants 
         Caption         =   "Table of Contents"
      End
      Begin VB.Menu mnuHelpAdministrator 
         Caption         =   "Administrator"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub frmBackOfficePeriodPaymentSummery_Click()
    With frmPeriodPaymentSummery
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
    
End Sub

Private Sub MDIForm_Load()
    lblDateTime.Caption = Format("Date : " & Format(Date, "dd MM yy") & "   Time : " & Format(Time, "H:M AMPM"))
End Sub

Private Sub DelUncompletedBills()
'    Dim rsTem As New ADODB.Recordset
'    Dim temSql As String
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "Delete from tblIncomeBill where Completed = 0 AND UserID = " & UserID
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .State = 1 Then .Close
'    End With
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
    i = MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo, "Exit?")
    If i = vbNo Then
        Cancel = True
    End If
End Sub



Private Sub mmnuBHTRoomOccupancy_Click()
    frmRoomPatients.Show
    frmRoomPatients.ZOrder 0
End Sub

Private Sub mnnuBackOfficeGSBSummery_Click()
    frmAllGSBSummeryF.Show
    frmAllGSBSummeryF.ZOrder 0
End Sub

Private Sub mnuAbout_Click()
    frmTem.Show
End Sub

Private Sub mnuAdmit_Click()
    frmAdmit.Show
    frmAdmit.ZOrder 0
    frmAdmit.Top = 0
    frmAdmit.Left = 0
End Sub

Private Sub mnuAgentCancellations_Click()
    With frmAgentBillCancellationSearch
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuAgentPayments_Click()
    With frmAgentPayments
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuAgentReprints_Click()
    With frmAgentBillReprint
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBackOfficeAgentPayments_Click()
    frmAgentBills.Show
    frmAgentBills.ZOrder 0
End Sub

Private Sub mnuBackOfficeAllShiftEndSummeries_Click()
    frmAllShiftEndSummeries.Show
    frmAllShiftEndSummeries.ZOrder 0
    
End Sub

Private Sub mnuBackOfficeBHTBookKeepingSummery_Click()
    frmBHTBookeepingSummery.Show
    frmBHTBookeepingSummery.ZOrder 0
End Sub

Private Sub mnuBackOfficeBHTBookKeepingSummeryExtended_Click()
    frmBHTBookeepingSummeryExtended.Show
    frmBHTBookeepingSummeryExtended.ZOrder 0
End Sub

Private Sub mnuBackOfficeBHTPaymentBookKeepingSummery_Click()
    frmBHTPaymentBookeepingSummery.Show
    frmBHTPaymentBookeepingSummery.ZOrder 0
End Sub

Private Sub mnuBackOfficeBHTPayments_Click()
    On Error Resume Next
    With frmHealthSchemeSupplierBillSettling
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With

End Sub

Private Sub mnuBackOfficeBHTServicesActualSummery_Click()
    frmBHTServicesActualSummery.Show
    frmBHTServicesActualSummery.ZOrder 0
End Sub

Private Sub mnuBackOfficeBHTServicesActualSummeryD_Click()
    frmBHTServicesActualSummeryD.Show
    frmBHTServicesActualSummeryD.ZOrder 0

End Sub

Private Sub mnuBackOfficeBHTServicesOPDValue_Click()
    frmBHTServicesOPDValue.Show
    frmBHTServicesOPDValue.ZOrder 0
End Sub

Private Sub mnuBackOfficeBHTServicesOPDValueD_Click()
    frmBHTServicesOPDValueD.Show
    frmBHTServicesOPDValueD.ZOrder 0
End Sub

Private Sub mnuBackOfficeBHTSummery_Click()
    frmAllBHTSummeryF.Show
    frmAllBHTSummeryF.ZOrder 0
End Sub

Private Sub mnuBackOfficeBookKeepingSummery_Click()
    frmGSBBookeepingSummery.Show
    frmGSBBookeepingSummery.ZOrder 0
End Sub

Private Sub mnuBackOfficeBookKeepingSummeryE_Click()
    frmGSBBookeepingSummeryExtended.Show
    frmGSBBookeepingSummeryExtended.ZOrder 0
End Sub

Private Sub mnuBackOfficeDailySummeryReport_Click()
    frmDailySummeryReport2.Show
    frmDailySummeryReport2.ZOrder 0
End Sub

Private Sub mnuBackOfficeDailySummeryReport2_Click()
    frmDailySummeryReport.Show
    frmDailySummeryReport.ZOrder 0
End Sub

Private Sub mnuBackOfficeDailySummeryReportHospital_Click()
    frmDailySummeryReportHospital.Show
    frmDailySummeryReportHospital.ZOrder o
End Sub

Private Sub mnuBackOfficeDayEndCounts_Click()
    With frmDayEndSummeryCounts
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBackOfficeDetailedSummery_Click()
    With frmSummeryDetails
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBackOfficeExpenceDetails_Click()
    frmExpenceDetails.Show
    frmExpenceDetails.ZOrder 0
End Sub

Private Sub mnuBackOfficeGSBAgeAnalysis_Click()
    frmAgeAnalysisPatientGSB.Show
    frmAgeAnalysisPatientGSB.ZOrder 0
End Sub

Private Sub mnuBackOfficeGSBPaymentBookKeepingSummery_Click()
    frmGSBPaymentBookeepingSummery.Show
    frmGSBPaymentBookeepingSummery.ZOrder 0
End Sub

Private Sub mnuBackOfficeHealthSchemeSuppliersAgeAnalysis_Click()
    With frmAgeAnalysisBHT
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBackOfficeHSSCurrentBalance_Click()
    With frmCurrentCompanyBalance
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBackOfficeHSSMTAgeAnalysis_Click()
    frmAgeAnalysisMT.Show
    frmAgeAnalysisMT.ZOrder 0
End Sub

Private Sub mnuBackOfficeHSSMTPayments_Click()
    frmMTHealthSchemeSupplierBillSettling.Show
    frmMTHealthSchemeSupplierBillSettling.ZOrder 0
End Sub

Private Sub mnuBackOfficeMedicalTestBillsWithLabServices_Click()
    frmMTBillsWithLabServices.Show
    frmMTBillsWithLabServices.ZOrder 0
End Sub

Private Sub mnuBackOfficeMTCreditLetters_Click()
    frmMTCreditLetters.Show
    frmMTCreditLetters.ZOrder 0
End Sub

Private Sub mnuBackOfficePatientsAgeAnalysis_Click()
    frmAgeAnalysisPatientBHT.Show
    frmAgeAnalysisPatientBHT.ZOrder 0
End Sub

Private Sub mnuBackOfficePatientsMTAgeAnalysis_Click()
    frmAgeAnalysisMT.Show
    frmAgeAnalysisMT.ZOrder 0
End Sub

'Private Sub mnuBackOfficePayments_Click()
'    On Error Resume Next
'    With frmHealthSchemeSupplierBillSettling
'        .Show
'        .ZOrder 0
'        .Top = 0
'        .Left = 0
'    End With
'End Sub

Private Sub mnuBackOfficePaymentsAgeAnalysis_Click()
    frmBHTProfessionalPaymentsAgeAnalysis.Show
    frmBHTProfessionalPaymentsAgeAnalysis.ZOrder 0
End Sub

Private Sub mnuBackOfficePaymentSummery_Click()
    With frmPaymentSummery
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBackOfficePeriodSummeryReport_Click()
    frmPeriodSummeryDetails.Show
    frmPeriodSummeryDetails.ZOrder 0
End Sub

Private Sub mnuBackOfficePeriodSummeryReportByDate_Click()
    frmPeriodSummeryDetailsByDate.Show
    frmPeriodSummeryDetailsByDate.ZOrder 0
End Sub

Private Sub mnuBackOfficeShiftEndCounts_Click()
    With frmShiftEndSUmmeryCounts
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBackOfficeViewServices_Click()
    With frmServiceDetails
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBackup_Click()
    With frmBackUp
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBHTAddDetailsOfSx_Click()
    frmAddDetailsOfSx.Show
    frmAddDetailsOfSx.ZOrder 0
End Sub

Private Sub mnuBHTAdditionalCharges_Click()
    With frmBHTCharge
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBHTAdmissionBook_Click()
    frmAdmissionBook.Show
    
End Sub

Private Sub mnuBHTAdvanceBulkDischarge_Click()
    frmBHTBulkDischarge.Show
    frmBHTBulkDischarge.ZOrder 0
End Sub

Private Sub mnuBHTBHTSummery_Click()
    With frmBHTSummeryF
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBHTCashierDischarge_Click()
    frmBHTSummeryFNoDischarge.Show
    frmBHTSummeryFNoDischarge.ZOrder 0
End Sub

Private Sub mnuBHTCashierSummery_Click()
    frmBHTSummeryCashier.Show
    frmBHTSummeryCashier.ZOrder 0
End Sub

Private Sub mnuBHTEditRoomDetails_Click()
    frmEditRoomDetails.Show
    frmEditRoomDetails.ZOrder 0
End Sub

Private Sub mnuBHTGetDetailsOfSx_Click()
    frmGetDetailsOfSx.Show
    frmGetDetailsOfSx.ZOrder 0
End Sub

Private Sub mnuBHTPaymentBillSearch_Click()
    frmSearchBHTBill.Show
    frmSearchBHTBill.ZOrder 0
End Sub

Private Sub mnuBHTPaymentCancellations_Click()
    With frmBHTBillCancellationSearch
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBHTPaymentsPay_Click()
    With frmBHTPayments
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBHTPaymentsReprints_Click()
    With frmBHTBillReprintSearch
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBHTProfessionalCHarges_Click()
    With frmBHTProfessionalCharges
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBHTProfessionalCHargesD_Click()
    With frmBHTProfessionalChargesD
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With

End Sub

Private Sub mnuBHTRefunds_Click()
    With frmBHTRefund
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBHTReverseDischarge_Click()
    frmReverseDischarge1.Show
    frmReverseDischarge1.ZOrder 0
End Sub

Private Sub mnuBHTServices_Click()
    With frmBHTServiceBills
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBHTServicesD_Click()
    With frmBHTServiceBillsD
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With

End Sub

Private Sub mnuBHTServiceValues_Click()
    frmBHTServiceValues.Show
    frmBHTServiceValues.ZOrder 0
End Sub

Private Sub mnuBHTServiceValuesByDOD_Click()
    frmBHTServiceValuesByDOD.Show
    frmBHTServiceValuesByDOD.ZOrder 0
End Sub

Private Sub mnuBHTSummeryA_Click()
    With frmBHTSummery
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuBHTSummeryAllF_Click()
    With frmBHTSummeryFD
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With

End Sub

Private Sub mnuChangeRoom_Click()
    With frmChangeRoom
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuDayEndSummery_Click()
    With frmDayEndSummery
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuDepartments_Click()
    With frmDepartments
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuDischarge_Click()
    With frmDischarge
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuEditAgents_Click()
    With frmAgents
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With

End Sub

Private Sub mnuEditBHT_Click()
    With frmBHTEdit
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuEditCategoriesServiceSecessions_Click()
    With frmServiceSecession
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuEditHospitalCharges_Click()
    frmHospitalCharges.Show
    frmHospitalCharges.ZOrder 0
End Sub

Private Sub mnuEditPatientCategory_Click()
    With frmPatientCategory
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuEditProfessionalChargesForServiceCategories_Click()
    With frmPreofessionalChargesForServiceCategories
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuEditPrevilagesMenuEnable_Click()
    With frmAuthorityPrevilagesMenuEnable
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuEditProfessionalChargesForServiceSubcategories_Click()
    With frmProfessionalChargesForServiceSubcategories
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuEditRoomCategory_Click()
    With frmRoomCategory
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuEditServiceCategory_Click()
    With frmServiceCategory
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuEditServiceSubcategory_Click()
    With frmServiceSubCategory
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuEditSpeciality_Click()
    frmSpeciality.Show
    frmSpeciality.ZOrder 0
End Sub

Private Sub mnuEditSurgeries_Click()
    frmSx.Show
    frmSx.ZOrder 0
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuGSBAdmit_Click()
    With frmAdmitGSBNew
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuGSBEDIT_Click()
    frmGSBEdit.Show
    frmGSBEdit.ZOrder 0
End Sub

Private Sub mnuGSBPay_Click()
    With frmGSBPayments
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuGSBPaymentCancellations_Click()
    With frmGSBBillCancellationSearch
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuGSBPaymentReprint_Click()
    With frmGSBBillReprintSearch
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuGSBPaymentSearch_Click()
    frmSearchGSBill.Show
    frmSearchGSBill.ZOrder 0
End Sub

Private Sub mnuGSBProfessionalCharges_Click()
    With frmGSBProfessionalCharges
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuGSBRefunds_Click()
    With frmGSBRefund
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuGSBReprint_Click()
    frmEditGSB.Show
    frmEditGSB.ZOrder 0
End Sub

Private Sub mnuGSBServices_Click()
    With frmGSBServiceBills
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuGSBServiceValuesByDOD_Click()
    frmGSBServiceValuesByDOD.Show
    frmGSBServiceValuesByDOD.ZOrder 0
End Sub

Private Sub mnuGSBSummery_Click()
    With frmGSBSummery
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuHealthSchemeSuppliers_Click()
    With frmHealthSchemeSuppliers
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuHelpAdministrator_Click()
    With frmAdministrator
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuLabBillCancellation_Click()
    With frmLabBillCancellationSearch
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuLabBillReprint_Click()
    With frmLabBillReprintSearch
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuLabBills_Click()
    With frmLabServiceBills
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub



Private Sub mnuPatientDetails_Click()
    With frmPatientsDetails
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuLabBillsList_Click()
    frmLabBillList.Show
    frmLabBillList.ZOrder 0
End Sub

Private Sub mnuLabLabServiceValues_Click()
    frmLabServiceValues.Show
    frmLabServiceValues.ZOrder 0
End Sub

Private Sub mnuLabPatientList_Click()
    frmLabServicePatientList.Show
    frmLabServicePatientList.ZOrder 0
End Sub

Private Sub mnuLabSearchBills_Click()
    frmSearchLabBill.Show
    frmSearchLabBill.ZOrder 0
End Sub

Private Sub mnuLabServiceBills_Click()
    frmLabServiceCategoryList.Show
    frmLabServiceCategoryList.ZOrder 0
End Sub

Private Sub mnuMedicalTestBillCancellation_Click()
    With frmMTBillCancellationSearch
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuMedicalTestBillRefunds_Click()
    frmMTRefund.Show
    frmMTRefund.ZOrder 0
End Sub

Private Sub mnuMedicalTestBillReprints_Click()
    With frmMTBillReprintSearch
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With

End Sub

Private Sub mnuMedicalTestBills_Click()
    With frmMTBills
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuMTBillSearch_Click()
    frmSearchMTBill.Show
    frmSearchMTBill.ZOrder 0
End Sub

Private Sub mnuMTServiceValues_Click()
    frmMTServiceValues.Show
    frmMTServiceValues.ZOrder 0
End Sub

Private Sub mnuOPDBillCancellation_Click()
    With frmOPDBillCancellationSearch
        .Show
        .ZOrder 0
    End With
End Sub

Private Sub mnuOPDBillReprints_Click()
    With frmOPDBillReprintSearch
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuOPDBills_Click()
    With frmOPDServiceBills
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub


Private Sub mnuOPDRefunds_Click()
    With frmOPDRefund
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuOPDSearchBills_Click()
    frmSearchOPDBill.Show
    frmSearchOPDBill.ZOrder 0
End Sub

Private Sub mnuOPDServiceBills_Click()
    With frmOPDServiceCategoryList
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuOPDSErviceCounts_Click()
    With frmOPDServiceCounts
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuOPDServiceValues_Click()
    With frmOPDServiceValues
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuPaymentsAddExpences_Click()
    frmExpenceBills.Show
    frmExpenceBills.ZOrder 0
End Sub

Private Sub mnuPaymentsCancelExpences_Click()
    frmExpenceBillCancellationSearch.Show
    frmExpenceBillCancellationSearch.ZOrder 0
End Sub

Private Sub mnuPaymentsHSSBHT_Click()
    With frmHSSBillSettlingCashier
        .Show
        .ZOrder 0
'        .Top = 0
'        .Left = 0
    End With
End Sub

Private Sub mnuPaymentsHSSMT_Click()
    frmMTHealthSchemeSupplierBillSettlingCashier.Show
    frmMTHealthSchemeSupplierBillSettlingCashier.ZOrder 0
End Sub

Private Sub mnuPaymentsProfessionalFeePaymentsForGSB_Click()
    With frmGSBProfessionalPayments
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuPaymentsProfessionalFeePaymentsForInwardPatients_Click()
    With frmBHTProfessionalPayments
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuPaymentsProfessionalFeePaymentsForLabPatients_Click()
    With frmLabProfessionalPayments
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuPaymentsProfessionalFeePaymentsForOPDPatients_Click()
    With frmOPDProfessionalPayments
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuPaymentsProfessionalFeePaymentsForRPatients_Click()
    frmRProfessionalPayments.Show
    frmRProfessionalPayments.ZOrder 0
End Sub

Private Sub mnuPharmacyBillCancellation_Click()
    With frmPharmacyBillCancellation
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuPharmacyBillReturn_Click()
    With frmPharmacyBillReturns
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuPharmacyBills_Click()
    With frmPharmacyBills
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuPreferanceHospitalCharges_Click()
    With frmInwardPatientCharges
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuPrevilagesMenuVisible_Click()
    With frmAuthorityPrevilagesMenuVisible
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuPrintingPreferances_Click()
    With frmPrintingPreferances
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuProgramPreferances_Click()
    With frmProgramPreferance
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub


Private Sub mnuRBillCancellation_Click()
    frmRBillCancellationSearch.Show
    frmRBillCancellationSearch.ZOrder 0
End Sub

Private Sub mnuRBillRefunds_Click()
    frmRRefund.Show
    frmRRefund.ZOrder 0
End Sub

Private Sub mnuRBillReprints_Click()
    frmRBillReprintSearch.Show
    frmRBillReprintSearch.ZOrder 0
End Sub

Private Sub mnuRBills_Click()
    frmRServiceBills.Show
    frmRServiceBills.ZOrder 0
End Sub

Private Sub mnuRestore_Click()
    With frmRestore
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuRBillsByCategories_Click()
    frmRServiceCategoryList.Show
    frmRServiceCategoryList.ZOrder 0
End Sub

Private Sub mnuRBillValues_Click()
    frmRServiceValues.Show
    frmRServiceValues.ZOrder 0
End Sub

Private Sub mnuRooms_Click()
    With frmRoom
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuRSarchBills_Click()
    frmSearchRBill.Show
    frmSearchRBill.ZOrder 0
End Sub

Private Sub mnuShiftEndSummery_Click()
    With frmShiftEndSummery
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuStaff_Click()
    With frmStaff
        .Show
        .ZOrder 0
        .Top = 0
        .Left = 0
    End With
End Sub

Private Sub mnuTipOfTheDay_Click()
    frmTem.Show
End Sub

Private Sub Timer1_Timer()
    lblDateTime.Caption = Format("Date : " & Format(Date, "dd MMMM yyyy") & "   Time : " & Format(Time, "H:M AMPM"))
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        
        Case 5:
        Case 6:
        Case 7: mnuExit_Click
    End Select
End Sub

