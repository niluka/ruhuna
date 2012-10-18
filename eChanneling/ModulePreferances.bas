Attribute VB_Name = "ModulePreferances"
Option Explicit
    
    Public cnnChannelling As New ADODB.Connection

    Public UserName As String
    Public UserID As Long
    Public UserAuthority  As Integer
    
    Public LongAd As String
    Public ShortAd As String
    
    Public InstitutionName As String
    Public InstitutionAddress As String
    Public InstitutionTelephone As String
    
    Public PrintingOnBlankPaper As Boolean
    Public PrintingOnPrintedPaper As Boolean
    
    Public AdLong As String
    Public AdShort As String
    
    Public DatabasePath As String
        
    Public AskBeforeAdding As Boolean
    Public AgentEssential As Boolean
    
    Public PreferanceColourScheme  As Integer
    
    Public BillPrinterName As String
    Public ReportPrinterName As String
    Public BillPaperName As String
    Public BillPaperHeight As Long
    Public BillPaperWidth As Long
    Public ReportPaperName As String
    Public ReportPaperHeight As Long
    Public ReportPaperWidth As Long
    Public AllowNameChange As Boolean
    Public AddForeignerSuffix As Boolean
    Public AutomaticCapitalization As Boolean
    Public AgentNameForCreditBookings As Boolean
    Public SurnameFirst As Boolean
    Public NoAllNames As Boolean
    Public CanSelectAgent As Boolean
    Public ChangeToCash As Boolean
    Public ClearAgentDetails As Boolean
    Public AllowReprint As Boolean
    Public BackUpPath As String
    Public PayToDoctor As Boolean
    Public AfterAddSpeciality As Boolean
    Public AfterAddConsultant As Boolean
    Public AfterAddDates As Boolean
    Public AfterAddPatient As Boolean
    Public AllowAbsent As Boolean
    Public HospitalDetails As Boolean
    Public DisplayPrintChkBox As Boolean
    Public AgentCashOnly As Boolean
    Public EnglishDateFormat As Boolean
    Public DefaultShortDate As String
    Public DefaultLongDate As String
    
    Public AgentBookingValidation As Boolean
    Public DoctorPaymentDetailedReport As Boolean
    Public AgentBillNumber As Boolean
    
    Public OnePrintForAgents As Boolean
    
    Public PaymentCash As Integer
    Public PaymentCredit As Integer
    Public PaymentAgent As Integer
    
    Public PartialRepayments As Boolean
    Public DetailedCount As Boolean
    
    Public Const AuthorityAdministrator  As Integer = 1
    Public Const AuthorityOwner  As Integer = 2
    Public Const AuthorityOwnerCOvered  As Integer = 3
    Public Const AuthorityHumanResources  As Integer = 4
    Public Const AuthorityAccount  As Integer = 5
    Public Const AuthorityUser  As Integer = 6
    Public Const AuthorityAnalyzer  As Integer = 7
    Public Const NoColourScheme  As Integer = 0
    Public Const ColourEnergy  As Integer = 1
    Public Const ColourAqua  As Integer = 2
    Public Const ColourSunny  As Integer = 3
    Public ColourScheme  As Integer
    
    Public BttnBackColour As Long
    Public BttnForeColour As Long
    Public FrmBackColour As Long
    Public FrmForeColour As Long
    Public FrameBackColour As Long
    Public FrameForeColour As Long
    Public TxtBackColour As Long
    Public TxtForeColour As Long
    Public LblBackColour As Long
    Public LblForeColour As Long
    
    Public GridBackColor As Long
    Public GridBackColorBkg As Long
    Public GridBackColorFixed As Long
    Public GridBackColorSel As Long
    
    Public GridForeColor As Long
    Public GridForeColorFixed As Long
    Public GridForeColorSel As Long
    
    Public Const MorningSecession  As Integer = 1
    Public Const EveningSecession  As Integer = 2
    Public Const NoSecessionPreferance  As Integer = 3
    Public Const NoReleventSecession  As Integer = 4
    Public Const Doctor      As Integer = 1
    Public Const Staff  As Integer = 2
    Public Const Investigation  As Integer = 3
    Public Const Other  As Integer = 4
    
    Public IncomeDeflation  As Integer
    Public AdvanceBookingDays As Long
    
    Public TemTotalCash As Double
    Public TemTotalPayment As Double
    
    Public TemUserPassward As String
    Public LoginSucceeded As Boolean
    Public CheckLogin As Boolean
