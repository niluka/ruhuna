VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lakmedipro - eHospital Assistant"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTemUsername 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   3645
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx cmdOK 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Login"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "&Password"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "&User Name"
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
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    Dim SuppliedWord As String
    
    
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Dim TemResponce As Byte
    Dim UserNameFound As Boolean
    UserNameFound = False
    TemUserPassward = txtPassword.Text
    If Trim(txtUserName.Text) = "" Then
        TemResponce = MsgBox("You have not entered a username", vbCritical, "Username")
        txtUserName.SetFocus
        Exit Sub
    End If
    If Trim(txtPassword.Text) = "" Then
        TemResponce = MsgBox("You have not entered a password", vbCritical, "Password")
        txtPassword.SetFocus
        Exit Sub
    End If
    With DataEnvironment1.rssqlStaff
        If .State = 1 Then .Close
        .Source = "Select tblstaff.* from tblstaff where (StaffUser = 1)"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            txtTemUsername.Text = DecreptedWord(!StaffUserName)
            If UCase(txtUserName.Text) = UCase(txtTemUsername.Text) Then
                UserNameFound = True
                If txtPassword.Text = DecreptedWord(!staffpassword) Then
                If CheckLogin = True Then
                        If !logged = 1 Then
                            TemResponce = MsgBox("There is a user currently logged using the same username. Please check on that. If no one is logged from the user name, it may be due to a sudden breakdown of the program, for example, after a power failure. In this case, advice the owner to and Reset All Users Login", vbCritical, "User Logged")
                            txtUserName.SetFocus
                            SendKeys "{home}+{end}"
                            Exit Sub
                        Else
                            !logged = 1
                            .Update
                        End If
                End If
                    UserName = UCase(txtUserName.Text)
                    UserID = !Staff_ID
                    
                    If Not IsNull(!StaffAuthority) Then
                        UserAuthority = !StaffAuthority
                    Else
                        UserAuthority = 0
                    End If
                    Exit Do
                Else
                    TemResponce = MsgBox("The username and password you entered are not matching. Please try again", vbCritical, "Wrong Username and Password")
                    txtUserName.SetFocus
                    SendKeys "{home}+{end}"
                    Exit Sub
                End If
            Else
            End If
            .MoveNext
        Loop
        .Close
        If UserNameFound = False Then
            TemResponce = MsgBox("There is no such  a username, Please try again", vbCritical, "Username")
            txtUserName.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        End With
        Unload Me
        
        MDIFrmReception.Show
        MDIFrmReception.Caption = MDIFrmReception.Caption & " - " & UserName
End Sub

Private Sub Form_Load()
    On Error GoTo EH
    Dim TemResponce As Byte
    Dim WillExpire As Boolean
    Dim ExpiaryDate As Date
    Dim constr As String
    
    WillExpire = False
    ExpiaryDate = #3/31/2012#
    If WillExpire = True And ExpiaryDate < Date Then
        TemResponce = MsgBox("The Program has expiared. Please contact Lakmedipro for Assistant", vbCritical, "Expired")
        End
    End If
    Dim TemPath As String
    Call LoadPreferances
'    If FSys.FileExists(DatabasePath) = False Then
'        TemResponce = MsgBox("The path of the database you have selected does not exist. Please select the database", vbCritical, "Wrong database path")
'        frmInitialPreferances.Show 1
'    End If

'    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & DatabasePath & " ;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
'    cnnChannelling.Open constr

    
'    constr = "Provider=SQLOLEDB.1;Password=Bud7Nil;Persist Security Info=True;User ID=sa;Initial Catalog=HospitalSQL;Data Source=SRV"
    
'    DataEnvironment1.cnnHospital.ConnectionString = "data source=" & DatabasePath  'GetSetting(App.EXEName, "Options", "DatabaseLocation", App.Path & "\hospital.mdb")
    
    
    constr = "Driver={SQLOLEDB.1};" & _
           "Server=SRV;" & _
           "Database=HospitalSQL;" & _
           "Uid=sa;" & _
           "Pwd=Bud7Nil"

    
    
    'cnnChannelling.Open constr
    
    cnnChannelling.Open DataEnvironment1.cnnHospital.ConnectionString
    
    'DataEnvironment1.cnnHospital.ConnectionString = constr
    
    
    
    Call LoadInstitutionDetails
    Call GetAds
    
    Exit Sub
    
EH:

    MsgBox "Error Connecting to server"
    End
    
End Sub


Private Sub GetAds()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "select * from tblads"
    .Open
    If .RecordCount = 0 Then Exit Sub
    Dim TemNum As Long
    Randomize
    TemNum = Round((Rnd * .RecordCount - 1), 0) + 1
    If TemNum < 1 Then TemNum = 0
    If TemNum > .RecordCount Then TemNum = .RecordCount
    If .State = 1 Then .Close
    .Source = "select * from tblads where id = " & TemNum
    .Open
    If .RecordCount = 0 Then Exit Sub
    LongAd = !AdLong
    ShortAd = !AdShort
    .Close
End With
End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtUserName.Text <> "" Then cmdOK_Click: Exit Sub
    If KeyAscii = 13 And txtUserName.Text = "" Then txtUserName.SetFocus: Exit Sub
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

Private Sub LoadPreferances()
'    DatabasePath = "D:\hospital.mdb"

    DatabasePath = GetSetting(App.EXEName, "Options", "DatabaseLocation", App.Path & "\hospital.mdb")
    PreferanceColourScheme = GetSetting(App.EXEName, "Options", "Colour Scheme", ColourEnergy)
    BillPrinterName = GetSetting(App.EXEName, "Options", "BillPrinterName", "Fax")
    BillPaperName = GetSetting(App.EXEName, "Options", "BillPaperName", "Fax")
    ReportPrinterName = GetSetting(App.EXEName, "Options", "ReportPrinterName", "Fax")
    ReportPaperName = GetSetting(App.EXEName, "Options", "ReportPaperName", "Fax")
    IncomeDeflation = Val(GetSetting(App.EXEName, "Options", "IncomeDeflation", "3"))
    ColourScheme = 1 'Val(GetSetting(App.EXEName, "Options", "ColourScheme", "1"))
    AdvanceBookingDays = Val(GetSetting(App.EXEName, "Options", "AdvanceBookingDays", "3"))
    PrintingOnBlankPaper = GetSetting(App.EXEName, "Options", "PrintingOnBlankPaper", False)
    PrintingOnPrintedPaper = GetSetting(App.EXEName, "Options", "PrintingOnPrintedPaper", True)
    AskBeforeAdding = GetSetting(App.EXEName, "Options", "AskBeforeAdding", True)
    AgentEssential = GetSetting(App.EXEName, "Options", "agentessential", True)
    AllowNameChange = GetSetting(App.EXEName, "Options", "AllowNameChange", True)
    AddForeignerSuffix = GetSetting(App.EXEName, "Options", "AddForeignerSuffix", False)
    AutomaticCapitalization = GetSetting(App.EXEName, "Options", "AutomaticCapitalization", True)
    AgentNameForCreditBookings = GetSetting(App.EXEName, "Options", "AgentNameForCreditBookings", True)
    NoAllNames = GetSetting(App.EXEName, "Options", "NoAllNames", False)
    SurnameFirst = GetSetting(App.EXEName, "Options", "SurnameFirst", True)
    CanSelectAgent = GetSetting(App.EXEName, "Options", "CanSelectAgent", True)
    ChangeToCash = GetSetting(App.EXEName, "Options", "ChangeToCash", False)
    AllowReprint = GetSetting(App.EXEName, "Options", "AllowReprint", True)
    BackUpPath = GetSetting(App.EXEName, "Options", "BackUpPath", App.Path)
    PayToDoctor = GetSetting(App.EXEName, "Options", "PayToDoctor", True)
    AllowAbsent = GetSetting(App.EXEName, "Options", "AllowAbsent", True)
    AfterAddSpeciality = GetSetting(App.EXEName, "Options", "AfterAddSpeciality", False)
    AfterAddConsultant = GetSetting(App.EXEName, "Options", "AfterAddConsultant", False)
    AfterAddDates = GetSetting(App.EXEName, "Options", "AfterAddDates", True)
    AfterAddPatient = GetSetting(App.EXEName, "Options", "AfterAddPatient", False)
    HospitalDetails = GetSetting(App.EXEName, "Options", "HospitalDetails", True)
    DisplayPrintChkBox = GetSetting(App.EXEName, "Options", "DisplayPrintChkBox", True)
    PaymentAgent = Val(GetSetting(App.EXEName, "Options", "PaymentAgent", 1))
    PaymentCredit = Val(GetSetting(App.EXEName, "Options", "PaymentCredit", 1))
    PaymentCash = Val(GetSetting(App.EXEName, "Options", "PaymentCash", 1))
    AgentCashOnly = GetSetting(App.EXEName, "Options", "AgentCashOnly", True)
    EnglishDateFormat = GetSetting(App.EXEName, "Options", "EnglishDateFormat", True)
    PartialRepayments = GetSetting(App.EXEName, "Options", "PartialRepayments", False)
    DetailedCount = GetSetting(App.EXEName, "Options", "DetailedCount", True)
    OnePrintForAgents = GetSetting(App.EXEName, "Options", "OnePrintForAgents", False)
    CheckLogin = GetSetting(App.EXEName, "Options", "CheckLogin", False)
    AgentBookingValidation = GetSetting(App.EXEName, "Options", "AgentBookingValidation", False)
    DoctorPaymentDetailedReport = GetSetting(App.EXEName, "Options", "DoctorPaymentDetailedReport", False)
    AgentBillNumber = GetSetting(App.EXEName, "Options", "AgentBillNumber", False)
    If EnglishDateFormat = True Then
        DefaultLongDate = "DD MMMM YYYY"
        DefaultShortDate = "dd mm yy"
    Else
        DefaultLongDate = "YYYY MMMM DD"
        DefaultShortDate = "yy mm dd"
    End If
End Sub

Private Sub LoadInstitutionDetails()
'    With DataEnvironment1.rscmmdInstitutionDetails
'        If .State = 0 Then .Open
'        SuppliedWord = !InstitutionName
'        InstitutionName = DecreptedWord(SuppliedWord)
'        SuppliedWord = !InstitutionAddress
'        InstitutionAddress = DecreptedWord(SuppliedWord)
'        SuppliedWord = !institutiontelephone1
'        InstitutionTelephone = DecreptedWord(SuppliedWord)
'        If .State = 1 Then .Close
'    End With
End Sub
