VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lakmedipro - eCashier"
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
   Begin MSDataListLib.DataCombo dtcDepartment 
      Height          =   360
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtTemUsername 
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   840
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
      Top             =   1440
      Width           =   3645
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   2160
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
      Top             =   2160
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
   Begin VB.Label Label3 
      Caption         =   "&Department"
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
      TabIndex        =   8
      Top             =   240
      Width           =   1815
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
      TabIndex        =   6
      Top             =   1440
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
      TabIndex        =   5
      Top             =   960
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
    Dim rsStaff As New ADODB.Recordset
    Dim rsStore As New ADODB.Recordset
    Dim rsAds As New ADODB.Recordset
    Dim rsHospital As New ADODB.Recordset
    Dim temSql As String
    Dim constr As String
    Dim TemUserPassward As String
    
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
    If Not IsNumeric(dtcDepartment.BoundText) Then
        TemResponce = MsgBox("You have not selected a department", vbCritical, "Department")
        dtcDepartment.SetFocus
        Exit Sub
    End If
    With rsStaff
        If .State = 1 Then .Close
        temSql = "Select tblstaff.* from tblstaff where (IsAUser = 1)"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 1 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            txtTemUsername.Text = DecreptedWord(!UserName)
            If UCase(txtUserName.Text) = UCase(txtTemUsername.Text) Then
                UserNameFound = True
                If txtPassword.Text = DecreptedWord(!Password) Then
                    UserName = UCase(txtUserName.Text)
                    UserID = !StaffID
                    UserFullName = !Name
                    If Not IsNull(!AuthorityID) Then
                        UserAuthority = !AuthorityID
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
        
        UserStore = dtcDepartment.Text
        UserStoreID = dtcDepartment.BoundText
        
        Unload Me
        MDIMain.Show
        MDIMain.Caption = MDIMain.Caption & " - " & UserStore & " - " & UserFullName

End Sub

Private Sub Form_Load()
    Dim TemResponce As Byte
    Dim WillExpire As Boolean
    Dim ExpiaryDate As Date
    WillExpire = False
    ExpiaryDate = #3/31/2012#
    If WillExpire = True And ExpiaryDate < Date Then
        TemResponce = MsgBox("The Program has expiared. Please contact Lakmedipro for Assistant", vbCritical, "Expired")
        End
    End If
    Dim TemPath As String
    Call LoadPreferances
    
'    If FSys.FileExists(Database) = False Then
'        TemResponce = MsgBox("The path of the database you have selected does not exist. Please select the database", vbCritical, "Wrong database path")
'        frmInitialPreferances.Show 1
'    End If
    
    
    DatabasePath = App.Path ' FSys.GetParentFolderName(Database)
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & Database & " ;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
    cnnStores.Open Dataenvironment1.Connection1.ConnectionString
    
    'Dataenvironment1.Connection1.ConnectionString = "data source=" & constr  'GetSetting(App.EXEName, "Options", "DatabaseLocation", App.Path & "\hospital.mdb")
    
    Call LoadInstitutionDetails
    Call WriteDailyDetails
    Call FillCombo
    Call GetAds
    dtcDepartment.Text = GetSetting(App.EXEName, "Options", "dtcDepartment", "")
End Sub

Private Sub FillCombo()
    With rsStore
        If .State = 1 Then .Close
        temSql = "SELECT * from tblStore order by store"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcDepartment
        Set .RowSource = rsStore
        .ListField = "Store"
        .BoundColumn = "StoreID"
    End With
End Sub

Private Sub WriteDailyDetails()
Dim rsDailyIssue As New ADODB.Recordset
Dim rsIssue As New ADODB.Recordset


End Sub

Private Sub GetAds()
With rsAds
    
    If .State = 1 Then .Close
    temSql = "select * from tblAds"
    .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    If .RecordCount > 0 Then
        Dim TemNum As Long
        Randomize
        TemNum = Round((Rnd * .RecordCount - 1), 0) + 1
        If TemNum < 1 Then TemNum = 0
        If TemNum > .RecordCount Then TemNum = .RecordCount
        If .State = 1 Then .Close
        temSql = "select * from tblAds where ID = " & TemNum
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 1 Then
            LongAd = !AdLong
            ShortAd = !AdShort
        End If
    End If
    .Close
End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "Options", "dtcDepartment", dtcDepartment.Text
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtUserName.Text <> "" Then cmdOK_Click: Exit Sub
    If KeyAscii = 13 And txtUserName.Text = "" Then txtUserName.SetFocus: Exit Sub
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

Private Sub LoadPreferances()
    Database = GetSetting(App.EXEName, "Options", "Database", App.Path & "\eStore.mdb")
    DoNotAllowExpireConsumption = GetSetting(App.EXEName, "Options", "DoNotAllowExpireConsumption", True)
    DoNotAllowExpireSale = GetSetting(App.EXEName, "Options", "DoNotAllowExpireSale", True)
    DoNotAllowExpireTransfer = GetSetting(App.EXEName, "Options", "DoNotAllowExpireTransfer", True)
    LongDateFormat = GetSetting(App.EXEName, "Options", "LongDateFormat", "dddd, dd MMMM yyyy")
    ShortDateFormat = GetSetting(App.EXEName, "Options", "ShortDateFormat", "dd MM yy")
    BillPrinterName = GetSetting(App.EXEName, "Options", "BillPrinterName", "")
    BillPaperName = GetSetting(App.EXEName, "Options", "BillPaperName", "")
    PrescreptionPrinterName = GetSetting(App.EXEName, "Options", "PrescreptionPrinterName", "")
    ReportPrinterName = GetSetting(App.EXEName, "Options", "ReportPrinterName", "")
    ReportPaperName = GetSetting(App.EXEName, "Options", "ReportPaperName", "")
    PrescreptionPaperName = GetSetting(App.EXEName, "Options", "PrescreptionPaperName", "")
    PrintingOnBlankPaper = GetSetting(App.EXEName, "Options", "PrintingOnBlankPaper", True)
    PrintingOnPrintedPaper = GetSetting(App.EXEName, "Options", "PrintingOnPrintedPaper", False)
    HighRate = GetSetting(App.EXEName, "Options", "HighRate", 1)

    HospitalName = "Ruhunu Hospital (Pvt) Ltd"
    HospitalDescreption = "Karapitiya, Galle. Tel. 091 22 34059/60, Fax. 091 22 34061"
    HospitalAddress = "Tel. 091 22 34059/60, Fax. 091 22 34061"
    DefaultFont.Name = "Lucida Console"

    LabName = "Nawaloka Metropolis"
    LabDescreption = "Ruhunu Hospital, Karapitiya, Galle. Tel. 091 22 34059/60, Fax. 091 22 34061"
    LabAddress = "Tel. 091 5622244 , Fax. 091 22 34061"

    RadiologyName = "Roentgents International (Pvt) Ltd"
    RadiologyDescreption = "Ruhunu Hospital, Karapitiya, Galle. Tel. 091 22 34059/60, Fax. 091 22 34061"
    RadiologyAddress = "Tel. 091 22 34059/60, Fax. 091 22 34061"




End Sub

Private Sub LoadInstitutionDetails()
'    With rsHospital
'        If .State = 1 Then .Close
'        temSql = "SELECT * from tblInstitutionDetail"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount < 1 Then Exit Sub
'        SuppliedWord = !InstitutionName
'        HospitalName = DecreptedWord(SuppliedWord)
'        SuppliedWord = !InstitutionDescription
'        HospitalDescreption = DecreptedWord(SuppliedWord)
'        SuppliedWord = !institutionAddress
'        HospitalAddress = DecreptedWord(SuppliedWord)
'        If .State = 1 Then .Close
'    End With
End Sub
