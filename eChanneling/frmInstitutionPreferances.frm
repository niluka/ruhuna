VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmInstitutionPreferances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Institution Preferances"
   ClientHeight    =   5535
   ClientLeft      =   4440
   ClientTop       =   1680
   ClientWidth     =   8880
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   8880
   Begin VB.Frame framInstutions 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8655
      Begin VB.TextBox txtInsname 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   6375
      End
      Begin VB.TextBox txtDiscription 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   720
         Width           =   6375
      End
      Begin VB.TextBox txtRegistration 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   1200
         Width           =   6375
      End
      Begin VB.TextBox txtAddress01 
         Height          =   975
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   6375
      End
      Begin VB.TextBox txtTelephone01 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtTelephone02 
         Height          =   375
         Left            =   5280
         TabIndex        =   7
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtEmail01 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtEmail02 
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtwbsite01 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   4200
         Width           =   3135
      End
      Begin VB.TextBox txtWebsite02 
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   4200
         Width           =   3135
      End
      Begin VB.TextBox txtFax 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   3240
         Width           =   6375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution &Name"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Discription"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Registration No"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Address"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Tele&phone No"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Email "
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "&Website"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fax"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   2295
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   4920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Save / Exit"
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
Attribute VB_Name = "frmInstitutionPreferances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim SuppliedWord As String

Private Sub Setcolours()
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
    framInstutions.BackColor = FrameBackColour
    framInstutions.ForeColor = FrameForeColour
    Me.BackColor = FrameBackColour
    Me.ForeColor = FrameForeColour
End Sub


Private Sub Form_Load()
    If UserAuthority = AuthorityAdministrator Then
        txtInsname.Locked = False
    Else
        txtInsname.Locked = True
    End If
    Call SetPreferances
    Call Setcolours
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub SetPreferances()
        Call GetInstitutionDetails
End Sub

Private Sub SavePreferancesToFile()
    Call SaveInstitutionDetails
End Sub

Private Sub SavePreferancesToMemory()
    Call SaveInstitutionDetailsToMemory
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SavePreferancesToFile
    Call SavePreferancesToMemory
End Sub
Private Sub SaveInstitutionDetails()
Dim TemResponce As Long

'If txtInsname.Text = "" Then
'    TemResponce = MsgBox("You have not entered an Institution Name", vbCritical, "? Institution Name")
'    txtInsname.SetFocus
'    Exit Sub
'End If

'On Error GoTo ErrorHandler

With DataEnvironment1.rscmmdInstitutionDetails
    
    If .State = 0 Then .Open
    

    SuppliedWord = txtInsname.Text
    !InstitutionName = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtDiscription.Text
    !InstitutionDescription = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtRegistration.Text
    !InstitutionRegistation = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtAddress01.Text
    !InstitutionAddress = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtTelephone01.Text
    !institutiontelephone1 = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtTelephone02.Text
    !InstitutionTelephone2 = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtFax.Text
    !InstitutionFax = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtEmail01.Text
    !InstitutionEmail = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtEmail02.Text
    !InstitutionEmail2 = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtwbsite01.Text
    !InstitutionWebSite1 = EncreptedWord(SuppliedWord)
    
    SuppliedWord = txtWebsite02.Text
    !InstitutionWebSite2 = EncreptedWord(SuppliedWord)
    
    
    .Update

    If .State = 1 Then .Close
    
End With

Exit Sub

ErrorHandler:
    MsgBox ("An Error Occured during Updating" & vbNewLine & Me.Caption & vbNewLine & Err.Number & vbNewLine & Err.Description)
    DataEnvironment1.rscmmdInstitutionDetails.CancelUpdate

End Sub

Private Sub SaveInstitutionDetailsToMemory()
    Dim TemResponce As Long
'    If txtInsname.Text = "" Then
'        TemResponce = MsgBox("You have not entered an Institution Name", vbCritical, "? Institution Name")
'        txtInsname.Enabled = True
'        txtInsname.SetFocus
'        Exit Sub
'    End If
    InstitutionName = txtInsname.Text
    InstitutionAddress = txtAddress01.Text
    InstitutionTelephone = txtTelephone01.Text
End Sub

Private Sub GetInstitutionDetails()

With DataEnvironment1.rscmmdInstitutionDetails
    If .State = 0 Then .Open
    
    
    SuppliedWord = !InstitutionName
    txtInsname.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionDescription
    txtDiscription.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionRegistation
    txtRegistration.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionAddress
    txtAddress01.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !institutiontelephone1
    txtTelephone01.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionTelephone2
    txtTelephone02.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionFax
    txtFax.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionEmail
    txtEmail01.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionEmail2
    txtEmail02.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionWebSite1
    txtwbsite01.Text = DecreptedWord(SuppliedWord)
    
    SuppliedWord = !InstitutionWebSite2
    txtWebsite02.Text = DecreptedWord(SuppliedWord)
    
    If .State = 1 Then .Close
    
End With

End Sub
