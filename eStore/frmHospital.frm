VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmHospital 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hospital Profile"
   ClientHeight    =   5640
   ClientLeft      =   3390
   ClientTop       =   3390
   ClientWidth     =   9000
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
   Icon            =   "frmHospital.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9000
   Begin VB.Frame framInstutions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   8655
      Begin VB.TextBox txtInsname 
         Height          =   375
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   6375
      End
      Begin VB.TextBox txtDiscription 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   720
         Width           =   6375
      End
      Begin VB.TextBox txtRegistration 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1200
         Width           =   6375
      End
      Begin VB.TextBox txtAddress01 
         Height          =   975
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1680
         Width           =   6375
      End
      Begin VB.TextBox txtTelephone01 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtTelephone02 
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtEmail01 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtEmail02 
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtwbsite01 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   4200
         Width           =   3135
      End
      Begin VB.TextBox txtWebsite02 
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   4200
         Width           =   3135
      End
      Begin VB.TextBox txtFax 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   3240
         Width           =   6375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital &Name"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Discription"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Registration No"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Address"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Tele&phone No"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Email "
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "&Website"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fax"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3240
         Width           =   2295
      End
   End
   Begin btButtonEx.ButtonEx bttnOK 
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Top             =   5040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Ok"
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
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Cancel"
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
End
Attribute VB_Name = "frmHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemResponce  As Integer
    Dim rsIns As New ADODB.Recordset
    Dim A As String
    Dim temSql As String

Private Sub bttnOk_Click()
    If Date > #12/18/2010# Then
        TemResponce = MsgBox("Please contact Lakmedipro for Assistant", vbCritical, "Expired")
        End
    End If
    If Trim(txtInsname.Text) = "" Then
        TemResponce = MsgBox("You have not entered the name of the institution", vbCritical, "Institution Name?")
        txtInsname.SetFocus
        Exit Sub
    End If
    Unload Me
End Sub
Private Sub Form_Load()
    With rsIns
        If .State = 1 Then .Close
        temSql = "SELECT tblInstitutionDetail.* FROM tblInstitutionDetail"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 1 Then
            TemResponce = MsgBox("Someone had altered the database outside the program. Please contact Lakmedipro for assistant" & vbNewLine & Me.Caption & vbNewLine & Err.Description, vbCritical, "Altered Database")
            Exit Sub
        End If
        .MoveFirst
        txtInsname.Text = DecreptedWord(!InstitutionName)
        txtDiscription.Text = DecreptedWord(!InstitutionDescription)
        txtRegistration.Text = DecreptedWord(!InstitutionRegistation)
        txtAddress01.Text = DecreptedWord(!institutionAddress)
        txtTelephone01.Text = DecreptedWord(!institutiontelephone1)
        txtTelephone02.Text = DecreptedWord(!InstitutionTelephone2)
        txtFax.Text = DecreptedWord(!InstitutionFax)
        txtEmail01.Text = DecreptedWord(!InstitutionEmail)
        txtEmail02.Text = DecreptedWord(!InstitutionEmail2)
        txtwbsite01.Text = DecreptedWord(!InstitutionWebSite1)
        txtWebsite02.Text = DecreptedWord(!InstitutionWebSite2)
        .Close
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If txtInsname.Text = "" Then
        TemResponce = MsgBox("You have not entered an Institution Name", vbCritical, "? Institution Name")
        txtInsname.SetFocus
        Exit Sub
    End If
    With rsIns
        If .State = 1 Then .Close
        temSql = "SELECT tblInstitutionDetail.* FROM tblInstitutionDetail"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount < 1 Then
            TemResponce = MsgBox("Someone had altered the database outside the program. Please contact Lakmedipro for assistant" & vbNewLine & Me.Caption & vbNewLine & Err.Description, vbCritical, "Altered Database")
            Exit Sub
        End If
        HospitalName = txtInsname.Text
        !InstitutionName = EncreptedWord(HospitalName)
        HospitalDescreption = txtDiscription.Text
        !InstitutionDescription = EncreptedWord(HospitalDescreption)
        !InstitutionRegistation = EncreptedWord(txtRegistration.Text)
        HospitalAddress = txtAddress01.Text
        !institutionAddress = EncreptedWord(HospitalAddress)
        !institutiontelephone1 = EncreptedWord(txtTelephone01.Text)
        !InstitutionTelephone2 = EncreptedWord(txtTelephone02.Text)
        !InstitutionFax = EncreptedWord(txtFax.Text)
        !InstitutionEmail = EncreptedWord(txtEmail01.Text)
        !InstitutionEmail2 = EncreptedWord(txtEmail02.Text)
        !InstitutionWebSite1 = EncreptedWord(txtwbsite01.Text)
        !InstitutionWebSite2 = EncreptedWord(txtWebsite02.Text)
        .Update
        .Close
    End With

End Sub

Private Sub bttnCancel_Click()
    Unload Me
End Sub

