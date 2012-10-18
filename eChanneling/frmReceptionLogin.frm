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
   Icon            =   "frmReceptionLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
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

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Dim TemResponce  As Integer
    Dim UserNameFound As Boolean
    UserNameFound = False
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
        .Source = "Select tblstaff.* from tblstaff where (StaffUser = true)"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            txtTemUsername.Text = DecreptedWord(!StaffUserName)
'            MsgBox "Username" & vbNewLine & txtTemUserName.Text
'            MsgBox "Password" & vbNewLine & DecreptedWord(!staffpassword)
            If UCase(txtUserName.Text) = UCase(txtTemUsername.Text) Then
                UserNameFound = True
                If txtPassword.Text = DecreptedWord(!staffpassword) Then
                    UserName = UCase(txtUserName.Text)
                    UserID = !staff_ID
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
End Sub

Private Sub Form_Load()
    Dim TemResponce  As Integer
    Dim WillExpire As Boolean
    Dim ExpiaryDate As Date
    
    WillExpire = False
    ExpiaryDate = #2/28/2008#

If WillExpire = True And ExpiaryDate < Date Then
    TemResponce = MsgBox("The Program has expiared. Please contact Lakmedipro for Assistant", vbCritical, "Expired")
    End
ElseIf WillExpire = True And ExpiaryDate > Date Then
    TemResponce = MsgBox("The is a trial program and it will expire in after unknown number of days, maximum 60 days ", vbCritical, "Demo Version")
End If

  DataEnvironment1.cnnHospital.ConnectionString = "data source=" & App.Path & "\hospital.mdb;"

 

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtUserName.Text <> "" Then cmdOK_Click: Exit Sub
    If KeyAscii = 13 And txtUserName.Text = "" Then txtUserName.SetFocus: Exit Sub
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPassword.SetFocus
End Sub
