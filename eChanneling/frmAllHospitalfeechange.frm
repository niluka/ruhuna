VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAllHospitalfeechange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Hospital Fee"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAllHospitalfeechange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8685
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Close"
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
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      Begin btButtonEx.ButtonEx bttnChangeAllHospitalfee 
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Change All"
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
      Begin btButtonEx.ButtonEx bttnForeigenHospitslfee 
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Change Foreign "
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
      Begin btButtonEx.ButtonEx bttnAgentHospitalChange 
         Height          =   375
         Left            =   5640
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Change Agent Fee"
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
      Begin btButtonEx.ButtonEx bttnHospialfeeChange 
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   " Change Local  Fee"
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
      Begin VB.TextBox txtForeign 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtAgent 
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtLocal 
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Foreign Hospital Fee"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Agent Hospital Fee"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Local Hospital Fee"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmAllHospitalfeechange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A, B, C

Private Sub bttnChangeAllHospitalfee_Click()
If txtLocal.Text = "" Then A = MsgBox("Enter Local Hospiatal Fee", vbCritical, "Local Hospital Fee Empty"): Exit Sub
If txtAgent.Text = "" Then B = MsgBox("Enter Agent Hospiatal Fee", vbCritical, "Agent Hospital Fee Empty"): Exit Sub
If txtForeign.Text = "" Then C = MsgBox("Enter Foreigner Hospiatal Fee", vbCritical, "Foreigner Hospital Fee Empty"): Exit Sub

With DataEnvironment1.rssqlTem2

    If .State = 1 Then .Close
    .Open "Select * From tblFacilitySecession"
    
    Do While .EOF = False
    !LocalHospitalFee = txtLocal.Text
    !AgentHospitalFee = txtAgent.Text
    !ForeignHospitalFee = txtForeign.Text
    .MoveNext
    Loop
    
    If .State = 1 Then .Close
    A = MsgBox("All Update successfully", vbCritical, "Updated")
    txtLocal.Text = ""
    txtAgent.Text = ""
    txtForeign.Text = ""

End With

End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub FillFee()
With DataEnvironment1.rsSqlTem1111

    If .State = 1 Then .Close
    .Open "Select* From tblFacilitySecession"
    .MoveFirst
    txtLocal.Text = !LocalHospitalFee
    txtAgent.Text = !AgentHospitalFee
    txtForeign.Text = !ForeignHospitalFee
    If .State = 1 Then .Close

End With

End Sub

Private Sub bttnHospialfeeChange_Click()

If txtLocal.Text = "" Then A = MsgBox("Enter Local Hospiatal Fee", vbCritical, "Local Hospital Fee Empty")

With DataEnvironment1.rssqlTem1

    If .State = 1 Then .Close
    .Open "Select* From tblFacilitySecession "
    
    Do While .EOF = False
    !LocalHospitalFee = txtLocal.Text
    .MoveNext
    Loop
    
    A = MsgBox("Local Hospital Fee Update successfully", vbCritical, "Updated")
    If .State = 1 Then .Close
    txtLocal.Text = ""
    
End With

End Sub

Private Sub bttnForeigenHospitslfee_Click()
If txtForeign.Text = "" Then A = MsgBox("Enter Agent Hospiatal Fee", vbCritical, "Agent Hospital Fee Empty"): Exit Sub

With DataEnvironment1.rssqlTem1

    If .State = 1 Then .Close
    .Open "Select* From tblFacilitySecession"
    
    Do While .EOF = False
    !ForeignHospitalFee = txtForeign.Text
    .MoveNext
    Loop
    
    A = MsgBox("Foreigner Hospital Fee Update successfully", vbCritical, "Updated")
    If .State = 1 Then .Close
    txtForeign.Text = ""
    
End With

End Sub

Private Sub bttnAgentHospitalChange_Click()
If txtForeign.Text = "" Then A = MsgBox("Enter Foreign Hospiatal Fee", vbCritical, "Foreign Hospital Fee Empty"): Exit Sub

With DataEnvironment1.rssqlTem1

    If .State = 1 Then .Close
    .Open "Select* From tblFacilitySecession"
    
    Do While .EOF = False
    !AgentHospitalFee = txtAgent.Text
    .MoveNext
    Loop
    
    A = MsgBox("Agent Hospital Fee Update successfully", vbCritical, "Updated")
    If .State = 1 Then .Close
    txtAgent.Text = ""
    
End With



End Sub

Private Sub Form_Load()
FillFee
End Sub
