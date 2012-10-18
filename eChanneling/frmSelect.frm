VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lakmedipro e-Hospital - Complete Computer Solution For Hospital Management"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8730
   ClipControls    =   0   'False
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin btButtonEx.ButtonEx bttnDoctor 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Doctor"
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
   Begin btButtonEx.ButtonEx bttnReception 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Reception"
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
   Begin btButtonEx.ButtonEx bttnLaboratory 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Laboratory"
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
   Begin btButtonEx.ButtonEx bttnStaff 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Staff"
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
   Begin btButtonEx.ButtonEx bttnFinance 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Finance"
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
   Begin btButtonEx.ButtonEx bttnExit 
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "EXIT"
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
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ToNewForm As Boolean


Private Sub bttnExit_Click()
ToNewForm = False
Unload Me
End Sub

Private Sub bttnReception_Click()
Dim A
ToNewForm = True
Unload Me
A = Shell(App.Path & "\eReception.exe")
'ToNewForm = True
'frmReception.Show
'
End Sub

Private Sub Form_Load()
Call SetColours
End Sub

Private Sub SetColours()
Dim PreferanceColour As Byte
PreferanceColour = 1

Select Case PreferanceColour
Case 1:

Case 2:

Case 3:

End Select








End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ToNewForm = True Then Exit Sub
Dim TemResponce As Byte
TemResponce = MsgBox("Are you sure you want to exit Lakmedipro eHospital?", vbCritical + vbYesNo, "Exit?")

If TemResponce = vbYes Then
    Cancel = False
Else
    Cancel = True
End If

End Sub

