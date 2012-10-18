VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmHospitalPreferance 
   Caption         =   "Hospital Preferance"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
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
   ScaleHeight     =   6150
   ScaleWidth      =   9030
   Begin VB.CheckBox chkDoNotAllowExpireSale 
      Caption         =   "Do NOT allow Sale of expired items"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin VB.CheckBox chkDoNotAllowExpireConsumption 
      Caption         =   "Do NOT allow consumption of expired items"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   7680
      TabIndex        =   0
      Top             =   5520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
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
Attribute VB_Name = "frmHospitalPreferance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public DoNotAllowExpireSale As Boolean
'Public DoNotAllowExpireConsumption As Boolean
'Public DoNotAllowExpireTransfer As Boolean

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub SaveToMemory()
    If chkDoNotAllowExpireConsumption.Value = 1 Then
        DoNotAllowExpireConsumption = True
    Else
        DoNotAllowExpireConsumption = False
    End If
    If chkDoNotAllowExpireSale.Value = 1 Then
        DoNotAllowExpireSale = True
    Else
        DoNotAllowExpireSale = False
    End If
    
End Sub

Private Sub SaveToReg()
    If chkDoNotAllowExpireConsumption.Value = 1 Then
        SaveSetting App.EXEName, "Options", "DoNotAllowExpireConsumption", True
    Else
        SaveSetting App.EXEName, "Options", "DoNotAllowExpireConsumption", False
    End If
    If chkDoNotAllowExpireSale.Value = 1 Then
        SaveSetting App.EXEName, "Options", "DoNotAllowExpireSale", True
    Else
        SaveSetting App.EXEName, "Options", "DoNotAllowExpireSale", False
    End If
End Sub

Private Sub GetFromMemory()
    If DoNotAllowExpireConsumption = True Then
        chkDoNotAllowExpireConsumption.Value = 1
    Else
        chkDoNotAllowExpireConsumption.Value = 0
    End If
    If DoNotAllowExpireSale = True Then
        chkDoNotAllowExpireSale.Value = 1
    Else
        chkDoNotAllowExpireSale.Value = 0
    End If
End Sub


Private Sub Form_Load()
    Call GetFromMemory
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveToMemory
    Call SaveToReg
End Sub
