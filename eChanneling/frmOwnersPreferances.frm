VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmOwnersPreferances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Owner Preferances"
   ClientHeight    =   2205
   ClientLeft      =   4440
   ClientTop       =   1680
   ClientWidth     =   4665
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
   ScaleHeight     =   2205
   ScaleWidth      =   4665
   Begin btButtonEx.ButtonEx btnDeleteData 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Delete Data"
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   1560
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
   Begin MSComctlLib.Slider SliderIncomeDeflation 
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   767
      _Version        =   393216
      Min             =   1
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Income Deflation"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmOwnersPreferances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Setcolours()
    Me.BackColor = FrameBackColour
    Me.ForeColor = FrameForeColour
    bttnClose.BackColor = BttnBackColour
    bttnClose.ForeColor = BttnForeColour
End Sub

Private Sub btnDeleteData_Click()
frmDeleteAllData.Show
frmDeleteAllData.ZOrder 0
End Sub

Private Sub Form_Load()
    Call SetPreferances
    If UserAuthority <> AuthorityOwner Then
        SliderIncomeDeflation.Visible = False
        Label1.Caption = "Hi"
    End If
    
End Sub
Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub SetPreferances()
    SliderIncomeDeflation.Value = IncomeDeflation
End Sub

Private Sub SavePreferancesToFile()
    SaveSetting App.EXEName, "Options", "IncomeDeflation", SliderIncomeDeflation.Value
End Sub

Private Sub SavePreferancesToMemory()
    IncomeDeflation = SliderIncomeDeflation.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SavePreferancesToFile
    Call SavePreferancesToMemory
End Sub
