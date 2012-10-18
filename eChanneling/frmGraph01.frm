VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmGraph01 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "frmGraph01.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   10425
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   255
      Left            =   8640
      TabIndex        =   0
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
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
   Begin VB.OLE OLE1 
      Height          =   7455
      Left            =   120
      SizeMode        =   2  'AutoSize
      TabIndex        =   1
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmGraph01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bttnClose_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    OLE1.OLETypeAllowed = 0
    OLE1.SizeMode = 1
    OLE1.CreateLink App.Path & "\graph01.xls"
    OLE1.Refresh
End Sub

Private Sub Form_Resize()
    OLE1.Top = 50
    OLE1.Left = 50
    OLE1.Width = Me.Width - 100
    OLE1.Height = Me.Height - 100
    bttnClose.Top = Me.Height - (bttnClose.Height * 4)
    bttnClose.Left = Me.Width - (bttnClose.Width * 2)
End Sub
