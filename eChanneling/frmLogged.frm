VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmLogged 
   Caption         =   "Logged Users"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
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
   ScaleHeight     =   3360
   ScaleWidth      =   6570
   Begin VB.ListBox List2 
      Height          =   2460
      Left            =   3000
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin btButtonEx.ButtonEx bttnMark 
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Mark as Logged Off"
      Enabled         =   0   'False
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
   Begin VB.ListBox List1 
      Height          =   2460
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin btButtonEx.ButtonEx bttnMarkAll 
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Appearance      =   3
      Enabled         =   0   'False
      Caption         =   "Mark All as Logged Off"
      Enabled         =   0   'False
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
   Begin btButtonEx.ButtonEx bttnExit 
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Exit"
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
   Begin VB.Label Label1 
      Caption         =   "Currently Logged Users"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmLogged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bttnExit_Click()
Unload Me
End Sub

Private Sub bttnMark_Click()
If List1.ListCount < 1 Then Exit Sub
If List2.ListCount < 1 Then Exit Sub
If IsNumeric(List2.Text) = False Then Exit Sub
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "select * from tblstaff where staff_ID = " & List2.Text
    .Open
    If .RecordCount <> 0 Then
        !logged = False
        .Update
    End If
    If .State = 1 Then .Close
End With
Call FillAllLogged
End Sub

Private Sub bttnMarkAll_Click()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "select * from tblstaff"
    .Open
    If .RecordCount <> 0 Then
        !logged = False
        .Update
    End If
    If .State = 1 Then .Close
End With
Call FillAllLogged
End Sub

Private Sub Form_Load()
Call FillAllLogged
End Sub

Private Sub FillAllLogged()
    With DataEnvironment1.rssqlTem
        List1.Clear
        List2.Clear
        If .State = 1 Then .Close
        .Source = "select * from tblstaff where logged = 1 order by StaffName "
        .Open
        If .RecordCount <> 0 Then
            While .EOF = False
                bttnMarkAll.Enabled = True
                List1.AddItem !StaffName
                List2.AddItem !Staff_ID
                .MoveNext
            Wend
        End If
        If .State = 1 Then .Close
    End With
End Sub

Private Sub List1_Click()
    
    If List1.ListCount < 1 Then Exit Sub
    List2.ListIndex = List1.ListIndex
    If List2.ListCount < 1 Then Exit Sub
    If IsNumeric(List2.Text) = False Then Exit Sub
    bttnMark.Enabled = True
    bttnMarkAll.Enabled = True
End Sub
