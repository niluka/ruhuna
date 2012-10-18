VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmLRMain 
   Caption         =   "New Investigation Requests"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin btButtonEx.ButtonEx ButtonEx1 
      Height          =   495
      Left            =   7200
      TabIndex        =   5
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "ButtonEx1"
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
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   240
      TabIndex        =   4
      Top             =   5040
      Width           =   6615
   End
   Begin VB.ListBox ListInvestigationID 
      Height          =   4350
      Left            =   6480
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.ListBox ListInvestigation 
      Height          =   4350
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.ListBox ListIxCatogeryID 
      Height          =   4350
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.ListBox ListIxCatogery 
      Height          =   4350
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "frmLRMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTem As New ADODB.Recordset

Private Sub ButtonEx1_Click()
Call TemTite
End Sub

Private Sub Form_Load()
    Call FillIxCatogery
End Sub

Private Sub FillIxCatogery()
    ListIxCatogery.Clear
    ListIxCatogeryID.Clear
    If NoAllNames = False Then
        ListIxCatogery.AddItem "All"
        ListIxCatogeryID.AddItem "All"
    End If
    Dim rsIxCatogery As New ADODB.Recordset
    With rsIxCatogery
        .Open "Select * from tblIxCatogery order by ixcatogery", dbHospital, adOpenForwardOnly, adLockReadOnly
        While .EOF = False
            ListIxCatogery.AddItem !ixcatogery
            ListIxCatogeryID.AddItem !ixcatogery_ID
            .MoveNext
        Wend
    End With
End Sub

Private Sub TemTite()
With rsTem
If .State = 1 Then .Close
.Open " SELECT tblDoctor.*, tblTitle.Title FROM tblTitle INNER JOIN tblDoctor ON tblTitle.Title_ID = tblDoctor.DoctorTitle_ID", dbHospital, 3, 3

MsgBox rsTem.Fields("Title") & " . " & rsTem.Fields("DoctorName")
End With
End Sub
