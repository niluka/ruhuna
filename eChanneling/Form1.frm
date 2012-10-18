VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAgentBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agent Balance"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   10095
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   7560
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
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Excel"
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
   Begin MSFlexGridLib.MSFlexGrid gridAgent 
      Height          =   7095
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12515
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAgentBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()

    GridToExcel gridAgent, "Agent Balance", Format(Date, "dd MMMM yyyy") & " - " & Format(Time, "hh:mm AMPM")

End Sub

Private Sub fillGrid()
    With gridAgent
        .Clear
        .Rows = 1
        .Cols = 4
        
        .Row = 0
        
        .col = 0
        .Text = "Code"
        
        .col = 1
        .Text = "Agent"
        
        .col = 2
        .Text = "Balance"
        
        .col = 3
        .Text = "Limit"
    
    End With
    
    Dim temSQL As String
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT     InstitutionCode, InstitutionName, InstitutionCredit, InstitutionMaxCredit " & _
                    "FROM         dbo.tblInstitutions " & _
                    "ORDER BY InstitutionName "
        .Open temSQL, cnnChannelling, adOpenStatic, adLockReadOnly
        While .EOF = False
        
            gridAgent.Rows = gridAgent.Rows + 1
            gridAgent.Row = gridAgent.Rows - 1
            
            gridAgent.col = 0
            gridAgent.Text = !InstitutionCode
            
            gridAgent.col = 1
            gridAgent.Text = !InstitutionName
            
            gridAgent.col = 2
            gridAgent.Text = Format(!InstitutionCredit, "0.00")
            
            gridAgent.col = 3
            gridAgent.Text = Format(!InstitutionMaxCredit, "0.00")
            
            .MoveNext
        Wend
        .Close
    
    End With
    
End Sub

Private Sub Form_Load()
    GetCommonSettings Me
    fillGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub
