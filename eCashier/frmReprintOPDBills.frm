VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReprintOPDBills 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprint OPD Bills"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
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
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   6480
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
   Begin MSFlexGridLib.MSFlexGrid gridBill 
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8493
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbUser 
      Height          =   360
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67239939
      CurrentDate     =   39992
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67239939
      CurrentDate     =   39992
   End
   Begin VB.Label Label3 
      Caption         =   "User"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmReprintOPDBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub Form_Load()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    cmbUser.BoundText = UserID
    Call FillCombos
End Sub

Private Sub FillCombos()
    Dim Staff As New clsFillCombos
    Staff.FillSpecificFieldBoolCombo cmbUser, "staff", "Name", "Name", "IsAUser", False
End Sub

Private Sub FillGrid()
    Call FormatGrid
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = ""
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridBill.Rows = gridBill.Rows + 1
            gridBill.Row = gridBill.Rows - 1
            
            gridBill.Col = 0
            gridBill.Text = !BillID
            
            
            .MoveNext
        Wend
        
        .Close
    End With
End Sub

Private Sub FormatGrid()
    With gridBill
        .Clear
        .Rows = 1
        .Cols = 5
        
        .Col = 0
        .Text = "ID"
        
        .Col = 1
        .Text = "Date"
        
        .Col = 2
        .Text = "Time"
        
        .Col = 3
        .Text = "Customer"
        
        .Col = 4
        .Text = "Value"
        
    End With
End Sub
