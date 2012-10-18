VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRoomPatients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Patients"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13680
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
   ScaleHeight     =   7110
   ScaleWidth      =   13680
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   12240
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Close"
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
   Begin MSFlexGridLib.MSFlexGrid gridRoom 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   9551
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   10920
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
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
      Left            =   9600
      TabIndex        =   5
      Top             =   6480
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
   Begin btButtonEx.ButtonEx btnRefresh 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "&Refresh"
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
   Begin VB.Label lblSubTopic 
      Alignment       =   2  'Center
      Caption         =   "SUbtopic"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   13455
   End
   Begin VB.Label lblTopic 
      Alignment       =   2  'Center
      Caption         =   "Topic"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13455
   End
End
Attribute VB_Name = "frmRoomPatients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub SaveSettings()
    SaveCommonSettings Me
    
    
End Sub

Private Sub GetSettings(): On Error Resume Next
    GetCommonSettings Me
    lblTopic.Caption = HospitalName
    lblSubtopic.Caption = "Room Occupancy"
    
    lblTopic.Left = 0
    lblTopic.Width = Me.Width
    
    lblSubtopic.Left = 0
    lblSubtopic.Width = Me.Width
    
'    btnClose.Left = Me.Width - btnClose.Width - 100
'    btnPrint.Left = Me.Width - btnClose.Width - btnPrint.Width - 200
'    btnExcel.Left = Me.Width - btnClose.Width - btnPrint.Width - btnExcel.Width - 300
'
'    btnClose.Left = Me.Height - btnClose.Height - 100
'    btnPrint.Left = btnClose.Top
'    btnExcel.Left = btnClose.Top

End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridRoom, "Room Occupancy"
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    GridPrint gridRoom, ThisReportFormat, "Room Occpancy of Patients", "On " & Format(Date, "dd MMMM yyyy")
    Printer.EndDoc
End Sub

Private Sub btnRefresh_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call GetSettings
    Call FillGrid
End Sub

Private Sub FormatGrid()
    With gridRoom
        .Clear
        
        .Cols = 6
        
        .Rows = 1
        
        .Col = 0
        .Text = "Room"
        
        .Col = 1
        .Text = "BHT"
        
        .Col = 2
        .Text = "Patient"
        
        .Col = 3
        .Text = "Guardian"
        
        .Col = 4
        .Text = "Address"
        
        .Col = 5
        .Text = "Phone"
        
    End With
End Sub

Private Sub FillGrid()
    Dim Col As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT     TOP 100 PERCENT dbo.tblRoom.Room, dbo.tblBHT.BHT, dbo.tblPatientMainDetails.FirstName, dbo.tblBHT.GuardianName, " & _
                        "dbo.tblBHT.GuardianAddress , dbo.tblBHT.GuardianPhone " & _
                    "FROM         dbo.tblPatientMainDetails RIGHT OUTER JOIN " & _
                        "dbo.tblBHT ON dbo.tblPatientMainDetails.PatientID = dbo.tblBHT.PatientID RIGHT OUTER JOIN " & _
                        "dbo.tblRoomPatient ON dbo.tblBHT.BHTID = dbo.tblRoomPatient.BHTID LEFT OUTER JOIN " & _
                        "dbo.tblRoom ON dbo.tblRoomPatient.RoomID = dbo.tblRoom.RoomID " & _
                    "Where (dbo.tblBHT.Discharge = 0) And (dbo.tblRoomPatient.ToTime Is Null) " & _
                    "ORDER BY dbo.tblRoom.Room "
        
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridRoom.Rows = gridRoom.Rows + 1
            gridRoom.Row = gridRoom.Rows - 1
            
            For Col = 0 To gridRoom.Cols - 1
                gridRoom.Col = Col
                gridRoom.Text = Format(.Fields(Col).Value, "")
            Next
            
            .MoveNext
        Wend
        .Close
    
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub
