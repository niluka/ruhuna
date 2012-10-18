VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGetDetailsOfSx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get details of Surgeries"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
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
   ScaleHeight     =   7005
   ScaleWidth      =   8895
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   77922307
      CurrentDate     =   40178
   End
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Process"
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
   Begin MSFlexGridLib.MSFlexGrid gridDetails 
      Height          =   4335
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7646
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   6360
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
   Begin MSDataListLib.DataCombo cmbSpeciality 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   195
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbConsultant 
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Top             =   765
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSx 
      Height          =   360
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   77922307
      CurrentDate     =   40178
   End
   Begin VB.Label Label4 
      Caption         =   "&Surgery"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Consultant"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "S&peciality"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmGetDetailsOfSx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsStaff As New ADODB.Recordset
    Dim temSql As String
    
Private Sub btnProcess_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbConsultant_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbSx.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbConsultant.Text = Empty
    End If

End Sub

Private Sub cmbSpeciality_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbConsultant.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbSpeciality.Text = Empty
    End If

End Sub

Private Sub cmbSx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnProcess_Click
    ElseIf KeyCode = vbKeyEscape Then
        cmbSx.Text = Empty
    End If

End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call FillCombos
    Call GetSettings
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpFrom.Value = DateSerial(Year(Date), 1, 1)
    dtpTo.Value = Date
    GetCommonSettings Me
End Sub

Private Sub FormatGrid()
    With gridDetails
        .Clear
        
        .Cols = 4
        .Rows = 1
        
        .Col = 0
        .Text = "ID"
        
        .Col = 1
        .Text = "BHT"
        
        .Col = 2
        .Text = "Patient"
        
        .Col = 3
        .Text = "Net Price"
    End With
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        
        temSql = "SELECT     dbo.tblBHT.BHTID, dbo.tblBHT.BHT, dbo.tblPatientMainDetails.FirstName, dbo.tblBHT.NetPrice, dbo.tblBHT.Discharge, dbo.tblBHTSx.SxID, dbo.tblBHTSx.DoctorID " & _
                    "FROM         dbo.tblBHT LEFT OUTER JOIN " & _
                      "dbo.tblPatientMainDetails ON dbo.tblBHT.PatientID = dbo.tblPatientMainDetails.PatientID RIGHT OUTER JOIN " & _
                      "dbo.tblBHTSx ON dbo.tblBHT.BHTID = dbo.tblBHTSx.BHTID " & _
                        "Where dbo.tblBHT.BHTID <> 1 "
        
        If IsNumeric(cmbConsultant.BoundText) = True Then
            temSql = temSql & " And dbo.tblBHTSx.DoctorID = " & Val(cmbConsultant.BoundText) & " "
        End If
        
        If IsNumeric(cmbSx.BoundText) = True Then
            temSql = temSql & " And dbo.tblBHTSx.SxID = " & Val(cmbSx.BoundText) & " "
        End If
        
        temSql = temSql & "AND dbo.tblBHT.DOD BETWEEN '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "'  AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridDetails.Rows = gridDetails.Rows + 1
            gridDetails.Row = gridDetails.Rows - 1
            
            gridDetails.Col = 0
            gridDetails.Text = !BHTID
            
            gridDetails.Col = 1
            gridDetails.Text = !BHT
        
            gridDetails.Col = 2
            gridDetails.Text = !FirstName
        
            gridDetails.Col = 3
            gridDetails.Text = Format(!NetPrice, "#,##0.00")
        
            .MoveNext
        Wend
        .Close
    End With
    
    gridDetails.ColWidth(0) = 0
    
End Sub

Private Sub FillCombos()
    Dim Speciality As New clsFillCombos
    Speciality.FillAnyCombo cmbSpeciality, "Speciality", False
    Dim Sx As New clsFillCombos
    Sx.FillAnyCombo cmbSx, "Sx", True
End Sub

Private Sub cmbSpeciality_Change()
    With rsStaff
        If .State = 1 Then .Close
        If IsNumeric(cmbSpeciality.BoundText) = True Then
            temSql = "SELECT tblStaff.Name AS NameWithTitle, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID where SpecialityID = " & Val(cmbSpeciality.BoundText) & " ORDER BY Name"
        Else
            temSql = "SELECT tblStaff.Name AS NameWithTitle, tblStaff.StaffID FROM tblStaff LEFT JOIN tblTitle ON tblStaff.TitleID = tblTitle.TitleID Order BY Name"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbConsultant
        Set .RowSource = rsStaff
        .ListField = "NameWithTitle"
        .BoundColumn = "StaffID"
        .Text = Empty
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub

Private Sub gridDetails_DblClick()
    Dim temBHTID As Long
    temBHTID = gridDetails.TextMatrix(gridDetails.Row, 0)
    frmBHTSummeryF.Show
    frmBHTSummeryF.cmbBHT.BoundText = temBHTID
    frmBHTSummeryF.ZOrder 0
    frmBHTSummeryF.btnProcess_Click
End Sub
