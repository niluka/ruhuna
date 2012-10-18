VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAddDetailsOfSx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add details of Surgeries"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
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
   ScaleHeight     =   7260
   ScaleWidth      =   9510
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   6720
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
      TabIndex        =   3
      Top             =   680
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
      TabIndex        =   5
      Top             =   1240
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
      TabIndex        =   7
      Top             =   1800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid gridDetails 
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7646
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   8760
      TabIndex        =   11
      Top             =   2280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Del"
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
   Begin VB.Label Label4 
      Caption         =   "&Surgery"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Consultant"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "S&peciality"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&BHT"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddDetailsOfSx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsStaff As New ADODB.Recordset
    Dim temSql As String
    
    
    
Private Sub FormatGrid()
    With gridDetails
        .Clear
        
        .Cols = 3
        .Rows = 1
        
        .Col = 0
        .Text = "Surgery"
        
        .Col = 1
        .Text = "Doctor"
        
        If .ColWidth(0) < 1000 Then .ColWidth(0) = 3600
        If .ColWidth(1) < 1000 Then .ColWidth(1) = 4600
        .ColWidth(2) = 0
        
    End With
End Sub

Private Sub fillGrid()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
       If Val(cmbBHT.BoundText) = 0 Then Exit Sub
        
    temSql = "    SELECT  dbo.tblBHTSx.BHTSxID,   dbo.tblSx.Sx, dbo.tblTitle.Title + ' ' + dbo.tblStaff.Name AS Doctor " & _
                "FROM         dbo.tblSx RIGHT OUTER JOIN " & _
                      "dbo.tblTitle RIGHT OUTER JOIN " & _
                      "dbo.tblStaff RIGHT OUTER JOIN " & _
                      "dbo.tblBHTSx ON dbo.tblStaff.StaffID = dbo.tblBHTSx.DoctorID ON dbo.tblTitle.TitleID = dbo.tblStaff.TitleID ON " & _
                      "dbo.tblSx.SxID = dbo.tblBHTSx.SxID " & _
                        "Where (dbo.tblBHTSx.BHTID = " & Val(cmbBHT.BoundText) & ")"
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
        While .EOF = False
            gridDetails.Rows = gridDetails.Rows + 1
            gridDetails.Row = gridDetails.Rows - 1
            
            gridDetails.Col = 0
            gridDetails.Text = !Sx
            
            gridDetails.Col = 1
            gridDetails.Text = !Doctor
        
            gridDetails.Col = 2
            gridDetails.Text = !BHTSxID
        
        
            .MoveNext
        Wend
        End If
        .Close
    End With
    
    
    
End Sub
    
    
Private Sub btnDelete_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblBHTSX where BHTSxID = " & Val(gridDetails.TextMatrix(gridDetails.Row, 2))
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            .Delete adAffectCurrent
        End If
        If .State = 1 Then .Close
    End With
    Call FormatGrid
    Call fillGrid
    'Call ClearDetails
    cmbBHT.SetFocus
End Sub

Private Sub btnSave_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblBHTSX where BHTID = 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !BHTID = Val(cmbBHT.BoundText)
        !SxID = Val(cmbSx.BoundText)
        !DoctorID = Val(cmbConsultant.BoundText)
        .Update
        .Close
    End With
    Call FormatGrid
    Call fillGrid
'    Call ClearDetails
    
    
    cmbBHT.SetFocus
End Sub

Private Sub ClearDetails()
    cmbBHT.Text = Empty
    cmbSpeciality.Text = Empty
    cmbConsultant.Text = Empty
    cmbSx.Text = Empty
End Sub

Private Sub cmbBHT_Change()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblBHTSX where BHTID = " & Val(cmbBHT.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbSpeciality.Text = Empty
            cmbSx.BoundText = !SxID
            cmbConsultant.BoundText = !DoctorID
        Else
            cmbSpeciality.Text = Empty
            cmbConsultant.Text = Empty
            cmbSx.Text = Empty
        End If
        .Close
    End With
    Call FormatGrid
    Call fillGrid
End Sub

Private Sub cmbBHT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbSpeciality.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbBHT.Text = Empty
    End If
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
        btnSave_Click
    ElseIf KeyCode = vbKeyEscape Then
        cmbSx.Text = Empty
    End If

End Sub

Private Sub Form_Load()
    Call FillCombos
    GetCommonSettings Me
End Sub

Private Sub FillCombos()
    Dim Speciality As New clsFillCombos
    Speciality.FillAnyCombo cmbSpeciality, "Speciality", False
    Dim BHT As New clsFillCombos
    BHT.FillBoolCombo cmbBHT, "BHT", "BHT", "IsBHT", False
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

Private Sub Form_Unload(Cancel As Integer)
    SaveCommonSettings Me
End Sub
