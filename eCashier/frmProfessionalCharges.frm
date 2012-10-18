VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmProfessionalCharges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Professional Charges"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8355
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
   ScaleHeight     =   8070
   ScaleWidth      =   8355
   Begin VB.TextBox txtTotal 
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   5880
      Width           =   1815
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Top             =   7440
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   6360
      TabIndex        =   13
      Top             =   7440
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Delete"
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
   Begin MSFlexGridLib.MSFlexGrid gridCharge 
      Height          =   3975
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7011
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.TextBox txtValue 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox txtComments 
      Height          =   1095
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   4575
   End
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
   Begin MSDataListLib.DataCombo cmbSpeciality 
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff 
      Height          =   360
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   77529090
      CurrentDate     =   39960
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   77529091
      CurrentDate     =   39960
   End
   Begin VB.Label Label7 
      Caption         =   "Total"
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Da&te"
      Height          =   255
      Left            =   6360
      TabIndex        =   18
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Tim&e"
      Height          =   255
      Left            =   6360
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "&Value"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "&Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Pro&fession"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "&BHT"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmProfessionalCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsStaff As New ADODB.Recordset

Private Sub btnAdd_Click()
    If IsNumeric(cmbBHT.BoundText) = False Then
        MsgBox "BHT?"
        cmbBHT.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbStaff.BoundText) = False Then
        MsgBox "Staff?"
        cmbStaff.SetFocus
        Exit Sub
    End If
    If Val(txtValue.Text) = 0 Then
        MsgBox "Value?"
        txtValue.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblProfessionalCharges "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !ProfessionalCharge = True
        !ForBHTID = Val(cmbBHT.BoundText)
        !Fee = Val(txtValue.Text)
        !Comments = txtComments.Text
        !StaffID = Val(cmbStaff.BoundText)
        !Date = dtpDate.Value
        !Time = dtpTime.Value
        .Update
        .Close
    End With
    Call ClearValues
    Call FormatGrid
    Call FillGrid
    cmbSpeciality.SetFocus
End Sub

Private Sub FillGrid()
    Dim TotalValue As Double
    
    gridCharge.Visible = False
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblSpeciality.Speciality, tblStaff.Name, tblProfessionalCharges.* " & _
                    "FROM (tblSpeciality RIGHT JOIN tblStaff ON tblSpeciality.SpecialityID = tblStaff.SpecialityID) RIGHT JOIN tblProfessionalCharges ON tblStaff.StaffID = tblProfessionalCharges.StaffID " & _
                    "WHERE (((tblProfessionalCharges.ProfessionalCharge)=True) AND ((tblProfessionalCharges.Cancelled)=False) AND ((tblProfessionalCharges.ForBHTID)=" & Val(cmbBHT.BoundText) & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        While .EOF = False
            gridCharge.Rows = gridCharge.Rows + 1
            gridCharge.Row = gridCharge.Rows - 1
            
            gridCharge.Col = 0
            gridCharge.Text = !ProfessionalChargesID
            
            gridCharge.Col = 1
            gridCharge.Text = Format(!Date, "dd MMM yyyy")
            
            gridCharge.Col = 2
            gridCharge.Text = Format(!Time, "HH MM")
            
            gridCharge.Col = 3
            gridCharge.Text = !Speciality
            
            gridCharge.Col = 4
            gridCharge.Text = !Name
            
            gridCharge.Col = 5
            gridCharge.Text = Format(!Comments, "")
            
            gridCharge.Col = 6
            gridCharge.Text = Format(!Fee, "0.00")
            
            TotalValue = TotalValue + !Fee
            
            .MoveNext
        
        Wend
        .Close
    End With
    txtTotal.Text = Format(TotalValue, "0.00")
    
    gridCharge.Visible = True
    
End Sub

Private Sub FormatGrid()
    With gridCharge
        .Rows = 1
        .Cols = 7
        
        .Col = 1
        .Text = "Date"
        
        .Col = 2
        .Text = "Time"
        
        .Col = 3
        .Text = "Speciality"
        
        .Col = 4
        .Text = "Name"
        
        .Col = 5
        .Text = "Comments"
        
        .Col = 6
        .Text = "Value"
        
        .ColWidth(0) = 0
    End With
End Sub


Private Sub ClearValues()
    cmbSpeciality.Text = Empty
    cmbStaff.Text = Empty
    txtComments.Text = Empty
    txtValue.Text = Empty
    
End Sub


Private Sub btnPrint_Click()
    
    Call GridPrint(gridCharge)
    
    Exit Sub
    
    Dim ProfessionalCharges As PrintReport
    Dim ProfessionalChargesCols(5) As PrintColumn
    Dim ColDate As PrintColumn
    Dim ColTime As PrintColumn
    Dim ColSpeciality As PrintColumn
    Dim ColDoc As PrintColumn
    Dim ColFee As PrintColumn
    
    Dim StrDate() As String
    Dim StrTime() As String
    Dim strSpeciality() As String
    Dim strDoc() As String
    Dim strFee() As String
    
    Dim i As Integer
    With gridCharge
    
        ReDim StrDate(.Rows - 1)
        ReDim StrTime(.Rows - 1)
        ReDim strSpeciality(.Rows - 1)
        ReDim strDoc(.Rows - 1)
        ReDim strFee(.Rows - 1)
        
        
        For i = 0 To .Rows - 2
            StrDate(i) = .TextMatrix(i + 1, 1)
            StrTime(i) = .TextMatrix(i + 1, 2)
            strSpeciality(i) = .TextMatrix(i + 1, 3)
            strDoc(i) = .TextMatrix(i + 1, 4)
            strFee(i) = .TextMatrix(i + 1, 6)
        Next i
    End With
    
    ColDate.ColText() = StrDate()
    ColTime.ColText() = StrTime()
    ColSpeciality.ColText() = strSpeciality()
    ColDoc.ColText() = strDoc()
    ColFee.ColText() = strFee()
    
    ColDate.Topic = "Date"
    ColTime.Topic = "Time"
    ColSpeciality.Topic = "Speciality"
    ColDoc.Topic = "Doctor"
    ColFee.Topic = "Fee"
    
    ColDate.TextAlignmant = CentreAlign
    ColTime.TextAlignmant = CentreAlign
    ColSpeciality.TextAlignmant = LeftAlign
    ColDoc.TextAlignmant = LeftAlign
    ColFee.TextAlignmant = RightAlign
    
    GetPrintDefaults ProfessionalCharges
    ProfessionalChargesCols(0) = ColDate
    ProfessionalChargesCols(1) = ColTime
    ProfessionalChargesCols(2) = ColSpeciality
    ProfessionalChargesCols(3) = ColDoc
    ProfessionalChargesCols(4) = ColFee
    
    ProfessionalCharges.Topic = HospitalName
    ProfessionalCharges.Subtopic = "Professional Charges"
    
    ProfessionalCharges.PrintColums() = ProfessionalChargesCols()
    
    Call PrintMyReport(ProfessionalCharges, True)

End Sub

Private Sub cmbBHT_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbSpeciality_Change()
    With rsStaff
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where SpecialityID = " & Val(cmbSpeciality.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbStaff
        Set .RowSource = rsStaff
        .ListField = "Name"
        .BoundColumn = "StaffID"
        .Text = Empty
    End With
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call GetSettings
    Call FormatGrid
End Sub

Private Sub FillCombos()
    Dim BHT As New clsFillCombos
    BHT.FillAnyCombo cmbBHT, "BHT"
    Dim Speciality As New clsFillCombos
    Speciality.FillAnyCombo cmbSpeciality, "Speciality", False
End Sub

Private Sub GetSettings()
    dtpDate.Value = Date
    dtpTime.Value = Time
End Sub
