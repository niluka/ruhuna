VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBHTCharge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Additional Charges"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
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
   ScaleHeight     =   5895
   ScaleWidth      =   10110
   Begin VB.TextBox txtDetails 
      Height          =   1815
      Left            =   6720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   2880
      Width           =   3255
   End
   Begin VB.TextBox txtEditID 
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtDelID 
      Height          =   360
      Left            =   9240
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   67174403
      CurrentDate     =   39956
   End
   Begin VB.TextBox txtComments 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   4575
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
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
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSFlexGridLib.MSFlexGrid gridService 
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7435
      _Version        =   393216
   End
   Begin VB.TextBox txtCharge 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   4575
   End
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   360
      Left            =   2040
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   375
      Left            =   7440
      TabIndex        =   17
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   67174402
      CurrentDate     =   39956
   End
   Begin VB.Label Label2 
      Caption         =   "Details"
      Height          =   255
      Left            =   6720
      TabIndex        =   20
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Time"
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Date"
      Height          =   255
      Left            =   6720
      TabIndex        =   16
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Total"
      Height          =   255
      Left            =   6720
      TabIndex        =   12
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Hospital Charge"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "BHT"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBHTCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsSC As New ADODB.Recordset
    Dim temSql As String
    Dim rsSPC As New ADODB.Recordset
    Dim rsStaff() As New ADODB.Recordset
    Dim PSCCount As Long
    Dim rsBHT As New ADODB.Recordset
    Dim MyBHT As New clsBHT

Private Sub btnAdd_Click()
    Dim n As Integer
    If IsNumeric(cmbBHT.BoundText) = False Then
        MsgBox "BHT?"
        cmbBHT.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtEditID.Text) = True Then
            temSql = "Select * from tblPatientCharge where PatientChargeID = " & Val(txtEditID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount <= 0 Then
                .AddNew
            End If
        Else
            temSql = "Select * from tblPatientCharge  where PatientChargeID = 0 "
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
        End If
        !BHTID = Val(cmbBHT.BoundText)
        !Comments = txtComments.Text
        !ServiceDate = dtpDate.Value
        !ServiceTime = dtpTime.Value
        !Charge = Val(txtCharge.Text)
        !HospitalCharge = Val(txtCharge.Text)
        !UserID = UserID
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        txtEditID.Text = !NewID
        .Close
    End With
    Call FillGrid
    Call ClearAddValues
    cmbBHT.SetFocus
End Sub

Private Sub ClearAddValues()
    Dim n As Long
    txtComments.Text = Empty
    txtCharge.Text = Empty
    txtEditID.Text = Empty
    txtDelID.Text = Empty
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    If Val(txtDelID.Text) = 0 Then
        MsgBox "Please select one to delete"
        Exit Sub
    End If
    
    
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtDelID.Text) = True Then
            temSql = "Select * from tblPatientCharge where PatientChargeID = " & Val(txtDelID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Deleted = True
                !DeletedUserID = UserID
                !DeletedDate = Date
                !DeletedTime = Now
                .Update
            End If
            .Close
        End If
    End With
    Call FillGrid
    Call ClearAddValues
    cmbBHT.SetFocus
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    
    GridPrint gridService, ThisReportFormat
    Printer.EndDoc
    
End Sub

Private Sub cmbBHT_Click(Area As Integer)
    If IsNumeric(cmbBHT.BoundText) = False Then Exit Sub
    Call FillGrid
    MyBHT.BHTID = Val(cmbBHT.BoundText)
    Call DisplayDetails
    
End Sub

Private Sub DisplayDetails(): On Error Resume Next
    Dim temText As String
    Dim r As Long
    temText = "Patient Name : " & MyBHT.FirstName & vbNewLine
    temText = temText & "Guardian : " & MyBHT.GuardianName & vbNewLine
    temText = temText & "Address : " & MyBHT.PtAddress & vbNewLine
    temText = temText & "BHT : " & MyBHT.BHT & vbNewLine
    temText = temText & "Age : " & MyBHT.AgeInWords & vbNewLine
    temText = temText & "Admitted : " & Format(MyBHT.DOA, "dd MMMM yyyy") & " at " & Format(MyBHT.TOA, "HH:MM AMPM") & vbNewLine
    If MyBHT.Discharge = True Then
        temText = temText & "Discharged :" & Format(MyBHT.DOD, "dd MMMM yyyy") & " at " & Format(MyBHT.TOD, "HH:MM AMPM") & vbNewLine
    Else
        temText = temText & "Not yet discharged" & vbNewLine
    End If
    temText = temText & "Payment Method : " & MyBHT.PaymentMethod
    If MyBHT.HealthSchemeSupplier <> "" Then
        temText = temText & " (" & MyBHT.HealthSchemeSupplier & ")" & vbNewLine
    Else
        temText = temText & vbNewLine
    End If
    If MyBHT.Comments <> "" Then
        temText = temText & MyBHT.Comments & vbNewLine
    End If
    
    txtDetails.Text = temText
End Sub


Private Sub ClearServiceValues()
    Dim n As Integer
    txtComments.Text = Empty
    txtCharge.Text = Empty
    txtEditID.Text = Empty
    txtDelID.Text = Empty
End Sub


Private Sub Form_Load()
    Call GetSettings
    Call FillCombos
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    dtpTime.Value = Time
End Sub

Private Sub FillCombos()
    With rsBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where IsBHT = 1 And Discharge = 0 order by BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
End Sub

Private Sub FillGrid()
    Call FormatGrid
    Dim rsTem As New ADODB.Recordset
    Dim TotalCharge As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientCharge.* FROM tblPatientCharge WHERE (((tblPatientCharge.Deleted)=0) AND ((tblPatientCharge.BHTID)=" & Val(cmbBHT.BoundText) & ")) " & _
                    "ORDER BY tblPatientCharge.PatientChargeID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridService.Rows = gridService.Rows + 1
            gridService.Row = gridService.Rows - 1
            gridService.Col = 0
            gridService.Text = !PatientChargeID
            gridService.Col = 1
            gridService.Text = !ServiceDate
            gridService.Col = 2
            gridService.Text = !Comments
            gridService.Col = 3
            gridService.Text = Format(!Charge, "0.00")
            TotalCharge = TotalCharge + !Charge
            .MoveNext
        Wend
    End With
    lblTotal.Caption = Format(TotalCharge, "0.00")
End Sub

Private Sub FormatGrid()
    '   0   ID
    '   1   Date
    '   2   Service
    '   3   Comments
    '   4   Charges
    With gridService
        .Cols = 5
        .Rows = 1
        .ColWidth(0) = 0
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        
        .Col = 1
        .Text = "Date"
        
        .Col = 2
        .Text = "Comments "
        
        .Col = 3
        .Text = "Charge"
    End With
    lblTotal.Caption = "0.00"
End Sub

Private Sub gridService_Click()
    With gridService
        txtDelID.Text = Val(.TextMatrix(.Row, 0))
        .Col = .Cols - 1
        .ColSel = 0
    End With
End Sub

Private Sub gridService_DblClick()
    Dim rsTem As New ADODB.Recordset
    With gridService
        txtEditID.Text = Val(.TextMatrix(.Row, 0))
        .Col = .Cols - 1
        .ColSel = 0
    End With
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtEditID.Text) = True Then
            temSql = "Select * from tblPatientCharge where PatientChargeID = " & Val(txtEditID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
                txtComments.Text = !Comments
                dtpDate.Value = !ServiceDate
                dtpTime.Value = !ServiceTime
                txtCharge.Text = Format(!Charge, "0.00")
            End If
            .Close
        End If
    End With
End Sub

