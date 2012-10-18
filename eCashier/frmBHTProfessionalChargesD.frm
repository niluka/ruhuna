VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBHTProfessionalChargesD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BHT Professional Charges"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12180
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
   ScaleWidth      =   12180
   Begin VB.TextBox txtDetails 
      Height          =   1815
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox txtTotal 
      Height          =   375
      Left            =   9960
      TabIndex        =   19
      Top             =   2280
      Width           =   2055
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   9360
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
      Left            =   10680
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
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSFlexGridLib.MSFlexGrid gridCharge 
      Height          =   3975
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   11895
      _ExtentX        =   20981
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
      Format          =   66977794
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
      Format          =   66977795
      CurrentDate     =   39960
   End
   Begin VB.Label Label8 
      Caption         =   "Details"
      Height          =   255
      Left            =   8760
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Total"
      Height          =   255
      Left            =   8760
      TabIndex        =   20
      Top             =   2280
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
Attribute VB_Name = "frmBHTProfessionalChargesD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsStaff As New ADODB.Recordset
    Dim rsBHT As New ADODB.Recordset
    Dim MyBHT As New clsBHT

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
        temSql = "Select * from tblProfessionalCharges where ProfessionalChargesID = 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !UserID = UserID
        !ProfessionalCharge = True
        !ForBHTID = Val(cmbBHT.BoundText)
        !Fee = Val(txtValue.Text)
        !Comments = txtComments.Text
        !StaffID = Val(cmbStaff.BoundText)
        !Date = dtpDate.Value
        !Time = dtpTime.Value
        !IsInwardPaymentBill = True
        .Update
        .Close
    End With
    Call ClearValues
    Call FormatGrid
    Call fillGrid
    cmbSpeciality.SetFocus
End Sub

Private Sub fillGrid()
    Dim TotalValue As Double
    
    gridCharge.Visible = False
    
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblProfessionalCharges.UserID, tblSpeciality.Speciality, tblStaff.Name, tblProfessionalCharges.* " & _
                    "FROM (tblSpeciality RIGHT JOIN tblStaff ON tblSpeciality.SpecialityID = tblStaff.SpecialityID) RIGHT JOIN tblProfessionalCharges ON tblStaff.StaffID = tblProfessionalCharges.StaffID " & _
                    "WHERE (((tblProfessionalCharges.ProfessionalCharge) = 1 ) AND ((tblProfessionalCharges.Cancelled)=0) AND ((tblProfessionalCharges.ForBHTID)=" & Val(cmbBHT.BoundText) & "))"
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
            gridCharge.Text = !Name & Space(10)
            
            gridCharge.Col = 5
            
            gridCharge.Text = Format(!Comments, "")
            
            If Trim(gridCharge.Text) = "" Then
                gridCharge.Text = Space(15)
            End If
            
            gridCharge.Col = 6
            gridCharge.Text = Format(!Fee, "0.00")
            
            TotalValue = TotalValue + !Fee
            
            gridCharge.Col = 7
            gridCharge.Text = FullStaffName(!UserID)
                        
            
            .MoveNext
        
        Wend
        .Close
    End With
    txtTotal.Text = Format(TotalValue, "0.00")
    gridCharge.Rows = gridCharge.Rows + 1
    gridCharge.Row = gridCharge.Rows - 1
    gridCharge.Col = 6
    gridCharge.Text = Format(TotalValue, "0.00")
    gridCharge.Col = 4
    gridCharge.Text = "Total"
    
    gridCharge.Visible = True
    
End Sub

Private Sub FormatGrid()
    With gridCharge
        .Rows = 1
        .Cols = 8
        
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
        
        .Col = 7
        .Text = "User"
        
        .ColWidth(0) = 0
        .ColWidth(3) = 0
        
    End With
End Sub


Private Sub ClearValues()
    cmbSpeciality.Text = Empty
    cmbStaff.Text = Empty
    txtComments.Text = Empty
    txtValue.Text = Empty
    
End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim temID As Long
    temID = Val(gridCharge.TextMatrix(gridCharge.Row, 0))
    If temID = 0 Then
        MsgBox "Please select what to delete?"
        gridCharge.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblProfessionalCharges Where ProfessionalChargesID = " & temID
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If !Paid = True And !PaidCancelled = False Then
                MsgBox "This is already paid to doctor. So you can't cancel"
                .Close
                Exit Sub
            End If
            !Cancelled = True
            !CancelledUserID = UserID
            !CancelledDate = Date
            !CancelledTime = Now
            !CancelledDateTime = Now
            .Update
        End If
        .Close
    End With
    Call ClearValues
    Call FormatGrid
    Call fillGrid
    cmbSpeciality.SetFocus

End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    
    GridPrint gridCharge, ThisReportFormat, "Professional Charges", "BHT : " & cmbBHT.Text
    Printer.EndDoc
End Sub

Private Sub cmbBHT_Change()
    Call FormatGrid
    Call fillGrid
    If IsNumeric(cmbBHT.BoundText) = False Then Exit Sub
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
    With rsBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where IsBHT = 1 order by BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
    Dim Speciality As New clsFillCombos
    Speciality.FillAnyCombo cmbSpeciality, "Speciality", False
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    dtpTime.Value = Time
End Sub

