VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBHTProfessionalPaymentsAgeAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Professional Fee Payments for BHT patients"
   ClientHeight    =   8250
   ClientLeft      =   855
   ClientTop       =   -2445
   ClientWidth     =   10965
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
   ScaleHeight     =   8250
   ScaleWidth      =   10965
   Begin VB.ComboBox cmbType 
      Height          =   360
      Left            =   6720
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   120
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid gridPay 
      Height          =   5895
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   10398
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbSpeciality 
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbStaff 
      Height          =   360
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   9480
      TabIndex        =   4
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   6960
      Width           =   5175
      Begin VB.ComboBox cmbPrinter 
         Height          =   360
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbPaper 
         Height          =   360
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label17 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Paper"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Process"
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
   Begin VB.Label Label3 
      Caption         =   "Type"
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Doctor"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Speciality"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBHTProfessionalPaymentsAgeAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim temSql As String
    Dim rsStaff As New ADODB.Recordset
    
    Dim CsetPrinter As New cSetDfltPrinter
    Dim NumForms As Long, i As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    Dim SuppliedWord As String
    
    
Private Sub FormatGrid()
    With gridPay
        .Clear
        
        .Rows = 1
        .Cols = 9
        
        .Row = 0
        
        .Col = 0
        .Text = "Name"
        
        .Col = 1
        .Text = "< 7"

        .Col = 2
        .Text = "8 - 14"

        .Col = 3
        .Text = "15 - 21"

        .Col = 4
        .Text = "22 - 30"

        .Col = 5
        .Text = "31 - 60"

        .Col = 6
        .Text = "61 - 90"

        .Col = 7
        .Text = "> 90"

        .Col = 8
        .Text = "Total"

    End With
End Sub


Private Sub btnProcess_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub bttnClose_Click()
    Unload Me
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
    With cmbStaff
        Set .RowSource = rsStaff
        .ListField = "NameWithTitle"
        .BoundColumn = "StaffID"
        .Text = Empty
    End With
End Sub

Private Sub cmbSpeciality_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbStaff.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbSpeciality.Text = Empty
    End If
End Sub

'Private Sub cmbStaff_Change()
'    Call FormatGrid
'    Call FillGrid
'End Sub
'
'Private Sub cmbStaff_Click(Area As Integer)
'    Call FormatGrid
'    Call FillGrid
'End Sub

Private Sub cmbStaff_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbType.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbStaff.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call PopulatePrinters
    Call GetSettings
    Call FormatGrid
    Call FillGrid
    
End Sub

Private Sub printBill()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle): DoEvents
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    
    Dim temText As String
    Dim CenterX As Long
    Dim FieldX As Long
    Dim NoX As Long
    Dim ValueX As Long
    Dim AllLines() As String
    Dim i As Integer
    Dim temY As Long
    Dim n As Long
    
    With MyFOnt
        .Name = DefaultFont.Name
        .Bold = False
        .Italic = False
        .Size = 12
        .Italic = False
        .Underline = False
    End With
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then

        Printer.EndDoc
        
    End If
End Sub


Private Sub FillCombos()
    Dim Speciality As New clsFillCombos
    Speciality.FillAnyCombo cmbSpeciality, "Speciality", False
    With cmbType
'        .AddItem "ALL"
        .AddItem "BHT"
        .AddItem "GSB"
        .AddItem "LAB"
        .AddItem "OPD"
        .AddItem "RON"
        .AddItem "MED"
        .AddItem "HST"
    End With
End Sub

Private Sub GetSettings(): On Error Resume Next
'    cmbPaymentMethod.BoundText = 1 'Val(GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1))
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "")
    GetCommonSettings Me
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
    SaveCommonSettings Me
End Sub


Private Sub FillGrid()
    Dim D0D7 As Double
    Dim D8D14 As Double
    Dim D15D21 As Double
    Dim D22D30 As Double
    Dim D31D60 As Double
    Dim D61D90 As Double
    Dim MoreD90 As Double
    
    Dim TD0D7 As Double
    Dim TD8D14 As Double
    Dim TD15D21 As Double
    Dim TD22D30 As Double
    Dim TD31D60 As Double
    Dim TD61D90 As Double
    Dim TMoreD90 As Double
    
    
    Dim temDays As Double
    Dim PreviousStaffID As Long
    
    Dim rsTem As New ADODB.Recordset
    Dim temText As String
    Dim temTotal As Double
    Dim temCount As Double
    
    Dim temS As String
    Dim temF As String
    Dim temW As String
    Dim temG As String
    Dim temH As String
    Dim temO As String
    
    
    With rsTem
        If .State = 1 Then .Close
        Select Case cmbType.Text
'            Case "ALL"
'                temW = "WHERE (dbo.tblProfessionalCharges.Cancelled = 0) "
            Case "OPD"
                temS = "SELECT     TOP 100 PERCENT SUM(dbo.tblProfessionalCharges.Fee) AS TotalProfessionalFee, dbo.tblProfessionalCharges.StaffID, dbo.tblProfessionalCharges.Date , dbo.tblStaff.SpecialityID "
                temF = "FROM         dbo.tblBHT RIGHT OUTER JOIN dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForOPDBillID = dbo.tblIncomeBill.IncomeBillID LEFT OUTER JOIN dbo.tblStaff ON dbo.tblProfessionalCharges.StaffID = dbo.tblStaff.StaffID ON dbo.tblBHT.BHTID = dbo.tblProfessionalCharges.ForBHTID "
                temW = "WHERE (dbo.tblProfessionalCharges.Cancelled = 0) AND (dbo.tblProfessionalCharges.IsOPDBill = 1) "
                temG = "GROUP BY dbo.tblProfessionalCharges.StaffID, dbo.tblStaff.SpecialityID, dbo.tblProfessionalCharges.Date, dbo.tblStaff.Name,  dbo.tblProfessionalCharges.ForOPDBillID "
            
            Case "LAB"
                temS = "SELECT     TOP 100 PERCENT SUM(dbo.tblProfessionalCharges.Fee) AS TotalProfessionalFee, dbo.tblProfessionalCharges.StaffID, dbo.tblProfessionalCharges.Date , dbo.tblStaff.SpecialityID "
                temF = "FROM         dbo.tblBHT RIGHT OUTER JOIN dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForLabBillID = dbo.tblIncomeBill.IncomeBillID LEFT OUTER JOIN dbo.tblStaff ON dbo.tblProfessionalCharges.StaffID = dbo.tblStaff.StaffID ON dbo.tblBHT.BHTID = dbo.tblProfessionalCharges.ForBHTID "
                temW = "WHERE (dbo.tblProfessionalCharges.Cancelled = 0) AND (dbo.tblProfessionalCharges.IsLabBill = 1) "
                temG = "GROUP BY dbo.tblProfessionalCharges.StaffID, dbo.tblStaff.SpecialityID, dbo.tblProfessionalCharges.Date, dbo.tblStaff.Name ,  dbo.tblProfessionalCharges.ForLabBillID "
    
            Case "BHT"
                temS = "SELECT     TOP 100 PERCENT SUM(dbo.tblProfessionalCharges.Fee) AS TotalProfessionalFee, dbo.tblProfessionalCharges.StaffID, dbo.tblProfessionalCharges.Date , dbo.tblStaff.SpecialityID "
                temF = "FROM         dbo.tblStaff RIGHT OUTER JOIN dbo.tblProfessionalCharges ON dbo.tblStaff.StaffID = dbo.tblProfessionalCharges.StaffID LEFT OUTER JOIN dbo.tblBHT ON dbo.tblProfessionalCharges.ForBHTID = dbo.tblBHT.BHTID "
                temW = "WHERE (dbo.tblProfessionalCharges.Cancelled = 0) AND (dbo.tblProfessionalCharges.IsInwardPaymentBill = 1) "
                temG = "GROUP BY dbo.tblProfessionalCharges.StaffID, dbo.tblStaff.SpecialityID, dbo.tblProfessionalCharges.Date, dbo.tblStaff.Name "
            
            Case "GSB"
                temS = "SELECT     TOP 100 PERCENT SUM(dbo.tblProfessionalCharges.Fee) AS TotalProfessionalFee, dbo.tblProfessionalCharges.StaffID, dbo.tblProfessionalCharges.Date , dbo.tblStaff.SpecialityID "
                temF = "FROM         dbo.tblStaff RIGHT OUTER JOIN dbo.tblProfessionalCharges ON dbo.tblStaff.StaffID = dbo.tblProfessionalCharges.StaffID LEFT OUTER JOIN dbo.tblBHT ON dbo.tblProfessionalCharges.ForBHTID = dbo.tblBHT.BHTID "
                temW = "WHERE (dbo.tblProfessionalCharges.Cancelled = 0) AND (dbo.tblProfessionalCharges.IsGSBill = 1) "
                temG = "GROUP BY dbo.tblProfessionalCharges.StaffID, dbo.tblStaff.SpecialityID, dbo.tblProfessionalCharges.Date, dbo.tblStaff.Name "
    
            Case "MED"
                temS = "SELECT     TOP 100 PERCENT SUM(dbo.tblProfessionalCharges.Fee) AS TotalProfessionalFee, dbo.tblProfessionalCharges.StaffID, dbo.tblProfessionalCharges.Date , dbo.tblStaff.SpecialityID "
                temF = "FROM         dbo.tblBHT RIGHT OUTER JOIN dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForMedicalTestBillID = dbo.tblIncomeBill.IncomeBillID LEFT OUTER JOIN dbo.tblStaff ON dbo.tblProfessionalCharges.StaffID = dbo.tblStaff.StaffID ON dbo.tblBHT.BHTID = dbo.tblProfessionalCharges.ForBHTID "
                temW = "WHERE (dbo.tblProfessionalCharges.Cancelled = 0) AND (dbo.tblProfessionalCharges.IsMedicalTestBill = 1) "
                temG = "GROUP BY dbo.tblProfessionalCharges.StaffID, dbo.tblStaff.SpecialityID, dbo.tblProfessionalCharges.Date, dbo.tblStaff.Name ,  dbo.tblProfessionalCharges.ForMedicalTestBillID "
        
            Case "RON"
                temS = "SELECT     TOP 100 PERCENT SUM(dbo.tblProfessionalCharges.Fee) AS TotalProfessionalFee, dbo.tblProfessionalCharges.StaffID, dbo.tblProfessionalCharges.Date , dbo.tblStaff.SpecialityID "
                temF = "FROM         dbo.tblBHT RIGHT OUTER JOIN dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForRBillID = dbo.tblIncomeBill.IncomeBillID LEFT OUTER JOIN dbo.tblStaff ON dbo.tblProfessionalCharges.StaffID = dbo.tblStaff.StaffID ON dbo.tblBHT.BHTID = dbo.tblProfessionalCharges.ForBHTID "
                temW = "WHERE (dbo.tblProfessionalCharges.Cancelled = 0) AND (dbo.tblProfessionalCharges.IsRBill = 1) "
                temG = "GROUP BY dbo.tblProfessionalCharges.StaffID, dbo.tblStaff.SpecialityID, dbo.tblProfessionalCharges.Date, dbo.tblStaff.Name ,  dbo.tblProfessionalCharges.ForRBillID "
    
            Case "HST"
                temS = "SELECT     TOP 100 PERCENT SUM(dbo.tblProfessionalCharges.Fee) AS TotalProfessionalFee, dbo.tblProfessionalCharges.StaffID, dbo.tblProfessionalCharges.Date , dbo.tblStaff.SpecialityID "
                temF = "FROM         dbo.tblBHT RIGHT OUTER JOIN dbo.tblProfessionalCharges LEFT OUTER JOIN dbo.tblIncomeBill ON dbo.tblProfessionalCharges.ForHSTBillID = dbo.tblIncomeBill.IncomeBillID LEFT OUTER JOIN dbo.tblStaff ON dbo.tblProfessionalCharges.StaffID = dbo.tblStaff.StaffID ON dbo.tblBHT.BHTID = dbo.tblProfessionalCharges.ForBHTID "
                temW = "WHERE (dbo.tblProfessionalCharges.Cancelled = 0) AND (dbo.tblProfessionalCharges.IsHSTBill = 1) "
                temG = "GROUP BY dbo.tblProfessionalCharges.StaffID, dbo.tblStaff.SpecialityID, dbo.tblProfessionalCharges.Date, dbo.tblStaff.Name,   dbo.tblProfessionalCharges.ForHSTBillID "
            
            Case Else
                
                Exit Sub
                
        End Select
        
        If IsNumeric(cmbStaff.BoundText) = True Then
            temW = temW & " AND dbo.tblProfessionalCharges.StaffID = " & Val(cmbStaff.BoundText)
        End If
        
        temO = "ORDER BY dbo.tblStaff.Name "
        

        If IsNumeric(cmbSpeciality.BoundText) = True Then
            temW = temW & " AND dbo.tblStaff.SpecialityID = " & Val(cmbSpeciality.BoundText) & "  "
        End If
        
        temSql = temS & temF & temW & temG & temH & temO
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        While .EOF = False
        
        
            temDays = Abs(DateDiff("d", Date, !Date))
            
            If PreviousStaffID <> !StaffID Then
                gridPay.Rows = gridPay.Rows + 1
                gridPay.Row = gridPay.Rows - 1
                gridPay.Col = 0
                gridPay.Text = FullStaffName(PreviousStaffID)
                gridPay.Col = 1
                gridPay.Text = Format(D0D7, "#,##0.00")
                gridPay.Col = 2
                gridPay.Text = Format(D8D14, "#,##0.00")
                gridPay.Col = 3
                gridPay.Text = Format(D15D21, "#,##0.00")
                gridPay.Col = 4
                gridPay.Text = Format(D22D30, "#,##0.00")
                gridPay.Col = 5
                gridPay.Text = Format(D31D60, "#,##0.00")
                gridPay.Col = 6
                gridPay.Text = Format(D61D90, "#,##0.00")
                gridPay.Col = 7
                gridPay.Text = Format(MoreD90, "#,##0.00")
                gridPay.Col = 8
                gridPay.Text = Format(D0D7 + D8D14 + D15D21 + D22D30 + D31D60 + D61D90 + MoreD90, "#,##0.00")
                TD0D7 = TD0D7 + D0D7
                TD8D14 = TD8D14 + D8D14
                TD15D21 = TD15D21 + D15D21
                TD22D30 = TD22D30 + D22D30
                TD31D60 = TD31D60 + D31D60
                TD61D90 = TD61D90 + D61D90
                TMoreD90 = TMoreD90 + MoreD90
                
                D0D7 = 0
                D8D14 = 0
                D15D21 = 0
                D22D30 = 0
                D31D60 = 0
                D61D90 = 0
                MoreD90 = 0
                PreviousStaffID = !StaffID
            Else
            
            End If
                        
            If temDays < 8 Then
                D0D7 = D0D7 + !TotalProfessionalFee
            ElseIf temDays < 14 Then
                D8D14 = D8D14 + !TotalProfessionalFee
            ElseIf temDays < 21 Then
                D15D21 = D15D21 + !TotalProfessionalFee
            ElseIf temDays < 31 Then
                D22D30 = D22D30 + !TotalProfessionalFee
            ElseIf temDays < 61 Then
                D31D60 = D31D60 + !TotalProfessionalFee
            ElseIf temDays < 91 Then
                D61D90 = D61D90 + !TotalProfessionalFee
            Else
                MoreD90 = MoreD90 + !TotalProfessionalFee
            End If
        
        
        
            
            .MoveNext
        Wend
        .Close
    End With
    
    gridPay.Rows = gridPay.Rows + 1
    gridPay.Row = gridPay.Rows - 1
    
    gridPay.Col = 0
    gridPay.Text = "Total"
    
    gridPay.Col = 1
    gridPay.Text = Format(TD0D7, "#,##0.00")
    
    gridPay.Col = 2
    gridPay.Text = Format(TD8D14, "#,##0.00")
    
    gridPay.Col = 3
    gridPay.Text = Format(TD15D21, "#,##0.00")
    gridPay.Col = 4
    gridPay.Text = Format(TD22D30, "#,##0.00")
    gridPay.Col = 5
    gridPay.Text = Format(TD31D60, "#,##0.00")
    gridPay.Col = 6
    gridPay.Text = Format(TD61D90, "#,##0.00")
    gridPay.Col = 7
    gridPay.Text = Format(TMoreD90, "#,##0.00")
    gridPay.Col = 8
    gridPay.Text = Format(TD0D7 + TD8D14 + TD15D21 + TD22D30 + TD31D60 + TD61D90 + TMoreD90, "#,##0.00")
    
    
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub


Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub PopulatePapers(): On Error Resume Next
    cmbPaper.Clear
    SetPrinter = False
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text): DoEvents
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
'        With FormSize
'            .cx = BillPaperHeight
'            .cy = BillPaperWidth
'        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                cmbPaper.AddItem PtrCtoVbString(.pName)
'                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle): DoEvents
    End If
End Sub

Private Sub cmbPrinter_Change()
    Call PopulatePapers
End Sub

Private Sub cmbPrinter_Click()
    Call PopulatePapers
End Sub

