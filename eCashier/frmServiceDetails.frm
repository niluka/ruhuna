VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmServiceDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Details"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13200
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
   ScaleWidth      =   13200
   Begin VB.CheckBox chkOPD 
      Caption         =   "For OPD"
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CheckBox chkBHT 
      Caption         =   "For BHT"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   720
      Width           =   1815
   End
   Begin VB.CheckBox chkGSB 
      Caption         =   "For GSB"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CheckBox chkLab 
      Caption         =   "For Lab"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CheckBox chkMT 
      Caption         =   "For MT"
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   720
      Width           =   1815
   End
   Begin VB.CheckBox chkHST 
      Caption         =   "For HST"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CheckBox chkR 
      Caption         =   "For Roentgents "
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   1440
      Width           =   1815
   End
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11880
      TabIndex        =   5
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
      Left            =   120
      TabIndex        =   4
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   7560
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
   Begin MSFlexGridLib.MSFlexGrid gridService 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbCat 
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Category"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmServiceDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridService, HospitalName
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    ThisReportFormat.ReportPrintOrientation = Landscape
    
    
    
    GetPrintDefaults ThisReportFormat
    
    With ThisReportFormat
        
        .LeftMargin = 0
        .ColSpace = 70
        
        .TopicFontSize = 11
        .TopicFontName = "Tahoma"
        
        .SubTopicFontSize = 10
        .SubTopicFontName = "Tahoma"
        
        .HeaderFontName = "Tahoma"
        .HeaderFontSize = 8
        .HeaderFontBold = False
        .HeaderFontUnderline = False
        
        .ColTopicFontName = "Tahoma"
        .ColTopicFontSize = 8
        .ColTopicFontBold = False
        .ColTopicFontUnderline = False
        
        .ColFontSize = 7
        .ColFontName = "Tahoma"
        
    End With
    
    
    
    GridPrint gridService, ThisReportFormat, HospitalName & " - Daily Summery Report"
    Printer.EndDoc
End Sub

Private Sub btnProcess_Click()
    Screen.MousePointer = vbHourglass
    gridService.Visible = False
    Call FormatGrid
    Call FillGrid
    gridService.Visible = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call FillCombos
    Call GetSettings
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
End Sub

Private Sub GetSettings(): On Error Resume Next
    GetCommonSettings Me
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillAnyCombo cmbCat, "ServiceCategory", True
End Sub

Private Sub FormatGrid()
    With gridService
        .Clear
        
        .Cols = 5
        .Rows = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "Category"
        
        .Col = 1
        .Text = "Sub category"
        
        .Col = 2
        .Text = "Hospital Fee"
        
        .Col = 3
        .Text = "Professional Fee"
        
        .Col = 4
        .Text = "Total Fee"
    End With
End Sub

Private Function ProfessionalPayment(SubCatID As Long) As Double
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT sum(tblServiceProfessionalCharges.Fee) as SumOfP  " & _
                    "FROM tblServiceProfessionalCharges  " & _
                    "Where tblServiceProfessionalCharges.ServiceSubcategoryID  = " & SubCatID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If IsNull(!SumOfP) = False Then
            ProfessionalPayment = !SumOfP
        Else
            ProfessionalPayment = 0
        End If
        .Close
    End With
End Function

Private Function HospitalFeeCat(CatID As Long) As Double
    HospitalFeeCat = 0
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceCategory where ServiceCategoryID = " & CatID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!Fee) = False Then
                HospitalFeeCat = Format(!Fee, "0.00")
            End If
        End If
        .Close
    End With
End Function


Private Function HospitalFeeSCat(CatID As Long) As Double
    HospitalFeeSCat = 0
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblServiceSubCategory where ServiceSubCategoryID = " & CatID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!Fee) = False Then
                HospitalFeeSCat = Format(!Fee, "0.00")
            End If
        End If
        .Close
    End With
End Function

Private Sub FillGrid()
    If IsNumeric(cmbCat.BoundText) = True Then
        Call SubCatFill
    Else
        Call CatFill
    End If

End Sub

Private Sub CatFill()
    GetSettings
    Dim rsTem As New ADODB.Recordset
    Dim rsSC As New ADODB.Recordset
    Dim i As Integer
    Dim CatH As Double
    Dim SCatH As Double
    Dim SCatP As Double
    
    Dim temSelect As String
    Dim temWhere As String
    Dim temOrder As String
    
    
    
    With gridService
        If .ColWidth(0) < 10 Then .ColWidth(0) = 1200
        temSql = "Select * from tblServiceCategory where Deleted= 0 order by ServiceCategory "
        
        
        
        
        
        
        
            temSelect = "Select * from tblServiceCategory   "
            temWhere = "where Deleted= 0  "
            temOrder = " Order By ServiceCategory  "
            
            If chkBHT.Value = 1 Then
                temWhere = temWhere & " AND ForBHT = 1 "
            Else
'                temWhere = temWhere & " AND ForBHT = 0 "
            
            End If
            
            If chkGSB.Value = 1 Then
                temWhere = temWhere & " AND ForGSB = 1 "
            Else
'                temWhere = temWhere & " AND ForGSB = 0 "
            End If
            
            If chkOPD.Value = 1 Then
                temWhere = temWhere & " AND ForOPD = 1 "
            Else
'                temWhere = temWhere & " AND ForOPD = 0 "
            End If
            
            If chkLab.Value = 1 Then
                temWhere = temWhere & " AND ForLab = 1 "
            Else
'                temWhere = temWhere & " AND ForLab = 0 "
            End If
            
            If chkMT.Value = 1 Then
                temWhere = temWhere & " AND ForMT = 1 "
            Else
'                temWhere = temWhere & " AND ForMT = 0 "
            End If
            
            If chkHST.Value = 1 Then
                temWhere = temWhere & " AND ForHST = 1 "
            Else
'                temWhere = temWhere & " AND ForHST = 0 "
            End If
            
            If chkR.Value = 1 Then
                temWhere = temWhere & " AND ForR = 1 "
            Else
'                temWhere = temWhere & " AND ForR = 0 "
            End If
            
            temSql = temSelect & temWhere & temOrder
        
        
        
        
        If rsTem.State = 1 Then rsTem.Close
        rsTem.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While rsTem.EOF = False
            .Rows = .Rows + 1
            .Row = .Rows - 1
            
            .Col = 0
            .Text = rsTem!ServiceCategory
            
            CatH = HospitalFeeCat(rsTem!ServiceCategoryID)
            
            .Col = 2
            .Text = Format(CatH, "#,##0.00")
            
            .Col = 4
            .Text = Format(CatH, "#,##0.00")
            
            
            If rsSC.State = 1 Then rsSC.Close
            
            temSelect = "Select * from tblServiceSubCategory  "
            temWhere = "where ServiceCategoryID = " & rsTem!ServiceCategoryID & " And Deleted = 0 "
            temOrder = " Order By ServiceSubcategory  "
            
            If chkBHT.Value = 1 Then
                temWhere = temWhere & " AND ForBHT = 1 "
            Else
'                temWhere = temWhere & " AND ForBHT = 0 "
            
            End If
            
            If chkGSB.Value = 1 Then
                temWhere = temWhere & " AND ForGSB = 1 "
            Else
'                temWhere = temWhere & " AND ForGSB = 0 "
            End If
            
            If chkOPD.Value = 1 Then
                temWhere = temWhere & " AND ForOPD = 1 "
            Else
'                temWhere = temWhere & " AND ForOPD = 0 "
            End If
            
            If chkLab.Value = 1 Then
                temWhere = temWhere & " AND ForLab = 1 "
            Else
'                temWhere = temWhere & " AND ForLab = 0 "
            End If
            
            If chkMT.Value = 1 Then
                temWhere = temWhere & " AND ForMT = 1 "
            Else
'                temWhere = temWhere & " AND ForMT = 0 "
            End If
            
            If chkHST.Value = 1 Then
                temWhere = temWhere & " AND ForHST = 1 "
            Else
'                temWhere = temWhere & " AND ForHST = 0 "
            End If
            
            If chkR.Value = 1 Then
                temWhere = temWhere & " AND ForR = 1 "
            Else
'                temWhere = temWhere & " AND ForR = 0 "
            End If
            
            temSql = temSelect & temWhere & temOrder
            
            rsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            
            While rsSC.EOF = False
                .Rows = .Rows + 1
                
                .Row = .Rows - 1
                
                .Col = 1
                .Text = rsSC!ServiceSubcategory
                
                SCatH = HospitalFeeSCat(rsSC!ServiceSubcategoryID)
                SCatP = ProfessionalPayment(rsSC!ServiceSubcategoryID)
                                
                .Col = 2
                .Text = Format(SCatH, "#,##0.00")
                
                .Col = 3
                .Text = Format(SCatP, "#,##0.00")
                
                .Col = 4
                .Text = Format(SCatH + SCatP, "#,##0.00")
                                
                                
                                
                rsSC.MoveNext
            Wend
            
            rsTem.MoveNext
        Wend
        
    End With
End Sub

Private Sub SubCatFill()
    GetSettings
    Dim rsTem As New ADODB.Recordset
    Dim rsSC As New ADODB.Recordset
    Dim i As Integer
    Dim CatH As Double
    Dim SCatH As Double
    Dim SCatP As Double
        
    Dim temSelect As String
    Dim temWhere As String
    Dim temOrder As String
    
        
        
    If rsSC.State = 1 Then rsSC.Close
    temSelect = "Select * from tblServiceSubCategory "
    temWhere = "where ServiceCategoryID = " & Val(cmbCat.BoundText) & " And Deleted = 0 "
    temOrder = "Order By ServiceSubcategory "
    

            If chkBHT.Value = 1 Then
                temWhere = temWhere & " AND ForBHT = 1 "
            Else
'                temWhere = temWhere & " AND ForBHT = 0 "
            
            End If
            
            If chkGSB.Value = 1 Then
                temWhere = temWhere & " AND ForGSB = 1 "
            Else
'                temWhere = temWhere & " AND ForGSB = 0 "
            End If
            
            If chkOPD.Value = 1 Then
                temWhere = temWhere & " AND ForOPD = 1 "
            Else
'                temWhere = temWhere & " AND ForOPD = 0 "
            End If
            
            If chkLab.Value = 1 Then
                temWhere = temWhere & " AND ForLab = 1 "
            Else
'                temWhere = temWhere & " AND ForLab = 0 "
            End If
            
            If chkMT.Value = 1 Then
                temWhere = temWhere & " AND ForMT = 1 "
            Else
'                temWhere = temWhere & " AND ForMT = 0 "
            End If
            
            If chkHST.Value = 1 Then
                temWhere = temWhere & " AND ForHST = 1 "
            Else
'                temWhere = temWhere & " AND ForHST = 0 "
            End If
            
            If chkR.Value = 1 Then
                temWhere = temWhere & " AND ForR = 1 "
            Else
'                temWhere = temWhere & " AND ForR = 0 "
            End If
    
    
    temSql = temSelect & temWhere & temOrder
    
    rsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    With gridService
    
        .ColWidth(0) = 0
        
        While rsSC.EOF = False
            .Rows = .Rows + 1
            
            .Row = .Rows - 1
            
            .Col = 1
            .Text = rsSC!ServiceSubcategory
            
            SCatH = HospitalFeeSCat(rsSC!ServiceSubcategoryID)
            SCatP = ProfessionalPayment(rsSC!ServiceSubcategoryID)
                            
            .Col = 2
            .Text = Format(SCatH, "#,##0.00")
            
            .Col = 3
            .Text = Format(SCatP, "#,##0.00")
            
            .Col = 4
            .Text = Format(SCatH + SCatP, "#,##0.00")
                            
                            
                            
            rsSC.MoveNext
        Wend
    End With
End Sub
