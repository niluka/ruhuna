VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBHTBookeepingSummery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BHT Bookeeping Summmery"
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13230
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
   ScaleHeight     =   10845
   ScaleWidth      =   13230
   Begin MSFlexGridLib.MSFlexGrid gridSummery 
      Height          =   9495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   16748
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   21954563
      CurrentDate     =   39960
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   21954563
      CurrentDate     =   39960
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   10200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "To &Excel"
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
      Left            =   120
      TabIndex        =   6
      Top             =   10200
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
      Left            =   11880
      TabIndex        =   7
      Top             =   10200
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
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   120
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
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmBHTBookeepingSummery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
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
    
    
    Dim temBHT As String
    Dim temBHTID As Long
    Dim temDOD As Date
    Dim temDOA As Date
    Dim temPt As String
    Dim temPM As String
    Dim temCC As String

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    
    Dim temMyBHT As New clsBHT
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblBHT.*, dbo.tblPatientMainDetails.FirstName FROM dbo.tblBHT LEFT OUTER JOIN dbo.tblPatientMainDetails ON dbo.tblBHT.PatientID = dbo.tblPatientMainDetails.PatientID Where IsBHT = 1 AND DOD between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            
            temBHT = !BHT
            
            temBHTID = !BHTID
            
            temMyBHT.BHTID = !BHTID
            
            temPM = temMyBHT.PaymentMethod
            temCC = temMyBHT.HealthSchemeSupplier
            
            temPt = !FirstName
            
            temDOA = Format(!DOA, "dd MMMM yyyy")
            
            temDOD = Format(!DOD, "dd MMMM yyyy")
            
            
            
            BHTDetails (!BHTID)
            
            
            
            .MoveNext
        Wend
        .Close
    End With

End Sub

Private Sub BHTDetails(BHTID As Long)
    Dim AdmissionCharge As Double
    Dim LinanCharge As Double
    Dim RoomCharge As Double
    Dim MaintananceCharge As Double
    Dim NursingCharge As Double
    Dim ProfessionalCharge As Double
    Dim MedicineCharge As Double
    
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select  Sum(Price) as SumOfPrice,  Sum(Discount) as SumOfDiscount,  Sum(NetPrice) as SumOfNetPrice,  Sum(Balance) as SumOfBalance,  Sum(AdmissionCharge) as SumOfAdmissionCharge,  Sum(LinanCharge) as SumOfLinanCharge,  Sum(RoomCharge) as SumOfRoomCharge,  Sum(ServicesCharge) as SumOfServicesCharge,  Sum(MaintananceCharge) as SumOfMaintananceCharge,  Sum(NursingCharge) as SumOfNursingCharge,  Sum(ProfessionalCharge) as SumOfProfessionalCharge,  Sum(AdditionalCharge) as SumOfAdditionalCharge,  Sum(MedicineCharge) as SumOfMedicineCharge,  Sum(TotalCharge) as SumOfTotalCharge,  Sum(Payments) as SumOfPayments,  Sum(FAdmissionRate) as SumOfFAdmissionRate,  Sum(FInitialLinanRate) as SumOfFInitialLinanRate,  Sum(FLaterLinanRate) as SumOfFLaterLinanRate,  Sum(FMaintananceRate) as SumOfFMaintananceRate,  Sum(FMaintainaceCashDiscountRate) as SumOfFMaintainaceCashDiscountRate, " & _
                    "Sum(FNursingRate) as SumOfFNursingRate,  Sum(FICUNursingRate) as SumOfFICUNursingRate,  Sum(FAdmissionFee) as SumOfFAdmissionFee, " & _
                    "Sum(FAdmissionCharge) as SumOfFAdmissionCharge,  Sum(FLinanCharge) as SumOfFLinanCharge,  Sum(FRoomCharge) as SumOfFRoomCharge,  Sum(FServicesCharge) as SumOfFServicesCharge,  Sum(FMaintananceCharge) as SumOfFMaintananceCharge,  Sum(FNursingCharge) as SumOfFNursingCharge,  Sum(FProfessionalCharge) as SumOfFProfessionalCharge,  Sum(FMedicineCharge) as SumOfFMedicineCharge,  Sum(FAdditionalCharge) as SumOfFAdditionalCharge,  Sum(FTotalCharge) as SumOfFTotalCharge,  Sum(FPayments) as SumOfFPayments " & _
                    "from tblBHT " & _
                    "Where BHTID = " & BHTID
                    
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfAdmissionCharge) = False Then AdmissionCharge = !SumOfAdmissionCharge
            If IsNull(!SumOfLinanCharge) = False Then LinanCharge = !SumOfLinanCharge
            If IsNull(!SumOfRoomCharge) = False Then RoomCharge = !SumOfRoomCharge
            If IsNull(!SumOfMaintananceCharge) = False Then MaintananceCharge = !SumOfMaintananceCharge
            If IsNull(!SumOfNursingCharge) = False Then NursingCharge = !SumOfNursingCharge
            If IsNull(!SumOfProfessionalCharge) = False Then ProfessionalCharge = !SumOfProfessionalCharge
            If IsNull(!SumOfMedicineCharge) = False Then MedicineCharge = !SumOfMedicineCharge
        
        End If
    End With
    
    NewRow "Admission Fee", AdmissionCharge
    NewRow "Room Charge", RoomCharge
    NewRow "Medicine Charge", MedicineCharge
    
    FillBHTServices (BHTID)
    
    NewRow "Professional Charges", ProfessionalCharge
    NewRow "Linen Charge", LinanCharge
    NewRow "Nursing Charges", NursingCharge
    NewRow "Maintainance Charges", MaintananceCharge
    

End Sub


Private Sub FillBHTServices(BHTID As Long)
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblServiceCategory.ServiceCategory, Sum(tblPatientService.Charge) AS SumOfCharge " & _
                    "FROM (tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID " & _
                    "WHERE (((tblPatientService.Deleted)=0) AND ((tblPatientService.BHTID)=" & BHTID & ")) " & _
                    "GROUP BY tblServiceCategory.ServiceCategory"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            NewRow !ServiceCategory, !SumOfCharge
            .MoveNext
        Wend
    End With
End Sub

Private Sub NewRow(Descreption As String, Value As Double)
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = temBHT
    gridSummery.Col = 1
    gridSummery.Text = temBHTID
    gridSummery.Col = 2
    gridSummery.Text = temPt
    
    gridSummery.Col = 3
    gridSummery.Text = temCC
    
    
    gridSummery.Col = 4
    gridSummery.Text = Format(temDOA, "dd MMM yy")
    
    gridSummery.Col = 5
    gridSummery.Text = Format(temDOD, "dd MMM yy")
    
    gridSummery.Col = 7
    gridSummery.Text = temPM
    
    gridSummery.Col = 8
    gridSummery.Text = Descreption
    
    gridSummery.Col = 10
    gridSummery.Text = Format(Value, "#,##0.00")
    
End Sub

Private Sub FormatGrid()
    With gridSummery
        .Clear
        
        .Cols = 11
        .Rows = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "BHT"
        
        .Col = 1
        .Text = "Final Bill No"
        
        .Col = 2
        .Text = "Patient"
        
        .Col = 3
        .Text = "Company"
       
        .Col = 4
        .Text = "DOA"
        
        .Col = 5
        .Text = "DOD"
        
        .Col = 6
        .Text = ""
        

        .Col = 7
        .Text = "Paid as"
        
        
        .Col = 8
        .Text = "Descreption"
        
        .Col = 9
        .Text = ""
        
        .Col = 10
        .Text = "Value"
        
    End With
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpFrom.Value = Date
    dtpTo.Value = Date
    GetCommonSettings Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridSummery, "Book Keeping Summery For BHTs", "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")

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
    
    
    
    GridPrint gridSummery, ThisReportFormat, "Book Keeping Summery For BHTs", "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    Printer.EndDoc

End Sub

Private Sub btnProcess_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call GetSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSettings
End Sub
