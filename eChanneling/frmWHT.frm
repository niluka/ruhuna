VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmWHT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WHT"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12900
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
   ScaleHeight     =   7500
   ScaleWidth      =   12900
   Begin btButtonEx.ButtonEx btnProcess 
      Height          =   375
      Left            =   5040
      TabIndex        =   5
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
   Begin MSFlexGridLib.MSFlexGrid gridDoc 
      Height          =   5775
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   10186
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   78446595
      CurrentDate     =   40188
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   78446595
      CurrentDate     =   40188
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   6960
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
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6960
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   6960
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
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmWHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnExcel_Click()
    GridToExcel gridDoc, "Withhalding Tax", "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
End Sub

Private Sub btnPrint_Click()
    Dim MyRP As PrintReport
    GetPrintDefaults MyRP
    
    GridPrint gridDoc, MyRP, "Withhalding Tax", "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    
End Sub

Private Sub btnProcess_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call GetSettings
End Sub

Private Sub FormatGrid()
    With gridDoc
        .Clear
        
        .Rows = 1
        .Cols = 5
        
        .Row = 0
        
        .Col = 0
        .Text = "No."
        
        .Col = 1
        .Text = "Doctor"
        
        .Col = 2
        .Text = "Channeling Fee"
        
        .Col = 3
        .Text = "Withhalding Tax"
        
        .Col = 4
        .Text = "Balance Paid"
        
        
    
    End With
    
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    
    If rsTem.State = 1 Then rsTem.Close
    temSql = "SELECT dbo.tblTitle.Title + ' ' + dbo.tblDoctor.DoctorName AS DocNameWithTitle, SUM(dbo.tblStaffPayment.PaidAmount) " & _
                      "AS SumOfPaidAmount, SUM(dbo.tblStaffPayment.TaxAmount) AS SumOfTaxAmount, SUM(dbo.tblStaffPayment.DoctorAmount) " & _
                      "AS SumOfDoctorAmount " & _
                        "FROM         dbo.tblTitle RIGHT OUTER JOIN " & _
                        "dbo.tblDoctor ON dbo.tblTitle.Title_ID = dbo.tblDoctor.DoctorTitle_ID RIGHT OUTER JOIN " & _
                        "dbo.tblStaffPayment ON dbo.tblDoctor.Doctor_ID = dbo.tblStaffPayment.Staff_ID " & _
                        "WHERE     (dbo.tblStaffPayment.PaidDate BETWEEN CONVERT(DATETIME, '" & dtpFrom.Value & "', 102) AND CONVERT(DATETIME, '" & dtpTo.Value & "', 102)) " & _
                        "GROUP BY dbo.tblTitle.Title + ' ' + dbo.tblDoctor.DoctorName, dbo.tblDoctor.DoctorName " & _
                        "ORDER BY dbo.tblDoctor.DoctorName"
    
    If rsTem.State = 1 Then rsTem.Close
    rsTem.Open temSql, cnnChannelling, adOpenStatic, adLockReadOnly
    
    
    
    With gridDoc
        
        While rsTem.EOF = False
            
            .Rows = .Rows + 1
            .Row = .Rows - 1
            
            .Col = 0
            .Text = .Row
            
            .Col = 1
            .Text = Format(rsTem!DocNameWithTitle, "")
            
            .Col = 2
            .Text = Format(rsTem!SumOfPaidAmount, "#,##0.00")
            
            .Col = 3
            .Text = Format(rsTem!SumOfTaxAmount, "#,##0.00")
            
            .Col = 4
            .Text = Format(rsTem!SumOfDoctorAmount, "#,##0.00")
            
            rsTem.MoveNext
        
        Wend
        rsTem.Close
    End With
    
End Sub

Private Sub SaveSettings()
    dtpFrom.Value = Date
    dtpTo.Value = Date
    SaveCommonSettings Me
End Sub

Private Sub GetSettings()
    GetCommonSettings Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub
