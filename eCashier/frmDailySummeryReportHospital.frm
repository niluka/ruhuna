VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDailySummeryReportHospital 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Summery Report - Hospital"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
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
   ScaleHeight     =   9285
   ScaleWidth      =   15270
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   5400
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6840
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   8760
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
      TabIndex        =   5
      Top             =   8760
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
      Left            =   13920
      TabIndex        =   4
      Top             =   8760
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
      Left            =   3480
      TabIndex        =   3
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67633155
      CurrentDate     =   40141
   End
   Begin MSFlexGridLib.MSFlexGrid gridSummery 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14208
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frmDailySummeryReportHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim CollectionForTheDay As Double
    
Private Sub OPDCollectionIncome()
    Dim SubTotal As Double
    Dim temTotal As Double
    Dim temCount As Long
    Dim rsTem As New ADODB.Recordset
    Dim rsCat As New ADODB.Recordset
    Dim rsIsSC As New ADODB.Recordset
    Dim rsSC As New ADODB.Recordset
    Dim temCatID As Long
    Dim temSCID As Long
    Dim SCExists As Boolean
    Dim SecTotal As Double
    Dim OPDCollection As Double
    
    OPDCollection = 0
    SecTotal = 0
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "OPD Collection"
    
    
    SubTotal = 0
    
        If rsCat.State = 1 Then rsCat.Close
        temSql = "SELECT ServiceCategory, ServiceCategoryID From dbo.tblServiceCategory Where (ForOPD = 1) ORDER BY ServiceCategory"
        rsCat.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While rsCat.EOF = False
            temCatID = rsCat!ServiceCategoryID
            
            
            If rsIsSC.State = 1 Then rsIsSC.Close
            temSql = "SELECT COUNT(ServiceSubcategory) AS CountOfServiceSubcategory FROM dbo.tblServiceSubcategory WHERE (ServiceCategoryID = " & temCatID & ")"
            rsIsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If rsIsSC.RecordCount > 0 Then
                If IsNull(rsIsSC!CountOfServiceSubcategory) = False Then
                    If rsIsSC!CountOfServiceSubcategory > 1 Then
                        SCExists = True
                    Else
                        SCExists = False
                    End If
                Else
                    SCExists = False
                End If
            Else
                SCExists = False
            End If
            rsIsSC.Close
            
            
            SubTotal = 0
            
            With rsTem
                If .State = 1 Then .Close
                temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge " & _
                            "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.OPDBillID " & _
                            "WHERE (dbo.tblPatientService.Deleted = 0) AND  (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                            "(dbo.tblIncomeBill.IsOPDBill = 1)  AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    If IsNull(!SumOfHospitalCharge) = False Then
                        SubTotal = !SumOfHospitalCharge
                    End If
                    SubTotal = Val(Format(!SumOfHospitalCharge, "0.00"))
                End If
                .Close
            End With
            
            If SubTotal <> 0 Then
            
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.Text = rsCat!ServiceCategory
                
                If SCExists = True Then
                
                    If rsSC.State = 1 Then rsSC.Close
                    temSql = "SELECT ServiceSubcategory, ServiceSubcategoryID FROM dbo.tblServiceSubcategory WHERE (ServiceCategoryID = " & temCatID & ")"
                    rsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    While rsSC.EOF = False
                        temSCID = rsSC!ServicesubcategoryID
                    
                        temTotal = 0
                    
                        With rsTem
                            If .State = 1 Then .Close
                            temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, Count(dbo.tblPatientService.HospitalCharge) AS SCount  " & _
                                        "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.OPDBillID " & _
                                        "WHERE    (dbo.tblPatientService.Deleted = 0) AND    (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                        "(dbo.tblIncomeBill.IsOPDBill = 1) AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceSubCategoryID = " & temSCID & ") "
                            
                            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                            If .RecordCount > 0 Then
                                If IsNull(!SumOfHospitalCharge) = False Then
                                    temTotal = !SumOfHospitalCharge
                                End If
                                temCount = Val(Format(!SCount, "0"))
                            End If
                            .Close
                        End With
                    
                        If temTotal <> 0 Then
                            gridSummery.Rows = gridSummery.Rows + 1
                            gridSummery.Row = gridSummery.Rows - 1
                            gridSummery.Col = 1
                            gridSummery.Text = Space(5) & rsSC!ServiceSubcategory
                            
                            gridSummery.Col = 2
                            gridSummery.Text = temCount
                            
                            gridSummery.Col = 3
                            gridSummery.Text = Format(temTotal, "###0.00")
                            
                        End If
                    
                        rsSC.MoveNext
                    Wend
                    rsSC.Close
                    
'                    gridSummery.Col = 4
'                    gridSummery.Text = Format(SubTotal, "###0.00")
                    
                    
                    With rsTem
                        
                        SubTotal = 0
                        
                        
                                            
                        If .State = 1 Then .Close
                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.OPDBillID  " & _
                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                    "(dbo.tblIncomeBill.IsOPDBill = 1) AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                        
                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                        If .RecordCount > 0 Then
                            If IsNull(!SumOfHospitalCharge) = False Then
                                If IsNull(!SumOfHospitalCharge) = False Then
                                    SubTotal = !SumOfHospitalCharge
                                End If
                                
                                
                                SubTotal = !SumOfHospitalCharge

                                
                                gridSummery.Col = 4
                                gridSummery.Text = Format(SubTotal, "###0.00")
                                
                                SecTotal = SecTotal + SubTotal
                                OPDCollection = OPDCollection + SubTotal
                                
                            End If
                        End If
                        .Close
                    End With
                    
                Else
                
                    
                    SecTotal = 0
                
                    SubTotal = 0
                    
                    With rsTem
                        If .State = 1 Then .Close
                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, COUNT(dbo.tblPatientService.HospitalCharge) AS SCount  " & _
                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.OPDBillID " & _
                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND (dbo.tblIncomeBill.Cancelled = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                    "(dbo.tblIncomeBill.IsOPDBill = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                        
                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                        If .RecordCount > 0 Then
                            
                            If IsNull(!SumOfHospitalCharge) = False Then
                                SubTotal = !SumOfHospitalCharge
                            End If
                            
                                
                            SecTotal = SecTotal + SubTotal
                            
                            OPDCollection = OPDCollection + SubTotal
                            
                            gridSummery.Col = 3
                            gridSummery.Text = Format(SubTotal, "###0.00")
                            
                            gridSummery.Col = 2
                            gridSummery.Text = Val(Format(!SCount, "0"))
                            
                            gridSummery.Col = 4
                            gridSummery.Text = Format(SecTotal, "###0.00")
                            
                            
                        End If
                        .Close
                    End With
                    
                End If
            
            End If
            
            rsCat.MoveNext
        Wend
        rsCat.Close

    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 5
    gridSummery.Text = Format(OPDCollection, "###0.00")

    CollectionForTheDay = CollectionForTheDay + OPDCollection

End Sub



'Private Sub RCollectionIncome()
'    Dim SubTotalHH As Double
'    Dim rsTem As New ADODB.Recordset
'    Dim rsCat As New ADODB.Recordset
'    Dim rsIsSC As New ADODB.Recordset
'    Dim rsSC As New ADODB.Recordset
'    Dim temCatID As Long
'    Dim temSCID As Long
'    Dim SCExists As Boolean
'    Dim SecTotal As Double
'
'    SecTotal = 0
'
'    gridSummery.Rows = gridSummery.Rows + 1
'    gridSummery.Row = gridSummery.Rows - 1
'    gridSummery.Col = 0
'    gridSummery.Text = "Roentgents Collection"
'
'
'    SubTotalHH = 0
'        If rsCat.State = 1 Then rsCat.Close
'        temSql = "SELECT ServiceCategory, ServiceCategoryID From dbo.tblServiceCategory Where (ForR = 1) ORDER BY ServiceCategory"
'        rsCat.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While rsCat.EOF = False
'            temCatID = rsCat!ServiceCategoryID
'            SubTotalHH = 0
'            If rsIsSC.State = 1 Then rsIsSC.Close
'            temSql = "SELECT COUNT(ServiceSubcategory) AS CountOfServiceSubcategory FROM dbo.tblServiceSubcategory WHERE (ServiceCategoryID = " & temCatID & ")"
'            rsIsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'            If rsIsSC.RecordCount > 0 Then
'                If IsNull(rsIsSC!CountOfServiceSubcategory) = False Then
'                    If rsIsSC!CountOfServiceSubcategory > 1 Then
'                        SCExists = True
'                    Else
'                        SCExists = False
'                    End If
'                Else
'                    SCExists = False
'                End If
'            Else
'                SCExists = False
'            End If
'            rsIsSC.Close
'
'
'            SubTotalHH = 0
'            With rsTem
'                If .State = 1 Then .Close
'                temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
'                            "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID " & _
'                            "WHERE (dbo.tblPatientService.Deleted = 0) AND  (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
'                            "(dbo.tblIncomeBill.IsRBill = 1)  AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
'
'                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                If .RecordCount > 0 Then
'                    If IsNull(!SumOfHospitalCharge) = False Then
'                        SubTotalHH = !SumOfHospitalCharge
'                    End If
'                End If
'                .Close
'            End With
'
'            If SubTotalHH <> 0 Then
'
'                gridSummery.Rows = gridSummery.Rows + 1
'                gridSummery.Row = gridSummery.Rows - 1
'                gridSummery.Col = 1
'                gridSummery.Text = rsCat!ServiceCategory
'
'                If SCExists = True Then
'
'                    If rsSC.State = 1 Then rsSC.Close
'                    temSql = "SELECT ServiceSubcategory, ServiceSubcategoryID FROM dbo.tblServiceSubcategory WHERE (ServiceCategoryID = " & temCatID & ")"
'                    rsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                    While rsSC.EOF = False
'                        temSCID = rsSC!ServiceSubcategoryID
'
'                        SubTotalHH = 0
'
'                        With rsTem
'                            If .State = 1 Then .Close
'                            temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
'                                        "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID " & _
'                                        "WHERE    (dbo.tblPatientService.Deleted = 0) AND    (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
'                                        "(dbo.tblIncomeBill.IsRBill = 1) AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceSubCategoryID = " & temSCID & ") "
'
'                            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                            If .RecordCount > 0 Then
'                                If IsNull(!SumOfHospitalCharge) = False Then
'                                    SubTotalHH = !SumOfHospitalCharge
'                                    SecTotal = SecTotal + SubTotalHH
'                                End If
'                            End If
'                            .Close
'                        End With
'
'                        If SubTotalHH <> 0 Then
'                            gridSummery.Rows = gridSummery.Rows + 1
'                            gridSummery.Row = gridSummery.Rows - 1
'                            gridSummery.Col = 1
'                            gridSummery.Text = Space(5) & rsSC!ServiceSubCategory
'                            gridSummery.Col = 5
'                            gridSummery.Text = Format(SubTotalHH, "###0.00")
'                        End If
'
'                        rsSC.MoveNext
'                    Wend
'                    rsSC.Close
'                    With rsTem
'
'                        SubTotalHH = 0
'
'                        If .State = 1 Then .Close
'                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
'                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID  " & _
'                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
'                                    "(dbo.tblIncomeBill.IsRBill = 1) AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
'
'                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                        If .RecordCount > 0 Then
'                            If IsNull(!SumOfHospitalCharge) = False Then
'                                SubTotalHH = !SumOfHospitalCharge
'                                gridSummery.Col = 5
'                                gridSummery.Text = Format(SubTotalHH, "###0.00")
'                            End If
'                        End If
'                        .Close
'                    End With
'
'                Else
'
'                    SubTotalHH = 0
'                    With rsTem
'                        If .State = 1 Then .Close
'                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
'                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID " & _
'                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND (dbo.tblIncomeBill.Cancelled = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
'                                    "(dbo.tblIncomeBill.IsRBill = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
'
'                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                        If .RecordCount > 0 Then
'                            If IsNull(!SumOfHospitalCharge) = False Then
'                                SubTotalHH = !SumOfHospitalCharge
'                                SecTotal = SecTotal + SubTotalHH
'                                gridSummery.Col = 5
'                                gridSummery.Text = Format(SubTotalHH, "###0.00")
'                            End If
'                        End If
'                        .Close
'                    End With
'
'                End If
'
'            End If
'
'            rsCat.MoveNext
'        Wend
'        rsCat.Close
'
'    gridSummery.Rows = gridSummery.Rows + 1
'    gridSummery.Row = gridSummery.Rows - 1
'    gridSummery.Col = 7
'    gridSummery.Text = Format(SecTotal, "###0.00")
'
'    CollectionForTheDay = CollectionForTheDay + SecTotal
'
'End Sub


Private Sub LabCollectionIncome()
    
    Dim SubTotalH As Double
    Dim SubTotal As Double
    Dim SubTotalP As Double
    
    Dim rsTem As New ADODB.Recordset
    Dim rsCat As New ADODB.Recordset
    Dim rsIsSC As New ADODB.Recordset
    Dim rsSC As New ADODB.Recordset
    Dim temCatID As Long
    Dim temSCID As Long
    Dim SCExists As Boolean
    Dim SecTotal As Double
    
    SecTotal = 0
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "Lab Collection"
    
    
    SubTotal = 0
    SubTotalH = 0
    SubTotalP = 0
    
        If rsCat.State = 1 Then rsCat.Close
        temSql = "SELECT ServiceCategory, ServiceCategoryID From dbo.tblServiceCategory Where (ForLab = 1) ORDER BY ServiceCategory"
        rsCat.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While rsCat.EOF = False
            temCatID = rsCat!ServiceCategoryID
            SubTotalH = 0
            If rsIsSC.State = 1 Then rsIsSC.Close
            temSql = "SELECT COUNT(ServiceSubcategory) AS CountOfServiceSubcategory FROM dbo.tblServiceSubcategory WHERE (ServiceCategoryID = " & temCatID & ")"
            rsIsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If rsIsSC.RecordCount > 0 Then
                If IsNull(rsIsSC!CountOfServiceSubcategory) = False Then
                    If rsIsSC!CountOfServiceSubcategory > 1 Then
                        SCExists = True
                    Else
                        SCExists = False
                    End If
                Else
                    SCExists = False
                End If
            Else
                SCExists = False
            End If
            rsIsSC.Close
            
            
            SubTotal = 0
            SubTotalH = 0
            SubTotalP = 0
            
            With rsTem
                If .State = 1 Then .Close
                temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge " & _
                            "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.LabBillID " & _
                            "WHERE (dbo.tblPatientService.Deleted = 0) AND  (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                            "(dbo.tblIncomeBill.IsLabBill = 1)  AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    If IsNull(!SumOfHospitalCharge) = False Then
                        SubTotalH = !SumOfHospitalCharge
                    End If
                    If IsNull(!SumOfProfessionalCharge) = False Then
                        SubTotalP = !SumOfProfessionalCharge
                    End If
                    If IsNull(!SumOfCharge) = False Then
                        SubTotal = !SumOfCharge
                    End If
                End If
                .Close
            End With
            
            If SubTotal <> 0 Then
            
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.Text = rsCat!ServiceCategory
                
                If SCExists = True Then
                
                    If rsSC.State = 1 Then rsSC.Close
                    temSql = "SELECT ServiceSubcategory, ServiceSubcategoryID FROM dbo.tblServiceSubcategory WHERE (ServiceCategoryID = " & temCatID & ")"
                    rsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    While rsSC.EOF = False
                        temSCID = rsSC!ServicesubcategoryID
                    
                        SubTotal = 0
                        SubTotalH = 0
                        SubTotalP = 0
                    
                        With rsTem
                            If .State = 1 Then .Close
                            temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
                                        "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.LabBillID " & _
                                        "WHERE    (dbo.tblPatientService.Deleted = 0) AND    (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                        "(dbo.tblIncomeBill.IsLabBill = 1) AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceSubCategoryID = " & temSCID & ") "
                            
                            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                            If .RecordCount > 0 Then
                                If IsNull(!SumOfHospitalCharge) = False Then
                                    SubTotalH = !SumOfHospitalCharge
                                End If
                                If IsNull(!SumOfProfessionalCharge) = False Then
                                    SubTotalP = !SumOfProfessionalCharge
                                End If
                                If IsNull(!SumOfCharge) = False Then
                                    SubTotal = !SumOfCharge
                                End If
                            End If
                            .Close
                        End With
                    
                        If SubTotal <> 0 Then
                            gridSummery.Rows = gridSummery.Rows + 1
                            gridSummery.Row = gridSummery.Rows - 1
                            gridSummery.Col = 1
                            gridSummery.Text = Space(5) & rsSC!ServiceSubcategory
                            
                            gridSummery.Col = 2
                            gridSummery.Text = Format(SubTotalH, "###0.00")
                            gridSummery.Col = 3
                            gridSummery.Text = Format(SubTotalP, "###0.00")
                            gridSummery.Col = 4
                            gridSummery.Text = Format(SubTotal, "###0.00")
                            
                        End If
                    
                        rsSC.MoveNext
                    Wend
                    rsSC.Close
                    
                    
                    
                    With rsTem
                        
                        SubTotal = 0
                        SubTotalH = 0
                        SubTotalP = 0
                                            
                        If .State = 1 Then .Close
                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.LabBillID  " & _
                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                    "(dbo.tblIncomeBill.IsLabBill = 1) AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                        
                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                        If .RecordCount > 0 Then
                            If IsNull(!SumOfHospitalCharge) = False Then
                                If IsNull(!SumOfHospitalCharge) = False Then
                                    SubTotalH = !SumOfHospitalCharge
                                End If
                                If IsNull(!SumOfProfessionalCharge) = False Then
                                    SubTotalP = !SumOfProfessionalCharge
                                End If
                                If IsNull(!SumOfCharge) = False Then
                                    SubTotal = !SumOfCharge
                                End If
                                gridSummery.Col = 3
                                gridSummery.Text = Format(SubTotal, "###0.00")
                                SecTotal = SecTotal + SubTotal
                            End If
                        End If
                        .Close
                    End With
                    
                Else
                
                    SubTotalH = 0
                    SubTotalP = 0
                    SubTotal = 0
                    
                    With rsTem
                        If .State = 1 Then .Close
                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.LabBillID " & _
                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND (dbo.tblIncomeBill.Cancelled = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                    "(dbo.tblIncomeBill.IsLabBill = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                        
                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                        If .RecordCount > 0 Then
                            
                            If IsNull(!SumOfHospitalCharge) = False Then
                                SubTotalH = !SumOfHospitalCharge
                            End If
                            If IsNull(!SumOfProfessionalCharge) = False Then
                                SubTotalP = !SumOfProfessionalCharge
                            End If
                            If IsNull(!SumOfCharge) = False Then
                                SubTotal = !SumOfCharge
                            End If
                                
                            SecTotal = SecTotal + SubTotal
                            
                            
                            gridSummery.Col = 2
                            gridSummery.Text = Format(SubTotalH, "###0.00")
                            gridSummery.Col = 3
                            gridSummery.Text = Format(SubTotalP, "###0.00")
                            gridSummery.Col = 4
                            gridSummery.Text = Format(SubTotal, "###0.00")
                            
                            gridSummery.Col = 3
                            gridSummery.Text = Format(SecTotal, "###0.00")
                            
                            
                        End If
                        .Close
                    End With
                    
                End If
            
            End If
            
            rsCat.MoveNext
        Wend
        rsCat.Close

    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 5
    gridSummery.Text = Format(SecTotal, "###0.00")

    CollectionForTheDay = CollectionForTheDay + SecTotal

End Sub


Private Sub RCollectionIncome()
    
    Dim SubTotal As Double
    Dim temCount As Long
    Dim temTotal As Double
    
    Dim rsTem As New ADODB.Recordset
    Dim rsCat As New ADODB.Recordset
    Dim rsIsSC As New ADODB.Recordset
    Dim rsSC As New ADODB.Recordset
    Dim temCatID As Long
    Dim temSCID As Long
    Dim SCExists As Boolean
    Dim SecTotal As Double
    
    SecTotal = 0
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "Ron. Collection"
    
    
    SubTotal = 0
    
        
        
        If rsCat.State = 1 Then rsCat.Close
        temSql = "SELECT ServiceCategory, ServiceCategoryID From dbo.tblServiceCategory Where (ForR = 1) ORDER BY ServiceCategory"
        rsCat.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While rsCat.EOF = False
            temCatID = rsCat!ServiceCategoryID
            If rsIsSC.State = 1 Then rsIsSC.Close
            temSql = "SELECT COUNT(ServiceSubcategory) AS CountOfServiceSubcategory FROM dbo.tblServiceSubcategory WHERE (ServiceCategoryID = " & temCatID & ")"
            rsIsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
            If rsIsSC.RecordCount > 0 Then
                If IsNull(rsIsSC!CountOfServiceSubcategory) = False Then
                    If rsIsSC!CountOfServiceSubcategory > 1 Then
                        SCExists = True
                    Else
                        SCExists = False
                    End If
                Else
                    SCExists = False
                End If
            Else
                SCExists = False
            End If
            rsIsSC.Close
            
            
            SubTotal = 0
            
            With rsTem
                If .State = 1 Then .Close

                temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, Count (dbo.tblPatientService.HospitalCharge) AS SCount " & _
                            "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID " & _
                            "WHERE (dbo.tblPatientService.Deleted = 0) AND  (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                            "(dbo.tblIncomeBill.IsRBill = 1)  AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                                
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                SubTotal = 0
                If .RecordCount > 0 Then
                    If IsNull(!SumOfHospitalCharge) = False Then
                        SubTotal = !SumOfHospitalCharge
                    End If
                    
                    
                
                
                End If
                .Close
            
            
            
            
            
            If .State = 1 Then .Close
                temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, Count (dbo.tblPatientService.HospitalCharge) AS SCount " & _
                            "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID " & _
                            "WHERE (dbo.tblPatientService.Deleted = 0) AND  (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CancelledDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                            "(dbo.tblIncomeBill.IsRBill = 1)  AND (dbo.tblIncomeBill.Cancelled = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    If IsNull(!SumOfHospitalCharge) = False Then
                        SubTotal = SubTotal - !SumOfHospitalCharge
                    End If
                    
                End If
                .Close
            
            
            
            
            
            
            
            End With
            
            If SubTotal <> 0 Then
            
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.Text = rsCat!ServiceCategory
                
                If SCExists = True Then
                
                    If rsSC.State = 1 Then rsSC.Close
                    temSql = "SELECT ServiceSubcategory, ServiceSubcategoryID FROM dbo.tblServiceSubcategory WHERE (ServiceCategoryID = " & temCatID & ")"
                    rsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                    While rsSC.EOF = False
                        temSCID = rsSC!ServicesubcategoryID
                    
                        temTotal = 0
                    
                        With rsTem
                            If .State = 1 Then .Close
                            temSql = "SELECT  SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge,  COUNT(dbo.tblPatientService.HospitalCharge) AS SCount  " & _
                                        "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID " & _
                                        "WHERE    (dbo.tblPatientService.Deleted = 0) AND    (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                        "(dbo.tblIncomeBill.IsRBill = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceSubCategoryID = " & temSCID & ") "
                            
                            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                            If .RecordCount > 0 Then
                                If IsNull(!SumOfHospitalCharge) = False Then
                                    temTotal = !SumOfHospitalCharge
                                End If
                                temCount = Val(Format(!SCount, "0"))

                                
                            End If
                            .Close
                        
                            If .State = 1 Then .Close
                            temSql = "SELECT  SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge,  COUNT(dbo.tblPatientService.HospitalCharge) AS SCount  " & _
                                        "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID " & _
                                        "WHERE    (dbo.tblPatientService.Deleted = 0) AND    (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CancelledDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                        "(dbo.tblIncomeBill.IsRBill = 1) AND (dbo.tblIncomeBill.Cancelled = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceSubCategoryID = " & temSCID & ") "
                            
                            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                            If .RecordCount > 0 Then
                                If IsNull(!SumOfHospitalCharge) = False Then
                                    temTotal = temTotal - !SumOfHospitalCharge
                                End If
                                temCount = temCount - Val(Format(!SCount, "0"))
                            End If
                            .Close
                                                
                        
                        
                        End With
                    
                        If temTotal <> 0 Then
                            gridSummery.Rows = gridSummery.Rows + 1
                            gridSummery.Row = gridSummery.Rows - 1
                            gridSummery.Col = 1
                            gridSummery.Text = Space(5) & rsSC!ServiceSubcategory
                            gridSummery.Col = 2
                            gridSummery.Text = temCount
                            gridSummery.Col = 3
                            gridSummery.Text = Format(temTotal, "###0.00")
                            
                        End If
                    
                        rsSC.MoveNext
                    Wend
                    rsSC.Close
                    
                    
                    
                    With rsTem
                        
                        SubTotal = 0
                                            
                        If .State = 1 Then .Close
                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID  " & _
                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                    "(dbo.tblIncomeBill.IsRBill = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                        
                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                        SubTotal = 0
                        If .RecordCount > 0 Then
                            If IsNull(!SumOfHospitalCharge) = False Then
                                If IsNull(!SumOfHospitalCharge) = False Then
                                    SubTotal = !SumOfHospitalCharge
                                End If
                                
'                                gridSummery.Col = 4
'                                gridSummery.Text = Format(SubTotal, "###0.00")
'                                SecTotal = SecTotal + SubTotal
                            End If
                        End If
                        .Close
                        
                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID  " & _
                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CancelledDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                    "(dbo.tblIncomeBill.IsRBill = 1) AND (dbo.tblIncomeBill.Cancelled = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                        
                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                        If .RecordCount > 0 Then
                            If IsNull(!SumOfHospitalCharge) = False Then
                                If IsNull(!SumOfHospitalCharge) = False Then
                                    SubTotal = SubTotal - !SumOfHospitalCharge
                                End If
                                
'                                gridSummery.Col = 4
'                                gridSummery.Text = Format(SubTotal, "###0.00")
'                                SecTotal = SecTotal + SubTotal
                            End If
                        End If
                        .Close
                        
                        
                        gridSummery.Col = 4
                        gridSummery.Text = Format(SubTotal, "###0.00")
                        SecTotal = SecTotal + SubTotal
                
                    End With
                    
                Else
                
                    SubTotal = 0
                    
                    With rsTem
                        If .State = 1 Then .Close
                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, Count(dbo.tblPatientService.HospitalCharge) AS SCount  " & _
                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID " & _
                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                    "(dbo.tblIncomeBill.IsRBill = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                        
                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                        SubTotal = 0
                        temCount = 0
                        If .RecordCount > 0 Then
                            
                            If IsNull(!SumOfHospitalCharge) = False Then
                                SubTotal = !SumOfHospitalCharge
                            End If
                            temCount = Val(Format(!SCount, "0"))
'                            SecTotal = SecTotal + SubTotal
'                            gridSummery.Col = 2
'                            gridSummery.Text = Count
'                            gridSummery.Col = 3
'                            gridSummery.Text = Format(SubTotal, "###0.00")
'                            gridSummery.Col = 4
'                            gridSummery.Text = Format(SubTotal, "###0.00")
                            
                            
                        End If
                        .Close
                    
                        If .State = 1 Then .Close
                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, Count(dbo.tblPatientService.HospitalCharge) AS SCount  " & _
                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.RBillID " & _
                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND (dbo.tblIncomeBill.Cancelled = 1) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CancelledDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
                                    "(dbo.tblIncomeBill.IsRBill = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
                        
                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                        If .RecordCount > 0 Then
                            
                            If IsNull(!SumOfHospitalCharge) = False Then
                                SubTotal = SubTotal - !SumOfHospitalCharge
                            End If
                            temCount = temCount - Val(Format(!SCount, "0"))
                            
                            
                        End If
                        .Close
                    
                        SecTotal = SecTotal + SubTotal
                        gridSummery.Col = 2
                        gridSummery.Text = Count
                        gridSummery.Col = 3
                        gridSummery.Text = Format(SubTotal, "###0.00")
                        gridSummery.Col = 4
                        gridSummery.Text = Format(SubTotal, "###0.00")
                    
                    
                    
                    End With
                    
                End If
            
            End If
            
            rsCat.MoveNext
        Wend
        rsCat.Close

    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 5
    gridSummery.Text = Format(SecTotal, "###0.00")

    CollectionForTheDay = CollectionForTheDay + SecTotal

End Sub


Private Sub MTCollectionIncome()
'    Dim SubTotalHH As Double
'    Dim rsTem As New ADODB.Recordset
'    Dim rsCat As New ADODB.Recordset
'    Dim rsIsSC As New ADODB.Recordset
'    Dim rsSC As New ADODB.Recordset
'    Dim temCatID As Long
'    Dim temSCID As Long
'    Dim SCExists As Boolean
'    Dim SecTotal As Double
'
'    SecTotal = 0
'
'    gridSummery.Rows = gridSummery.Rows + 1
'    gridSummery.Row = gridSummery.Rows - 1
'    gridSummery.Col = 0
'    gridSummery.Text = "Medical Tests Collection"
'
'
'    SubTotalHH = 0
'        If rsCat.State = 1 Then rsCat.Close
'        temSql = "SELECT ServiceCategory, ServiceCategoryID From dbo.tblServiceCategory Where (ForMT = 1) ORDER BY ServiceCategory"
'        rsCat.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        While rsCat.EOF = False
'            temCatID = rsCat!ServiceCategoryID
'            SubTotalHH = 0
'            If rsIsSC.State = 1 Then rsIsSC.Close
'            temSql = "SELECT COUNT(ServiceSubcategory) AS CountOfServiceSubcategory FROM dbo.tblServiceSubcategory WHERE (ServiceCategoryID = " & temCatID & ")"
'            rsIsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'            If rsIsSC.RecordCount > 0 Then
'                If IsNull(rsIsSC!CountOfServiceSubcategory) = False Then
'                    If rsIsSC!CountOfServiceSubcategory > 1 Then
'                        SCExists = True
'                    Else
'                        SCExists = False
'                    End If
'                Else
'                    SCExists = False
'                End If
'            Else
'                SCExists = False
'            End If
'            rsIsSC.Close
'
'
'            SubTotalHH = 0
'            With rsTem
'                If .State = 1 Then .Close
'                temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
'                            "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.MedicalTestBillID " & _
'                            "WHERE (dbo.tblPatientService.Deleted = 0) AND  (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
'                            "(dbo.tblIncomeBill.IsMedicalTestBill = 1)  AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
'
'                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                If .RecordCount > 0 Then
'                    If IsNull(!SumOfHospitalCharge) = False Then
'                        SubTotalHH = !SumOfHospitalCharge
'                    End If
'                End If
'                .Close
'            End With
'
'            If SubTotalHH <> 0 Then
'
'                gridSummery.Rows = gridSummery.Rows + 1
'                gridSummery.Row = gridSummery.Rows - 1
'                gridSummery.Col = 1
'                gridSummery.Text = rsCat!ServiceCategory
'
'                If SCExists = True Then
'
'                    If rsSC.State = 1 Then rsSC.Close
'                    temSql = "SELECT ServiceSubcategory, ServiceSubcategoryID FROM dbo.tblServiceSubcategory WHERE (ServiceCategoryID = " & temCatID & ")"
'                    rsSC.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                    While rsSC.EOF = False
'                        temSCID = rsSC!ServiceSubcategoryID
'
'                        SubTotalHH = 0
'
'                        With rsTem
'                            If .State = 1 Then .Close
'                            temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
'                                        "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.MedicalTestBillID " & _
'                                        "WHERE    (dbo.tblPatientService.Deleted = 0) AND    (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
'                                        "(dbo.tblIncomeBill.IsMedicalTestBill = 1) AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceSubCategoryID = " & temSCID & ") "
'
'                            .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                            If .RecordCount > 0 Then
'                                If IsNull(!SumOfHospitalCharge) = False Then
'                                    SubTotalHH = !SumOfHospitalCharge
'                                    SecTotal = SecTotal + SubTotalHH
'                                End If
'                            End If
'                            .Close
'                        End With
'
'                        If SubTotalH <> 0 Then
'                            gridSummery.Rows = gridSummery.Rows + 1
'                            gridSummery.Row = gridSummery.Rows - 1
'                            gridSummery.Col = 1
'                            gridSummery.Text = Space(5) & rsSC!ServiceSubCategory
'                            gridSummery.Col = 5
'                            gridSummery.Text = Format(SubTotalH, "###0.00")
'                        End If
'
'                        rsSC.MoveNext
'                    Wend
'                    rsSC.Close
'                    With rsTem
'
'                        SubTotalH = 0
'
'                        If .State = 1 Then .Close
'                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
'                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.MedicalTestBillID  " & _
'                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
'                                    "(dbo.tblIncomeBill.IsMedicalTestBill = 1) AND (dbo.tblIncomeBill.Cancelled = 0) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
'
'                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                        If .RecordCount > 0 Then
'                            If IsNull(!SumOfHospitalCharge) = False Then
'                                SubTotalH = !SumOfHospitalCharge
'                                gridSummery.Col = 5
'                                gridSummery.Text = Format(SubTotalH, "###0.00")
'                            End If
'                        End If
'                        .Close
'                    End With
'
'                Else
'
'                    SubTotalH = 0
'                    With rsTem
'                        If .State = 1 Then .Close
'                        temSql = "SELECT     SUM(dbo.tblPatientService.HospitalCharge) AS SumOfHospitalCharge, SUM(dbo.tblPatientService.ProfessionalCharge) AS SumOfProfessionalCharge, SUM(dbo.tblPatientService.Charge) AS SumOfCharge  " & _
'                                    "FROM         dbo.tblIncomeBill RIGHT OUTER JOIN dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.MedicalTestBillID " & _
'                                    "WHERE   (dbo.tblPatientService.Deleted = 0) AND (dbo.tblIncomeBill.Cancelled = 0) AND     (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.CompletedDate = CONVERT(DATETIME, '" & Format(dtpDate.Value, "dd MMMM yyyy") & "', 102)) AND " & _
'                                    "(dbo.tblIncomeBill.IsMedicalTestBill = 1) AND (dbo.tblIncomeBill.PaymentMethodID <> 4) AND (dbo.tblPatientService.ServiceCategoryID = " & temCatID & ") "
'
'                        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'                        If .RecordCount > 0 Then
'                            If IsNull(!SumOfHospitalCharge) = False Then
'                                SubTotalH = !SumOfHospitalCharge
'                                SecTotal = SecTotal + SubTotalH
'                                gridSummery.Col = 5
'                                gridSummery.Text = Format(SubTotalH, "###0.00")
'                            End If
'                        End If
'                        .Close
'                    End With
'
'                End If
'
'            End If
'
'            rsCat.MoveNext
'        Wend
'        rsCat.Close
'
'    gridSummery.Rows = gridSummery.Rows + 1
'    gridSummery.Row = gridSummery.Rows - 1
'    gridSummery.Col = 7
'    gridSummery.Text = Format(SecTotal, "###0.00")


End Sub


Private Sub BHTCollectionPayments()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "BHT Collection"
    
    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT     dbo.tblIncomeBill.NetTotal, dbo.tblBHT.BHT, dbo.tblPatientMainDetails.FirstName FROM         dbo.tblIncomeBill LEFT OUTER JOIN                       dbo.tblBHT ON dbo.tblIncomeBill.BHTID = dbo.tblBHT.BHTID LEFT OUTER JOIN                       dbo.tblPatientMainDetails ON dbo.tblBHT.PatientID = dbo.tblPatientMainDetails.PatientID WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.IsInwardPaymentBill)=1) AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.PaymentMethodID)<>4))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !NetTotal
                
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.CellAlignment = 1
                gridSummery.Text = Format(!BHT, "") & " - " & Format(!FirstName, "")
                
                gridSummery.Col = 3
                gridSummery.Text = Format(!NetTotal, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT     dbo.tblIncomeReturnBill.ReturnValue, dbo.tblBHT.BHT, dbo.tblPatientMainDetails.FirstName " & _
                        "FROM         dbo.tblIncomeReturnBill LEFT OUTER JOIN " & _
                        "dbo.tblBHT ON dbo.tblIncomeReturnBill.BHTID = dbo.tblBHT.BHTID LEFT OUTER JOIN " & _
                        "dbo.tblPatientMainDetails ON dbo.tblBHT.PatientID = dbo.tblPatientMainDetails.PatientID " & _
                        "Where (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)<>4)  AND ((tblBHT.IsBHT)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl - !ReturnValue
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.CellAlignment = 1
                gridSummery.Text = "Repaid to " & Format(!BHT, "") & " - " & Format(!FirstName, "")
                gridSummery.Col = 3
                gridSummery.Text = "(" & Format(!ReturnValue, "###0.00") & ")"
                .MoveNext
            Wend
        End If
        .Close
    End With
'    gridSummery.Rows = gridSummery.Rows + 1
'    gridSummery.Row = gridSummery.Rows - 1
'    gridSummery.Col = 0
'    gridSummery.Text = "BHT Collection Payments"
    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")
    CollectionForTheDay = CollectionForTheDay + temDbl
End Sub

Private Sub GSBCollectionPayments()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "GSB Collection"
    
    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT     dbo.tblIncomeBill.NetTotal, dbo.tblBHT.BHT, dbo.tblPatientMainDetails.FirstName FROM         dbo.tblIncomeBill LEFT OUTER JOIN                       dbo.tblBHT ON dbo.tblIncomeBill.BHTID = dbo.tblBHT.BHTID LEFT OUTER JOIN                       dbo.tblPatientMainDetails ON dbo.tblBHT.PatientID = dbo.tblPatientMainDetails.PatientID WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.IsGSBill)=1) AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.PaymentMethodID)<>4))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !NetTotal
                
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.CellAlignment = 1
                gridSummery.Text = Format(!BHT, "") & " - " & Format(!FirstName, "")
                
                gridSummery.Col = 3
                gridSummery.Text = Format(!NetTotal, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT     dbo.tblIncomeReturnBill.ReturnValue, dbo.tblBHT.BHT, dbo.tblPatientMainDetails.FirstName " & _
                        "FROM         dbo.tblIncomeReturnBill LEFT OUTER JOIN " & _
                        "dbo.tblBHT ON dbo.tblIncomeReturnBill.BHTID = dbo.tblBHT.BHTID LEFT OUTER JOIN " & _
                        "dbo.tblPatientMainDetails ON dbo.tblBHT.PatientID = dbo.tblPatientMainDetails.PatientID " & _
                        "Where (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)<>4)  AND ((tblBHT.IsGSB)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl - !ReturnValue
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.CellAlignment = 1
                gridSummery.Text = "Repaid to " & Format(!BHT, "") & " - " & Format(!FirstName, "")
                gridSummery.Col = 3
                gridSummery.Text = "(" & Format(!ReturnValue, "###0.00") & ")"
                .MoveNext
            Wend
        End If
        .Close
    End With
'    gridSummery.Rows = gridSummery.Rows + 1
'    gridSummery.Row = gridSummery.Rows - 1
'    gridSummery.Col = 0
'    gridSummery.Text = "GSB Collection Payments"
    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")
    CollectionForTheDay = CollectionForTheDay + temDbl
End Sub

Private Sub PharmacyCollectionIncome()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeBill.NetTotal) AS SumOfNetTotal From tblIncomeBill WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.IsPharmacyBill)=1)AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.PaymentMethodID)<>4))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumofNetTotal) = False Then
                temDbl = !SumofNetTotal
            End If
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue FROM tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)<>4) AND ((tblIncomeBill.IsPharmacyBill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfReturnValue) = False Then
                temDbl = temDbl - !SumOfReturnValue
            End If
        End If
        .Close
    End With
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "Pharmacy Collection"
    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")
    CollectionForTheDay = CollectionForTheDay + temDbl
End Sub

Private Sub AgentCollectionIncome()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "Agent Collection"
    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblIncomeBill.NetTotal, dbo.tblIncomeBill.DisplayBillID , dbo.tblIncomeBill.IncomeBillID, dbo.tblAgent.Agent, dbo.tblAgent.Code FROM dbo.tblIncomeBill LEFT OUTER JOIN dbo.tblAgent ON dbo.tblIncomeBill.AgentID = dbo.tblAgent.AgentID WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.IsAgentBill)=1)AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.PaymentMethodID)<>4))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !NetTotal
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.CellAlignment = 1
                gridSummery.Text = !DisplayBillID & " " & Format(!Agent, "") & " (" & Format(!Code, "") & ")"
                gridSummery.Col = 3
                gridSummery.Text = Format(!NetTotal, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblIncomeReturnBill.ReturnValue, dbo.tblAgent.Agent FROM dbo.tblAgent RIGHT OUTER JOIN dbo.tblIncomeBill ON dbo.tblAgent.AgentID = dbo.tblIncomeBill.AgentID RIGHT OUTER JOIN dbo.tblIncomeReturnBill ON dbo.tblIncomeBill.IncomeBillID = dbo.tblIncomeReturnBill.IncomeBillID WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)<>4) AND ((tblIncomeBill.IsAgentBill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl - !ReturnValue
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.CellAlignment = 1
                gridSummery.Text = Format(!Agent, "")
                gridSummery.Col = 3
                gridSummery.Text = "(" & Format(!ReturnValue, "###0.00") & ")"
                .MoveNext
            Wend
        End If
        .Close
    End With
    
    
    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")
    CollectionForTheDay = CollectionForTheDay + temDbl
End Sub

Private Sub HSSCollectionPayments()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    Dim temMyBHT As New clsBHT
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "Company Collection"
    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblIncomeBill.NetTotal, dbo.tblIncomeBill.BHTID , dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierName FROM dbo.tblIncomeBill LEFT OUTER JOIN dbo.tblHealthSchemeSuppliers ON dbo.tblIncomeBill.PaidHSSID = dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierID WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.IsHSSPaymentBill)=1)AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.PaymentMethodID)<>4))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                
                If !BHTID <> 0 Then
                    temMyBHT.BHTID = !BHTID
                    gridSummery.Col = 1
                    gridSummery.CellAlignment = 1
                    gridSummery.Text = Format(!HealthSchemeSupplierName, "") & " (BHT " & temMyBHT.BHT & ")"
                Else
                    gridSummery.Col = 1
                    gridSummery.CellAlignment = 1
                    gridSummery.Text = Format(!HealthSchemeSupplierName, "")
                End If
                
                temDbl = temDbl + !NetTotal
                gridSummery.Col = 4
                gridSummery.Text = Format(!NetTotal, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblIncomeReturnBill.ReturnValue, dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierName FROM dbo.tblHealthSchemeSuppliers RIGHT OUTER JOIN dbo.tblIncomeBill ON dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierID = dbo.tblIncomeBill.PaidHSSID RIGHT OUTER JOIN dbo.tblIncomeReturnBill ON dbo.tblIncomeBill.IncomeBillID = dbo.tblIncomeReturnBill.IncomeBillID WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)<>4) AND ((tblIncomeBill.IsHSSPaymentBill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl - !ReturnValue
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.CellAlignment = 1
                gridSummery.Text = Format(!HealthSchemeSupplierName, "")
                gridSummery.Col = 4
                gridSummery.Text = "(" & Format(!ReturnValue, "###0.00") & ")"
                .MoveNext
            Wend
        End If
        .Close
    End With
'    gridSummery.Rows = gridSummery.Rows + 1
'    gridSummery.Row = gridSummery.Rows - 1
'    gridSummery.Col = 0
'    gridSummery.Text = "Company Collection Payments"
    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")
    
    CollectionForTheDay = CollectionForTheDay + temDbl
    
End Sub

Private Sub BHTProfessionalPayments()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "     BHT"
    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblStaff.Name, dbo.tblProfessionalPaymentBill.Value " & _
                    "FROM dbo.tblProfessionalPaymentBill LEFT OUTER JOIN dbo.tblStaff ON dbo.tblProfessionalPaymentBill.StaffID = dbo.tblStaff.StaffID " & _
                    "WHERE (dbo.tblProfessionalPaymentBill.Date = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND (dbo.tblProfessionalPaymentBill.Cancelled = 0) AND (dbo.tblProfessionalPaymentBill.IsInwardPaymentBill = 1) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !Value
'                gridSummery.Rows = gridSummery.Rows + 1
'                gridSummery.Row = gridSummery.Rows - 1
'                gridSummery.Col = 1
'                gridSummery.CellAlignment = 1
'                gridSummery.Text = Format(!Name, "")
'                gridSummery.Col = 3
'                gridSummery.Text = Format(!Value, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")
    CollectionForTheDay = CollectionForTheDay + temDbl
End Sub

Private Sub OPDProfessionnalPayments()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "     OPD"
    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblStaff.Name, dbo.tblProfessionalPaymentBill.Value " & _
                    "FROM dbo.tblProfessionalPaymentBill LEFT OUTER JOIN dbo.tblStaff ON dbo.tblProfessionalPaymentBill.StaffID = dbo.tblStaff.StaffID " & _
                    "WHERE (dbo.tblProfessionalPaymentBill.Date = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND (dbo.tblProfessionalPaymentBill.Cancelled = 0) AND (dbo.tblProfessionalPaymentBill.IsOPDBill = 1) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !Value
'                gridSummery.Rows = gridSummery.Rows + 1
'                gridSummery.Row = gridSummery.Rows - 1
'                gridSummery.Col = 1
'                gridSummery.CellAlignment = 1
'                gridSummery.Text = Format(!Name, "")
'                gridSummery.Col = 3
'                gridSummery.Text = Format(!Value, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")


End Sub

Private Sub GSBProfessionnalPayments()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "     GSB"
    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblStaff.Name, dbo.tblProfessionalPaymentBill.Value " & _
                    "FROM dbo.tblProfessionalPaymentBill LEFT OUTER JOIN dbo.tblStaff ON dbo.tblProfessionalPaymentBill.StaffID = dbo.tblStaff.StaffID " & _
                    "WHERE (dbo.tblProfessionalPaymentBill.Date = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND (dbo.tblProfessionalPaymentBill.Cancelled = 0) AND (dbo.tblProfessionalPaymentBill.IsGSBill = 1) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !Value
'                gridSummery.Rows = gridSummery.Rows + 1
'                gridSummery.Row = gridSummery.Rows - 1
'                gridSummery.Col = 1
'                gridSummery.CellAlignment = 1
'                gridSummery.Text = Format(!Name, "")
'                gridSummery.Col = 3
'                gridSummery.Text = Format(!Value, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")



End Sub


Private Sub LabProfessionnalPayments()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "     Lab"
    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblStaff.Name, dbo.tblProfessionalPaymentBill.Value " & _
                    "FROM dbo.tblProfessionalPaymentBill LEFT OUTER JOIN dbo.tblStaff ON dbo.tblProfessionalPaymentBill.StaffID = dbo.tblStaff.StaffID " & _
                    "WHERE (dbo.tblProfessionalPaymentBill.Date = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND (dbo.tblProfessionalPaymentBill.Cancelled = 0) AND (dbo.tblProfessionalPaymentBill.IsLabBill = 1) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !Value
'                gridSummery.Rows = gridSummery.Rows + 1
'                gridSummery.Row = gridSummery.Rows - 1
'                gridSummery.Col = 1
'                gridSummery.CellAlignment = 1
'                gridSummery.Text = Format(!Name, "")
'                gridSummery.Col = 3
'                gridSummery.Text = Format(!Value, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")


End Sub

Private Sub RProfessionnalPayments()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "     Ron."
    
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT dbo.tblStaff.Name, dbo.tblProfessionalPaymentBill.Value " & _
                    "FROM dbo.tblProfessionalPaymentBill LEFT OUTER JOIN dbo.tblStaff ON dbo.tblProfessionalPaymentBill.StaffID = dbo.tblStaff.StaffID " & _
                    "WHERE (dbo.tblProfessionalPaymentBill.Date = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND (dbo.tblProfessionalPaymentBill.Cancelled = 0) AND (dbo.tblProfessionalPaymentBill.IsRBill = 1) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !Value
'                gridSummery.Rows = gridSummery.Rows + 1
'                gridSummery.Row = gridSummery.Rows - 1
'                gridSummery.Col = 1
'                gridSummery.CellAlignment = 1
'                gridSummery.Text = Format(!Name, "")
'                gridSummery.Col = 3
'                gridSummery.Text = Format(!Value, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")


End Sub


Private Sub CreditIncome()
'    Dim rsTem As New ADODB.Recordset
'    Dim temDbl As Double
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "SELECT Sum(tblIncomeBill.NetTotal) AS SumOfNetTotal From tblIncomeBill WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.PaymentMethodID)=4))"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            If IsNull(!SumOfNetTotal) = False Then
'                temDbl = !SumOfNetTotal
'            End If
'        End If
'        .Close
'    End With
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "SELECT Sum(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue FROM tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=4) AND ((tblIncomeReturnBill.Cancelled)=0))"
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
'            If IsNull(!SumOfReturnValue) = False Then
'                temDbl = temDbl - !SumOfReturnValue
'            End If
'        End If
'        .Close
'    End With
'    gridSummery.Rows = gridSummery.Rows + 1
'    gridSummery.Row = gridSummery.Rows - 1
'    gridSummery.Col = 0
'    gridSummery.Text = "Credit Transactions"
'    gridSummery.Col = 5
'    gridSummery.Text = Format(temDbl, "###0.00")
'
End Sub

Private Sub CheqeIncome()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    Dim temText As String
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "Cheques Transactions"
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT NetTotal, IsOPDBill, IsRBill, IsLabBill, IsPharmacyBill, IsInwardPaymentBill, IsGSBill, IsAgentBill, IsHSSPaymentBill, PaymentComments, DisplayBillID FROM dbo.tblIncomeBill WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.PaymentMethodID)=5)) ORDER BY IsOPDBill, IsRBill, IsLabBill, IsPharmacyBill, IsInwardPaymentBill, IsGSBill, IsAgentBill, IsHSSPaymentBill, DisplayBillID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !NetTotal
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.CellAlignment = 1
                If !IsOPDBill = True Then
                    temText = "OPD Bill No - "
                ElseIf !IsRBill = True Then
                    temText = "Rontgens Bill No - "
                ElseIf !IsLabBill = True Then
                    temText = "Lab Bill No - "
                ElseIf !IsPharmacyBill = True Then
                    temText = "Pharmacy Bill No - "
                ElseIf !IsInwardPaymentBill = True Then
                    temText = "BHT Payment No - "
                ElseIf !IsGSBill = True Then
                    temText = "GSB Payment No - "
                ElseIf !IsAgentBill = True Then
                    temText = "Agent Payment Bill No - "
                ElseIf !IsHSSPaymentBill = True Then
                    temText = "Company Payment Bill No - "
                Else
                    temText = "Unknown Bill No  - "
                End If
                gridSummery.Text = temText & Format(!DisplayBillID, "")
                gridSummery.Col = 2
                gridSummery.CellAlignment = 1
                gridSummery.Text = Left(Format(!PaymentComments, ""), 15)
                gridSummery.Col = 3
                gridSummery.Text = Format(!NetTotal, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue FROM tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=5)  AND ((tblIncomeReturnBill.Cancelled)=0))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfReturnValue) = False Then
                temDbl = temDbl - !SumOfReturnValue
            End If
        End If
        .Close
    End With

    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")

End Sub

Private Sub CardIncome()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    Dim temText As String
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "Credit Card Transactions"
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT NetTotal, IsOPDBill, IsRBill, IsLabBill, IsPharmacyBill, IsInwardPaymentBill, IsGSBill, IsAgentBill, IsHSSPaymentBill, PaymentComments, DisplayBillID FROM dbo.tblIncomeBill WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.PaymentMethodID)=3)) ORDER BY IsOPDBill, IsRBill, IsLabBill, IsPharmacyBill, IsInwardPaymentBill, IsGSBill, IsAgentBill, IsHSSPaymentBill, DisplayBillID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !NetTotal
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.CellAlignment = 1
                If !IsOPDBill = True Then
                    temText = "OPD Bill No - "
                ElseIf !IsRBill = True Then
                    temText = "Rontgens Bill No - "
                ElseIf !IsLabBill = True Then
                    temText = "Lab Bill No - "
                ElseIf !IsPharmacyBill = True Then
                    temText = "Pharmacy Bill No - "
                ElseIf !IsInwardPaymentBill = True Then
                    temText = "BHT Payment No - "
                ElseIf !IsGSBill = True Then
                    temText = "GSB Payment No - "
                ElseIf !IsAgentBill = True Then
                    temText = "Agent Payment Bill No - "
                ElseIf !IsHSSPaymentBill = True Then
                    temText = "Company Payment Bill No - "
                Else
                    temText = "Unknown Bill No  - "
                End If
                gridSummery.Text = temText & Format(!DisplayBillID, "")
                gridSummery.Col = 2
                gridSummery.CellAlignment = 1
                gridSummery.Text = Left(Format(!PaymentComments, ""), 20)
                gridSummery.Col = 3
                gridSummery.Text = Format(!NetTotal, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue FROM tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=3)  AND ((tblIncomeReturnBill.Cancelled)=0))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfReturnValue) = False Then
                temDbl = temDbl - !SumOfReturnValue
            End If
        End If
        .Close
    End With

    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")

End Sub

Private Sub SlipsIncome()
    Dim rsTem As New ADODB.Recordset
    Dim temDbl As Double
    Dim temText As String
    
    gridSummery.Rows = gridSummery.Rows + 1
    gridSummery.Row = gridSummery.Rows - 1
    gridSummery.Col = 0
    gridSummery.Text = "Slips Transactions"
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT IncomeBillID, NetTotal, IsOPDBill, IsRBill, IsLabBill, IsPharmacyBill, IsInwardPaymentBill, IsGSBill, IsAgentBill, IsHSSPaymentBill, PaymentComments, DisplayBillID FROM dbo.tblIncomeBill WHERE (((tblIncomeBill.CompletedDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.Cancelled)=0) AND ((tblIncomeBill.PaymentMethodID)=7)) ORDER BY IsOPDBill, IsRBill, IsLabBill, IsPharmacyBill, IsInwardPaymentBill, IsGSBill, IsAgentBill, IsHSSPaymentBill, DisplayBillID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                temDbl = temDbl + !NetTotal
                gridSummery.Rows = gridSummery.Rows + 1
                gridSummery.Row = gridSummery.Rows - 1
                gridSummery.Col = 1
                gridSummery.CellAlignment = 1
'                If !IsOPDBill = True Then
'                    temText = "OPD Bill No - "
'                ElseIf !IsRBill = True Then
'                    temText = "Rontgens Bill No - "
'                ElseIf !IsLabBill = True Then
'                    temText = "Lab Bill No - "
'                ElseIf !IsPharmacyBill = True Then
'                    temText = "Pharmacy Bill No - "
'                ElseIf !IsInwardPaymentBill = True Then
'                    temText = "BHT Payment No - "
'                ElseIf !IsGSBill = True Then
'                    temText = "GSB Payment No - "
'                ElseIf !IsAgentBill = True Then
'                    temText = "Agent Payment Bill No - "
'                ElseIf !IsHSSPaymentBill = True Then
'                    temText = "Company Payment Bill No - "
'                Else
'                    temText = "Unknown Bill No  - "
'                End If
                
                
                If !IsOPDBill = True Then
                    temText = "OPD - "
                ElseIf !IsRBill = True Then
                    temText = "Ron - "
                ElseIf !IsLabBill = True Then
                    temText = "Lab - "
                ElseIf !IsPharmacyBill = True Then
                    temText = "Pha - "
                ElseIf !IsInwardPaymentBill = True Then
                    temText = "BHT - "
                ElseIf !IsGSBill = True Then
                    temText = "GSB - "
                ElseIf !IsAgentBill = True Then
                    temText = "Agn - "
                ElseIf !IsHSSPaymentBill = True Then
                    temText = "Com - "
                Else
                    temText = "Unknown - "
                End If
                
                
                'gridSummery.Text = temText & Format(!DisplayBillID, "")
                
                gridSummery.Text = temText & Format(!IncomeBillID, "")
                
                gridSummery.Col = 2
                gridSummery.CellAlignment = 1
                gridSummery.Text = Format(!PaymentComments, "")
                gridSummery.Col = 3
                gridSummery.Text = Format(!NetTotal, "###0.00")
                .MoveNext
            Wend
        End If
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT Sum(tblIncomeReturnBill.ReturnValue) AS SumOfReturnValue FROM tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID WHERE (((tblIncomeReturnBill.ReturnDate)='" & Format(dtpDate.Value, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=7)  AND ((tblIncomeReturnBill.Cancelled)=0))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfReturnValue) = False Then
                temDbl = temDbl - !SumOfReturnValue
            End If
        End If
        .Close
    End With

    gridSummery.Col = 5
    gridSummery.Text = Format(temDbl, "###0.00")

End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridSummery, HospitalName & " - Daily Summery Report", Format(dtpDate.Value, "dd MMMM yyyy")

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
    
    
    
    GridPrint gridSummery, ThisReportFormat, HospitalName & " - Daily Summery Report", Format(dtpDate.Value, "dd MMMM yyyy")
    Printer.EndDoc
End Sub


Private Function ExpenceCollectionIncome() As Double
    ExpenceCollectionIncome = 0
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT sum(tblPatientService.Charge) as SumOfCharge " & _
                    "FROM tblIncomeBill RIGHT JOIN ((tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID) ON tblIncomeBill.IncomeBillID = tblPatientService.ExpenceBillID " & _
                    "WHERE tblIncomeBill.Completed = 1 AND tblIncomeBill.CompletedDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' "
        temSql = temSql & " AND tblPatientService.ExpenceBillID <> 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfCharge) = False Then
                ExpenceCollectionIncome = !SumOfCharge
            End If
            .Close
        End If
        If .State = 1 Then .Close
        temSql = "SELECT sum(tblPatientService.Charge) as SumOfCharge " & _
                    "FROM tblIncomeBill RIGHT JOIN ((tblPatientService LEFT JOIN tblServiceCategory ON tblPatientService.ServiceCategoryID = tblServiceCategory.ServiceCategoryID) LEFT JOIN tblServiceSubcategory ON tblPatientService.ServiceSubcategoryID = tblServiceSubcategory.ServiceSubcategoryID) ON tblIncomeBill.IncomeBillID = tblPatientService.ExpenceBillID " & _
                    "WHERE tblIncomeBill.Cancelled = 1  AND tblIncomeBill.Completed = 1 AND tblIncomeBill.CancelledDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' "
        temSql = temSql & " AND tblPatientService.ExpenceBillID <> 0 "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If IsNull(!SumOfCharge) = False Then
                ExpenceCollectionIncome = ExpenceCollectionIncome - !SumOfCharge
            End If
            .Close
        End If
        If .State = 1 Then .Close
    End With
    
End Function


Private Sub btnProcess_Click()
    Call FormatGrid
    
    CollectionForTheDay = 0
    
    Call OPDCollectionIncome
    Call RCollectionIncome
    Call PharmacyCollectionIncome
    'Call LabCollectionIncome
    'Call MTCollectionIncome
    Call BHTCollectionPayments
    Call GSBCollectionPayments
    Call AgentCollectionIncome
    Call HSSCollectionPayments
    
    With gridSummery
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = "Collection for the day"
        .Col = 5
        .Text = Format(CollectionForTheDay, "###0.00")
    
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = "C/F Cash Balance"
        .Col = 5
        .Text = "............"
    
    
    
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = "Less"
    
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .Text = "Professional Payments"
    
    
    End With
    
    Call BHTProfessionalPayments
'    Call GSBProfessionnalPayments
'    Call OPDProfessionnalPayments
'    Call LabProfessionnalPayments
'    Call RProfessionnalPayments
    
    With gridSummery
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = "Petty Cash Payments"
        .Col = 5
        .Text = Abs(ExpenceCollectionIncome)
    
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = "Income to Bank"
        .Col = 5
        .Text = "............"
        
        .Rows = .Rows + 2
        .Row = .Rows - 1
    
    End With
    
    
    Call CreditIncome
    Call CheqeIncome
    Call CardIncome
    Call SlipsIncome

    With gridSummery
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = "Bank Deposit- Cash"
    
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = Space(5) & "Peoples Bank"
        .Col = 5
        .Text = "............"
        
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = Space(5) & "HNB 1"
        .Col = 5
        .Text = "............"
        
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = Space(5) & "HNB 2"
        .Col = 5
        .Text = "............"
        
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = Space(5) & "BOC"
        .Col = 5
        .Text = "............"
        
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = Space(5) & "NSB"
        .Col = 5
        .Text = "............"
        
   
    
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = "B/F Cash Balance"
        .Col = 5
        .Text = "===="
    
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 0
        .Text = "Prepaired By " & UserFullName
        .Col = 5
        .Text = "---------"
    
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .Text = "Printed Date"
        .Col = 5
        .Text = Format(Date, "dd MMM yyyy")
    
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .Text = "Printed Time"
        .Col = 5
        .Text = Time
    
    
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .Text = "Checked By "
        .Col = 5
        .Text = "---------"
    
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .Text = "Book keeper"
        .Col = 5
        .Text = "---------"
    
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .Text = "Accountant "
        .Col = 5
        .Text = "---------"
    
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .Text = "Manager"
        .Col = 5
        .Text = "---------"
    
    
        .ColWidth(3) = 1600
        .ColWidth(4) = 1600
    
    End With


End Sub

Private Sub FormatGrid()
    With gridSummery
        .Clear
        .Rows = 1
        
        .Cols = 6
        
        .Col = 0
        .Text = "Section"
        
        .Col = 1
        .Text = "Discreption"
        
        .Col = 2
        .Text = "Count"
        
        .Col = 3
        .Text = "Hos. Fee"
               
        .Col = 4
        .Text = "Sub Total"
        
        .Col = 5
        .Text = "Totals"
    
    
    End With
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    Dim MaxWidth As Double
    Dim MaxRow As Double
    For i = 0 To gridSummery.Rows - 1
        If MaxWidth < Printer.TextWidth(gridSummery.TextMatrix(i, Val(Text1.Text))) Then
            MaxWidth = Printer.TextWidth(gridSummery.TextMatrix(i, Val(Text1.Text)))
            MaxRow = i
        End If
        
    Next
    MsgBox MaxWidth & vbTab & MaxRow
            gridSummery.Col = Val(Text1.Text)
            gridSummery.Row = MaxRow
            gridSummery.CellForeColor = vbRed

End Sub

Private Sub Form_Load()
    Call GetSettings
    GetCommonSettings Me
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub
