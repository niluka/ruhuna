VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPeriodSummeryDetailsByDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Period vise Summery Details"
   ClientHeight    =   8985
   ClientLeft      =   3930
   ClientTop       =   -180
   ClientWidth     =   11145
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
   ScaleHeight     =   8985
   ScaleWidth      =   11145
   Begin VB.OptionButton optPayments 
      Caption         =   "Payments"
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   3120
      TabIndex        =   18
      Top             =   8040
      Width           =   2415
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   3120
      TabIndex        =   16
      Top             =   8520
      Width           =   2415
   End
   Begin VB.OptionButton optRefunds 
      Caption         =   "Refunds"
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   2520
      Width           =   2655
   End
   Begin VB.OptionButton optCancelled 
      Caption         =   "Cancelled"
      Height          =   255
      Left            =   7080
      TabIndex        =   13
      Top             =   2160
      Width           =   2655
   End
   Begin VB.OptionButton optBilled 
      Caption         =   "Billed"
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   2160
      Value           =   -1  'True
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid gridBill 
      Height          =   4935
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8705
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbUser 
      Height          =   360
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.ComboBox cmbType 
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   3615
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      Top             =   8040
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   8520
      TabIndex        =   0
      Top             =   8040
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
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   1800
      TabIndex        =   9
      Top             =   2520
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   77922307
      CurrentDate     =   39969
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      Top             =   1080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   77922307
      CurrentDate     =   39969
   End
   Begin btButtonEx.ButtonEx btnExcel 
      Height          =   495
      Left            =   7200
      TabIndex        =   23
      Top             =   8040
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
   Begin VB.Label Label5 
      Caption         =   "To"
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "From"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Total Value"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   8520
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Total Count"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   8040
      Width           =   2775
   End
   Begin VB.Label Label42 
      Caption         =   "Paid as"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblTopic 
      Alignment       =   2  'Center
      Caption         =   "Topic"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   10815
   End
   Begin VB.Label lblSubtopic 
      Alignment       =   2  'Center
      Caption         =   "Topic"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   10815
   End
   Begin VB.Label Label8 
      Caption         =   "Bill Type"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Cashier"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmPeriodSummeryDetailsByDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsTem As New ADODB.Recordset
    Dim MyGrid As Grid
    Dim MyGridRow() As GridRow
    Dim MyGridCell() As GridCell
    Dim i As Integer
    Dim n As Integer
    Dim TotalValue As Double
    Dim TotalCount As Long

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnExcel_Click()
    GridToExcel gridBill, "Detailed Summery - " & cmbUser.Text & " - " & cmbType.Text & " - " & cmbPaymentMethod.Text, "From : " & Format(dtpFrom.Value, "dd MM yy") & " To : " & Format(dtpTo.Value, "dd MM yy")
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    
    
    With ThisReportFormat
        .ColSpace = 500
        
        .SubTopicFontSize = 8
    
        .ColFontSize = 7
        .ColSpace = 200
    
    End With
    
    GridPrint gridBill, ThisReportFormat, "Detailed Summery - " & cmbUser.Text & " - " & cmbType.Text & " - " & cmbPaymentMethod.Text, "From : " & Format(dtpFrom.Value, "dd MM yy") & " To : " & Format(dtpTo.Value, "dd MM yy")
    Printer.EndDoc
    
End Sub

Private Sub cmbPaymentMethod_Change()
    Call FormatGrid
    Call FillGrid

End Sub


Private Sub cmbPaymentMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbPaymentMethod.Text = Empty
    End If
End Sub

Private Sub cmbType_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbType_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbUser_Change()
    Call FormatGrid
    Call FillGrid

End Sub

Private Sub cmbUser_Click(Area As Integer)
    Call FormatGrid
    Call FillGrid

End Sub

Private Sub cmbUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbUser.Text = Empty
    End If
End Sub

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub



Private Sub CatIncome(ByVal IncomeCategory As String, PaymentMethodID As Long)
    TotalValue = 0
    TotalCount = 0
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(cmbPaymentMethod.BoundText) = True Then
            temSql = "SELECT tblIncomeBill.*, tblAgent.Agent, tblOPDPatient.FirstName as OPDFirstName, tblBHTPatient.FirstName as BHTFirstName, tblBHT.BHT "
            temSql = temSql & "FROM (((tblIncomeBill LEFT JOIN tblPatientMainDetails AS tblOPDPatient ON tblIncomeBill.PatientID = tblOPDPatient.PatientID) LEFT JOIN tblBHT ON tblIncomeBill.BHTID = tblBHT.BHTID) LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID) LEFT JOIN tblAgent ON tblIncomeBill.AgentID = tblAgent.AgentID "
            If IsNumeric(cmbUser.BoundText) = True Then
                If IncomeCategory = "All" Then
                    temSql = temSql & "where (((tblIncomeBill.CompletedDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND  ((tblIncomeBill.CompletedUserID)=" & Val(cmbUser.BoundText) & ") AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
                Else
                    temSql = temSql & "where (((tblIncomeBill.CompletedDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeBill.CompletedUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1)AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
                End If
            Else
                If IncomeCategory = "All" Then
                    temSql = temSql & "where (((tblIncomeBill.CompletedDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
                Else
                    temSql = temSql & "where (((tblIncomeBill.CompletedDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1)AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.PaymentMethodID)=" & PaymentMethodID & "))"
                End If
            End If
            temSql = temSql & " ORDER by tblIncomeBill.CompletedDate, tblIncomeBill.CompletedTime"
        Else
            temSql = "SELECT tblIncomeBill.*, tblAgent.Agent, tblOPDPatient.FirstName as OPDFirstName, tblBHTPatient.FirstName as BHTFirstName, tblBHT.BHT "
            temSql = temSql & "FROM (((tblIncomeBill LEFT JOIN tblPatientMainDetails AS tblOPDPatient ON tblIncomeBill.PatientID = tblOPDPatient.PatientID) LEFT JOIN tblBHT ON tblIncomeBill.BHTID = tblBHT.BHTID) LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID) LEFT JOIN tblAgent ON tblIncomeBill.AgentID = tblAgent.AgentID "
            If IsNumeric(cmbUser.BoundText) = True Then
                If IncomeCategory = "All" Then
                    temSql = temSql & "where (((tblIncomeBill.CompletedDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND  ((tblIncomeBill.CompletedUserID)=" & Val(cmbUser.BoundText) & ") AND ((tblIncomeBill.Completed)=1))"
                Else
                    temSql = temSql & "where (((tblIncomeBill.CompletedDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeBill.CompletedUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1)AND ((tblIncomeBill.Completed)=1))"
                End If
            Else
                If IncomeCategory = "All" Then
                    temSql = temSql & "where (((tblIncomeBill.CompletedDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Completed)=1))"
                Else
                    temSql = temSql & "where (((tblIncomeBill.CompletedDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1)AND ((tblIncomeBill.Completed)=1))"
                End If
            End If
            temSql = temSql & " ORDER by tblIncomeBill.CompletedDate, tblIncomeBill.CompletedTime"
        
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            ReDim MyGridRow(.RecordCount)
            TotalCount = .RecordCount
            For i = 0 To .RecordCount - 1
                MyGridCell(0).Col = 0
                MyGridCell(0).Text = ![IncomeBillID]
                MyGridCell(1).Col = 0
                MyGridCell(1).Text = ![DisplayBillID]
               
                MyGridCell(2).Col = 1
                MyGridCell(2).CellAlignment = 1
                  
                If !PharmacyBillID <> 0 Then
                    MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "") & " " & Format(![Agent], "") & " " & !PharmacyBillID
                Else
                    MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "") & " " & Format(![Agent], "")
                End If
                
                MyGridCell(3).Col = 2
                MyGridCell(3).Text = Format(!NetTotal, "0.00")
                MyGridCell(4).Col = 3
                MyGridCell(4).Text = Format(!CompletedDate, "dd MM yy") 'Format(!Time, "hh:mm AMPM")
                MyGridCell(5).Col = 4
                If !Cancelled = True Then
                    MyGridCell(5).Text = "Cancelled on " & Format(!CancelledDate, "dd MM yy") & " at " & Format(!CancelledTime, "mm:hh AMPM")
                Else
                    MyGridCell(5).Text = ""
                End If
                MyGridCell(6).Col = 5
                MyGridCell(6).Text = ""
                MyGridRow(i).RowCells = MyGridCell
                TotalValue = TotalValue + !NetTotal
                .MoveNext
            Next i
            
            With gridBill
                .Rows = UBound(MyGridRow) + 1
                For i = 1 To .Rows - 1
                    For n = 0 To UBound(MyGridRow(i - 1).RowCells)
                        .TextMatrix(i, n) = MyGridRow(i - 1).RowCells(n).Text
                    Next n
                Next i
            End With
        End If
        .Close
    End With
End Sub

Private Sub CatPay(ByVal IncomeCategory As String, PaymentMethodID As Long)
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT ProfessionalPaymentBillID, StaffID, Value, Time, PaymentComments  " & _
                    "FROM tblProfessionalPaymentBill "
        If IncomeCategory <> "All" Then
            temSql = temSql & "WHERE tblProfessionalPaymentBill.Date between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' AND tblProfessionalPaymentBill.Is" & IncomeCategory & "Bill =1 "
        Else
            temSql = temSql & "WHERE tblProfessionalPaymentBill.Date between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo.Value, "dd MMMM yyyy") & "' "
        End If
        If IsNumeric(cmbPaymentMethod.BoundText) = True Then
            temSql = temSql & "AND tblProfessionalPaymentBill.PaymentMethodID = " & PaymentMethodID & " "
        End If
        If IsNumeric(cmbUser.BoundText) = True Then
            temSql = temSql & "AND tblProfessionalPaymentBill.UserID = " & Val(cmbUser.BoundText) & " "
        End If
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            ReDim MyGridRow(.RecordCount)
            TotalCount = .RecordCount
            For i = 0 To .RecordCount - 1
                MyGridCell(0).Col = 0
                MyGridCell(0).Text = ![ProfessionalPaymentBillID]
                MyGridCell(1).Col = 0
                MyGridCell(1).Text = ![ProfessionalPaymentBillID]
               
                MyGridCell(2).Col = 1
                MyGridCell(2).CellAlignment = 1
                MyGridCell(2).Text = FullStaffName(!StaffID)
                
                MyGridCell(3).Col = 2
                MyGridCell(3).Text = Format(!Value, "0.00")
                MyGridCell(4).Col = 3
                MyGridCell(4).Text = Format(!Date, "dd MM yy") 'Format(!Time, "hh:mm AMPM")
                MyGridCell(5).Col = 4
                If !Value < 0 Then
                    MyGridCell(5).Text = "A Cancellation "
                Else
                    MyGridCell(5).Text = ""
                End If
                
                MyGridCell(6).Col = 5
                MyGridCell(6).Text = !PaymentComments
                MyGridRow(i).RowCells = MyGridCell
                TotalValue = TotalValue + !Value
                .MoveNext
            Next i
            
            With gridBill
                .Rows = UBound(MyGridRow) + 1
                For i = 1 To .Rows - 1
                    For n = 0 To UBound(MyGridRow(i - 1).RowCells)
                        .TextMatrix(i, n) = MyGridRow(i - 1).RowCells(n).Text
                    Next n
                Next i
            End With
        End If
        
        .Close
    End With
End Sub




Private Sub CatReturn(ByVal IncomeCategory As String, PaymentMethodID As Long)
    
    TotalValue = 0
    TotalCount = 0
    With rsTem
        If .State = 1 Then .Close
        
        If IsNumeric(cmbPaymentMethod.BoundText) = True Then
            temSql = "SELECT tblIncomeReturnBill.*, tblIncomeBill.*,  tblIncomeReturnBill.IncomeBillID as RIncomeBillID, tblIncomeBill.IncomeBillID, tblBHT.BHT, tblOPDPatient.FirstName as OPDFirstName, tblBHTPatient.FirstName as BHTFirstName " & _
                        "FROM ((tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID) LEFT JOIN tblPatientMainDetails AS tblOPDPatient ON tblIncomeBill.PatientID = tblOPDPatient.PatientID) LEFT JOIN (tblBHT LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID) ON tblIncomeReturnBill.BHTID = tblBHT.BHTID "
            If IsNumeric(cmbUser.BoundText) = True Then
                If IncomeCategory = "All" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblIncomeReturnBill.ReturnUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeReturnBill.Cancelled)=0))"
                ElseIf IncomeCategory = "GS" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblIncomeReturnBill.ReturnUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeReturnBill.IncomeBillID)=0) AND  ((tblBHT.ISGSB)= 1 ) AND ((tblIncomeReturnBill.Cancelled)=0))"
                ElseIf IncomeCategory = "InwardPayment" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblIncomeReturnBill.ReturnUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeReturnBill.IncomeBillID)=0) AND  ((tblBHT.ISBHT)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
                Else
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND ((tblIncomeReturnBill.ReturnUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
                End If
            Else
                If IncomeCategory = "All" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND  ((tblIncomeReturnBill.Cancelled)=0))"
                ElseIf IncomeCategory = "GS" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND  ((tblIncomeReturnBill.IncomeBillID)=0) AND  ((tblBHT.ISGSB)= 1 ) AND ((tblIncomeReturnBill.Cancelled)=0))"
                ElseIf IncomeCategory = "InwardPayment" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND  ((tblIncomeReturnBill.IncomeBillID)=0) AND  ((tblBHT.ISBHT)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
                Else
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "') AND ((tblIncomeReturnBill.PaymentMethodID)=" & PaymentMethodID & ")  AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
                End If
            End If
        Else
            temSql = "SELECT tblIncomeReturnBill.*, tblIncomeBill.*, tblBHT.BHT, tblOPDPatient.FirstName, tblBHTPatient.FirstName " & _
                        "FROM ((tblIncomeReturnBill LEFT JOIN tblIncomeBill ON tblIncomeReturnBill.IncomeBillID = tblIncomeBill.IncomeBillID) LEFT JOIN tblPatientMainDetails AS tblOPDPatient ON tblIncomeBill.PatientID = tblOPDPatient.PatientID) LEFT JOIN (tblBHT LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID) ON tblIncomeReturnBill.BHTID = tblBHT.BHTID "
            If IsNumeric(cmbUser.BoundText) = True Then
                If IncomeCategory = "All" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate)between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')   AND ((tblIncomeReturnBill.ReturnUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeReturnBill.Cancelled)=0))"
                ElseIf IncomeCategory = "GS" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND ((tblIncomeReturnBill.ReturnUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeReturnBill.IncomeBillID)=0) AND  ((tblBHT.ISGSB)= 1 ) AND ((tblIncomeReturnBill.Cancelled)=0))"
                ElseIf IncomeCategory = "InwardPayment" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND ((tblIncomeReturnBill.ReturnUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeReturnBill.IncomeBillID)=0) AND  ((tblBHT.ISBHT)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
                Else
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND ((tblIncomeReturnBill.ReturnUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
                End If
            Else
                If IncomeCategory = "All" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND  ((tblIncomeReturnBill.Cancelled)=0))"
                ElseIf IncomeCategory = "GS" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND  ((tblIncomeReturnBill.IncomeBillID)=0) AND  ((tblBHT.ISGSB)= 1 ) AND ((tblIncomeReturnBill.Cancelled)=0))"
                ElseIf IncomeCategory = "InwardPayment" Then
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND ((tblIncomeReturnBill.IncomeBillID)=0) AND  ((tblBHT.ISBHT)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
                Else
                    temSql = temSql & "Where (((tblIncomeReturnBill.ReturnDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeReturnBill.Cancelled)=0))"
                End If
            End If
        
        End If
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            TotalCount = .RecordCount
            ReDim MyGridRow(.RecordCount)
            For i = 0 To .RecordCount - 1
                MyGridCell(0).Col = 0
                If IsNull(![RIncomeBillID]) = True Then
                    MyGridCell(0).Text = "BHT/GSB"
                Else
                    MyGridCell(0).Text = ![RIncomeBillID]
                End If
                MyGridCell(1).Col = 1
                
                If IsNull(![DisplayBillID]) = True Then
                    MyGridCell(1).Text = "BHT/GSB"
                Else
                    MyGridCell(1).Text = ![DisplayBillID]
                End If
                MyGridCell(2).Col = 1
                MyGridCell(2).CellAlignment = 1
                MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "")
                
                
                If !PharmacyBillID <> 0 Then
                    MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "") & " " & Format(![Agent], "") & " " & !PharmacyBillID
                Else
                    MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "") & " " & Format(![Agent], "")
                End If
                
                
                
                MyGridCell(3).Col = 2
                MyGridCell(3).Text = Format(!NetTotal, "0.00")
                MyGridCell(4).Col = 3
                MyGridCell(4).Text = Format(!CompletedTime, "HH:MM AMPM")
                MyGridCell(5).Col = 4
                MyGridCell(5).Text = "Returned on " & Format(!ReturnDate, "dd MMMM yyyy") & " at " & !ReturnTime
                MyGridCell(6).Col = 5
                MyGridCell(6).Text = "Returned Value " & Format(!ReturnValue, "0.00")
                MyGridRow(i).RowCells = MyGridCell
                TotalValue = TotalValue + !ReturnValue
                MyGridRow(i).RowCells = MyGridCell
                .MoveNext
            Next i
            
            With gridBill
                .Rows = UBound(MyGridRow) + 1
                For i = 1 To .Rows - 1
                    For n = 0 To UBound(MyGridRow(i - 1).RowCells) - 1
                        .TextMatrix(i, n) = MyGridRow(i - 1).RowCells(n).Text
                    Next n
                Next i
            End With
        End If
        .Close
    End With
End Sub

Private Sub CatCancellation(ByVal IncomeCategory As String, PaymentMethodID As Long)
    
    TotalCount = 0
    TotalValue = 0
    With rsTem
        If .State = 1 Then .Close
                
        If IsNumeric(cmbPaymentMethod.BoundText) = True Then
            temSql = "SELECT tblIncomeBill.*, tblOPDPatient.FirstName as OPDFirstName, tblBHTPatient.FirstName as BHTFirstName, tblBHT.BHT  "
            temSql = temSql & "FROM ((tblIncomeBill LEFT JOIN tblPatientMainDetails AS tblOPDPatient ON tblIncomeBill.PatientID = tblOPDPatient.PatientID) LEFT JOIN tblBHT ON tblIncomeBill.BHTID = tblBHT.BHTID) LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID "
            If IsNumeric(cmbUser.BoundText) = True Then
                If IncomeCategory <> "All" Then
                    temSql = temSql & "Where (((tblIncomeBill.CancelledDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND ((tblIncomeBill.CancelledUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeBill.Cancelled)=1) AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
                Else
                    temSql = temSql & "Where (((tblIncomeBill.CancelledDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND ((tblIncomeBill.Cancelled)=1)  AND ((tblIncomeBill.CancelledUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
                End If
            Else
                If IncomeCategory <> "All" Then
                    temSql = temSql & "Where (((tblIncomeBill.CancelledDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeBill.Cancelled)=1) AND ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
                Else
                    temSql = temSql & "Where (((tblIncomeBill.CancelledDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND ((tblIncomeBill.Cancelled)=1)  AND  ((tblIncomeBill.Completed)=1) AND ((tblIncomeBill.CancelledPaymentMethodID)=" & PaymentMethodID & "))"
                End If
            End If
        Else
            temSql = "SELECT tblIncomeBill.*, tblOPDPatient.FirstName  as OPDFirstName, tblBHTPatient.FirstName  as BHTFirstName, tblBHT.BHT  "
            temSql = temSql & "FROM ((tblIncomeBill LEFT JOIN tblPatientMainDetails AS tblOPDPatient ON tblIncomeBill.PatientID = tblOPDPatient.PatientID) LEFT JOIN tblBHT ON tblIncomeBill.BHTID = tblBHT.BHTID) LEFT JOIN tblPatientMainDetails AS tblBHTPatient ON tblBHT.PatientID = tblBHTPatient.PatientID "
            If IsNumeric(cmbUser.BoundText) = True Then
                If IncomeCategory <> "All" Then
                    temSql = temSql & "Where (((tblIncomeBill.CancelledDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND ((tblIncomeBill.CancelledUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeBill.Cancelled)=1) AND ((tblIncomeBill.Completed)=1))"
                Else
                    temSql = temSql & "Where (((tblIncomeBill.CancelledDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND ((tblIncomeBill.Cancelled)=1)  AND ((tblIncomeBill.CancelledUserID)=" & Val(cmbUser.BoundText) & ") AND  ((tblIncomeBill.Completed)=1))"
                End If
            Else
                If IncomeCategory <> "All" Then
                    temSql = temSql & "Where (((tblIncomeBill.CancelledDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND  ((tblIncomeBill.Is" & IncomeCategory & "Bill)=1) AND ((tblIncomeBill.Cancelled)=1) AND ((tblIncomeBill.Completed)=1))"
                Else
                    temSql = temSql & "Where (((tblIncomeBill.CancelledDate) between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' AND '" & Format(dtpTo, "dd MMMM yyyy") & "')  AND ((tblIncomeBill.Cancelled)=1)  AND  ((tblIncomeBill.Completed)=1))"
                End If
            End If
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            ReDim MyGridRow(.RecordCount)
            TotalCount = .RecordCount
            For i = 0 To .RecordCount - 1
                MyGridCell(0).Col = 0
                MyGridCell(0).Text = ![IncomeBillID]
                MyGridCell(1).Col = 0
                MyGridCell(1).Text = ![DisplayBillID]
                MyGridCell(2).Col = 1
                MyGridCell(2).CellAlignment = 1
                MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "")
                
                If !PharmacyBillID <> 0 Then
                    MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "") & " "  ' & Format(![Agent], "") & " " & !PharmacyBillID
                Else
                    MyGridCell(2).Text = Format(![OPDFirstName], "") & Format(![BHT], "") & " " & Format(![BHTFirstName], "") & " "  ' & Format(![Agent], "")
                End If
                
                
                MyGridCell(3).Col = 2
                MyGridCell(3).Text = Format(!NetTotal, "0.00")
                MyGridCell(4).Col = 3
                MyGridCell(4).Text = Format(!CompletedDate, "dd MM yy") 'Format(!Time, "dd MM yy") 'Format(!Time, "hh:mm AMPM")
                MyGridCell(5).Col = 4
                If !Cancelled = True Then
                    MyGridCell(5).Text = "Cancelled on " & Format(!CancelledDate, "dd MM yy") & " at " & Format(!CancelledTime, "mm:hh AMPM")
                Else
                    MyGridCell(5).Text = ""
                End If
                MyGridCell(6).Col = 5
                MyGridCell(6).Text = ""
                TotalValue = TotalValue + !NetTotal
                MyGridRow(i).RowCells = MyGridCell
                .MoveNext
            Next i
            
            With gridBill
                .Rows = UBound(MyGridRow) + 1
                For i = 1 To .Rows - 1
                    For n = 0 To UBound(MyGridRow(i - 1).RowCells) - 1
                        .TextMatrix(i, n) = MyGridRow(i - 1).RowCells(n).Text
                    Next n
                Next i
            End With
        End If
        .Close
    End With
End Sub

Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
End Sub


Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    Call GetSettings
    Call FillGrid
End Sub

Private Sub FormatGrid()
    With gridBill
        .Cols = 7

        .Rows = 1

        .Row = 0

        .Col = 0
        .Text = "ID"

        .Col = 1
        .Text = "Bill ID"

        .Col = 2
        .Text = "Customer/BHT/GSB/Doctor"

        .Col = 3
        .Text = "Value"

        .Col = 4
        .Text = "Date"

        .Col = 5
        .Text = "Remarks"

        .Col = 6
        .Text = "Comments"


        .ColWidth(0) = 0
        .ColWidth(6) = 0
        
        ReDim MyGridCell(6)

    End With
End Sub

Private Sub FillGrid()
    Dim TemString As String
    
    TotalValue = 0
    
    Select Case cmbType.Text
        Case "OPD Bills": TemString = "OPD"
        Case "Lab Bills": TemString = "Lab"
        Case "Pharmacy Bills": TemString = "Pharmacy"
        Case "Inward Bills": TemString = "InwardPayment"
        Case "Medical Test Bills": TemString = "MedicalTest"
        Case "Green Sheet Bills": TemString = "GS"
        Case "All Bills": TemString = "All"
        Case "Agent Bills": TemString = "Agent"
        Case "Expence Bills": TemString = "Expence"
        Case "Roentgents Bills": TemString = "R"
        Case "Health Scheme Supplier Payments": TemString = "HSSPayment"
        Case Else:  Exit Sub
    End Select
    
    
    If optBilled.Value = True Then
         CatIncome TemString, Val(cmbPaymentMethod.BoundText)
    ElseIf optCancelled.Value = True Then
         CatCancellation TemString, Val(cmbPaymentMethod.BoundText)
    ElseIf optRefunds.Value = True Then
         CatReturn TemString, Val(cmbPaymentMethod.BoundText)
    ElseIf optPayments.Value = True Then
        CatPay TemString, Val(cmbPaymentMethod.BoundText)
    End If
    
    With gridBill
        .Rows = .Rows + 2
        .Row = .Rows - 1
        .Col = 1
        .Text = "Total"
        .Col = 2
        .Text = Format(TotalValue, "0.00")
        
    End With
    
    
    txtValue.Text = Format(TotalValue, "0.00")
    txtCount.Text = TotalCount
End Sub

Private Sub FillCombos()
    Dim Cashier As New clsFillCombos
    Cashier.FillSpecificField cmbUser, "Staff", "Name", False
    Dim PM As New clsFillCombos
    PM.FillAnyCombo cmbPaymentMethod, "PaymentMethod", False
    
    cmbType.AddItem "Inward Bills"
    cmbType.AddItem "Green Sheet Bills"
    cmbType.AddItem "OPD Bills"
    cmbType.AddItem "Roentgents Bills"
    cmbType.AddItem "Lab Bills"
    cmbType.AddItem "Pharmacy Bills"
    cmbType.AddItem "Medical Test Bills"
    cmbType.AddItem "Agent Bills"
    cmbType.AddItem "Expence Bills"
    cmbType.AddItem "Health Scheme Supplier Payments"
    cmbType.AddItem "Health Screening Test Bills"
    cmbType.AddItem "All Bills"

End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpFrom.Value = GetSetting(App.EXEName, Me.Name, dtpFrom.Name, "00:00:00")
    dtpTo.Value = GetSetting(App.EXEName, Me.Name, dtpTo.Name, "00:00:00")
    GetCommonSettings Me
    dtpFrom.Value = Date
    dtpTo.Value = Date
    lblTopic.Caption = HospitalName
    lblSubtopic.Caption = "Detiled Cashier Transactions - From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " To " & Format(dtpFrom.Value, "dd MMMM yyyy")
    cmbUser.BoundText = UserID
    cmbType.Text = "All Bills"
    cmbPaymentMethod.BoundText = GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1)
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
    SaveSetting App.EXEName, Me.Name, dtpFrom.Name, dtpFrom.Value
    SaveSetting App.EXEName, Me.Name, dtpTo.Name, dtpTo.Value
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub optBilled_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optCancelled_Click()
    Call FormatGrid
    Call FillGrid

End Sub

Private Sub optPayments_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optRefunds_Click()
    Call FormatGrid
    Call FillGrid
End Sub
