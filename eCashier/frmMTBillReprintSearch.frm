VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMTBillReprintSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medical Test Bill Reprint"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10995
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
   ScaleHeight     =   8625
   ScaleWidth      =   10995
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2040
      TabIndex        =   7
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   9600
      TabIndex        =   4
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
      Left            =   8280
      TabIndex        =   3
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
   Begin MSFlexGridLib.MSFlexGrid gridBill 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11033
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   78512131
      CurrentDate     =   40006
   End
   Begin MSDataListLib.DataCombo cmbUser 
      Height          =   360
      Left            =   2040
      TabIndex        =   8
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnToExcel 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&To Excel"
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
      Caption         =   "Payment Method"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "User"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMTBillReprintSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
    
Private Sub SaveSettings()
    dtpDate.Value = Date
    SaveCommonSettings Me
End Sub

Private Sub GetSettings(): On Error Resume Next
    dtpDate.Value = Date
    GetCommonSettings Me
End Sub

Private Sub FormatGrid()
    With gridBill
        .Rows = 1
        .Cols = 7
        
        .Row = 0
        
        .Col = 0
        .Text = "ID"
        
        .Col = 1
        .Text = "Bill ID"
        
        .Col = 2
        .Text = "Time"
        
        .Col = 3
        .Text = "Patient"
        
        .Col = 4
        .Text = "Payment"
        
        .Col = 5
        .Text = "Value"
        
        .Col = 6
        .Text = "Remarks"
        
        .ColWidth(0) = 0
    End With
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeBill.IncomeBillID,  tblIncomeBill.DisplayBillID, tblIncomeBill.CompletedTime, tblPatientMainDetails.FirstName, tblPaymentMethod.PaymentMethod as MyPM, tblIncomeBill.NetTotal, tblBookedUser.Name as BName, tblCancelledUser.Name as CName, tblIncomeBill.Cancelled, tblIncomeBill.CancelledDate, tblIncomeBill.CancelledTime, tblRefundMethod.PaymentMethod as MyRPM , tblHealthSchemeSuppliers.HealthSchemeSupplierName " & _
                    "FROM  dbo.tblIncomeBill LEFT OUTER JOIN                       dbo.tblHealthSchemeSuppliers ON dbo.tblIncomeBill.HSSID = dbo.tblHealthSchemeSuppliers.HealthSchemeSupplierID LEFT OUTER JOIN                      dbo.tblPaymentMethod ON dbo.tblIncomeBill.PaymentMethodID = dbo.tblPaymentMethod.PaymentMethodID LEFT OUTER JOIN                       dbo.tblStaff tblBookedUser ON dbo.tblIncomeBill.CompletedUserID = tblBookedUser.StaffID LEFT OUTER JOIN " & _
                    "dbo.tblStaff tblCancelledUser ON dbo.tblIncomeBill.CancelledUserID = tblCancelledUser.StaffID LEFT OUTER JOIN                      dbo.tblPaymentMethod tblRefundMethod ON dbo.tblIncomeBill.CancelledPaymentMethodID = tblRefundMethod.PaymentMethodID LEFT OUTER JOIN                      dbo.tblPatientMainDetails ON dbo.tblIncomeBill.PatientID = dbo.tblPatientMainDetails.PatientID " & _
                    "WHERE tblIncomeBill.Completed = 1  AND tblIncomeBill.IsMedicalTestBill = 1  AND tblIncomeBill.CompletedDate = '" & Format(dtpDate.Value, "dd MMMM yyyy") & "' "
        If IsNumeric(cmbUser.BoundText) = True Then temSql = temSql & " AND tblIncomeBill.CompletedUserID = " & Val(cmbUser.BoundText)
        If IsNumeric(cmbPaymentMethod.BoundText) = True Then temSql = temSql & " And tblIncomeBill.PaymentMethodID = " & Val(cmbPaymentMethod.BoundText)
        temSql = temSql & " Order by DisplayBillID"
        
        
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridBill.Rows = gridBill.Rows + 1
            gridBill.Row = gridBill.Rows - 1
            
            gridBill.Col = 0
            gridBill.Text = !IncomeBillID
            
            gridBill.Col = 1
            gridBill.Text = Format(!DisplayBillID)
            
            gridBill.Col = 2
            gridBill.Text = Format(!CompletedTime, "hh:mm AMPM")
            
            gridBill.Col = 3
            gridBill.Text = Format(!FirstName, "")
            
            gridBill.Col = 4
            gridBill.Text = Format(!HealthSchemeSupplierName, "") & " (" & ![MyPM] & ")"
            
            gridBill.Col = 5
            gridBill.Text = Format(!NetTotal, "0.00")
            
            
            gridBill.Col = 6
            If ![Cancelled] = True Then
                gridBill.Text = "Cancelled at " & ![CancelledTime] & " on " & ![CancelledDate] & " by " & ![CName] & "(" & ![MyRPM] & ")"
            End If
        
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub FillCombos()
    Dim Staff As New clsFillCombos
    Staff.FillSpecificFieldBoolCombo cmbUser, "Staff", "Name", "Name", "IsAUser", False
    Dim PM As New clsFillCombos
    PM.FillSpecificFieldBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "PaymentMethod", "CanReceive", False
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    GetPrintDefaults ThisReportFormat
    GridPrint gridBill, ThisReportFormat
    Printer.EndDoc
End Sub

Private Sub btnToExcel_Click()
    GridToExcel gridBill, "Medical Test Bills", Format(dtpDate.Value, "dd MMMM yyyy") & vbTab & cmbUser.Text & vbTab & cmbPaymentMethod.Text
    
End Sub

Private Sub cmbPaymentMethod_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbPaymentMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnPrint.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbPaymentMethod.Text = Empty
    End If
End Sub

Private Sub cmbUser_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPaymentMethod.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbUser.Text = Empty
    End If
End Sub

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbUser.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtpDate.Value = Date
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call GetSettings
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub gridBill_DblClick()
    Dim temBillID As Long
    temBillID = Val(gridBill.TextMatrix(gridBill.Row, 0))
    If temBillID <> 0 Then
        frmMTBillReprint.txtBillID.Text = temBillID
        frmMTBillReprint.Show
    End If
End Sub
