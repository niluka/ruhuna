VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmTem 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "SQL Commend"
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton btnDischarge 
      Caption         =   "Reset Balance"
      Height          =   495
      Left            =   1200
      TabIndex        =   20
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4680
      TabIndex        =   19
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   495
      Left            =   8280
      TabIndex        =   18
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton btnDblP 
      Caption         =   "Double Prof"
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   6840
      Width           =   8295
   End
   Begin VB.TextBox txtDID 
      Height          =   495
      Left            =   4200
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDbl 
      Caption         =   "Update"
      Height          =   495
      Left            =   8520
      TabIndex        =   15
      Top             =   6840
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   7440
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txtP 
      Height          =   495
      Left            =   7200
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtH 
      Height          =   495
      Left            =   7200
      TabIndex        =   12
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtT 
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox txtBT 
      Height          =   495
      Left            =   5880
      TabIndex        =   10
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton btnUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   8520
      TabIndex        =   9
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton btnPrevious 
      Caption         =   "<"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton btnNext 
      Caption         =   ">"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton btnRun 
      Caption         =   "Run"
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblCardA 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   6540
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblCashA 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblCreditA 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   4710
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblChequeA 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   8370
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblSlipsA 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   10200
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label42 
      Caption         =   "Total"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "frmTem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsTem As New ADODB.Recordset

Private Sub btnDischarge_Click()
    Dim rsTem As New ADODB.Recordset
    Dim i As Integer
    Dim TemPay As Double
    i = MsgBox("Are you sure you want to discharge this patient", vbYesNo)
    If i = vbNo Then Exit Sub
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where Discharge = 1"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            TemPay = FFillPayments(!BHTID)
            If TemPay <> 0 Then
                If !DOD > #1/31/2010# Then
                    DoEvents
                End If
                If !Payments <> TemPay Then
                    !Payments = TemPay
                End If
                !Balance = !FTotalCharge - TemPay - !Discount
                !FPayments = TemPay
                .Update
            End If
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Function FFillPayments(ThisBHTID As Long)
    Dim TotalPayments As Double
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill where Completed = 1 AND IsInwardPaymentBill = 1 AND Cancelled = 0  AND BHTID = " & ThisBHTID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            TotalPayments = TotalPayments + !NetTotal
            .MoveNext
        Wend
        .Close
    End With
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeBill.IncomeBillID, dbo.tblIncomeBill.DisplayBillID, dbo.tblIncomeBill.NetTotal, dbo.tblIncomeBill.CompletedDate, dbo.tblIncomeBill.CompletedTime, dbo.tblIncomeBill.PaymentComments " & _
                    "FROM         dbo.tblIncomeBill LEFT OUTER JOIN dbo.tblPaymentMethod ON dbo.tblIncomeBill.PaymentMethodID = dbo.tblPaymentMethod.PaymentMethodID " & _
                    "WHERE (dbo.tblIncomeBill.IsHSSPaymentBill = 1) AND (dbo.tblIncomeBill.Completed = 1) AND (dbo.tblIncomeBill.BHTID = " & ThisBHTID & ") AND (dbo.tblIncomeBill.Cancelled = 0) " & _
                    "ORDER BY dbo.tblIncomeBill.CompletedDate, dbo.tblIncomeBill.CompletedTime"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            TotalPayments = TotalPayments + !NetTotal
            .MoveNext
        Wend
        .Close
    End With
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIncomeReturnBill.* FROM tblIncomeReturnBill WHERE tblIncomeReturnBill.BHTID =" & ThisBHTID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            TotalPayments = TotalPayments - !ReturnValue
            .MoveNext
        Wend
        .Close
    End With
    FFillPayments = TotalPayments
End Function



Private Sub btnDblP_Click()
    Dim rsTem As New ADODB.Recordset
    Dim temSql As String
    With rsTem
        If .State = 1 Then .Close
        
            temSql = "Select * from tblPatientService where DeletedDate = '" & Date & "' AND DeletedUserID = " & 180
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                While .EOF = False
                
                !Deleted = False
                !DeletedUserID = UserID
                !DeletedDate = Date
                !DeletedTime = Now
                .Update
                    .MoveNext
                Wend
            .Close
        End If
    End With
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblProfessionalCharges where CancelledDate = '" & Date & "' AND CancelledUserID = " & 180
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !Cancelled = False
            !CancelledDate = Date
            !CancelledTime = Now
            !CancelledDateTime = Now
            !CancelledUserID = UserID
            .Update
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub btnNext_Click()
        rsTem.MoveNext
        Call DD
        
        While Val(txtT.Text) = Val(txtBT.Text) Or Val(txtT.Text) = 0
            rsTem.MoveNext
            Call DD
        Wend
End Sub

Private Sub btnPrevious_Click()
        rsTem.MovePrevious
        Call DD
End Sub

Private Sub btnRun_Click()
    With rsTem
        If .State = 1 Then .Close
        
        temSql = "SELECT     TOP 100 PERCENT dbo.tblIncomeBill.IncomeBillID, dbo.tblIncomeBill.NetTotal, dbo.tblPatientService.Charge, " & _
                      "dbo.tblPatientService.ProfessionalCharge, dbo.tblPatientService.HospitalCharge, dbo.tblIncomeBill.IsOPDBill, dbo.tblIncomeBill.IsLabBill, " & _
                      "dbo.tblIncomeBill.IsPharmacyBill, dbo.tblIncomeBill.IsInwardPaymentBill, dbo.tblIncomeBill.IsMedicalTestBill, dbo.tblIncomeBill.IsGSBill, " & _
                      "dbo.tblIncomeBill.IsAgentBill, dbo.tblIncomeBill.IsOtherBill, dbo.tblIncomeBill.IsRBill, dbo.tblIncomeBill.IsIncomeBill, dbo.tblIncomeBill.IsExpenceBill, " & _
                      "dbo.tblIncomeBill.DisplayBillID " & _
                    "FROM         dbo.tblIncomeBill LEFT OUTER JOIN " & _
                      "dbo.tblPatientService ON dbo.tblIncomeBill.IncomeBillID = dbo.tblPatientService.MedicalTestBillID " & _
                      "Where isOPDBill = 1 " & _
                    "ORDER BY dbo.tblIncomeBill.MedicalTestBillID DESC "
                    
                    
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        
        ProgressBar1.Min = 0
        ProgressBar1.Max = .RecordCount
        
    End With
    
    Call DD
    
End Sub

Private Sub DD()
With rsTem
    txtT.Text = Format(!Charge, "0.00")
    txtH.Text = Format(!HospitalCharge, "0.00")
    txtP.Text = Format(!ProfessionalCharge, "0.00")
    txtBT.Text = Format(!NetTotal, "0.00")
    ProgressBar1.Value = .AbsolutePosition
    txtDID.Text = !DisplayBillID
End With
End Sub

Private Sub btnUpdate_Click()
    With rsTem
    !Charge = !Charge * 2
    !HospitalCharge = !HospitalCharge * 2
    !ProfessionalCharge = !ProfessionalCharge * 2
    End With
    rsTem.Update
    Call DD
End Sub

Private Sub cmdDbl_Click()
    With rsTem
    !Charge = txtT.Text
    !HospitalCharge = txtH.Text
    !ProfessionalCharge = txtP.Text
    End With
    rsTem.Update
    Call DD

End Sub

Private Sub Command1_Click()
    Call PrintThisBill("1111", "Cash", "Buddhika", "10/10/2010", "08:52 PM", "OPD", "Ruhunu Hospital")
    
    Printer.EndDoc
    
End Sub

Private Sub Command3_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        'temSql = "DBCC SHRINKDATABASE (HospitalSQl)"
'        temSql = "DBCC INDEXDEFRAG (HospitalSQl, 'dbo.tblPatientMainDetails', 1)"
        
        temSql = "USE HospitalSQl; " & vbNewLine & _
        "DBCC DBREINDEX ('dbo.tblPatientMainDetails', ' '); "
        
        '.Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        
        cnnStores.Execute temSql
        
        'DBCC DBREINDEX ('Categories')
        'DBCC DBREINDEX ('Categories','Categoryname',80)
        'The following example defragments all partitions of the PK_Product_ProductID index in the Production.Product table in the AdventureWorks database.
        'Copy
        'DBCC INDEXDEFRAG(AdventureWorks, "Production.Product", PK_Product_ProductID)
        'GO

        
        .Close
    End With
End Sub

