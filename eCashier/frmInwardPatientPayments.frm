VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInwardPatientPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inward Patient Payments"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
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
   ScaleHeight     =   4440
   ScaleWidth      =   7590
   Begin VB.CheckBox Check1 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtPayment 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   3360
      Width           =   2655
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2040
      TabIndex        =   11
      Top             =   3360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   2880
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16187394
      CurrentDate     =   39962
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16187395
      CurrentDate     =   39962
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtDetails 
      Height          =   975
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   4815
   End
   Begin MSDataListLib.DataCombo cmbBHT 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
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
      Left            =   6240
      TabIndex        =   14
      Top             =   3840
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
   Begin btButtonEx.ButtonEx btnPay 
      Height          =   495
      Left            =   4920
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Pay"
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
   Begin VB.Label Label6 
      Caption         =   "&Time"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Payment &Method"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "&Value"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Date"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&BHT"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "&Details"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmInwardPatientPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim MyBHT As New clsBHT
    Dim rsBHT As New ADODB.Recordset
    
Private Sub FillCombos()
    With rsBHT
        If .State = 1 Then .Close
        temSql = "Select * from tblBHT where IsBHT = true And Discharge = False order by BHT"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbBHT
        Set .RowSource = rsBHT
        .ListField = "BHT"
        .BoundColumn = "BHTID"
    End With
    Dim Pay As New clsFillCombos
    Pay.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "CanReceive", False
End Sub

Private Sub GetSettings()
    dtpDate.Value = Date
    dtpTime.Value = Time
End Sub
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPay_Click()
    If IsNumeric(cmbBHT.BoundText) = False Then
        MsgBox "BHT?"
        cmbBHT.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Payment Method?"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    If Val(txtValue.Text) = 0 Then
        MsgBox "Value?"
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    Dim rstem As New ADODB.Recordset
    With rstem
        If .State = 1 Then .Close
        temSql = "Select * from tblIncomeBill"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !BHTID = Val(cmbBHT.BoundText)
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        !Date = dtpDate.Value
        !Time = dtpTime.Value
        !UserID = UserID
        !StoreID = UserStoreID
        !GrossTotal = Val(txtValue.Text)
        !Completed = True
        !CompletedDate = Date
        !CompletedTime = Time
        !CompletedUserID = UserID
        !IsInwardPaymentBill = True
        !NetTotal = Val(txtValue.Text)
        .Update
        .Close
    End With
    Call ClearValues
    cmbBHT.SetFocus
End Sub

Private Sub ClearValues()
    txtPayment.Text = Empty
    cmbBHT.Text = Empty
    txtValue.Text = Empty
    txtDetails.Text = Empty
    dtpDate.Value = Date
    dtpTime.Value = Time
End Sub

Private Sub cmbBHT_Change()
    If IsNumeric(cmbBHT.BoundText) = False Then Exit Sub
    MyBHT.BHTID = Val(cmbBHT.BoundText)
    Call DisplayDetails
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call GetSettings
    
End Sub

Private Sub DisplayDetails()
    Dim temText As String
    temText = "Name : " & MyBHT.FirstName & vbNewLine
    temText = temText & "Age : " & MyBHT.AgeInWords & vbNewLine
    temText = temText & "Admitted : " & Format(MyBHT.DOA, "dd MMMM yyyy") & " at " & Format(MyBHT.TOA, "HH:MM AMPM") & vbNewLine
    If MyBHT.Discharge = True Then
        temText = temText & "Discharged :" & Format(MyBHT.DOD, "dd MMMM yyyy") & " at " & Format(MyBHT.TOD, "HH:MM AMPM")
    Else
        temText = temText & "Not yet discharged"
    End If
    txtDetails.Text = temText
End Sub

