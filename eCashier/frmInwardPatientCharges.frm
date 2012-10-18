VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmInwardPatientCharges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inward Patient Charges"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
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
   ScaleHeight     =   4500
   ScaleWidth      =   7530
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   5760
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.TextBox txtICUNursingRate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtMaintananceRate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtMaintananceCashDiscountRate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtNursingRate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtLaterLinanRate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtInitialLinanRate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtAdmissionRate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Save"
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
   Begin VB.Label Label7 
      Caption         =   "ICU Nursing Rate"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Maintanance Rate"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Maintanance Discount Rate for Cash Patients"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Label Label4 
      Caption         =   "Nursing Rate"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Linan Charge after 3 days"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "First 3 day Linan Charge"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Admission Fee"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmInwardPatientCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    

Private Sub GetData()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblInwardPatientRates.AdmissionRate, tblInwardPatientRates.InitialLinanRate, tblInwardPatientRates.LaterLinanRate, tblInwardPatientRates.MaintananceRate, tblInwardPatientRates.MaintananceCashDiscountRate, tblInwardPatientRates.NursingRate, tblInwardPatientRates.ICUNursingRate " & _
                    "From tblInwardPatientRates"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtAdmissionRate.Text = Format(!AdmissionRate, "0.00")
            txtInitialLinanRate.Text = Format(!InitialLinanRate, "0.00")
            txtLaterLinanRate.Text = Format(!LaterLinanRate, "0.00")
            txtMaintananceRate.Text = Format(!MaintananceRate, "0.00")
            txtMaintananceCashDiscountRate.Text = Format(!MaintananceCashDiscountRate, "0.00")
            txtNursingRate.Text = Format(!NursingRate, "0.00")
            txtICUNursingRate.Text = Format(!ICUNursingRate, "0.00")
        End If
        .Close
    End With
End Sub

Private Sub SaveData()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT * From tblInwardPatientRates"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !AdmissionRate = Val(txtAdmissionRate.Text)
            !InitialLinanRate = Val(txtInitialLinanRate.Text)
            !LaterLinanRate = Val(txtLaterLinanRate.Text)
            !MaintananceRate = Val(txtMaintananceRate.Text)
            !MaintananceCashDiscountRate = Val(txtMaintananceCashDiscountRate.Text)
            !NursingRate = Val(txtNursingRate.Text)
            !ICUNursingRate = Val(txtICUNursingRate.Text)
            .Update
        End If
        .Close
    End With
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Call SaveData
End Sub

Private Sub Form_Load()
    Call GetData
End Sub
