VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmHospitalCharges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hospital Charges"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
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
   ScaleHeight     =   4665
   ScaleWidth      =   5865
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
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
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   6
      Left            =   3600
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   5
      Left            =   3600
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   4
      Left            =   3600
      TabIndex        =   9
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   240
      Index           =   6
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   555
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   555
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   555
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmHospitalCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTem As New ADODB.Recordset
    Dim temSql As String
    
Private Sub FillLables()
    lblValue(0).Caption = "Admission Rate"
    lblValue(1).Caption = "Initial Linan Rate"
    lblValue(2).Caption = "Later Linan Rate"
    lblValue(3).Caption = "Maintanance Rate"
    lblValue(4).Caption = "Maintanance Cash Discount Rate"
    lblValue(5).Caption = "Normal Nursing Rate"
    lblValue(6).Caption = "ICU Nursing Rate"
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call FillLables
    Call FillValues
End Sub

Private Sub FillValues()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT     AdmissionRate, InitialLinanRate, LaterLinanRate, MaintananceRate, MaintananceCashDiscountRate, NursingRate, ICUNursingRate FROM         dbo.tblInwardPatientRates"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtValue(0).Text = !AdmissionRate
            txtValue(1).Text = !InitialLinanRate
            txtValue(2).Text = !LaterLinanRate
            txtValue(3).Text = !MaintananceRate
            txtValue(4).Text = !MaintananceCashDiscountRate
            txtValue(5).Text = !NursingRate
            txtValue(6).Text = !ICUNursingRate
        End If
        .Close
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveValues
End Sub

Private Sub SaveValues()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT     AdmissionRate, InitialLinanRate, LaterLinanRate, MaintananceRate, MaintananceCashDiscountRate, NursingRate, ICUNursingRate FROM         dbo.tblInwardPatientRates"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !AdmissionRate = Val(txtValue(0).Text)
            !InitialLinanRate = Val(txtValue(1).Text)
            !LaterLinanRate = Val(txtValue(2).Text)
            !MaintananceRate = Val(txtValue(3).Text)
            !MaintananceCashDiscountRate = Val(txtValue(4).Text)
            !NursingRate = Val(txtValue(5).Text)
            !ICUNursingRate = Val(txtValue(6).Text)
            .Update
        End If
        .Close
    End With
End Sub

