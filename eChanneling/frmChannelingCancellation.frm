VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmChannelingCancellation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Channeling Cancellation"
   ClientHeight    =   9345
   ClientLeft      =   555
   ClientTop       =   1335
   ClientWidth     =   6960
   Icon            =   "frmChannelingCancellation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   6960
   Begin VB.OptionButton OptionNo 
      Caption         =   "No Prints"
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   8520
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton OptionTwo 
      Caption         =   "Two Prints"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   9000
      Width           =   1095
   End
   Begin VB.OptionButton OptionOne 
      Caption         =   "One Print"
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   8760
      Width           =   1095
   End
   Begin VB.Frame FrameFacilityDetails 
      Caption         =   "Channelling Details"
      Height          =   3255
      Left            =   240
      TabIndex        =   28
      Top             =   840
      Width           =   6375
      Begin VB.Label lblBookedOn 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   40
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label lblSecession 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   39
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label lblDoctor 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   38
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   37
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Booking Done on :"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Date :"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Doctor :"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Patient Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblSerial 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   32
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label15 
         Caption         =   "Secession Serial No. :"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Secession :"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblAppointmentDate 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   29
         Top             =   1200
         Width           =   3135
      End
   End
   Begin VB.Frame frameRepay 
      Caption         =   "Repayment"
      Height          =   4215
      Left            =   240
      TabIndex        =   22
      Top             =   4200
      Width           =   6375
      Begin VB.OptionButton OptionRepayPatient 
         Caption         =   "Repay Patient"
         Height          =   195
         Left            =   4080
         TabIndex        =   1
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtStaffRepay 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtInstitutionRepay 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtOtherRepay 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtRepayTotal 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtRepayComments 
         Height          =   615
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   3360
         Width           =   4215
      End
      Begin VB.OptionButton OptionRepayAgent 
         Caption         =   "Repay Agent"
         Height          =   195
         Left            =   2040
         TabIndex        =   23
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Doctor Fee :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Institution Fee:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Other Fee:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Paid Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Re-Payment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblStaffFeePaid 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblInstitutionFeePaid 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblOtherFeePaid 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblTotalPaid 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblPreviousStaffRepay 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblPreviousInstitutionRepay 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblPreviousOtherRepay 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblPreviousTotalRepay 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Previous Repays"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame FrameSearch 
      Caption         =   "Search"
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtBookingID 
         Height          =   345
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin btButtonEx.ButtonEx bttnSearchBookingID 
         Height          =   375
         Left            =   4320
         TabIndex        =   26
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Search Booking ID"
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
         Caption         =   "By Booking ID"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
   End
   Begin btButtonEx.ButtonEx bttnConfirmRepay 
      Height          =   375
      Left            =   3120
      TabIndex        =   44
      Top             =   8520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cancel Booking"
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
   Begin btButtonEx.ButtonEx bttnCancelRepay 
      Height          =   375
      Left            =   4920
      TabIndex        =   45
      Top             =   8520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Exit"
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
End
Attribute VB_Name = "frmChannelingCancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim TemPatientFacilityID As Long
    Dim TemPatientID As Long
    Dim TemDoctorID As Long
    Dim TemDoctorFee As Double
    Dim TemInstitutionFee As Double
    Dim TemOtherFee As Double
    Dim TemAgentId As Long

Private Sub bttnCancelRepay_Click()
    Unload Me
End Sub
Private Sub Setcolours()
    bttnCancelRepay.BackColor = BttnBackColour
    bttnCancelRepay.ForeColor = BttnForeColour
    bttnConfirmRepay.BackColor = BttnBackColour
    bttnConfirmRepay.ForeColor = BttnForeColour
    bttnSearchBookingID.BackColor = BttnBackColour
    bttnSearchBookingID.ForeColor = BttnForeColour
    frmChannelingCancellation.BackColor = FrmBackColour
    frmChannelingCancellation.ForeColor = FrmForeColour
    FrameSearch.BackColor = FrmBackColour
    FrameSearch.ForeColor = FrmForeColour
    FrameFacilityDetails.BackColor = FrmBackColour
    FrameFacilityDetails.ForeColor = FrmForeColour
    frameRepay.BackColor = FrmBackColour
    frameRepay.ForeColor = FrmForeColour
    OptionTwo.BackColor = FrmBackColour
    OptionTwo.ForeColor = FrmForeColour
    OptionOne.BackColor = FrmBackColour
    OptionOne.ForeColor = FrmForeColour
    
    OptionRepayAgent.BackColor = FrmBackColour
    OptionRepayAgent.ForeColor = FrmForeColour
    
    OptionRepayPatient.BackColor = FrmBackColour
    OptionRepayPatient.ForeColor = FrmForeColour
    
    
    
    OptionNo.BackColor = FrmBackColour
    OptionNo.ForeColor = FrmForeColour
    Label1.BackColor = LblBackColour
    Label1.ForeColor = LblForeColour
    Label10.BackColor = LblBackColour
    Label10.ForeColor = LblForeColour
    Label11.BackColor = LblBackColour
    Label11.ForeColor = LblForeColour
    Label12.BackColor = LblBackColour
    Label12.ForeColor = LblForeColour
    Label13.BackColor = LblBackColour
    Label13.ForeColor = LblForeColour
    Label14.BackColor = LblBackColour
    Label14.ForeColor = LblForeColour
    Label15.BackColor = LblBackColour
    Label15.ForeColor = LblForeColour
    Label2.BackColor = LblBackColour
    Label2.ForeColor = LblForeColour
    Label3.BackColor = LblBackColour
    Label3.ForeColor = LblForeColour
    Label4.BackColor = LblBackColour
    Label4.ForeColor = LblForeColour
    Label4.BackColor = LblBackColour
    Label4.ForeColor = LblForeColour
    Label5.BackColor = LblBackColour
    Label5.ForeColor = LblForeColour
    Label6.BackColor = LblBackColour
    Label6.ForeColor = LblForeColour
    Label7.BackColor = LblBackColour
    Label7.ForeColor = LblForeColour
    Label8.BackColor = LblBackColour
    Label8.ForeColor = LblForeColour
    Label9.BackColor = LblBackColour
    Label9.ForeColor = LblForeColour
End Sub

Private Sub bttnConfirmRepay_Click()
    Dim TemResponce  As Integer
    If Val(lblPreviousTotalRepay.Caption) + Val(txtRepayTotal.Text) > Val(lblTotalPaid.Caption) Then
        TemResponce = MsgBox("You can't repay an amount grater than that paid initially by the patient", vbCritical, "Exceeds Payment")
        txtStaffRepay.SetFocus
        Exit Sub
    End If
    If OptionRepayAgent.Value = False And OptionRepayPatient.Value = False And OptionRepayAgent.Visible = True And OptionRepayPatient.Visible = True Then
        TemResponce = MsgBox("You have not selected wether to repay the patient or the agent. Please select one.", vbQuestion, "Repay to whom?")
        Exit Sub
    End If

    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "select * from tblpatientrepay"
        If .State = 0 Then .Open
        .AddNew
        !patient_ID = TemPatientID
        !HospitalFacility_ID = 10
        !repayUser_ID = UserID
        !repaydate = Date
        !repaytime = Time
        !StaffRepay = Val(txtStaffRepay.Text)
        !InstitutionRepay = Val(txtInstitutionRepay.Text)
        !OtherRepay = Val(txtOtherRepay.Text)
        !TotalRepay = Val(txtRepayTotal.Text)
        !Staff_ID = TemDoctorID
        !isadoctor = True
        If Trim(txtRepayComments.Text) = "" Then
            !repaycomments = "Cancellation"
        Else
            !repaycomments = txtRepayComments.Text
        End If
        !patientfacility_ID = TemPatientFacilityID
        
        If OptionRepayAgent.Value = True Then
            !RefundToAgent = 1
            !RefundToPatient = False
        ElseIf OptionRepayPatient = True Then
            !RefundToAgent = False
            !RefundToPatient = 1
        Else
            !RefundToAgent = False
            !RefundToPatient = 1
        End If
        
        
        .Update
        .Close
        .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & TemPatientFacilityID
        If .State = 0 Then .Open
        If .RecordCount = 0 Then Exit Sub
            If IsNull(!Personalrefund) Then
                !personaldue = !personalfee - Val(txtStaffRepay.Text)
                !Personalrefund = Val(txtStaffRepay.Text)
            Else
                !personaldue = !personalfee - (Val(!Personalrefund) + Val(txtStaffRepay.Text))
                !Personalrefund = Val(!Personalrefund) + Val(txtStaffRepay.Text)
            End If
            If IsNull(!institutionrefund) Then
                !institutiondue = !institutionfee - Val(txtInstitutionRepay.Text)
                !institutionrefund = Val(txtInstitutionRepay.Text)
            Else
                !institutiondue = !institutionfee - (Val(!institutionrefund) + Val(txtInstitutionRepay.Text))
                !institutionrefund = Val(!institutionrefund) + Val(txtInstitutionRepay.Text)
            End If
            If IsNull(!otherrefund) Then
                !otherdue = !otherfee - Val(txtOtherRepay.Text)
                !otherrefund = Val(txtOtherRepay.Text)
            Else
                !otherdue = !otherfee - (Val(!otherrefund) + Val(txtOtherRepay.Text))
                !otherrefund = Val(!otherrefund) + Val(txtOtherRepay.Text)
            End If
            If IsNull(!totalrefund) Then
                !TotalDue = !totalfee - Val(txtRepayTotal.Text)
                !totalrefund = Val(txtRepayTotal.Text)
            Else
                !TotalDue = !totalfee - (Val(!totalrefund) + Val(txtRepayTotal.Text))
                !totalrefund = Val(!totalrefund) + Val(txtRepayTotal.Text)
            End If
            If Trim(txtRepayComments.Text) = "" Then
                !repaycomments = "Cancellation"
            Else
                !repaycomments = txtRepayComments.Text
            End If
            !repaydate = Date
            !repaytime = Time
            !cancelled = True
            !cancellednull = 1
            !repayUser_ID = UserID
            If OptionRepayAgent.Value = True Then
                !RefundToAgent = 1
                !RefundToPatient = False
            ElseIf OptionRepayPatient = True Then
                !RefundToAgent = False
                !RefundToPatient = 1
            Else
                !RefundToAgent = False
                !RefundToPatient = 1
            End If
            .Update
       
        .Close
    
        If OptionRepayAgent.Value = True Then
            If .State = 1 Then .Close
            .Source = "SELECT tblinstitutions.* from tblinstitutions where institution_ID =" & TemAgentId
            If .State = 0 Then .Open
            If .RecordCount = 0 Then Exit Sub
            !InstitutionCredit = !InstitutionCredit + Val(TemDoctorFee + TemOtherFee + TemInstitutionFee)
            .Update
            .Close
        End If
    
    
    End With
    
    If OptionOne.Value = True Then
        PrintOne
    ElseIf OptionTwo.Value = True Then
        PrintOne
        PrintOne
    End If

    ClearValues
    
    Unload Me

End Sub


Private Sub ClearValues()
    TemPatientFacilityID = Empty
    TemPatientID = Empty
    lblAppointmentDate.Caption = Empty
    lblBookedOn.Caption = Empty
    lblDoctor.Caption = Empty
    lblInstitutionFeePaid.Caption = Empty
    lblName.Caption = Empty
    lblOtherFeePaid.Caption = Empty
    lblPreviousInstitutionRepay.Caption = Empty
    lblPreviousOtherRepay.Caption = Empty
    lblPreviousStaffRepay.Caption = Empty
    lblPreviousTotalRepay.Caption = Empty
    lblSecession.Caption = Empty
    lblSerial.Caption = Empty
    lblStaffFeePaid.Caption = Empty
    lblTotalPaid.Caption = Empty
    txtBookingID.Text = Empty
    txtInstitutionRepay.Text = Empty
    txtOtherRepay.Text = Empty
    txtRepayComments.Text = Empty
    txtRepayTotal.Text = Empty
    txtStaffRepay.Text = Empty
End Sub

Private Sub bttnSearchbookingID_Click()

TemPatientFacilityID = Empty

If Not IsNumeric(txtBookingID.Text) Then
    Dim TemResponce  As Integer
    TemResponce = MsgBox("Please enter a valied booking ID to search", vbCritical, "Wrong ID")
    txtBookingID.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(txtBookingID.Text)
    .Open
    If .RecordCount = 0 Then
        TemResponce = MsgBox("There is no such a booking ID in the database. Please recheck", vbCritical, "ID Not found")
        txtBookingID.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If !HospitalFacility_ID <> 10 Then
        TemResponce = MsgBox("There booking ID is not for a channeling. Please recheck", vbCritical, "ID Not for channeling")
        txtBookingID.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If UserAuthority = AuthorityUser Then
        If !paidtostaff = True Then
            TemResponce = MsgBox("The money is already paid to the doctor. Therefore no refund can be done by a user. An accountant can pay if it is essential", vbCritical, "Already paid to the doctor")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    Else
        If !paidtostaff = True Then
            TemResponce = MsgBox("The money is already paid to the doctor. Are you sure you want to refund ?", vbCritical + vbYesNo, "Already paid to the doctor")
            If TemResponce = vbNo Then
                txtBookingID.SetFocus
                SendKeys "{home}+{end}"
                Exit Sub
            End If
        End If
    End If
    If !cancelled = True Then
        TemResponce = MsgBox("The booking is already cancelled. You can't cancel it again", vbCritical, "Already cancelled")
        txtBookingID.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If !Refund = True Then
        TemResponce = MsgBox("The booking has already repaied. You can't cancel it", vbCritical, "Repaied")
        txtBookingID.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If !FullyPaid = 0 Then
        TemResponce = MsgBox("The patient has not completed the payment. You can't cancel it", vbCritical, "Repaied")
        txtBookingID.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If !PaymentMode = "Agent" Then
        OptionRepayAgent.Visible = True
        OptionRepayPatient.Visible = True
        OptionRepayAgent.Value = False
        OptionRepayPatient.Value = False
    Else
        OptionRepayAgent.Visible = False
        OptionRepayPatient.Visible = False
        OptionRepayAgent.Value = False
        OptionRepayPatient.Value = False
    End If

    TemPatientFacilityID = !patientfacility_ID
    TemPatientID = !patientid
    TemDoctorID = !Staff_ID
    TemAgentId = !Agent_ID
    
    lblSerial.Caption = !DaySerial
    lblName.Caption = FindPatientByID(Val(!patientid))
    lblDoctor.Caption = FindDoctorFromID(!Staff_ID)
    lblAppointmentDate.Caption = Format(!AppointmentDate, DefaultLongDate)
    
    lblStaffFeePaid.Caption = Format(!personalfee, "0.00")
    lblInstitutionFeePaid.Caption = Format(!institutionfee, "0.00")
    lblOtherFeePaid.Caption = Format(!otherfee, "0.00")
    lblSecession.Caption = FindSecessionFromID(!Secession)
    lblBookedOn.Caption = Format(!BookingDate, DefaultLongDate)
    
    If Not IsNull(!Personalrefund) Then
        lblPreviousStaffRepay.Caption = Format(!Personalrefund, "0.00")
    Else
        lblPreviousStaffRepay.Caption = "0.00"
    End If
    If Not IsNull(!institutionrefund) Then
        lblPreviousInstitutionRepay.Caption = Format(!institutionrefund, "0.00")
    Else
        lblPreviousInstitutionRepay.Caption = "0.00"
    End If
    If Not IsNull(!otherrefund) Then
        lblPreviousOtherRepay.Caption = Format(!otherrefund, "0.00")
    Else
        lblPreviousOtherRepay.Caption = "0.00"
    End If
    
    .Close
    
End With

    txtStaffRepay.Text = lblStaffFeePaid.Caption
    txtInstitutionRepay.Text = lblInstitutionFeePaid.Caption
    
    bttnConfirmRepay.SetFocus
    
End Sub


Private Sub bttnSearchBookingID_KeyPress(ByVal KeyAscii As Integer)
    If KeyAscii = 13 Then txtStaffRepay.SetFocus
    
End Sub

Private Sub Form_Load()
If SetPrinter = False Then Unload Me: Exit Sub

    Me.Width = 7080
    Me.Height = 9390
    Me.Top = 0 ' (Screen.Height / 2) - (Me.Height / 2)
    Me.Left = (Screen.Width / 2) - (Me.Width / 3)
    Call Setcolours
End Sub

Private Sub txtBookingID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then bttnSearchbookingID_Click
End Sub

Private Function SetPrinter() As Boolean
SetPrinter = False
Dim MyPrinter As Printer

For Each MyPrinter In Printers
    If MyPrinter.DeviceName = ReportPrinterName Then
        Set Printer = MyPrinter
        SetPrinter = True
    End If
Next

If SetPrinter = False Then
        Dim TemResponce  As Integer
        TemResponce = MsgBox("You have not selected a valied printer for bill printing, Please select a printer", vbCritical, "No printer")
        frmPrintingPreferances.Show
        frmPrintingPreferances.ZOrder 0
        frmPrintingPreferances.SSTab1.Tab = 1
        frmPrintingPreferances.ComboBillPrinter.SetFocus
End If


End Function


Private Sub PrintOne()


With Printer
        
        .Font = "Bernard MT Condensed"
        Printer.Print
        .FontSize = 14
        Printer.Print 'Tab(2); InstitutionName
        .FontSize = 12
        Printer.Print ' Tab(3); InstitutionAddress
        Printer.Print 'Tab(3); InstitutionTelephone
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        
        .FontName = "Courier"
        .FontSize = 10
        Printer.Print
        
        Dim TemTab1 As Long
        Dim TemTab2 As Long
        Dim TemTab3 As Long
        Dim TemTab4 As Long
        Dim TemTab5 As Long
        Dim TemTab6 As Long
        
        TemTab1 = 2
        TemTab2 = 6
        TemTab3 = 35
        TemTab4 = 15
        TemTab5 = 36
        TemTab6 = 30
        Printer.Print Tab(TemTab4); "Cancellation"
        Printer.Print
       
        Printer.Print Tab(TemTab1); "Patient";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); lblName.Caption
        Printer.Print
        Printer.Print Tab(TemTab1); "Consultant";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); lblDoctor.Caption
        Printer.Print
        Printer.Print Tab(TemTab1); "Appo. Date ";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); lblAppointmentDate.Caption
        
        Printer.Print Tab(TemTab1); "Appo. No.";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); lblSerial.Caption
        
        Printer.Print Tab(TemTab1); "Appo. ID";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3); TemPatientFacilityID
        Printer.Print
        Printer.Print Tab(TemTab1); "Doctor Fee Refund";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3 + 8 - Len(Format(txtStaffRepay.Text, "0.00"))); Format(txtStaffRepay.Text, "0.00")
        Printer.Print Tab(TemTab1); "Hospital Fee Refund";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3 + 8 - Len(Format(txtInstitutionRepay.Text, "0.00"))); Format(txtInstitutionRepay.Text, "0.00")
        Printer.Print Tab(TemTab1); "Total Refund";
        Printer.Print Tab(TemTab6); " : ";
        Printer.Print Tab(TemTab3 + 8 - Len(Format(txtInstitutionRepay.Text + txtStaffRepay.Text, "0.00"))); Format(txtInstitutionRepay.Text + txtStaffRepay.Text, "0.00")
        Printer.Print
        Printer.Print Tab(TemTab2); "--------------------"
        Printer.Print Tab(TemTab2); UserName
        Printer.Print Tab(TemTab2); Format(Date, DefaultShortDate)
                
        .EndDoc
    End With
End Sub


Private Sub txtInstitutionRepay_Change()
    Call CalculateRefundTotals
End Sub

Private Sub txtInstitutionRepay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOtherRepay.SetFocus
    
End Sub

Private Sub txtOtherRepay_Change()
    Call CalculateRefundTotals
End Sub

Private Sub txtOtherRepay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRepayTotal.SetFocus
    
End Sub

Private Sub txtStaffRepay_Change()
    Call CalculateRefundTotals
End Sub

Private Sub CalculateRefundTotals()
    txtRepayTotal.Text = Format((Val(txtStaffRepay.Text) + Val(txtInstitutionRepay.Text) + Val(txtOtherRepay.Text)), "0.00")
    lblTotalPaid.Caption = Format((Val(lblStaffFeePaid.Caption) + Val(lblInstitutionFeePaid.Caption) + Val(lblOtherFeePaid.Caption)), "0.00")
    lblPreviousTotalRepay.Caption = Format((Val(lblPreviousStaffRepay.Caption) + Val(lblPreviousInstitutionRepay.Caption) + Val(lblPreviousOtherRepay.Caption)), "0.00")
End Sub

Private Sub txtStaffRepay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtInstitutionRepay.SetFocus
End Sub
