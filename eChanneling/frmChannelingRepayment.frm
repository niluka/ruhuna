VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmChannelingRepayment 
   Caption         =   "Channeling Cancellation / Refund"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameSearch 
      Caption         =   "Search"
      Height          =   975
      Left            =   240
      TabIndex        =   36
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtBookingID 
         Height          =   345
         Left            =   1560
         TabIndex        =   37
         Top             =   360
         Width           =   2655
      End
      Begin btButtonEx.ButtonEx bttnSearchVisitID 
         Height          =   375
         Left            =   4320
         TabIndex        =   38
         Top             =   360
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
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.OptionButton OptionOne 
      Caption         =   "One Print"
      Height          =   255
      Left            =   1560
      TabIndex        =   33
      Top             =   8400
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton OptionTwo 
      Caption         =   "Two Prints"
      Height          =   255
      Left            =   2760
      TabIndex        =   32
      Top             =   8400
      Width           =   1095
   End
   Begin VB.OptionButton OptionNo 
      Caption         =   "No Prints"
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Frame FrameFacilityDetails 
      Caption         =   "Channelling Details"
      Height          =   2895
      Left            =   240
      TabIndex        =   22
      Top             =   1080
      Width           =   6375
      Begin VB.Label Label15 
         Caption         =   "Secession Serial No. :"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblSerial 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   40
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   "Patient Name :"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Doctor :"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Date :"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Secession :"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblName 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   26
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblDoctor 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   25
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblDate 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblSecession 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   23
         Top             =   1800
         Width           =   3135
      End
   End
   Begin VB.Frame frameRepay 
      Caption         =   "Repayment"
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   4080
      Width           =   6375
      Begin VB.TextBox txtRepayComments 
         Height          =   1095
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox txtRepayTotal 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtOtherRepay 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtInstitutionRepay 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtStaffRepay 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4920
         TabIndex        =   1
         Top             =   840
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
      Begin VB.Label lblPreviousTotalRepay 
         Alignment       =   1  'Right Justify
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
      Begin VB.Label lblPreviousOtherRepay 
         Alignment       =   1  'Right Justify
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
      Begin VB.Label lblPreviousInstitutionRepay 
         Alignment       =   1  'Right Justify
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
      Begin VB.Label lblPreviousStaffRepay 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   1935
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
         TabIndex        =   15
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblTotalPaid 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblOtherFeePaid 
         Alignment       =   1  'Right Justify
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
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblInstitutionFeePaid 
         Alignment       =   1  'Right Justify
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
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblStaffFeePaid 
         Alignment       =   1  'Right Justify
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
         Top             =   840
         Width           =   1335
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
         TabIndex        =   10
         Top             =   240
         Width           =   1215
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
         TabIndex        =   9
         Top             =   240
         Width           =   1215
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
         TabIndex        =   8
         Top             =   1800
         Width           =   1695
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
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
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
   End
   Begin btButtonEx.ButtonEx bttnConfirmRepay 
      Height          =   375
      Left            =   3960
      TabIndex        =   34
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "R&epay"
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
      Left            =   5400
      TabIndex        =   35
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Ca&ncel"
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
Attribute VB_Name = "frmChannelingRepayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TemPatientFacilityID As Long

Private Sub bttnSearchbookingID_Click()

TemPatientFacilityID = Empty

If Not IsNumeric(txtBookingID.Text) Then
    Dim TemResponce As Byte
    TemResponce = MsgBox("Please enter a valied booking ID to search", vbCritical, "Wrong ID")
    txtBookingID.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT * from tblpatientfacility where patientfacility_ID = " & Val(txtBookingID.Text)
    If .RecordCount = 0 Then
        TemResponce = MsgBox("There is no such a booking ID in the database. Please recheck", vbCritical, "ID Not found")
        txtBookingID.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If !hospitalfacility_ID <> 10 Then
        TemResponce = MsgBox("There booking ID is not for a channeling. Please recheck", vbCritical, "ID Not for channeling")
        txtBookingID.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If UserAuthority = AuthorityUser Then
        If !PaidToSTaff = True Then
            TemResponce = MsgBox("The money is already paid to the doctor. Therefore no refund can be done by a user. An accountant can pay if it is essential", vbCritical, "Already paid to the doctor")
            txtBookingID.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    End If
    If !cancelled = True Then
        TemResponce = MsgBox("The booking is already cancelled. You can't cancel it again", vbCritical, "Already cancelled")
        txtBookingID.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    If !refund = True Then
        TemResponce = MsgBox("The booking has already repaied. You can't cancel it", vbCritical, "Repaied")
        txtBookingID.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
'    If !fullypaid = True Then
'        TemResponce = MsgBox("The patient has not completed the payment. You can't cancel it", vbCritical, "Repaied")
'        txtBookingID.SetFocus
'        SendKeys "{home}+{end}"
'        Exit Sub
'    End If

    TemPatientFacilityID = !patientfacility_ID
    
    lblSerial.Caption = !DaySerial
    lblName.Caption = FindPatientByID(Val(!patientid))
            
    lblStaffFeePaid.Caption = Format(!Personalfee, "0.00")
    lblInstitutionFeePaid.Caption = Format(!institutionfee, "0.00")
    lblOtherFeePaid.Caption = Format(!otherfee, "0.00")
    
    If Not IsNull(!personalrefund) Then
        lblPreviousStaffRepay.Caption = Format(!personalrefund, "0.00")
    Else
        lblPreviousStaffRepay.Caption = "0.00"
    End If
    If Not IsNull(!InstitutionRefund) Then
        lblPreviousInstitutionRepay.Caption = Format(!InstitutionRefund, "0.00")
    Else
        lblPreviousInstitutionRepay.Caption = "0.00"
    End If
    If Not IsNull(!OtherRefund) Then
        lblPreviousOtherRepay.Caption = Format(!OtherRefund, "0.00")
    Else
        lblPreviousOtherRepay.Caption = "0.00"
    End If
    
    .Close
    
End With

End Sub

Private Sub Form_Load()

End Sub
