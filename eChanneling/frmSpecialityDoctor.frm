VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSpecialtyDoctor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Speciality"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   ClipControls    =   0   'False
   Icon            =   "frmSpecialityDoctor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11280
   Begin VB.TextBox txtSearch 
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   4335
   End
   Begin VB.Frame FrameIx 
      Caption         =   "Speciality"
      Height          =   3375
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtName 
         Height          =   345
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   4215
      End
      Begin btButtonEx.ButtonEx bttnSave 
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "Sa&ve"
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
      Begin btButtonEx.ButtonEx bttnChange 
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         Appearance      =   3
         Caption         =   "C&hange"
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
      Begin btButtonEx.ButtonEx bttnCancel 
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Speciality Name"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx bttnAdd 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "A&dd"
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
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5655
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9975
      _Version        =   393216
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
   End
   Begin btButtonEx.ButtonEx bttnEdit 
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "Ed&it"
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   6360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Cl&ose"
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
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Appearance      =   3
      Caption         =   "&Delete"
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
Attribute VB_Name = "frmSpecialtyDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FromGrid As Boolean
    Dim TemSpecialityId As Long
    Dim BorderMargin As Long
Private Sub SetColour()



Select Case ColourScheme

Case 1:

BttnBackColour = 5341695
BttnForeColour = 1314458
FrmBackColour = 11066623
FrmForeColour = 1314458
FrameBackColour = 11066623
FrameForeColour = 1314458
TxtBackColour = 9881851
TxtForeColour = 1314458
LblBackColour = 11066623
LblForeColour = 1314458



GridBackColor = 9881855
GridBackColorBkg = 10474239
GridBackColorFixed = 8566015
GridBackColorSel = 5341695

GridForeColor = 1314458
GridForeColorFixed = 11944
GridForeColorSel = 3014824

'GridCellBackColor = 5853695
'GridCellForeColor = 658120


Case 2:

BttnBackColour = 14803300
BttnForeColour = 5539362
FrmBackColour = 16766120
FrmForeColour = 5539362
FrameBackColour = 16766120
FrameForeColour = 5539362
TxtBackColour = 16760450
TxtForeColour = 5539362
LblBackColour = 16766120
LblForeColour = 5539362

GridBackColor = 16760450
GridBackColorBkg = 16771260
GridBackColorFixed = 16105620
GridBackColorSel = 16737380

GridForeColor = 5539362
GridForeColorFixed = 5539362
GridForeColorSel = 16765588


Case 3:

BttnBackColour = 51455
BttnForeColour = 942490
FrmBackColour = 11070719
FrmForeColour = 942490
FrameBackColour = 11070719
FrameForeColour = 942490
TxtBackColour = 11528439
TxtForeColour = 1314458
LblBackColour = 11070719
LblForeColour = 942490

GridBackColor = 16760450
GridBackColorBkg = 16771260
GridBackColorFixed = 16105620
GridBackColorSel = 16737380

GridForeColor = 5539362
GridForeColorFixed = 5539362
GridForeColorSel = 16765588

End Select

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour

bttnAdd.BackColor = BttnBackColour
bttnAdd.ForeColor = BttnForeColour

'bttnCancel.BackColor = BttnBackColour
'bttnAddLeave.ForeColor = BttnForeColour

bttnCancel.BackColor = BttnBackColour
bttnCancel.ForeColor = BttnForeColour

bttnChange.BackColor = BttnBackColour
bttnChange.ForeColor = BttnForeColour

bttnClose.BackColor = BttnBackColour
bttnClose.ForeColor = BttnForeColour

bttnDelete.BackColor = BttnBackColour
bttnDelete.ForeColor = BttnForeColour

bttnEdit.BackColor = BttnBackColour
bttnEdit.ForeColor = BttnForeColour

'bttnLeaveDelete.BackColor = BttnBackColour
'bttnLeaveDelete.ForeColor = BttnForeColour

bttnSave.BackColor = BttnBackColour
bttnSave.ForeColor = BttnForeColour

'bttnChange.BackColor = BttnBackColour
'bttnChange.ForeColor = BttnForeColour


'OptionNoSecessions.BackColor = TxtBackColour
'OptionNoSecessions.ForeColor = TxtForeColour
'
'OptionTwoSecessions.BackColor = TxtBackColour
'OptionTwoSecessions.ForeColor = TxtForeColour

'bttnRemove.BackColor = BttnBackColour
'bttnRemove.ForeColor = BttnForeColour


frmSpecialtyDoctor.BackColor = FrmBackColour
frmSpecialtyDoctor.ForeColor = FrmForeColour

FrameIx.BackColor = FrameBackColour
FrameIx.ForeColor = FrameForeColour

'FrameAgent.BackColor = FrameBackColour
'FrameAgent.ForeColor = FrameForeColour

'FrameBooking.BackColor = FrameBackColour
'FrameBooking.ForeColor = FrameForeColour

'FrameCash.BackColor = FrameBackColour
'FrameCash.ForeColor = FrameForeColour

'FrameCheque.BackColor = FrameBackColour
'FrameCheque.ForeColor = FrameForeColour
'FrameCredit.BackColor = FrameBackColour
'FrameCredit.ForeColor = FrameForeColour
'FrameCreditCard.BackColor = FrameBackColour
'FrameCreditCard.ForeColor = FrameForeColour
'FramePatient.BackColor = FrameBackColour
'FramePatient.ForeColor = FrameForeColour
'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour
'
'FramePaymentMethod.BackColor = FrameBackColour
'FramePaymentMethod.ForeColor = FrameForeColour
'frameSearchPatient.BackColor = FrameBackColour
'frameSearchPatient.ForeColor = FrameForeColour
'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour



'chkEvening.BackColor = LblBackColour
'chkEvening.ForeColor = LblForeColour
'
'chkFridayEvening.BackColor = LblBackColour
'chkFridayEvening.ForeColor = LblForeColour
'
'chkMondayEvening.BackColor = LblBackColour
'chkMondayEvening.ForeColor = LblForeColour
'
'chkMondayFullLeave.BackColor = LblBackColour
'chkMondayFullLeave.ForeColor = LblForeColour
'
'chkMondayMorning.BackColor = LblBackColour
'chkMondayMorning.ForeColor = LblForeColour
'
'chkMoning.BackColor = LblBackColour
'chkMoning.ForeColor = LblForeColour
'
'chkSaturdayEvening.BackColor = LblBackColour
'chkSaturdayEvening.ForeColor = LblForeColour
'
'chkSaturdayFullLeave.BackColor = LblBackColour
'chkSaturdayFullLeave.ForeColor = LblForeColour
'
'chkSaturdayMorning.BackColor = LblBackColour
'chkSaturdayMorning.ForeColor = LblForeColour
'
'chkSundayEvening.BackColor = LblBackColour
'chkSundayEvening.ForeColor = LblForeColour
'
'chkSundayFullLeave.BackColor = LblBackColour
'chkSundayFullLeave.ForeColor = LblForeColour
'
'chkSundayEvening.BackColor = LblBackColour
'chkSundayEvening.ForeColor = LblForeColour
'
'chkSundayMorning.BackColor = LblBackColour
'chkSundayMorning.ForeColor = LblForeColour
'
'chkThursdayEvening.BackColor = LblBackColour
'chkThursdayEvening.ForeColor = LblForeColour
'
'chkThursdayFullLeave.BackColor = LblBackColour
'chkThursdayFullLeave.ForeColor = LblForeColour
'
'chkThursdayEvening.BackColor = LblBackColour
'chkThursdayEvening.ForeColor = LblForeColour
'
'chkThursdayMorning.BackColor = LblBackColour
'chkThursdayMorning.ForeColor = LblForeColour
'
'chkTuesdayEvening.BackColor = LblBackColour
'chkTuesdayEvening.ForeColor = LblForeColour
'
'chkTuesdayFullLeave.BackColor = LblBackColour
'chkTuesdayFullLeave.ForeColor = LblForeColour
'
'chkTuesdayEvening.BackColor = LblBackColour
'chkTuesdayEvening.ForeColor = LblForeColour
'
'chkTuesdayMorning.BackColor = LblBackColour
'chkTuesdayMorning.ForeColor = LblForeColour
'
'chkWednesdayEvening.BackColor = LblBackColour
'chkWednesdayEvening.ForeColor = LblForeColour
'
'chkWednesdayFullLeave.BackColor = LblBackColour
'chkWednesdayFullLeave.ForeColor = LblForeColour
'
'chkWednesdayEvening.BackColor = LblBackColour
'chkWednesdayEvening.ForeColor = LblForeColour
'
'chkWednesdayMorning.BackColor = LblBackColour
'chkWednesdayMorning.ForeColor = LblForeColour

'DataComboDoctorStaff.BackColor = TxtBackColour
'DataComboDoctorStaff.ForeColor = TxtForeColour
'
'DataComboFacility.BackColor = TxtBackColour
'DataComboFacility.ForeColor = TxtForeColour

'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboDoctorStaff.BackColor = TxtBackColour
'DataComboDoctorStaff.ForeColor = TxtForeColour
'
'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour
'
'DataCombo.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour




grid1.BackColor = GridBackColor
grid1.ForeColor = GridForeColor

grid1.BackColorBkg = GridBackColorBkg
grid1.BackColorFixed = GridBackColorFixed
grid1.BackColorSel = GridBackColorSel

grid1.ForeColor = GridForeColor
grid1.ForeColorFixed = GridForeColorFixed
grid1.ForeColorSel = GridForeColorSel

grid1.ForeColor = GridForeColor




'Label1.BackColor = LblBackColour
'Label1.ForeColor = LblForeColour

'LblDoctorStaff.BackColor = LblBackColour
'LblDoctorStaff.ForeColor = LblForeColour
'lblInstitutionFee.BackColor = LblBackColour
'lblInstitutionFee.ForeColor = LblForeColour
'Lbl.BackColor = LblBackColour
'LblCommentsLX.ForeColor = LblForeColour
'Label13.BackColor = LblBackColour
'Label13.ForeColor = LblForeColour
'Label14.BackColor = LblBackColour
'Label14.ForeColor = LblForeColour
'Label15.BackColor = LblBackColour
'Label15.ForeColor = LblForeColour
'Label16.BackColor = LblBackColour
'Label16.ForeColor = LblForeColour
'Label2.BackColor = LblBackColour
'Label2.ForeColor = LblForeColour
'Label18.BackColor = LblBackColour
'Label18.ForeColor = LblForeColour
'Label3.BackColor = LblBackColour
'Label3.ForeColor = LblForeColour
'Label20.BackColor = LblBackColour
'Label20.ForeColor = LblForeColour
'Label21.BackColor = LblBackColour
'Label21.ForeColor = LblForeColour
'Label4.BackColor = LblBackColour
'Label4.ForeColor = LblForeColour
'Label23.BackColor = LblBackColour
'Label23.ForeColor = LblForeColour
'Label24.BackColor = LblBackColour
'Label24.ForeColor = LblForeColour
'Label25.BackColor = LblBackColour
'Label25.ForeColor = LblForeColour
'Label26.BackColor = LblBackColour
'Label26.ForeColor = LblForeColour
'Label27.BackColor = LblBackColour
'Label27.ForeColor = LblForeColour
'Label4.BackColor = LblBackColour
'Label4.ForeColor = LblForeColour
'Label5.BackColor = LblBackColour
'Label5.ForeColor = LblForeColour
'Label6.BackColor = LblBackColour
'Label6.ForeColor = LblForeColour
'Label7.BackColor = LblBackColour
'Label7.ForeColor = LblForeColour
'
'Label8.BackColor = LblBackColour
'Label8.ForeColor = LblForeColour
'Label9.BackColor = LblBackColour
'Label9.ForeColor = LblForeColour
'
'lblAmount.BackColor = LblBackColour
'lblAmount.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour
'
'lblCashPaid.BackColor = LblBackColour
'lblCashPaid.ForeColor = LblForeColour
'
'lblChequeAmount.BackColor = LblBackColour
'lblChequeAmount.ForeColor = LblForeColour
'
'lblThisTimeCredit.BackColor = LblBackColour
'lblThisTimeCredit.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour
'
'lblCashBalance.BackColor = LblBackColour
'lblCashBalance.ForeColor = LblForeColour

'chkFridayFullLeave.BackColor = FrameBackColour
'chkFridayFullLeave.ForeColor = FrameForeColour
'
'chkFridayMorning.BackColor = FrameBackColour
'chkFridayMorning.ForeColor = FrameForeColour
'
'chkFullDayLeave.BackColor = FrameBackColour
'chkFullDayLeave.ForeColor = FrameForeColour

'chkHLine2.BackColor = FrameBackColour
'chkHLine2.ForeColor = FrameForeColour
'
'chkHLine3.BackColor = FrameBackColour
'chkHLine3.ForeColor = FrameForeColour
''
'chkHLine4.BackColor = FrameBackColour
'chkHLine4.ForeColor = FrameForeColour

'chk.BackColor = FrameBackColour
'chkParaethesia.ForeColor = FrameForeColour
'
'chkMuscleWeak.BackColor = FrameBackColour
'chkMuscleWeak.ForeColor = FrameForeColour
'
'chkSleep.BackColor = FrameBackColour
'chkSleep.ForeColor = FrameForeColour
'
'chkVisual.BackColor = FrameBackColour
'chkVisual.ForeColor = FrameForeColour
''
'chkSmell.BackColor = FrameBackColour
'chkSmell.ForeColor = FrameForeColour
'
'chkTaste.BackColor = FrameBackColour
'chkTaste.ForeColor = FrameForeColour
''
'chkSpeech.BackColor = FrameBackColour
'chkSpeech.ForeColor = FrameForeColour
'
'chkPsychiatric.BackColor = FrameBackColour
'chkPsychiatric.ForeColor = FrameForeColour
'
'chkThinHair.BackColor = FrameBackColour
'chkThinHair.ForeColor = FrameForeColour
'
'chkHoarseVoice.BackColor = FrameBackColour
'chkHoarseVoice.ForeColor = FrameForeColour
'
'chkUrgency.BackColor = FrameBackColour
'chkUrgency.ForeColor = FrameForeColour
'
'chkUrinaryFrequency.BackColor = FrameBackColour
'chkUrinaryFrequency.ForeColor = FrameForeColour
'
'chkUrgeIncontinence.BackColor = FrameBackColour
'chkUrgeIncontinence.ForeColor = FrameForeColour
'
'txt.BackColor = TxtBackColour
'txtAddress.ForeColor = TxtForeColour
'
'txtAge.BackColor = TxtBackColour
'txtAge.ForeColor = TxtForeColour
'
'txtAgentBalance.BackColor = TxtBackColour
'txtAgentBalance.ForeColor = TxtForeColour
'txtAuthorizationCode.BackColor = TxtBackColour
'txtAuthorizationCode.ForeColor = TxtForeColour
'txtCashDue.BackColor = TxtBackColour
'txtCashDue.ForeColor = TxtForeColour
'txtChequeNo.BackColor = TxtBackColour
'txtChequeNo.ForeColor = TxtForeColour
'txtDiscount.BackColor = TxtBackColour
'txtDiscount.ForeColor = TxtForeColour
'txtEmail.BackColor = TxtBackColour
'txtEmail.ForeColor = TxtForeColour
'txtFax.BackColor = TxtBackColour
'txtFax.ForeColor = TxtForeColour
'txtFirstName.BackColor = TxtBackColour
'txtFirstName.ForeColor = TxtForeColour
'txtGrossTotal.BackColor = TxtBackColour
'txtGrossTotal.ForeColor = TxtForeColour
'txtNetTotal.BackColor = TxtBackColour
'txtNetTotal.ForeColor = TxtForeColour
'
'txtNIC.BackColor = TxtBackColour
'txtNIC.ForeColor = TxtForeColour
'txtNotes.BackColor = TxtBackColour
'txtNotes.ForeColor = TxtForeColour
'txtOtherName.BackColor = TxtBackColour
'txtOtherName.ForeColor = TxtForeColour
'txtPaidForCredit.BackColor = TxtBackColour
'txtPaidForCredit.ForeColor = TxtForeColour
'txtSearchFirstName.BackColor = TxtBackColour
'txtSearchFirstName.ForeColor = TxtForeColour
'
'
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'
'Text2.BackColor = TxtBackColour
'Text2.ForeColor = TxtForeColour
'
'Text3.BackColor = TxtBackColour
'Text3.ForeColor = TxtForeColour
'
'Text4.BackColor = TxtBackColour
'Text4.ForeColor = TxtForeColour
'
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'
'Text1.BackColor = TxtBackColour
'Text1.ForeColor = TxtForeColour
'
'
'
'OptionAgent.BackColor = TxtBackColour
'OptionAgent.ForeColor = TxtForeColour
'
'OptionCash.BackColor = TxtBackColour
'OptionCash.ForeColor = TxtForeColour
'
'OptionCheque.BackColor = TxtBackColour
'OptionCheque.ForeColor = TxtForeColour
'OptionDoNotPrint.BackColor = TxtBackColour
'OptionDoNotPrint.ForeColor = TxtForeColour
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCreditCard.BackColor = TxtBackColour
'OptionCreditCard.ForeColor = TxtForeColour
'
'OptionMaster.BackColor = TxtBackColour
'OptionMaster.ForeColor = TxtForeColour
'
'OptionPrintOne.BackColor = TxtBackColour
'OptionPrintOne.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'OptionCredit.BackColor = TxtBackColour
'OptionCredit.ForeColor = TxtForeColour
'
'
'
'txtSearchID.BackColor = TxtBackColour
'txtSearchID.ForeColor = TxtForeColour
'txtSearchSurname.BackColor = TxtBackColour
'txtSearchSurname.ForeColor = TxtForeColour
'txtSurname.BackColor = TxtBackColour
'txtSurname.ForeColor = TxtForeColour
'txtTelephone.BackColor = TxtBackColour
'txtTelephone.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour
'txtPrivateTel.ForeColor = TxtForeColour





End Sub

Private Sub bttnAdd_Click()
    Call AfterAdd
    Call ClearValues
    txtName.SetFocus
End Sub

Private Sub bttnCancel_Click()
    Call ClearValues
    Call BeforeAddEdit
End Sub

Private Sub bttnChange_Click()
Dim TemResponce  As Integer

If Trim(txtName.Text) = "" Then
    TemResponce = MsgBox("You must enter a name for the investigation", vbCritical, "No Name")
    txtName.SetFocus
    Exit Sub
End If

Call EditData
Call ClearValues
Call BeforeAddEdit


End Sub

Private Sub EditData()
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT tblSpeciality.* from tblSpeciality Where (Speciality_ID = " & TemSpecialityId & ")"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then .Close: Exit Sub
    !speciality = Trim(txtName.Text)
    .Update
    grid1.Col = 1
    grid1.Text = Trim(txtName.Text)
    grid1.Col = 2
    grid1.Text = !speciality_ID
    TemSpecialityId = !speciality_ID
End With
End Sub

Private Sub FormatGrid()

Dim BorderMargin As Long
BorderMargin = 100


With grid1
    .Clear
    
    .Rows = 1
    .Cols = 3
    
    .ColWidth(0) = 600
    .ColWidth(2) = 1
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + BorderMargin)
    
    .Col = 0
    .CellAlignment = 4
    .Text = "No."
    
    .Col = 1
    .CellAlignment = 4
    .Text = "Speciality"
    
End With
End Sub

Private Sub FillGrid()
Dim NowRow As Long
With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT tblSpeciality.* from tblSpeciality order by Speciality"
    If .State = 0 Then .Open
    If .RecordCount = 0 Then .Close: Exit Sub
    .MoveFirst
    NowRow = 0
    While .EOF = False
        NowRow = NowRow + 1
        grid1.Rows = NowRow + 1
        grid1.Row = NowRow
        grid1.Col = 0
        grid1.Text = NowRow
        grid1.Col = 1
        grid1.Text = !speciality
        grid1.Col = 2
        grid1.Text = !speciality_ID
        .MoveNext
    Wend

End With
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnDelete_Click()
    With DataEnvironment1.rssqlTem
        If .State = 1 Then .Close
        .Source = "SELECT tblSpeciality.* from tblSpeciality Where (Speciality_ID = " & TemSpecialityId & ")"
        If .State = 0 Then .Open
        If .RecordCount = 0 Then .Close:   Exit Sub
        .Delete adAffectCurrent
        .Close
    End With
    Dim TemNum As Long
    With grid1
        .RemoveItem (grid1.Row)
        .Col = 0
        
        .Sort = 1
        
        For TemNum = 1 To .Rows - 1
            .Row = TemNum
            .Text = TemNum
        Next
    End With
    
Call ClearValues
Call BeforeAddEdit

End Sub

Private Sub bttnEdit_Click()
FromGrid = True
    Call AfterEdit
End Sub

Private Sub bttnSave_Click()
Dim TemResponce  As Integer

If Trim(txtName.Text) = "" Then
    TemResponce = MsgBox("You must enter a name for the investigation", vbCritical, "No Name")
    txtName.SetFocus
    Exit Sub
End If

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT tblSpeciality.* from tblSpeciality Where (Speciality = '" & Trim(txtName.Text) & "')"
    If .State = 0 Then .Open
    If .RecordCount <> 0 Then
        TemResponce = MsgBox("The Investigation you entered already exist.", vbCritical, "Name Exists")
        txtName.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
End With

Call SaveDetails
Call ClearValues
Call BeforeAddEdit

End Sub


Private Sub SaveDetails()
Dim TemNum As Long

With DataEnvironment1.rssqlTem
    If .State = 1 Then .Close
    .Source = "SELECT tblSpeciality.* from tblSpeciality"
    If .State = 0 Then .Open
    .AddNew
    !speciality = Trim(txtName.Text)
    .Update
    TemSpecialityId = !speciality_ID
    If .State = 1 Then .Close
    
    grid1.Rows = grid1.Rows + 1
    grid1.Row = grid1.Rows - 1
    grid1.Col = 1
    grid1.Text = Trim(txtName.Text)
    grid1.Col = 2
    grid1.Text = TemSpecialityId
    grid1.Col = 1
    grid1.Sort = 1
    grid1.Col = 0
    For TemNum = 1 To grid1.Rows - 1
        grid1.Row = TemNum
        grid1.Text = TemNum
    Next

End With
End Sub

Private Sub Form_Load()
    Call FormatGrid
    Call FillGrid
    Call ClearValues
    Call BeforeAddEdit
    Call SetColour
End Sub

Private Sub Grid1_Click()
    FromGrid = True
            
            bttnAdd.Enabled = True
            
            With grid1
                If .Row < 1 Then FromGrid = False: Exit Sub
                .Col = 2
                If Not IsNumeric(.Text) Then FromGrid = False: Exit Sub
                TemSpecialityId = Val(.Text)
                .Col = 1
                txtSearch.Text = .Text
                
                Call GetData
                
                .Col = 0
                .ColSel = .Cols - 1
                
                txtSearch.SetFocus
                SendKeys "{home}+{end}"
            FromGrid = False
            bttnAdd.Enabled = False
            bttnEdit.Enabled = True
            bttnDelete.Enabled = True
        End With
End Sub

Private Sub ClearValues()
    txtName.Text = Empty
End Sub

Private Sub GetData()
With DataEnvironment1.rssqlInvestigations
    If .State = 1 Then .Close
    .Source = "SELECT tblSpeciality.* from tblSpeciality where Speciality_ID = " & TemSpecialityId
    If .State = 0 Then .Open
    If .RecordCount = 0 Then Exit Sub
    If Not IsNull(!speciality) Then txtName.Text = !speciality
    If .State = 1 Then .Close
End With
End Sub


Private Sub BeforeAddEdit()

txtSearch.Text = Empty

grid1.Enabled = True
FrameIx.Enabled = False
bttnAdd.Enabled = True
bttnEdit.Enabled = True
bttnDelete.Enabled = False

bttnSave.Visible = False
bttnChange.Visible = False
bttnCancel.Visible = False

    FromGrid = False

End Sub


Private Sub AfterAdd()

txtSearch.Text = Empty

Call ClearValues

grid1.Enabled = False
FrameIx.Enabled = True
bttnAdd.Enabled = False
bttnEdit.Enabled = False
bttnDelete.Enabled = False

bttnSave.Visible = True
bttnChange.Visible = False
bttnCancel.Visible = True

End Sub

Private Sub AfterEdit()

txtSearch.Text = Empty

grid1.Enabled = False
FrameIx.Enabled = True
bttnAdd.Enabled = True
bttnEdit.Enabled = False
bttnDelete.Enabled = False

bttnSave.Visible = False
bttnChange.Visible = True
bttnCancel.Visible = True

End Sub


Private Sub txtSearch_Change()

    
' **************************************

    If FromGrid = True Then Exit Sub
    Dim TemFRows As Long
    Dim TemNowRow As Long
    Dim TemArray As Long
    Dim SearchSuccess As Boolean
    Dim TemLength As Single
    TemFRows = grid1.Rows
    grid1.Col = 1
    SearchSuccess = False
    If Len(txtSearch.Text) = 0 Then GoTo MeasureSuccess
    For TemArray = 1 To (TemFRows - 1)
        grid1.Row = TemArray
        If Len(txtSearch.Text) > Len(grid1.Text) Then
            GoTo FinishLoop
        Else
            TemLength = Len(txtSearch.Text)
        End If
        If UCase(Left((grid1.Text), TemLength)) = UCase(txtSearch.Text) Then
            SearchSuccess = True
            Exit For
        Else
            SearchSuccess = False
        End If
FinishLoop:
    Next
    
MeasureSuccess:
    
    If SearchSuccess = True Then
        grid1.TopRow = TemArray
        grid1.Row = TemArray
        grid1.Col = 0
        grid1.ColSel = (grid1.Cols - 1)
        bttnEdit.Enabled = True
        bttnDelete.Enabled = True
        bttnAdd.Enabled = False
        grid1.Col = 2
        TemSpecialityId = grid1.Text
        Call GetData
        grid1.Col = 0
        grid1.ColSel = grid1.Cols - 1
    Else
        grid1.TopRow = 1
        grid1.Row = 0
        grid1.Col = 0
        grid1.ColSel = 0
        bttnAdd.Enabled = True
        bttnEdit.Enabled = False
        bttnDelete.Enabled = False
    End If
'**************************************



End Sub
