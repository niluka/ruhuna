VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPaitionlistbyDoctor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patients List By Doctor"
   ClientHeight    =   7200
   ClientLeft      =   3315
   ClientTop       =   2340
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPaitionlistbyDoctor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9045
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin btButtonEx.ButtonEx bttnPrintView 
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   6240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Appearance      =   3
         Enabled         =   0   'False
         Caption         =   "Print &View"
         Enabled         =   0   'False
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
      Begin VB.ListBox listDoctor 
         Height          =   4620
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtp2 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61603841
         CurrentDate     =   39455
      End
      Begin MSFlexGridLib.MSFlexGrid msfGrid1 
         Height          =   4695
         Left            =   3720
         TabIndex        =   8
         Top             =   1200
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   8281
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61603841
         CurrentDate     =   39455
      End
      Begin MSDataListLib.DataCombo dtcSpeciality 
         Height          =   360
         Left            =   3720
         TabIndex        =   3
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin btButtonEx.ButtonEx bttnPrint 
         Height          =   375
         Left            =   4680
         TabIndex        =   6
         Top             =   6240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Appearance      =   3
         Enabled         =   0   'False
         Caption         =   "&Print"
         Enabled         =   0   'False
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
         Left            =   6840
         TabIndex        =   7
         Top             =   6240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "&Close"
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
      Begin MSDataListLib.DataCombo dtcDoctor 
         Height          =   360
         Left            =   360
         TabIndex        =   12
         Top             =   5040
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label4 
         Caption         =   "Doctor Name"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&To"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Speciality"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date  &From"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmPaitionlistbyDoctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long
Dim r As Long


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

'bttnClose.BackColor = BttnBackColour
'bttnClose.ForeColor = BttnForeColour
'
'bttnAdd.BackColor = BttnBackColour
'bttnAdd.ForeColor = BttnForeColour
'
'bttnAddLeave.BackColor = BttnBackColour
'bttnAddLeave.ForeColor = BttnForeColour
'
'bttnCancel.BackColor = BttnBackColour
'bttnCancel.ForeColor = BttnForeColour
'
'bttnChange.BackColor = BttnBackColour
'bttnChange.ForeColor = BttnForeColour
'
'bttnClose.BackColor = BttnBackColour
'bttnClose.ForeColor = BttnForeColour
'
'bttnDelete.BackColor = BttnBackColour
'bttnDelete.ForeColor = BttnForeColour
'
'bttnEdit.BackColor = BttnBackColour
'bttnEdit.ForeColor = BttnForeColour

'ButtonEx3.BackColor = BttnBackColour
'ButtonEx3.ForeColor = BttnForeColour

'bttnSerchBill.BackColor = BttnBackColour
'bttnSerchBill.ForeColor = BttnForeColour

'bttnUpdate.BackColor = BttnBackColour
'bttnUpdate.ForeColor = BttnForeColour


'OptionABC.BackColor = FrmBackColour
'OptionABC.ForeColor = FrmForeColour

'OptionAmEx.BackColor = FrmBackColour
'OptionAmEx.ForeColor = FrmForeColour
'
'OptionCash.BackColor = FrmBackColour
'OptionCash.ForeColor = FrmForeColour
'
'OptionVISA.BackColor = FrmBackColour
'OptionVISA.ForeColor = FrmForeColour
'
'OptionMaster.BackColor = FrmBackColour
'OptionMaster.ForeColor = FrmForeColour
'
'OptionCheque.BackColor = FrmBackColour
'OptionCheque.ForeColor = FrmForeColour
'
'OptionCreditCard.BackColor = FrmBackColour
'OptionCreditCard.ForeColor = FrmForeColour
'
'OptionCreditCard.BackColor = FrmBackColour
'OptionCreditCard.ForeColor = FrmForeColour

'FrameCash.BackColor = FrameBackColour
'FrameCash.ForeColor = FrameForeColour
'
Frame1.BackColor = FrmBackColour
Frame1.ForeColor = FrmForeColour
'
frmPaitionlistbyDoctor.BackColor = FrameBackColour
frmPaitionlistbyDoctor.ForeColor = FrameForeColour
'
'
'FrameCreditCard.BackColor = FrameBackColour
'FrameCreditCard.ForeColor = FrameForeColour
'
'FramePaymentMethod.BackColor = FrameBackColour
'FramePaymentMethod.ForeColor = FrameForeColour
'
'framCoustomerbill.BackColor = FrameBackColour
'framCoustomerbill.ForeColor = FrameForeColour
'
'frmPayment.BackColor = FrameBackColour
'frmPayment.ForeColor = FrameForeColour

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

'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour

'DataComboDoctorStaff.BackColor = TxtBackColour
'DataComboDoctorStaff.ForeColor = TxtForeColour

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




'grid1.BackColor = GridBackColor
'grid1.ForeColor = GridForeColor
'
'grid1.BackColorBkg = GridBackColorBkg
'grid1.BackColorFixed = GridBackColorFixed
'grid1.BackColorSel = GridBackColorSel
'
'grid1.ForeColor = GridForeColor
'grid1.ForeColorFixed = GridForeColorFixed
'grid1.ForeColorSel = GridForeColorSel
'
'grid1.ForeColor = GridForeColor




Label1.BackColor = LblBackColour
Label1.ForeColor = LblForeColour

'lblDoctorStaff.BackColor = LblBackColour
'lblDoctorStaff.ForeColor = LblForeColour
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

Private Sub FormatGrid()
With msfGrid1
    .Cols = 3
    .ColWidth(0) = 700
    .ColWidth(2) = 1
    .ColWidth(1) = .Width - (.ColWidth(0) + .ColWidth(2) + 350)
    .Row = 0
    .Col = 0
    .Text = "ID"
    .Col = 1
    .Text = "Patient Name"
    .Col = 2
    .Text = ""
End With
End Sub

Private Sub FindPatients()

With DataEnvironment1.rssqlTem1
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.Staff_ID = " & DtcDoctor.BoundText & " and (tblPatientFacility.BookingDate Between '" & dtp1.Value & "' and '" & dtp2.Value & "') and tblPatientFacility.refund = 0 and tblPatientFacility.cancelled = 0) Order by tblPatientFacility.PatientFacility_ID "
    .Open
    If .RecordCount = 0 Then Exit Sub
    bttnPrint.Enabled = True
    bttnPrintView.Enabled = True
    Call FormatGrid
    
    i = 1
    r = 1
        Do While .EOF = False
            With msfGrid1
            r = r + 1
            .Rows = r
            .Row = i
            .Col = 0
            .Text = i
            .CellAlignment = 7
            .Col = 1
            .Text = DataEnvironment1.rssqlTem1.Fields("FirstName")
            .CellAlignment = 1
            .Col = 2
            .Text = DataEnvironment1.rssqlTem1.Fields(0)
            i = i + 1
            End With
            
        .MoveNext
    Loop
End With

End Sub

Private Sub bttnClose_Click()
Unload Me
End Sub

Private Sub bttnPrint_Click()

With DataEnvironment1.rssqlTem1
dtrPatientsListbyDoctor.Visible = False
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.Staff_ID = " & DtcDoctor.BoundText & " and (tblPatientFacility.BookingDate Between '" & dtp1.Value & "' and '" & dtp2.Value & "') and tblPatientFacility.refund = 0 and tblPatientFacility.cancelled = 0) Order by tblPatientFacility.PatientFacility_ID "
    .Open
    If .RecordCount = 0 Then A = MsgBox("No Patients to Display", vbCritical + vbInformation, "No Patients"): Exit Sub
    
    Set dtrPatientsListbyDoctor.DataSource = DataEnvironment1.rssqlTem1
    
    With dtrPatientsListbyDoctor
    .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress
    
    .Sections("Section2").Controls.Item("RptDocterName").Caption = DtcDoctor.Text
    .Sections("Section2").Controls.Item("rptFromdate").Caption = dtp1.Value
    .Sections("Section2").Controls.Item("rptTodate").Caption = dtp2.Value
    .Sections("Section5").Controls.Item("rptlRcount").Caption = DataEnvironment1.rssqlTem1.RecordCount
    
    .PrintReport False
    Unload Me
    End With
    
End With
'bttnPrint.Enabled = False
End Sub

Private Sub bttnPrintView_Click()
With DataEnvironment1.rssqlTem2
    If .State = 1 Then .Close
    .Source = "Select tblPatientFacility.*,tblPatientMainDetails.* From tblPatientFacility Left Join tblPatientMainDetails On tblPatientFacility.PatientID = tblPatientMainDetails.Patient_ID Where (tblPatientFacility.Staff_ID = " & DtcDoctor.BoundText & " and (tblPatientFacility.BookingDate Between '" & dtp1.Value & "' and '" & dtp2.Value & "') and tblPatientFacility.refund = 0 and tblPatientFacility.cancelled = 0) Order by tblPatientFacility.PatientFacility_ID "
    .Open
    If .RecordCount = 0 Then A = MsgBox("No Patients to Display", vbCritical + vbInformation, "No Patients"): Exit Sub
    
    Set dtrPatientsListbyDoctor.DataSource = DataEnvironment1.rssqlTem2
    
    With dtrPatientsListbyDoctor
    .Sections("Section4").Controls.Item("RptName").Caption = InstitutionName
    .Sections("Section4").Controls.Item("RptAddress").Caption = InstitutionAddress

    
    .Sections("Section2").Controls.Item("RptDocterName").Caption = DtcDoctor.Text
    .Sections("Section2").Controls.Item("rptFromdate").Caption = dtp1.Value
    .Sections("Section2").Controls.Item("rptTodate").Caption = dtp2.Value
    .Sections("Section5").Controls.Item("rptlRcount").Caption = DataEnvironment1.rssqlTem2.RecordCount

    .Show
    End With
End With
'bttnPrintView.Enabled = False
End Sub

Private Sub dtcSpeciality_Click(Area As Integer)
If IsNumeric(dtcSpeciality.BoundText) = False Then Exit Sub
Call FindDoctors
End Sub

Private Sub FindDoctors()
With DataEnvironment1.rssqlTem10
    If .State = 1 Then .Close
    
    .Open "SELECT tblDoctor.Doctor_ID, tblDoctor.DoctorSpeciality_ID, tblDoctor.DoctorName, tblSpeciality.* FROM tblSpeciality LEFT JOIN tblDoctor ON tblSpeciality.Speciality_ID = tblDoctor.DoctorSpeciality_ID Where (tblDoctor.DoctorSpeciality_ID = " & dtcSpeciality.BoundText & ")Order by tblDoctor.DoctorName"
    If .RecordCount = 0 Then Exit Sub
    
    listDoctor.Clear
    If .RecordCount = 0 Then Exit Sub
    
        Do While .EOF = False
        listDoctor.AddItem !doctorname
        
        .MoveNext
        Loop

End With
End Sub

Private Sub dtp1_Change()
msfGrid1.Clear
If IsNumeric(DtcDoctor.BoundText) = False Then Exit Sub
Call FindPatients
End Sub

Private Sub dtp2_Change()
msfGrid1.Clear
If IsNumeric(DtcDoctor.BoundText) = False Then Exit Sub
Call FindPatients
End Sub

Private Sub Form_Load()
dtp1 = Date
dtp2 = Date
FillDatacomboSpeciality
FillDatacomboDoctor
Call SetColour
    If UserAuthority <> AuthorityOwner Then
        dtp1.Enabled = False
        dtp2.Enabled = False
    End If

End Sub

Private Sub FillDatacomboDoctor()

With DataEnvironment1.rssqlTem17
    If .State = 1 Then .Close
    .Source = "Select tblDoctor.* From tblDoctor Order by DoctorName"
    .Open
    Set DtcDoctor.RowSource = DataEnvironment1.rssqlTem17
    DtcDoctor.BoundColumn = "Doctor_ID"
    DtcDoctor.ListField = "DoctorName"
    
End With

End Sub


Private Sub FillDatacomboSpeciality()
With DataEnvironment1.rssqlTem15
    If .State = 1 Then .Close
    .Source = "Select tblSpeciality.* From tblSpeciality Order by Speciality"
    .Open
    Set dtcSpeciality.RowSource = DataEnvironment1.rssqlTem15
    dtcSpeciality.BoundColumn = "Speciality_ID"
    dtcSpeciality.ListField = "Speciality"
    
End With

End Sub

Private Sub listDoctor_Click()
If listDoctor.Text = "" Then Exit Sub
DtcDoctor.Text = listDoctor.Text
msfGrid1.Clear
If IsNumeric(DtcDoctor.BoundText) = False Then Exit Sub
Call FindPatients
End Sub

Private Sub listDoctor_KeyPress(KeyAscii As Integer)
If listDoctor.Text = "" Then Exit Sub
DtcDoctor.Text = listDoctor.Text
msfGrid1.Clear
If IsNumeric(DtcDoctor.BoundText) = False Then Exit Sub
Call FindPatients

End Sub


