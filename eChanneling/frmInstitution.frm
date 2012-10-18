VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmInstitution 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Institution's Profile"
   ClientHeight    =   5640
   ClientLeft      =   3390
   ClientTop       =   3390
   ClientWidth     =   9000
   ClipControls    =   0   'False
   Icon            =   "frmInstitution.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   9000
   Begin VB.Frame framInstutions 
      Height          =   4815
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   8655
      Begin VB.TextBox txtInsname 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   6375
      End
      Begin VB.TextBox txtDiscription 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   720
         Width           =   6375
      End
      Begin VB.TextBox txtRegistration 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1200
         Width           =   6375
      End
      Begin VB.TextBox txtAddress01 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1680
         Width           =   6375
      End
      Begin VB.TextBox txtTelephone01 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtTelephone02 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox txtEmail01 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtEmail02 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   3720
         Width           =   3135
      End
      Begin VB.TextBox txtwbsite01 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   4200
         Width           =   3135
      End
      Begin VB.TextBox txtWebsite02 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   10
         Top             =   4200
         Width           =   3135
      End
      Begin VB.TextBox txtFax 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   3240
         Width           =   6375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Institution &Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Discription"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Registration No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Tele&phone No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Email "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "&Website"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   2295
      End
   End
   Begin btButtonEx.ButtonEx bttnOK 
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Top             =   5040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnCancel 
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmInstitution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SuppliedWord As String
Dim TemResponce  As Integer
Dim A As String
Private Sub Setcolours()


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

bttnOK.BackColor = BttnBackColour
bttnOK.ForeColor = BttnForeColour

bttnCancel.BackColor = BttnBackColour
bttnCancel.ForeColor = BttnForeColour

'bttnChange.BackColor = BttnBackColour
'bttnChange.ForeColor = BttnForeColour
'
'bttnClose.BackColor = BttnBackColour
'bttnClose.ForeColor = BttnForeColour
'
'bttnEdit.BackColor = BttnBackColour
'bttnEdit.ForeColor = BttnForeColour
'
'bttnSave.BackColor = BttnBackColour
'bttnSave.ForeColor = BttnForeColour
'
frmInstitution.BackColor = FrmBackColour
frmInstitution.ForeColor = FrmForeColour

framInstutions.BackColor = FrameBackColour
framInstutions.ForeColor = FrameForeColour
'
'FramePrivate.BackColor = FrameBackColour
'FramePrivate.ForeColor = FrameForeColour
'
'FrameOfficial.BackColor = FrameBackColour
'FrameOfficial.ForeColor = FrameForeColour
'
'FramePayment.BackColor = FrameBackColour
'FramePayment.ForeColor = FrameForeColour
'
'chkCurrentlyChanneling.BackColor = LblBackColour
'chkCurrentlyChanneling.ForeColor = LblForeColour
'
'DataComboBank.BackColor = TxtBackColour
'DataComboBank.ForeColor = TxtForeColour
'
'DataComboPaymenyMethod.BackColor = TxtBackColour
'DataComboPaymenyMethod.ForeColor = TxtForeColour
'
'DataComboSex.BackColor = TxtBackColour
'DataComboSex.ForeColor = TxtForeColour
'
'DataComboSpeciality.BackColor = TxtBackColour
'DataComboSpeciality.ForeColor = TxtForeColour
'
'DataComboTitle.BackColor = TxtBackColour
'DataComboTitle.ForeColor = TxtForeColour
'
''DataCombo.BackColor = TxtBackColour
''DataComboBank.ForeColor = TxtForeColour
''
''DataComboBank.BackColor = TxtBackColour
''DataComboBank.ForeColor = TxtForeColour
''DataComboBank.BackColor = TxtBackColour
''DataComboBank.ForeColor = TxtForeColour
'
'
'
'
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
''grid1.ForeColor = Grid
'
'
'
'Label1.BackColor = LblBackColour
'Label1.ForeColor = LblForeColour

Label10.BackColor = LblBackColour
Label10.ForeColor = LblForeColour
Label11.BackColor = LblBackColour
Label11.ForeColor = LblForeColour
Label12.BackColor = LblBackColour
Label12.ForeColor = LblForeColour
'Label13.BackColor = LblBackColour
'Label13.ForeColor = LblForeColour
Label14.BackColor = LblBackColour
Label14.ForeColor = LblForeColour
'Label15.BackColor = LblBackColour
'Label15.ForeColor = LblForeColour
Label6.BackColor = LblBackColour
Label6.ForeColor = LblForeColour
Label2.BackColor = LblBackColour
Label2.ForeColor = LblForeColour
'Label18.BackColor = LblBackColour
'Label18.ForeColor = LblForeColour
Label3.BackColor = LblBackColour
Label3.ForeColor = LblForeColour
'Label20.BackColor = LblBackColour
'Label20.ForeColor = LblForeColour
'Label21.BackColor = LblBackColour
'Label21.ForeColor = LblForeColour
Label4.BackColor = LblBackColour
Label4.ForeColor = LblForeColour
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
'lblOfficialEmail.BackColor = LblBackColour
'lblOfficialEmail.ForeColor = LblForeColour
'
'lblOfficialWebsite.BackColor = LblBackColour
'lblOfficialWebsite.ForeColor = LblForeColour
'
'
'txtAccount.BackColor = TxtBackColour
'txtAccount.ForeColor = TxtForeColour
'
'txtBankBranch.BackColor = TxtBackColour
'txtBankBranch.ForeColor = TxtForeColour
'
'txtComments.BackColor = TxtBackColour
'txtComments.ForeColor = TxtForeColour
'txtCredit.BackColor = TxtBackColour
'txtCredit.ForeColor = TxtForeColour
'txtDesignation.BackColor = TxtBackColour
'txtDesignation.ForeColor = TxtForeColour
'txtListedName.BackColor = TxtBackColour
'txtListedName.ForeColor = TxtForeColour
'txtName.BackColor = TxtBackColour
'txtName.ForeColor = TxtForeColour
'txtOfficialAddress.BackColor = TxtBackColour
'txtOfficialAddress.ForeColor = TxtForeColour
'txtOfficialEMail.BackColor = TxtBackColour
'txtOfficialEMail.ForeColor = TxtForeColour
'txtOfficialFax.BackColor = TxtBackColour
'txtOfficialFax.ForeColor = TxtForeColour
'txtOfficialTel.BackColor = TxtBackColour
'txtOfficialTel.ForeColor = TxtForeColour
'txtOfficialWebsite.BackColor = TxtBackColour
'txtOfficialWebsite.ForeColor = TxtForeColour
'
'txtPrivateAddress.BackColor = TxtBackColour
'txtPrivateAddress.ForeColor = TxtForeColour
'txtPrivateEmail.BackColor = TxtBackColour
'txtPrivateEmail.ForeColor = TxtForeColour
'txtPrivateFax.BackColor = TxtBackColour
'txtPrivateFax.ForeColor = TxtForeColour
'txtPrivateMobile.BackColor = TxtBackColour
'txtPrivateMobile.ForeColor = TxtForeColour
'txtPrivateTel.BackColor = TxtBackColour
'txtPrivateTel.ForeColor = TxtForeColour
'
'
'txtQualifications.BackColor = TxtBackColour
'txtQualifications.ForeColor = TxtForeColour
'txtRegistation.BackColor = TxtBackColour
'txtRegistation.ForeColor = TxtForeColour
'txtSearch.BackColor = TxtBackColour
'txtSearch.ForeColor = TxtForeColour
End Sub







Private Sub bttnOk_Click()

If Date > #12/18/2008# Then
    TemResponce = MsgBox("Please contact Lakmedipro for Assistant", vbCritical, "Expired")
    End
End If

If Trim(txtInsname.Text) = "" Then
    TemResponce = MsgBox("You have not entered the name of the institution", vbCritical, "Institution Name?")
    txtInsname.SetFocus
    Exit Sub
End If

Call SaveInstitution
Unload Me
End Sub

Private Sub Form_Load()
With DataEnvironment1.rscmmdInstitutionDetails
    If .State = 0 Then .Open
    If .RecordCount = 0 Then
        TemResponce = MsgBox("Someone had altered the database outside the program. Please contact Lakmedipro for assistant" & vbNewLine & Me.Caption & vbNewLine & Err.Description, vbCritical, "Altered Database")
        .AddNew
        Exit Sub
    End If
    .MoveFirst
Call Display
Call Setcolours

End With
End Sub

Private Sub Form_Unload(cancel As Integer)
    With DataEnvironment1.rscmmdInstitutionDetails
        .Update
        If .State = 1 Then .Close
    End With
End Sub


Private Sub bttnCancel_Click()
Unload Me
End Sub

Private Sub Display()

With DataEnvironment1.rscmmdInstitutionDetails

SuppliedWord = !InstitutionName
txtInsname.Text = DecreptedWord(SuppliedWord)

SuppliedWord = !InstitutionDescription
txtDiscription.Text = DecreptedWord(SuppliedWord)

SuppliedWord = !InstitutionRegistation
txtRegistration.Text = DecreptedWord(SuppliedWord)

SuppliedWord = !InstitutionAddress
txtAddress01.Text = DecreptedWord(SuppliedWord)

SuppliedWord = !institutiontelephone1
txtTelephone01.Text = DecreptedWord(SuppliedWord)

SuppliedWord = !InstitutionTelephone2
txtTelephone02.Text = DecreptedWord(SuppliedWord)

SuppliedWord = !InstitutionFax
txtFax.Text = DecreptedWord(SuppliedWord)

SuppliedWord = !InstitutionEmail
txtEmail01.Text = DecreptedWord(SuppliedWord)

SuppliedWord = !InstitutionEmail2
txtEmail02.Text = DecreptedWord(SuppliedWord)

SuppliedWord = !InstitutionWebSite1
txtwbsite01.Text = DecreptedWord(SuppliedWord)

SuppliedWord = !InstitutionWebSite2
txtWebsite02.Text = DecreptedWord(SuppliedWord)


End With


End Sub


Private Sub SaveInstitution()

If txtInsname.Text = "" Then
    TemResponce = MsgBox("You have not entered an Institution Name", vbCritical, "? Institution Name")
    txtInsname.SetFocus
    Exit Sub
End If

'On Error GoTo ErrorHandler

With DataEnvironment1.rscmmdInstitutionDetails

SuppliedWord = txtInsname.Text
!InstitutionName = EncreptedWord(SuppliedWord)

SuppliedWord = txtDiscription.Text
!InstitutionDescription = EncreptedWord(SuppliedWord)

SuppliedWord = txtRegistration.Text
!InstitutionRegistation = EncreptedWord(SuppliedWord)

SuppliedWord = txtAddress01.Text
!InstitutionAddress = EncreptedWord(SuppliedWord)

'SuppliedWord = txtAddress02.Text
'!InstitutionAddressLine2 = EncreptedWord(SuppliedWord)
'
'SuppliedWord = txtAddress03.Text
'!InstitutionAddressLine3 = EncreptedWord(SuppliedWord)
'
'SuppliedWord = txtAddress04.Text
'!InstitutionAddressLine4 = EncreptedWord(SuppliedWord)

SuppliedWord = txtTelephone01.Text
!institutiontelephone1 = EncreptedWord(SuppliedWord)

SuppliedWord = txtTelephone02.Text
!InstitutionTelephone2 = EncreptedWord(SuppliedWord)

SuppliedWord = txtFax.Text
!InstitutionFax = EncreptedWord(SuppliedWord)

SuppliedWord = txtEmail01.Text
!InstitutionEmail = EncreptedWord(SuppliedWord)

SuppliedWord = txtEmail02.Text
!InstitutionEmail2 = EncreptedWord(SuppliedWord)

SuppliedWord = txtwbsite01.Text
!InstitutionWebSite1 = EncreptedWord(SuppliedWord)

SuppliedWord = txtWebsite02.Text
!InstitutionWebSite2 = EncreptedWord(SuppliedWord)


.Update

End With

Exit Sub
ErrorHandler:
If Err.Number = -2147217887 Then
    TemResponce = MsgBox("The Doctor name, " & txtInsname.Text & " is already there in the database. If you want to make changes, click the Edit button", , "Alredy in the database")
    DataEnvironment1.rscmmdInstitutionDetails.CancelUpdate
    
Else
    MsgBox ("An Error Occured during Updating" & vbNewLine & Me.Caption & vbNewLine & Err.Number & vbNewLine & Err.Description)
    DataEnvironment1.rscmmdInstitutionDetails.CancelUpdate
    
End If

End Sub


Private Sub ClearValues()

txtInsname.Text = Empty
txtRegistration.Text = Empty
txtAddress01.Text = Empty
'txtAddress02.Text = Empty
'txtAddress03.Text = Empty
'txtAddress04.Text = Empty
txtTelephone01.Text = Empty
txtTelephone02.Text = Empty
txtFax.Text = Empty
txtEmail01.Text = Empty
txtEmail02.Text = Empty
txtwbsite01.Text = Empty
txtWebsite02.Text = Empty

End Sub


Private Sub NameEmpty()
A = MsgBox("Enter Correct Name", vbCritical + vbExclamation, "Name Empty")
End Sub

