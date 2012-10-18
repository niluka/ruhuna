VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIFrmLab 
   BackColor       =   &H8000000C&
   Caption         =   "Lakmedipro eHospital - Reception"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1958
      ButtonWidth     =   2831
      ButtonHeight    =   1799
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Investigation"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Trace Investigations"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Patients"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Investigation Format"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Printing Preferances"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmLab.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmLab.frx":9A314
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmLab.frx":134628
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFrmLab.frx":1CE93C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuBackUp 
         Caption         =   "Back Up"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuInvestigation 
      Caption         =   "Investigation"
      Begin VB.Menu mnuNewInvestigation 
         Caption         =   "New Investigation"
      End
      Begin VB.Menu mnuTraceInvestigation 
         Caption         =   "Trace Investigation"
      End
      Begin VB.Menu mnuInvestigationNames 
         Caption         =   "Investigation Names"
      End
      Begin VB.Menu mnuInvestigationFormat 
         Caption         =   "Investigation Format"
      End
   End
   Begin VB.Menu mnuDataEntry 
      Caption         =   "Data Entry"
      Begin VB.Menu mnuPatients 
         Caption         =   "Patients"
      End
      Begin VB.Menu mnuDoctors 
         Caption         =   "Doctors"
      End
      Begin VB.Menu mnuStaff 
         Caption         =   "Staff"
      End
      Begin VB.Menu mnuInstitutions 
         Caption         =   "Institutions"
      End
      Begin VB.Menu mnuDepartments 
         Caption         =   "Institution Departments"
      End
      Begin VB.Menu mnuTitles 
         Caption         =   "Titles"
      End
      Begin VB.Menu mnuSpecimans 
         Caption         =   "Specimans"
      End
   End
   Begin VB.Menu mnuPreferances 
      Caption         =   "Preferances"
      Begin VB.Menu mnuInstitutionDetails 
         Caption         =   "Institution Details"
      End
      Begin VB.Menu mnuPrintingPreferances 
         Caption         =   "Printing Preferance"
      End
   End
End
Attribute VB_Name = "MDIFrmLab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_QueryUnload(cancel As Integer, UnloadMode As Integer)
    Dim TemResponce As Byte
    Dim TemForm As Form
    Dim AllFOrms As Form
    TemResponce = MsgBox("Are you aure you want to exit Lakmedipro e-Lab ?", vbInformation + vbYesNo, "EXIT?")
    If TemResponce = vbNo Then cancel = True: Exit Sub
    For Each TemForm In Forms
        Unload TemForm
    Next
End Sub

Private Sub mnuBackUp_Click()
    frmBackUp.Show
    frmBackUp.ZOrder 0
    frmBackUp.Top = 0
    frmBackUp.Left = 0
End Sub

Private Sub mnuDepartments_Click()
    frmInstitutionDepartments.Show
    frmInstitutionDepartments.ZOrder 0
    frmInstitutionDepartments.Top = 0
    frmInstitutionDepartments.Left = 0
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInstitutions_Click()
    FrmInstitutions1.Show
    FrmInstitutions1.ZOrder 0
    FrmInstitutions1.Top = 0
    FrmInstitutions1.Left = 0
End Sub

Private Sub mnuInvestigationFormat_Click()
    frmIxDetails.Show
    frmIxDetails.ZOrder 0
    frmIxDetails.Top = 0
    frmIxDetails.Left = 0
End Sub

Private Sub mnuInvestigationNames_Click()
    frmInvestigations.Show
    frmInvestigations.ZOrder 0
    frmInvestigations.Top = 0
    frmInvestigations.Left = 0
End Sub

Private Sub mnuNewInvestigation_Click()
    frmNewIx.Show
    frmNewIx.ZOrder 0
    frmNewIx.Top = 0
    frmNewIx.Left = 0
    
End Sub

Private Sub mnuPatients_Click()
    frmPatientMain.Show
    frmPatientMain.ZOrder 0
    frmPatientMain.Top = 0
    frmPatientMain.Left = 0
End Sub

Private Sub mnuRestore_Click()
frmRestore.Show
frmRestore.ZOrder 0
frmRestore.Top = 0
frmRestore.Left = 0
End Sub

Private Sub mnuSpecimans_Click()
    frmSpeciman.Show
    frmSpeciman.ZOrder 0
    frmSpeciman.Top = 0
    frmSpeciman.Left = 0
End Sub

Private Sub mnuTitles_Click()
    frmTitles.Show
    frmTitles.ZOrder 0
    frmTitles.Top = 0
    frmTitles.Left = 0
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 0: MsgBox 0
    Case 1: MsgBox 1
    Case 2: MsgBox 2
    Case 3: mnuPatients_Click
    Case 4:

End Select

End Sub
