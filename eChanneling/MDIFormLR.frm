VERSION 5.00
Begin VB.MDIForm MDIFormLR 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9285
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13290
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuInvestigationNames 
         Caption         =   "Investigation Names"
      End
      Begin VB.Menu mnuIxCatogery 
         Caption         =   "Investigation Catogery"
      End
      Begin VB.Menu mnuIxFormat 
         Caption         =   "Investigation Format"
      End
   End
   Begin VB.Menu mnuInvestigations 
      Caption         =   "Investigations"
   End
   Begin VB.Menu mnuBackoffice 
      Caption         =   "Backoffice"
   End
   Begin VB.Menu mnuPreferances 
      Caption         =   "Preferances"
      Begin VB.Menu mnuProgramPreferances 
         Caption         =   "Program Preferances"
      End
      Begin VB.Menu mnuPrintingPreferances 
         Caption         =   "Printing Preferances"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnutem 
      Caption         =   "Tem"
   End
End
Attribute VB_Name = "MDIFormLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuInvestigationNames_Click()
frmAddNewIxName.Show
End Sub

Private Sub mnuIxCatogery_Click()
frmIxCatogery.Show
End Sub

Private Sub mnuIxFormat_Click()
frmIxFormat.Show
End Sub

Private Sub mnuPrintingPreferances_Click()
frmLRPrintingPreferances.Show
End Sub

Private Sub mnuProgramPreferances_Click()
frmLRProgramPreferances.Show
End Sub

Private Sub mnutem_Click()
frmLRMain.Show
End Sub
