VERSION 5.00
Begin VB.MDIForm MDIFrmMain 
   BackColor       =   &H8000000C&
   Caption         =   "eHospital"
   ClientHeight    =   9315
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13425
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
   End
   Begin VB.Menu mnuBackOffice 
      Caption         =   "Back Office"
      Begin VB.Menu mnuTotalIncome 
         Caption         =   "Total Income"
      End
   End
   Begin VB.Menu mnuTem 
      Caption         =   "tem"
   End
End
Attribute VB_Name = "MDIFrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    DatabasePath = App.Path
End Sub

Private Sub mnuTem_Click()
    frmMainForm.Show
End Sub

Private Sub mnuTotalIncome_Click()
    frmTotalIncomeHos.Show
End Sub
