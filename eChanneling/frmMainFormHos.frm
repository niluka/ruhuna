VERSION 5.00
Begin VB.Form frmMainForm 
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   10230
End
Attribute VB_Name = "frmMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbHospital As New ADODB.Connection
Dim cnnStr As String
Dim rsPatientFacility As New ADODB.Recordset



Private Sub FindIncome()
    cnnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Hospital.mdb;Mode=ReadWrite;Persist Security Info=False"
    dbHospital.Open cnnStr
End Sub

Private Sub Form_Load()
    Call FindIncome
    rsPatientFacility.Open "Select* From tblPatientFacility", dbHospital, adOpenStatic, adLockOptimistic
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If dbHospital.State = 1 Then dbHospital.Close: Set dbHospital = Nothing
    If rsPatientFacility.State = 1 Then rsPatientFacility.Close: Set dbHospital = Nothing
End Sub

