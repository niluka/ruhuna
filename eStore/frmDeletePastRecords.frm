VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDeletePastRecords 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Past Records"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3945
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
   ScaleHeight     =   2355
   ScaleWidth      =   3945
   Begin btButtonEx.ButtonEx bttnDelete 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   255
      Caption         =   "Delete All Records"
      ForeColor       =   65535
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
      Cancel          =   -1  'True
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      BackColor       =   65535
      Caption         =   "C&lose"
      ForeColor       =   255
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
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67043331
      CurrentDate     =   39626
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   255
      CalendarForeColor=   255
      CalendarTitleBackColor=   255
      CalendarTitleForeColor=   255
      CalendarTrailingForeColor=   255
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   67043331
      CurrentDate     =   39626
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmDeletePastRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bttnDelete_Click()
    Dim tr As Long
    tr = MsgBox("Are you sure you want to Delete all Records", vbCritical + vbYesNo, "Are You Sure?")
    If tr = vbNo Then Exit Sub
    tr = MsgBox("Can You Please double check the time Peroid?" & vbNewLine & "From : " & Format(dtpFrom.Value, "dd MMMM yyyy") & vbNewLine & "To : " & Format(dtpTo.Value, "dd MMMM yyyy") & vbNewLine & "Correct?", vbQuestion + vbYesNo, "Please check")
    If tr = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "DELETE  from tblSaleBill where Date between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .State = 1 Then .Close
        temSql = "DELETE  from tblSale where Date between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .State = 1 Then .Close
        temSql = "DELETE  from tblReturnBill where Date between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .State = 1 Then .Close
        temSql = "DELETE  from tblReturn where Date between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .State = 1 Then .Close
        temSql = "DELETE  from tblIncome where Date between '" & Format(dtpFrom.Value, "dd MMMM yyyy") & "' And '" & Format(dtpTo.Value, "dd MMMM yyyy") & "'"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
    End With
    MsgBox "Selected records Deleted"
End Sub

Private Sub Form_Load()
    dtpFrom.Value = DateSerial((Year(Date) - 1), 1, 1)
    dtpTo.Value = DateSerial((Year(Date) - 1), 21, 31)
    dtpFrom.MaxDate = DateSerial((Year(Date) - 1), Month(Date), 1)
    dtpTo.MaxDate = DateSerial((Year(Date) - 1), Month(Date), 1)
End Sub
