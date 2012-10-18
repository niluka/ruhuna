VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdmit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admit Patients"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13830
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
   ScaleHeight     =   6045
   ScaleWidth      =   13830
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   615
      Left            =   12360
      TabIndex        =   21
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Appearance      =   3
      Caption         =   "C&lose"
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
   Begin btButtonEx.ButtonEx bttnAdmit 
      Height          =   615
      Left            =   10920
      TabIndex        =   20
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      Appearance      =   3
      Caption         =   "&Admit"
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
   Begin VB.Frame Frame2 
      Caption         =   "Admission Details"
      Height          =   4935
      Left            =   7320
      TabIndex        =   23
      Top             =   120
      Width           =   6375
      Begin MSDataListLib.DataCombo dtcRoom 
         Height          =   360
         Left            =   2400
         TabIndex        =   19
         Top             =   4440
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   "Select Room"
      End
      Begin VB.TextBox txtBHT 
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   3015
      End
      Begin MSComCtl2.MonthView mvDOA 
         Height          =   2820
         Left            =   2400
         TabIndex        =   15
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   72679425
         CurrentDate     =   39589
      End
      Begin MSComCtl2.DTPicker dtpTOA 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "h:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   3840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "hour MIN sec"
         Format          =   72679426
         CurrentDate     =   39589
      End
      Begin VB.Label Label9 
         Caption         =   "R&oom "
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "BH&T"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "&Date of Admission"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "&Time of Admission"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   3840
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Patient Details"
      Height          =   4935
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtAddress 
         Height          =   1815
         Left            =   2880
         TabIndex        =   3
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtPhone 
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   2760
         Width           =   3975
      End
      Begin VB.TextBox txtNIC 
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   3360
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   3840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         _Version        =   393216
         Format          =   72679425
         CurrentDate     =   39589
      End
      Begin MSDataListLib.DataCombo dtcHealthSchemeSupplier 
         Height          =   360
         Left            =   2880
         TabIndex        =   11
         Top             =   4320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label10 
         Caption         =   "&Health Scheme Supplier"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Date Of &Birth"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "&Phone"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "&Address"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "First &Name"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "&NIC No"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAdmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsPatients As New ADODB.Recordset
    Dim rsHSS As New ADODB.Recordset
    Dim rsBHT As New ADODB.Recordset
    Dim rsRoom As New ADODB.Recordset
    Dim rsTemRoom As New ADODB.Recordset
    Dim rsRoomPatient As New ADODB.Recordset
    Dim temSql As String
    Dim temPatientID As Long
    Dim TemBHTID As Long
    
Private Sub bttnAdmit_Click()
    If CanAdmit = False Then Exit Sub
    
    With rsBHT
    If .State = 1 Then .Close
        temSql = "SELECT * FROM tblBHT where BHT = '" & Trim(txtBHT.Text) & "'"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            MsgBox "The BHT Number " & txtBHT.Text & " already exists" & vbNewLine & "Please enter another BHT number"
            .Close
            txtBHT.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
        .Close
    End With
    With rsPatients
        If .State = 1 Then .Close
        temSql = "SELECT * FROM tblPatientMainDetails"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !FirstName = UCase(txtName.Text)
        !Address = txtAddress.Text
        !Phone = txtPhone.Text
        !NICNo = txtNIC.Text
        !DateOfBirth = dtpDOB.Value
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        temPatientID = !NewID
        .Close
    End With
    With rsBHT
    If .State = 1 Then .Close
        temSql = "SELECT * FROM tblBHT"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !BHT = txtBHT.Text
        !PatientID = temPatientID
        !DOA = mvDOA.Value
        !DIscharge = False
        !TOA = dtpTOA.Value
        !RoomID = Val(dtcRoom.BoundText)
        !HealthSchemeSupplierID = Val(dtcHealthSchemeSupplier.BoundText)
        .Update
        temSql = "SELECT @@IDENTITY AS NewID"
        .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        TemBHTID = !NewID
        .Close
    End With
    If IsNumeric(dtcRoom.BoundText) = True Then Call AddToRoom
    Dim i As Integer
    i = MsgBox("Admitted. Do You want to issue drugs for this patient now", vbYesNo)
    If i = vbYes Then
        frmHospitalSale.FillPatientCombo
    End If
    
    Unload Me
End Sub

Private Sub AddToRoom()
    With rsRoom
        If .State = 1 Then .Close
        temSql = "SELECT * FROM * tblRoom where RoomID = " & dtcRoom.BoundText
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If IsNull(!BHTID) = False Then
                If !BHTID <> 0 Then
                    With rsRoomPatient
                        If .State = 1 Then .Close
                        temSql = "SELECT * from tblRoomPatient where RoomID = " & dtcRoom.BoundText & " AND BHTID = " & rsRoom!BHTID
                        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                        If .RecordCount > 0 Then
                            !ToDate = Format(mvDOA.Value, "dd MMMM yyyy")
                            !ToTime = dtpTOA.Value
                            .Update
                        End If
                    End With
                End If
            End If
            !PatientID = temPatientID
            !BHTID = TemBHTID
            .Update
             With rsRoomPatient
                If .State = 1 Then .Close
                temSql = "SELECT * from tblRoomPatient"
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                .AddNew
                !PatientID = temPatientID
                !BHTID = TemBHTID
                !RoomID = dtcRoom.BoundText
                !FromTime = dtpTOA.Value
                !FromDate = Format(mvDOA.Value, "dd MMMM yyyy")
                .Update
            End With
        End If
        .Close
    End With
End Sub


Private Sub bttnClose_Click()
    Unload Me
End Sub



Private Sub dtcHealthSchemeSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 27 Then
'        dtcHealthSchemeSupplier.Text = Empty
'    End If
    If KeyCode = vbKeyEscape Then
        dtcHealthSchemeSupplier.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtBHT.SetFocus
    End If
End Sub

Private Sub dtcRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        bttnAdmit_Click
    End If
End Sub

Private Sub dtpDOB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtpDOB.Value = Date
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcHealthSchemeSupplier.SetFocus
    End If
End Sub

Private Sub dtpTOA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        dtpTOA.Value = Date
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtcRoom.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    dtpDOB.Value = Date
    mvDOA.Value = Date
End Sub

Private Sub FillCombos()
    With rsRoom
        If .State = 1 Then .Close
        temSql = "SELECT * from tblRoom order by Room"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcRoom
        Set .RowSource = rsRoom
        .ListField = "Room"
        .BoundText = "RoomID"
    End With
    With rsHSS
        If .State = 1 Then .Close
        temSql = "SELECT * from tblHealthSchemeSuppliers order by HealthSchemeSupplierName"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With dtcHealthSchemeSupplier
        Set .RowSource = rsHSS
        .ListField = "HealthSchemeSupplierName"
        .BoundColumn = "HealthSchemeSupplierID"
    End With
End Sub

Private Function CanAdmit() As Boolean
    CanAdmit = False
    Dim tr As Integer
    If Trim(txtName.Text) = Empty Then
        tr = MsgBox("You have not entered the name of the patient", vbCritical, "Name?")
        txtName.SetFocus
        Exit Function
    End If
    If Trim(txtBHT.Text) = Empty Then
        tr = MsgBox("You have not entered the BHT number", vbCritical, "BHT?")
        txtBHT.SetFocus
        Exit Function
    End If
    CanAdmit = True
End Function

Private Sub mvDOA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        mvDOA.Value = Date
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpTOA.SetFocus
    End If
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtAddress.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPhone.SetFocus
    End If
End Sub

Private Sub txtBHT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtBHT.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        mvDOA.SetFocus
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtName.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtAddress.SetFocus
    End If
End Sub

Private Sub txtNIC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtNIC.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        dtpDOB.SetFocus
    End If
End Sub

Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtPhone.Text = Empty
        KeyCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtNIC.SetFocus
    End If
End Sub
